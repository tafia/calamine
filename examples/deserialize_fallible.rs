// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! An example of deserializing cells that may be empty or contain unexpected
//! types.
//!
//! Compared with `deserialize_struct`, which expects every cell to contain a
//! value of the correct type, this example shows how to handle cells that may
//! be empty or hold a string where a number is expected.
//!
//! Calamine provides two helper functions for this:
//!
//! - [`calamine::deserialize_as_f64_or_none`] — returns `None` for any cell
//!   that cannot be read as `f64`, discarding the original value.
//! - [`calamine::deserialize_as_f64_or_string`] — returns `Err(String)` for any
//!   cell that cannot be read as `f64`, preserving the original value as a
//!   string so the caller can decide how to handle it.
//!
//! Both are used with Serde's `deserialize_with` field attribute.
//!
//! Note: empty cells are silently skipped during map deserialization (they are
//! not yielded as keys). Any struct field that may correspond to an empty cell
//! therefore also needs `#[serde(default)]`, so that Serde uses the field's
//! default value when the key is absent rather than returning a missing-field
//! error.
//!
//! The sample Excel file `temperature-fallible.xlsx` contains a single sheet
//! named "Sheet1" with the following data:
//!
//! ```text
//!  ___________________________________________________________
//! |         ||                |           |                   |
//! |         ||       A        |     B     |         C         |
//! |_________||________________|___________|___________________|
//! |    1    || label          | value     | notes             |
//! |_________||________________|___________|___________________|
//! |    2    || celsius        | 22.2222   | measured          |
//! |_________||________________|___________|___________________|
//! |    3    || fahrenheit     | 72        | converted         |
//! |_________||________________|___________|___________________|
//! |    4    || kelvin         | (empty)   | (empty)           |
//! |_________||________________|___________|___________________|
//! |    5    || rankine        | N/A       | derived           |
//! |_________||________________|___________|___________________|
//! |_          ________________________________________________|
//!   \ Sheet1 /
//!     ------
//! ```
//!
//! Next: `deserialize_no_headers` — positional deserialization when there is no
//! header row.

use calamine::{
    deserialize_as_f64_or_none, deserialize_as_f64_or_string, open_workbook,
    RangeDeserializerBuilder, Reader, Xlsx,
};
use serde::Deserialize;

// Variant A: discard values that cannot be read as f64.
//
// Both `value` and `notes` may be absent (empty cell skipped in map), so both
// need `#[serde(default)]`. `Option<f64>` defaults to `None`; `String`
// defaults to `""`.
#[derive(Debug, Deserialize)]
struct RowOrNone {
    label: String,
    #[serde(default, deserialize_with = "deserialize_as_f64_or_none")]
    value: Option<f64>,
    #[serde(default)]
    notes: String,
}

// Variant B: preserve values that cannot be read as f64 as their string
// representation, so the caller can log or handle them explicitly.
//
// `Result<f64, String>` does not implement `Default`, so a custom default
// function is required.
#[derive(Debug, Deserialize)]
struct RowOrString {
    label: String,
    #[serde(
        default = "missing_value",
        deserialize_with = "deserialize_as_f64_or_string"
    )]
    value: Result<f64, String>,
    #[serde(default)]
    notes: String,
}

fn missing_value() -> Result<f64, String> {
    Err(String::new())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/temperature-fallible.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")?;

    //
    // Variant A: Option<f64>
    //
    let rows_a: Vec<RowOrNone> = RangeDeserializerBuilder::with_deserialize_headers::<RowOrNone>()
        .from_range(&range)?
        .collect::<Result<_, _>>()?;

    // Normal float cells deserialize to Some(value).
    assert_eq!(rows_a[0].value, Some(22.2222));
    assert_eq!(rows_a[1].value, Some(72.0));

    // An empty cell produces None via the serde default; a string cell
    // produces None via the helper function.
    assert_eq!(rows_a[2].value, None); // empty cell → default
    assert_eq!(rows_a[3].value, None); // "N/A" → helper returns None

    // An empty notes cell also produces the default value for String.
    assert_eq!(rows_a[2].notes, "");
    assert_eq!(rows_a[3].notes, "derived");

    println!("or_none:");
    for row in &rows_a {
        println!("  {} {:?} {:?}", row.label, row.value, row.notes);
    }

    //
    // Variant B: Result<f64, String>
    //
    let rows_b: Vec<RowOrString> =
        RangeDeserializerBuilder::with_deserialize_headers::<RowOrString>()
            .from_range(&range)?
            .collect::<Result<_, _>>()?;

    // Normal float cells deserialize to Ok(value).
    assert_eq!(rows_b[0].value, Ok(22.2222));
    assert_eq!(rows_b[1].value, Ok(72.0));

    // An empty cell produces Err("") via the custom default function. A string
    // cell reaches the helper, which preserves the original string.
    assert_eq!(rows_b[2].value, Err(String::new()));
    assert_eq!(rows_b[3].value, Err("N/A".to_string()));

    println!("or_string:");
    for row in &rows_b {
        println!("  {} {:?} {:?}", row.label, row.value, row.notes);
    }

    Ok(())
}
