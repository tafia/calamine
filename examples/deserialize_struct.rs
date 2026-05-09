// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! An example of deserializing spreadsheet rows into a named struct.
//!
//! Compared with `deserialize_range`, which uses an anonymous tuple, this
//! example demonstrates how to deserialize spreadsheet data into a named struct
//! derived with [`serde::Deserialize`] and how to collect all rows in a single
//! call.
//!
//! [`RangeDeserializerBuilder::with_deserialize_headers`] reads the field names
//! directly from the struct definition, so the column order in the spreadsheet
//! does not need to match the field order in the struct.
//!
//! The sample Excel file `temperature.xlsx` contains a single sheet named
//! "Sheet1" with the following data:
//!
//! ```text
//!  ____________________________________________
//! |         ||                |                |
//! |         ||       A        |       B        |
//! |_________||________________|________________|
//! |    1    || label          | value          |
//! |_________||________________|________________|
//! |    2    || celsius        | 22.2222        |
//! |_________||________________|________________|
//! |    3    || fahrenheit     | 72             |
//! |_________||________________|________________|
//! |_          _________________________________|
//!   \ Sheet1 /
//!     ------
//! ```
//!
//! Next: `deserialize_flatten` — capturing extra columns into a `HashMap`
//! with `#[serde(flatten)]`.

use calamine::{open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Deserialize;

#[derive(Debug, Deserialize, PartialEq)]
struct Row {
    label: String,
    value: f64,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/temperature.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")?;

    // Collect all rows into a Vec<Row>. The field names on the struct must
    // match the header row of the spreadsheet. Column order does not matter.
    let rows: Vec<Row> = RangeDeserializerBuilder::with_deserialize_headers::<Row>()
        .from_range(&range)?
        .collect::<Result<_, _>>()?;

    assert_eq!(rows.len(), 2);

    assert_eq!(
        rows[0],
        Row {
            label: "celsius".into(),
            value: 22.2222,
        }
    );

    assert_eq!(
        rows[1],
        Row {
            label: "fahrenheit".into(),
            value: 72.0,
        }
    );

    for row in &rows {
        println!("{row:?}");
    }

    Ok(())
}
