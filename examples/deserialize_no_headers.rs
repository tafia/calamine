// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! An example of positional deserialization without a header row.
//!
//! Compared with earlier examples, which assume the first row contains column
//! names, this example shows [`RangeDeserializerBuilder::has_headers`] set to
//! `false`. Every row, including the header row, is then yielded as a data
//! item and the caller is responsible for handling it.
//!
//! Using [`calamine::Data`] as the element type of each row preserves the
//! original cell values without any type coercion. This is useful when the
//! schema is not known at compile time, or when you want to inspect the raw
//! cell values before deciding how to process them.
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
//! Next: `deserialize_seed` — runtime-driven deserialization using
//! `DeserializeSeed` when column names are only known after reading the header
//! row.

use calamine::{open_workbook, Data, DataType, RangeDeserializerBuilder, Reader, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/temperature.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")?;

    // With has_headers(false) every row, including any header row, is yielded
    // as a data item.
    let mut iter = RangeDeserializerBuilder::new()
        .has_headers(false)
        .from_range(&range)?;

    // The first row contains the column names. Capture it to confirm the
    // expected layout before processing the data rows.
    let headers: Vec<Data> = iter.next().ok_or("empty sheet")??;
    assert_eq!(headers, [Data::from("label"), Data::from("value")]);

    // Collect remaining rows as raw cell values.
    let rows: Vec<Vec<Data>> = iter.collect::<Result<_, _>>()?;
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0][0], Data::from("celsius"));
    assert_eq!(rows[0][1].as_f64(), Some(22.2222));
    assert_eq!(rows[1][0], Data::from("fahrenheit"));
    assert_eq!(rows[1][1].as_f64(), Some(72.0));

    for row in &rows {
        println!("{row:?}");
    }

    Ok(())
}
