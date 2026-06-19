// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! This example demonstrates the simplest way to deserialize a spreadsheet row.
//! An anonymous tuple is read one row at a time using [`Range::deserialize`].
//!
//! This is the starting point for the deserialization examples. The first row
//! of the range is consumed as a header and the iterator yields the remaining
//! rows. Each row is deserialized into the tuple type inferred from the call
//! site, with columns matched positionally.
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
//! Next: `deserialize_struct` — deserializing all rows into a named struct
//! without depending on column order.

use calamine::{open_workbook, Error, Reader, Xlsx};

fn main() -> Result<(), Error> {
    let path = "tests/temperature.xlsx";

    // Open the workbook.
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // Get the data range from the first sheet.
    let sheet_range = workbook.worksheet_range("Sheet1")?;

    // Get an iterator over data in the range.
    let mut iter = sheet_range.deserialize()?;

    // Get the next record in the range. The first row is assumed to be the
    // header.
    if let Some(result) = iter.next() {
        let (label, value): (String, f64) = result?;

        assert_eq!(label, "celsius");
        assert_eq!(value, 22.2222);

        Ok(())
    } else {
        Err(From::from("Expected at least one record but got none"))
    }
}
