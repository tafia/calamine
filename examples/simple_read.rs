// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! This is a minimal "hello world" example that demonstrates how to open a
//! workbook and read some information from it using the `calamine` crate.
//!
//! Note that the `calamine::Reader` trait is imported to bring the
//! `sheet_names` and `worksheet_range` methods into scope.
//!
//! The sample Excel file `readings.xlsx` contains a single sheet named "Sheet1"
//! with the following data:
//!
//! ```text
//!  ______________________________________________________________________________
//! |         ||                |                |                |                |
//! |         ||       A        |       B        |       C        |       D        |
//! |_________||________________|________________|________________|________________|
//! |    1    || Station        | Celsius        | Humidity       | Pressure       |
//! |_________||________________|________________|________________|________________|
//! |    2    || London         | 15             | 72             | 1013           |
//! |_________||________________|________________|________________|________________|
//! |    3    || Paris          | 18.5           | 65             | 1009           |
//! |_________||________________|________________|________________|________________|
//! |    4    || Berlin         | 12             | 80             | 1017           |
//! |_________||________________|________________|________________|________________|
//! |_          ___________________________________________________________________|
//!   \ Sheet1 /
//!     ------
//! ```
//!
//! Run the example like this:
//!
//! ```text
//! $ cargo run -q --example simple_read
//!
//! Sheet: Sheet1
//! Headers: Station, Celsius, Humidity, Pressure
//! Data:
//! London, 15, 72, 1013
//! Paris, 18.5, 65, 1009
//! Berlin, 12, 80, 1017
//! ```

use calamine::{open_workbook, Reader, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Read the workbook from a file path. The type annotation (`Xlsx<_>`) is
    // required to specify the format of the workbook. The `open_workbook()`
    // function doesn't infer the format from the file extension.
    let mut workbook: Xlsx<_> = open_workbook("tests/readings.xlsx")?;

    // Note: There is also an `open_workbook_auto()` function that can infer the
    // format from the file extension, but it is better to be explicit about the
    // format when possible.

    // Read the worksheet names from the workbook.
    let sheet_names = workbook.sheet_names();

    // Get the first sheet name.
    let sheet_name = sheet_names.first().ok_or("no sheets found")?.clone();

    println!("Sheet: {sheet_name}");

    // Get the data "range" for the worksheet. A range is a rectangular area of
    // cells, and it is the basic unit of deserialization in `calamine`.
    let range = workbook.worksheet_range(&sheet_name)?;
    let mut rows = range.rows();

    // In this example, the first row contains the column headers.
    let headers: Vec<String> = rows
        .next()
        .ok_or("empty sheet")?
        .iter()
        .map(|c| c.to_string())
        .collect();

    println!("Headers: {}", headers.join(", "));

    // Each subsequent row is a data row.
    println!("Data:");
    for row in rows {
        let values: Vec<String> = row.iter().map(|c| c.to_string()).collect();
        println!("{}", values.join(", "));
    }

    Ok(())
}
