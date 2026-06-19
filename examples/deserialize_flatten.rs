// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! An example of deserializing rows that have a fixed set of named columns
//! plus an unknown number of extra columns captured with `#[serde(flatten)]`.
//!
//! Compared with `deserialize_struct`, which requires every column to be
//! declared as a struct field, this example shows how to use Serde's
//! `#[serde(flatten)]` attribute to absorb any remaining columns into a
//! [`std::collections::HashMap`]. This is useful when the spreadsheet may
//! contain extra columns that are not known at compile time.
//!
//! The sample Excel file `readings.xlsx` contains a single sheet named
//! "Sheet1" with the following data:
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
//! Next: `deserialize_fallible` — handling cells that may be empty or contain
//! unexpected types.

use std::collections::HashMap;

use calamine::{open_workbook, RangeDeserializerBuilder, Reader, Xlsx};
use serde::Deserialize;

// The `station` field is matched by name to the "Station" header. All
// remaining columns are absorbed into `readings` by `#[serde(flatten)]`.
#[derive(Debug, Deserialize)]
struct StationRow {
    #[serde(rename = "Station")]
    station: String,
    #[serde(flatten)]
    readings: HashMap<String, f64>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/readings.xlsx";
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")?;

    // Use the default builder so that all columns — including those captured
    // by the flattened HashMap — are passed to the deserializer. The first
    // row of the range is consumed as the header.
    let rows: Vec<StationRow> = RangeDeserializerBuilder::new()
        .from_range(&range)?
        .collect::<Result<_, _>>()?;

    assert_eq!(rows.len(), 3);

    // The fixed field is populated from the "Station" column.
    assert_eq!(rows[0].station, "London");
    assert_eq!(rows[1].station, "Paris");
    assert_eq!(rows[2].station, "Berlin");

    // The remaining columns are captured by the flattened HashMap.
    assert_eq!(rows[0].readings.len(), 3);
    assert_eq!(rows[0].readings["Celsius"], 15.0);
    assert_eq!(rows[0].readings["Humidity"], 72.0);
    assert_eq!(rows[0].readings["Pressure"], 1013.0);

    for row in &rows {
        println!("{} {:?}", row.station, row.readings);
    }

    Ok(())
}
