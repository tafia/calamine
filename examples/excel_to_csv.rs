// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! An example for using the `calamine` crate to convert an Excel file to CSV.
//!
//! Converts XLSX, XLSM, XLSB, and XLS files. The filename and sheet name must
//! be specified as command line arguments. The output CSV will be written to a
//! file with the same name as the input file, but with a `.csv` extension.

use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

use calamine::{open_workbook_auto, Data, Range, Reader};

// usage: cargo run --example excel_to_csv file.xls[xmb] sheet_name
//
// Where:
// - `file.xls[xmb]` is the Excel file to convert. Required.
// - `sheet_name` is the name of the sheet to convert. Required.
//
// The output will be written to a file with the same name as the input file,
// including the path, but with a `.csv` extension.
//
fn main() {
    let excel_file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");

    let sheet_name = env::args()
        .nth(2)
        .expect("Expecting a sheet name as second argument");

    let excel_path = PathBuf::from(excel_file);
    match excel_path.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let csv_path = excel_path.with_extension("csv");
    let mut csv_file = BufWriter::new(File::create(csv_path).unwrap());
    let mut workbook = open_workbook_auto(&excel_path).unwrap();
    let range = workbook.worksheet_range(&sheet_name).unwrap();

    write_to_csv(&mut csv_file, &range).unwrap();
}

// Write the Excel data as strings to a CSV file. Uses a semicolon (`;`) as the
// field separator.
//
// Note, this is a simplified version of CSV and doesn't handle quoting of
// separators or other special cases. See the `csv.rs` crate for a more robust
// solution.
fn write_to_csv<W: Write>(output_file: &mut W, range: &Range<Data>) -> std::io::Result<()> {
    let max_column = range.get_size().1 - 1;

    for rows in range.rows() {
        for (col_number, cell_data) in rows.iter().enumerate() {
            match cell_data {
                Data::Empty => Ok(()),
                Data::Int(i) => write!(output_file, "{i}"),
                Data::Bool(b) => write!(output_file, "{b}"),
                Data::Error(e) => write!(output_file, "{e:?}"),
                Data::Float(f) => write!(output_file, "{f}"),
                Data::DateTime(d) => write!(output_file, "{}", d.as_f64()),
                Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => {
                    write!(output_file, "{s}")
                }
            }?;

            // Write the field separator except for the last column.
            if col_number != max_column {
                write!(output_file, ";")?;
            }
        }

        write!(output_file, "\r\n")?;
    }

    Ok(())
}
