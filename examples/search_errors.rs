// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! An example for using the `calamine` crate to to search a directory for Excel
//! files and check them for errors.
//!
//! Recursively searches for XLSX, XLSM, XLSB, and XLS Excel files and parses
//! them to check for any `calamine` errors. Also checks for and counts the
//! number of missing VBA references or the number of cells with Excel errors.
//!

use glob::{glob, GlobError};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

use calamine::{open_workbook_auto, Data, Error, Reader};

#[derive(Debug)]
#[allow(dead_code)]
// Simple error type to handle the various errors that can occur.
enum FileStatus {
    VbaError(Error),
    RangeError(Error),
    WorkbookOpenError(Error),
    Glob(GlobError),
}

// usage: cargo run --example search_errors [dir]
//
// Where:
//
// - `[dir]` is the directory to search for Excel files. Defaults to `.`.
//
// The analysis is written to an output file called `{dir}_errors.csv`. If no
// directory is specified, it defaults to the current directory. The output file
// will contain the file path, the number of missing VBA references, and the
// number of cells with errors for each Excel file found. Alternatively, if an
// error occurs while processing a file, it will be logged in the output.
//
fn main() -> Result<(), FileStatus> {
    let search_dir = env::args().nth(1).unwrap_or_else(|| ".".to_string());
    let file_pattern = format!("{search_dir}/**/*.xl*");
    let mut file_count = 0;

    // Strip/convert any directory characters to create an output filename.
    let mut output_filename = file_pattern
        .chars()
        .take_while(|c| *c != '*')
        .filter_map(|c| match c {
            ':' => None,
            '/' | '\\' | ' ' => Some('_'),
            c => Some(c),
        })
        .collect::<String>();

    // Append "_errors.csv" to the output filename.
    output_filename.push_str("errors.csv");

    // Use a default output filename for the default search directory.
    if search_dir == "." {
        output_filename = "errors.csv".to_string();
    }

    let mut output_file = BufWriter::new(File::create(&output_filename).unwrap());

    // Iterate through any Excel files that were found.
    for file in glob(&file_pattern).unwrap() {
        file_count += 1;
        let file = file.map_err(FileStatus::Glob)?;

        match analyze(&file) {
            Ok((missing_vba_refs, cell_errors)) => {
                writeln!(
                    output_file,
                    "{file:?}: Missing VBA refs = {missing_vba_refs:?}. Cell errors = {cell_errors}."
                )
            }
            Err(e) => writeln!(output_file, "{file:?}: Error = {e:?}."),
        }
        .unwrap_or_else(|e| println!("{e:?}"));
    }

    println!("Analyzed {file_count} excel files. See '{output_filename}' for analysis.");

    Ok(())
}

// Function to analyze a single Excel file for errors, missing VBA references
// and cell errors.
fn analyze(file: &PathBuf) -> Result<(Option<usize>, usize), FileStatus> {
    let mut workbook = open_workbook_auto(file).map_err(FileStatus::WorkbookOpenError)?;
    let mut num_cell_errors = 0;
    let mut num_missing_vba_refs = None;

    // Check if the workbook has a VBA project and count missing references.
    if let Some(vba) = workbook.vba_project().map_err(FileStatus::VbaError)? {
        num_missing_vba_refs = Some(
            vba.get_references()
                .iter()
                .filter(|r| r.is_missing())
                .count(),
        );
    }

    // Iterate through all sheets and count cell errors.
    for sheet_name in workbook.sheet_names() {
        let range = workbook
            .worksheet_range(&sheet_name)
            .map_err(FileStatus::RangeError)?;

        num_cell_errors += range
            .rows()
            .flat_map(|r| r.iter().filter(|c| matches!(**c, Data::Error(_))))
            .count();
    }

    Ok((num_missing_vba_refs, num_cell_errors))
}
