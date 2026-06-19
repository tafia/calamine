// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Demonstrates reading pictures/images and their metadata from an XLSX file
//! using the `picture` feature.
//!
//! Each [`Picture`] returned by [`Reader::pictures_with_metadata`] contains the
//! raw image bytes, the file extension, the worksheet name, and the anchor cell
//! position (row and column, both 0-based).
//!
//! The sample file used here contains three PNG images on Sheet1:
//!
//! - Two images inserted **in-cell** (Excel 365 rich-data style) at rows 0 and 2.
//! - One image inserted **over-cell** (classic DrawingML anchor) at row 8, col 4.
//!
//! Run the example like this (the `picture` feature must be enabled):
//!
//! ```text
//! $ cargo run -q --example read_picture_data --features picture
//!
//! Found 3 picture(s) in "tests/pictures_in_cell_and_over_cell.xlsx":
//!
//!   Sheet: "Sheet1", Cell: (8, 4), Type: png, Size: 200 bytes
//!   Sheet: "Sheet1", Cell: (0, 0), Type: png, Size: 200 bytes
//!   Sheet: "Sheet1", Cell: (2, 0), Type: png, Size: 178 bytes
//! ```

use calamine::{open_workbook, Picture, Reader, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/pictures_in_cell_and_over_cell.xlsx";

    let workbook: Xlsx<_> = open_workbook(path)?;

    // The `pictures_with_metadata()` method returns a Vec<Picture>. Each
    // Picture carries the raw image bytes, its file extension, the sheet it
    // lives on, and the 0-based row/column of its anchor cell.
    let pictures = workbook.pictures_with_metadata();

    println!("Found {} picture(s) in {path:?}:\n", pictures.len());
    for picture in &pictures {
        print_picture(picture);
    }

    Ok(())
}

fn print_picture(picture: &Picture) {
    println!(
        "  Sheet: {:?}, Cell: ({}, {}), Type: {}, Size: {} bytes",
        picture.sheet_name,
        picture.row,
        picture.col,
        picture.extension,
        picture.data.len(),
    );
}
