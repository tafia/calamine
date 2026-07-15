// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Example of reading cell style/formatting information from a worksheet in
//! an XLSX file.

use calamine::{open_workbook, Error, Reader, Xlsx};

fn main() -> Result<(), Error> {
    let path = "tests/styles.xlsx";

    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let sheet_name = workbook.sheet_names()[0].clone();

    // Read the styles of all explicitly formatted cells in the worksheet.
    let styles = workbook.worksheet_style(&sheet_name)?;

    println!(
        "'{}': styled range {:?}..={:?}",
        sheet_name,
        styles.start(),
        styles.end(),
    );

    // Iterate over the cells and print a summary of any visible formatting.
    for (row, col, style) in styles.cells() {
        if style.is_empty() {
            continue;
        }

        let mut summary = Vec::new();

        if let Some(font) = &style.font {
            if font.is_bold() {
                summary.push("bold".to_string());
            }
            if font.is_italic() {
                summary.push("italic".to_string());
            }
            if let Some(color) = &font.color {
                summary.push(format!("font color {color}"));
            }
        }

        if let Some(fill) = &style.fill {
            if let Some(color) = fill.get_color() {
                summary.push(format!("fill {color}"));
            }
        }

        if let Some(borders) = &style.borders {
            if borders.has_visible_borders() {
                summary.push("borders".to_string());
            }
        }

        if let Some(number_format) = &style.number_format {
            summary.push(format!("format '{}'", number_format.format_code));
        }

        if !summary.is_empty() {
            println!("({row}, {col}): {}", summary.join(", "));
        }
    }

    Ok(())
}
