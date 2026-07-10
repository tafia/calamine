// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Demonstrates streaming XLSX cell values and cell styles in one worksheet
//! pass.

use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = "tests/styles.xlsx";

    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let mut reader = workbook.worksheet_cells_reader("Sheet 1")?;

    // Stream the cell values and their style ids in a single pass. The style
    // id resolves to a full style via the workbook style palette.
    while let Some((cell, style_id)) = reader.next_cell_with_style_id()? {
        let style = &reader.styles()[style_id];

        let mut summary = Vec::new();

        if let Some(font) = &style.font {
            if font.is_bold() {
                summary.push("bold".to_string());
            }
            if font.is_italic() {
                summary.push("italic".to_string());
            }
        }

        if let Some(fill) = &style.fill {
            if let Some(color) = fill.get_color() {
                summary.push(format!("fill {color}"));
            }
        }

        if let Some(number_format) = &style.number_format {
            if number_format.format_code != "General" {
                summary.push(format!("format '{}'", number_format.format_code));
            }
        }

        let (row, col) = cell.get_position();
        println!(
            "row={}, col={}, value={:?}, style=[{}]",
            row + 1,
            col + 1,
            cell.get_value(),
            summary.join(", ")
        );
    }

    Ok(())
}
