// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use calamine::{open_workbook, Reader, Xlsx};

/// Example demonstrating how to capture column widths and row heights from Excel files
fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open an Excel file
    let path = format!("{}/tests/styles.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // Get the first sheet name
    let sheet_names = workbook.sheet_names();
    if let Some(sheet_name) = sheet_names.first() {
        println!("Getting layout information for sheet: {}", sheet_name);

        // Get the worksheet layout information (column widths and row heights)
        let layout = workbook.worksheet_layout(sheet_name)?;

        // Display default dimensions
        if let Some(default_col_width) = layout.default_column_width {
            println!("Default column width: {} characters", default_col_width);
        }
        if let Some(default_row_height) = layout.default_row_height {
            println!("Default row height: {} points", default_row_height);
        }

        // Display custom column widths
        if !layout.column_widths.is_empty() {
            println!("\nCustom column widths:");
            for (_, col_width) in &layout.column_widths {
                println!(
                    "  Column {}: {} characters (custom: {}, hidden: {}, best_fit: {})",
                    col_width.column,
                    col_width.width,
                    col_width.custom_width,
                    col_width.hidden,
                    col_width.best_fit
                );
            }
        }

        // Display custom row heights
        if !layout.row_heights.is_empty() {
            println!("\nCustom row heights:");
            for (_, row_height) in &layout.row_heights {
                println!(
                    "  Row {}: {} points (custom: {}, hidden: {})",
                    row_height.row, row_height.height, row_height.custom_height, row_height.hidden
                );
            }
        }

        // Example of using the helper methods
        println!("\nExample queries:");
        let effective_width_0 = layout.get_effective_column_width(0);
        let effective_height_0 = layout.get_effective_row_height(0);
        println!(
            "Effective width of column 0: {} characters",
            effective_width_0
        );
        println!("Effective height of row 0: {} points", effective_height_0);

        // Check if a specific column has custom width
        if let Some(col_width) = layout.get_column_width(0) {
            println!("Column 0 has custom width: {}", col_width.width);
        } else {
            println!("Column 0 uses default width");
        }

        // Check if layout has any custom dimensions
        if layout.has_custom_dimensions() {
            println!("This worksheet has custom column widths or row heights");
        } else {
            println!("This worksheet uses all default dimensions");
        }
    }

    Ok(())
}
