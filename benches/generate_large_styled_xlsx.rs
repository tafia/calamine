// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Generator for large styled xlsx files for benchmarking.
//!
//! Run with: cargo run --bin generate_large_styled_xlsx
//!
//! This creates `tests/large_styled.xlsx` with 1000 copies of style patterns.

use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, FormatUnderline, Workbook, XlsxError,
};

fn main() -> Result<(), XlsxError> {
    let output_path = format!("{}/tests/styles_1M.xlsx", env!("CARGO_MANIFEST_DIR"));
    println!(
        "Generating styles_1M.xlsx (1M styled cells) at: {}",
        output_path
    );

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Sheet 1")?;

    // Define formats matching styles.xlsx patterns
    let bold = Format::new().set_bold();
    let italic = Format::new().set_italic();
    let underline = Format::new().set_underline(FormatUnderline::Single);
    let strikethrough = Format::new().set_font_strikethrough();
    let red_font = Format::new().set_font_color(Color::Red);
    let fill_yellow = Format::new().set_background_color(Color::Yellow);
    let align_center = Format::new().set_align(FormatAlign::Center);
    let align_right = Format::new().set_align(FormatAlign::Right);
    let number_format = Format::new().set_num_format("0.00%");
    let currency_format = Format::new().set_num_format("$#,##0.00");
    let date_format = Format::new().set_num_format("yyyy-mm-dd");

    // Border formats
    let thin_border = Format::new()
        .set_border(FormatBorder::Thin)
        .set_border_color(Color::Black);
    let thick_border = Format::new()
        .set_border(FormatBorder::Thick)
        .set_border_color(Color::Blue);
    let dashed_border = Format::new()
        .set_border(FormatBorder::Dashed)
        .set_border_color(Color::Green);

    // Combined formats
    let bold_italic = Format::new().set_bold().set_italic();
    let bold_red = Format::new().set_bold().set_font_color(Color::Red);
    let italic_underline = Format::new()
        .set_italic()
        .set_underline(FormatUnderline::Single);
    let center_yellow = Format::new()
        .set_align(FormatAlign::Center)
        .set_background_color(Color::Yellow);
    let bold_border = Format::new().set_bold().set_border(FormatBorder::Thin);

    // Font size variations
    let size_8 = Format::new().set_font_size(8.0);
    let size_12 = Format::new().set_font_size(12.0);
    let size_16 = Format::new().set_font_size(16.0);
    let size_24 = Format::new().set_font_size(24.0);

    // Font name variations
    let arial = Format::new().set_font_name("Arial");
    let times = Format::new().set_font_name("Times New Roman");
    let courier = Format::new().set_font_name("Courier New");

    // Color variations
    let blue_font = Format::new().set_font_color(Color::Blue);
    let green_font = Format::new().set_font_color(Color::Green);
    let purple_font = Format::new().set_font_color(Color::Purple);
    let fill_cyan = Format::new().set_background_color(Color::Cyan);
    let fill_magenta = Format::new().set_background_color(Color::Magenta);
    let fill_orange = Format::new().set_background_color(Color::Orange);

    // The pattern of styles to repeat (20 columns x 50 rows = 1000 cells per block)
    // 1000 repetitions = 1M cells, ~3.2MB file
    let block_rows = 50;
    let block_cols = 20;
    let repetitions = 1000;

    println!(
        "Creating {} blocks of {}x{} = {} total cells",
        repetitions,
        block_rows,
        block_cols,
        repetitions * block_rows * block_cols
    );

    for rep in 0..repetitions {
        let row_offset = (rep * block_rows) as u32;

        for row in 0..block_rows as u32 {
            let actual_row = row_offset + row;

            // Column 0: Bold text
            worksheet.write_string_with_format(actual_row, 0, "Bold", &bold)?;

            // Column 1: Italic text
            worksheet.write_string_with_format(actual_row, 1, "Italic", &italic)?;

            // Column 2: Underline text
            worksheet.write_string_with_format(actual_row, 2, "Underline", &underline)?;

            // Column 3: Strikethrough
            worksheet.write_string_with_format(actual_row, 3, "Strike", &strikethrough)?;

            // Column 4: Red font
            worksheet.write_string_with_format(actual_row, 4, "Red", &red_font)?;

            // Column 5: Yellow fill
            worksheet.write_string_with_format(actual_row, 5, "Yellow", &fill_yellow)?;

            // Column 6: Center aligned
            worksheet.write_string_with_format(actual_row, 6, "Center", &align_center)?;

            // Column 7: Right aligned
            worksheet.write_string_with_format(actual_row, 7, "Right", &align_right)?;

            // Column 8: Number with percentage format
            worksheet.write_number_with_format(
                actual_row,
                8,
                0.1234 + (row as f64 * 0.001),
                &number_format,
            )?;

            // Column 9: Currency format
            worksheet.write_number_with_format(
                actual_row,
                9,
                1234.56 + (row as f64),
                &currency_format,
            )?;

            // Column 10: Date format
            worksheet.write_number_with_format(
                actual_row,
                10,
                45000.0 + (row as f64),
                &date_format,
            )?;

            // Column 11: Thin border
            worksheet.write_string_with_format(actual_row, 11, "Thin", &thin_border)?;

            // Column 12: Thick border
            worksheet.write_string_with_format(actual_row, 12, "Thick", &thick_border)?;

            // Column 13: Dashed border
            worksheet.write_string_with_format(actual_row, 13, "Dashed", &dashed_border)?;

            // Column 14: Bold + Italic
            worksheet.write_string_with_format(actual_row, 14, "Bold+Ital", &bold_italic)?;

            // Column 15: Bold + Red
            worksheet.write_string_with_format(actual_row, 15, "Bold+Red", &bold_red)?;

            // Column 16: Italic + Underline
            worksheet.write_string_with_format(actual_row, 16, "Ital+Uline", &italic_underline)?;

            // Column 17: Center + Yellow
            worksheet.write_string_with_format(actual_row, 17, "Ctr+Yellow", &center_yellow)?;

            // Column 18: Bold + Border
            worksheet.write_string_with_format(actual_row, 18, "Bold+Bdr", &bold_border)?;

            // Column 19: Mixed - rotate through variations
            match row % 10 {
                0 => worksheet.write_string_with_format(actual_row, 19, "Size8", &size_8)?,
                1 => worksheet.write_string_with_format(actual_row, 19, "Size12", &size_12)?,
                2 => worksheet.write_string_with_format(actual_row, 19, "Size16", &size_16)?,
                3 => worksheet.write_string_with_format(actual_row, 19, "Size24", &size_24)?,
                4 => worksheet.write_string_with_format(actual_row, 19, "Arial", &arial)?,
                5 => worksheet.write_string_with_format(actual_row, 19, "Times", &times)?,
                6 => worksheet.write_string_with_format(actual_row, 19, "Courier", &courier)?,
                7 => worksheet.write_string_with_format(actual_row, 19, "Blue", &blue_font)?,
                8 => worksheet.write_string_with_format(actual_row, 19, "Green", &green_font)?,
                _ => worksheet.write_string_with_format(actual_row, 19, "Purple", &purple_font)?,
            };
        }

        if rep % 100 == 0 {
            println!("Progress: {}/{} blocks", rep, repetitions);
        }
    }

    // Set some column widths
    for col in 0..block_cols as u16 {
        worksheet.set_column_width(col, 12.0)?;
    }

    workbook.save(&output_path)?;

    println!("Done! File saved to: {}", output_path);
    println!(
        "Total cells with styles: {}",
        repetitions * block_rows * block_cols
    );

    Ok(())
}
