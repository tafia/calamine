// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use calamine::{open_workbook, Cell, Color, Data, Font, FontWeight, Reader, Style};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Example of creating a cell with style
    let style = Style::new().with_font(
        Font::new()
            .with_name("Arial".to_string())
            .with_size(12.0)
            .with_weight(FontWeight::Bold)
            .with_color(Color::rgb(255, 0, 0)),
    );

    let cell = Cell::with_style((0, 0), Data::String("Hello World".to_string()), style);

    println!("Created cell with style:");
    if let Some(cell_style) = cell.get_style() {
        if let Some(font) = cell_style.get_font() {
            println!(
                "  Font: {} (size: {})",
                font.name.as_deref().unwrap_or("Unknown"),
                font.size.unwrap_or(0.0)
            );
            println!("  Bold: {}", font.is_bold());
            if let Some(color) = font.color {
                println!("  Color: {}", color);
            }
        }
    }

    // Example of creating CellData with style
    use calamine::CellData;

    let cell_data = CellData::with_style(
        Data::Int(42),
        Style::new().with_font(Font::new().with_weight(FontWeight::Bold)),
    );

    println!("\nCreated CellData with style:");
    if cell_data.has_style() {
        if let Some(style) = cell_data.get_style() {
            if let Some(font) = style.get_font() {
                println!("  Bold: {}", font.is_bold());
            }
        }
    }

    // Example of creating a more complex style
    let complex_style = Style::new()
        .with_font(
            Font::new()
                .with_name("Times New Roman".to_string())
                .with_size(14.0)
                .with_weight(FontWeight::Bold)
                .with_color(Color::rgb(0, 0, 255)),
        )
        .with_fill(calamine::Fill::solid(Color::rgb(255, 255, 0)))
        .with_borders(calamine::Borders::new());

    let styled_cell = Cell::with_style((1, 1), Data::Float(3.14), complex_style);

    println!("\nCreated cell with complex style:");
    if let Some(style) = styled_cell.get_style() {
        if let Some(font) = style.get_font() {
            println!(
                "  Font: {} (size: {})",
                font.name.as_deref().unwrap_or("Unknown"),
                font.size.unwrap_or(0.0)
            );
            println!("  Bold: {}", font.is_bold());
            if let Some(color) = font.color {
                println!("  Font color: {}", color);
            }
        }

        if let Some(fill) = style.get_fill() {
            if fill.is_visible() {
                println!("  Has fill");
                if let Some(color) = fill.get_color() {
                    println!("  Fill color: {}", color);
                }
            }
        }
    }

    println!("\nStyle system is working correctly!");

    Ok(())
}
