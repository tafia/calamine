use calamine::{open_workbook, Data, DataType, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("tests/styles.xlsx")?;

    println!("Testing color parsing with styles.xlsx...");

    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        for (row, col, cell) in range.cells() {
            match cell {
                Data::String(s) if !s.is_empty() => {
                    println!(
                        "Cell ({}, {}): {} - Has style: {}",
                        row,
                        col,
                        s,
                        cell.has_style()
                    );

                    if cell.has_style() {
                        if let Some(style) = cell.get_style() {
                            if let Some(font) = style.get_font() {
                                match font.color {
                                    Some(color) => println!(
                                        "  Font color: RGB({}, {}, {})",
                                        color.red, color.green, color.blue
                                    ),
                                    None => println!("  Font color: None"),
                                }
                            }
                        }
                    }
                }
                Data::Float(f) => {
                    println!(
                        "Cell ({}, {}): {} - Has style: {}",
                        row,
                        col,
                        f,
                        cell.has_style()
                    );

                    if cell.has_style() {
                        if let Some(style) = cell.get_style() {
                            if let Some(font) = style.get_font() {
                                match font.color {
                                    Some(color) => println!(
                                        "  Font color: RGB({}, {}, {})",
                                        color.red, color.green, color.blue
                                    ),
                                    None => println!("  Font color: None"),
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
        }
    }

    Ok(())
}
