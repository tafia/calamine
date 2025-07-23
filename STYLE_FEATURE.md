# Style Support in Calamine

This document describes the new style extraction functionality added to calamine, inspired by the umya-spreadsheet library.

## Overview

Calamine now supports extracting style information from Excel files, including:
- Font properties (name, size, weight, color, etc.)
- Fill properties (background colors, patterns)
- Border properties (style, color, position)
- Alignment properties (horizontal, vertical, text rotation, etc.)
- Protection properties (locked, hidden)

## Data Structures

### Style Components

The style system is built around several key data structures:

#### Color
```rust
use calamine::Color;

let red_color = Color::rgb(255, 0, 0);
let custom_color = Color::new(255, 128, 64, 32); // ARGB
```

#### Font
```rust
use calamine::{Font, FontWeight, FontStyle, UnderlineStyle};

let font = Font::new()
    .with_name("Arial".to_string())
    .with_size(12.0)
    .with_weight(FontWeight::Bold)
    .with_style(FontStyle::Italic)
    .with_color(Color::rgb(255, 0, 0));
```

#### Fill
```rust
use calamine::{Fill, FillPattern};

let fill = Fill::solid(Color::rgb(255, 255, 0));
let pattern_fill = Fill::new()
    .with_pattern(FillPattern::DarkGray)
    .with_foreground_color(Color::rgb(255, 0, 0))
    .with_background_color(Color::rgb(0, 0, 255));
```

#### Borders
```rust
use calamine::{Borders, Border, BorderStyle};

let borders = Borders::new();
let border = Border::with_color(BorderStyle::Thick, Color::rgb(0, 0, 0));
```

#### Alignment
```rust
use calamine::{Alignment, HorizontalAlignment, VerticalAlignment, TextRotation};

let alignment = Alignment::new()
    .with_horizontal(HorizontalAlignment::Center)
    .with_vertical(VerticalAlignment::Middle)
    .with_wrap_text(true);
```

#### Complete Style
```rust
use calamine::Style;

let style = Style::new()
    .with_font(font)
    .with_fill(fill)
    .with_borders(borders)
    .with_alignment(alignment);
```

## Usage Examples

### Reading Styles from Excel Files

```rust
use calamine::{open_workbook, Reader, Data};

let mut workbook = open_workbook("file.xlsx")?;

if let Ok(range) = workbook.worksheet_range("Sheet1") {
    for (row, col, cell) in range.cells() {
        if let Some(cell_data) = cell {
            if cell_data.has_style() {
                if let Some(style) = cell_data.get_style() {
                    // Access font properties
                    if let Some(font) = style.get_font() {
                        println!("Font: {}", font.name.as_deref().unwrap_or("Unknown"));
                        println!("Size: {}", font.size.unwrap_or(0.0));
                        println!("Bold: {}", font.is_bold());
                        if let Some(color) = font.color {
                            println!("Color: {}", color);
                        }
                    }
                    
                    // Access fill properties
                    if let Some(fill) = style.get_fill() {
                        if fill.is_visible() {
                            println!("Has fill");
                            if let Some(color) = fill.get_color() {
                                println!("Fill color: {}", color);
                            }
                        }
                    }
                    
                    // Access border properties
                    if let Some(borders) = style.get_borders() {
                        if borders.has_visible_borders() {
                            println!("Has borders");
                            if borders.left.is_visible() {
                                println!("Left border");
                            }
                        }
                    }
                }
            }
        }
    }
}
```

### Creating Cells with Styles

```rust
use calamine::{Cell, Data, Style, Font, FontWeight, Color};

let style = Style::new()
    .with_font(Font::new()
        .with_name("Arial".to_string())
        .with_size(12.0)
        .with_weight(FontWeight::Bold)
        .with_color(Color::rgb(255, 0, 0)));

let cell = Cell::with_style((0, 0), Data::String("Hello".to_string()), style);
```

### Working with CellData

```rust
use calamine::{CellData, Data, Style};

let cell_data = CellData::with_style(
    Data::Int(42),
    Style::new().with_font(Font::new().with_weight(FontWeight::Bold))
);

if cell_data.has_style() {
    if let Some(style) = cell_data.get_style() {
        // Access style properties
    }
}
```

## Supported Formats

Currently, style extraction is supported for:
- **XLSX**: Full style support including fonts, fills, borders, and alignment
- **XLSB**: Basic style support (format-based)
- **XLS**: Basic style support (format-based)
- **ODS**: Basic style support (format-based)

## Style Parsing

The style parser extracts information from the Excel styles.xml file, including:

### Font Properties
- Font name
- Font size
- Font weight (bold/normal)
- Font style (italic/normal)
- Underline style
- Strikethrough
- Font color
- Font family

### Fill Properties
- Fill pattern (solid, patterns, etc.)
- Foreground color
- Background color

### Border Properties
- Border style (thin, medium, thick, etc.)
- Border color
- Border position (left, right, top, bottom, diagonal)

### Alignment Properties
- Horizontal alignment (left, center, right, justify, etc.)
- Vertical alignment (top, center, bottom, justify, etc.)
- Text rotation
- Wrap text
- Indent level
- Shrink to fit

### Protection Properties
- Cell locked
- Cell hidden

## Limitations

1. **Theme Colors**: Theme color support is limited and may not fully match Excel's rendering
2. **Indexed Colors**: Indexed color support is basic
3. **Complex Patterns**: Some complex fill patterns may not be fully supported
4. **Conditional Formatting**: Conditional formatting styles are not yet supported

## Future Enhancements

Planned improvements include:
- Full theme color support
- Conditional formatting style extraction
- Style writing capabilities
- Enhanced pattern support
- Better color space handling

## API Reference

### Core Types

- `Style`: Complete cell style container
- `Font`: Font properties
- `Fill`: Fill properties
- `Borders`: Border properties
- `Alignment`: Alignment properties
- `Protection`: Protection properties
- `Color`: Color representation
- `CellData`: Cell value with optional style

### Key Methods

- `Cell::with_style()`: Create a cell with style
- `Cell::get_style()`: Get cell style
- `Cell::has_style()`: Check if cell has style
- `Style::is_empty()`: Check if style has any properties
- `Style::has_visible_properties()`: Check if style has visible properties

## Migration Guide

For existing code, the new style functionality is backward compatible. Existing code will continue to work without changes. To add style support:

1. Update your cell iteration to check for styles
2. Use `Cell::with_style()` when creating cells with styles
3. Access style properties through the style getter methods

## Examples

See the `examples/style_example.rs` file for a complete working example of style extraction and usage. 