# Style Reading Support in Calamine

This document describes the style extraction functionality in calamine for reading style information from Excel files.

## Overview

Calamine supports extracting style information from Excel files, including:
- Font properties (name, size, weight, color, etc.)
- Fill properties (background colors, patterns)
- Border properties (style, color, position)
- Alignment properties (horizontal, vertical, text rotation, etc.)
- Protection properties (locked, hidden)

## Data Structures

When reading styles from Excel files, you'll work with these data structures:

### Color
```rust
// Colors are returned when reading style information
// You can check color properties like:
if let Some(color) = font.color {
    let (r, g, b, a) = color.rgba();
    println!("Color RGBA: {}, {}, {}, {}", r, g, b, a);
}
```

### Font
```rust
// Font information extracted from Excel files
if let Some(font) = style.get_font() {
    // Access font properties
    if let Some(name) = &font.name {
        println!("Font name: {}", name);
    }
    if let Some(size) = font.size {
        println!("Font size: {}", size);
    }
    println!("Bold: {}", font.is_bold());
    println!("Italic: {}", font.is_italic());
    if let Some(color) = font.color {
        println!("Font color: {}", color);
    }
}
```

### Fill
```rust
// Fill information extracted from Excel files
if let Some(fill) = style.get_fill() {
    if fill.is_visible() {
        println!("Cell has background fill");
        if let Some(color) = fill.get_color() {
            println!("Fill color: {}", color);
        }
        if let Some(pattern) = fill.pattern {
            println!("Fill pattern: {:?}", pattern);
        }
    }
}
```

### Borders
```rust
// Border information extracted from Excel files
if let Some(borders) = style.get_borders() {
    if borders.has_visible_borders() {
        println!("Cell has borders");
        if borders.left.is_visible() {
            println!("Left border style: {:?}", borders.left.style);
            if let Some(color) = borders.left.color {
                println!("Left border color: {}", color);
            }
        }
        // Similar for right, top, bottom borders
    }
}
```

### Alignment
```rust
// Alignment information extracted from Excel files
if let Some(alignment) = style.get_alignment() {
    if let Some(horizontal) = alignment.horizontal {
        println!("Horizontal alignment: {:?}", horizontal);
    }
    if let Some(vertical) = alignment.vertical {
        println!("Vertical alignment: {:?}", vertical);
    }
    if alignment.wrap_text.unwrap_or(false) {
        println!("Text wrapping enabled");
    }
}
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
                    println!("Cell ({}, {}) has style:", row, col);
                    
                    // Access font properties
                    if let Some(font) = style.get_font() {
                        println!("  Font: {}", font.name.as_deref().unwrap_or("Unknown"));
                        println!("  Size: {}", font.size.unwrap_or(0.0));
                        println!("  Bold: {}", font.is_bold());
                        println!("  Italic: {}", font.is_italic());
                        if let Some(color) = font.color {
                            let (r, g, b, _) = color.rgba();
                            println!("  Color: rgb({}, {}, {})", r, g, b);
                        }
                    }
                    
                    // Access fill properties
                    if let Some(fill) = style.get_fill() {
                        if fill.is_visible() {
                            println!("  Has background fill");
                            if let Some(color) = fill.get_color() {
                                let (r, g, b, _) = color.rgba();
                                println!("  Fill color: rgb({}, {}, {})", r, g, b);
                            }
                        }
                    }
                    
                    // Access border properties
                    if let Some(borders) = style.get_borders() {
                        if borders.has_visible_borders() {
                            println!("  Has borders:");
                            if borders.left.is_visible() {
                                println!("    Left: {:?}", borders.left.style);
                            }
                            if borders.right.is_visible() {
                                println!("    Right: {:?}", borders.right.style);
                            }
                            if borders.top.is_visible() {
                                println!("    Top: {:?}", borders.top.style);
                            }
                            if borders.bottom.is_visible() {
                                println!("    Bottom: {:?}", borders.bottom.style);
                            }
                        }
                    }
                    
                    // Access alignment properties
                    if let Some(alignment) = style.get_alignment() {
                        if let Some(horizontal) = alignment.horizontal {
                            println!("  Horizontal alignment: {:?}", horizontal);
                        }
                        if let Some(vertical) = alignment.vertical {
                            println!("  Vertical alignment: {:?}", vertical);
                        }
                        if alignment.wrap_text.unwrap_or(false) {
                            println!("  Text wrapping: enabled");
                        }
                    }
                }
            }
        }
    }
}
```

### Checking for Specific Style Properties

```rust
use calamine::{open_workbook, Reader, FontWeight, HorizontalAlignment};

let mut workbook = open_workbook("file.xlsx")?;

if let Ok(range) = workbook.worksheet_range("Sheet1") {
    for (row, col, cell) in range.cells() {
        if let Some(cell_data) = cell {
            if cell_data.has_style() {
                if let Some(style) = cell_data.get_style() {
                    // Check for bold text
                    if let Some(font) = style.get_font() {
                        if font.is_bold() {
                            println!("Cell ({}, {}) has bold text", row, col);
                        }
                    }
                    
                    // Check for center alignment
                    if let Some(alignment) = style.get_alignment() {
                        if alignment.horizontal == Some(HorizontalAlignment::Center) {
                            println!("Cell ({}, {}) is center-aligned", row, col);
                        }
                    }
                    
                    // Check for background color
                    if let Some(fill) = style.get_fill() {
                        if fill.is_visible() {
                            println!("Cell ({}, {}) has background color", row, col);
                        }
                    }
                }
            }
        }
    }
}
```

### Working with CellData

```rust
use calamine::{CellData, Data};

// When iterating through cells, you get CellData which may contain style information
fn process_cell(cell_data: &CellData) {
    // Check if cell has any style information
    if cell_data.has_style() {
        println!("Cell value: {:?}", cell_data.get_value());
        
        if let Some(style) = cell_data.get_style() {
            if !style.is_empty() {
                println!("Cell has formatting");
                
                // Process style information as shown in previous examples
                if let Some(font) = style.get_font() {
                    if font.is_bold() {
                        println!("Text is bold");
                    }
                }
            }
        }
    }
}
```

## Supported Formats

Style extraction is supported for:
- [x] **XLSX**: Full style support including fonts, fills, borders, and alignment
- [ ] **XLSB**: Basic style support (format-based)
- [ ] **XLS**: Basic style support (format-based)
- [ ] **ODS**: Basic style support (format-based)

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
4. **Conditional Formatting**: Conditional formatting styles are not extracted

## Future Enhancements

Planned improvements for style reading include:
- Full theme color support
- Conditional formatting style extraction
- Enhanced pattern support
- Better color space handling

## API Reference

### Key Types for Reading Styles

- `Style`: Complete cell style container (read-only)
- `Font`: Font properties (read-only)
- `Fill`: Fill properties (read-only)
- `Borders`: Border properties (read-only)
- `Alignment`: Alignment properties (read-only)
- `Protection`: Protection properties (read-only)
- `Color`: Color representation (read-only)
- `CellData`: Cell value with optional style information

### Key Methods for Reading Styles

- `CellData::get_style()`: Get cell style information
- `CellData::has_style()`: Check if cell has style information
- `Style::get_font()`: Get font properties
- `Style::get_fill()`: Get fill properties
- `Style::get_borders()`: Get border properties
- `Style::get_alignment()`: Get alignment properties
- `Style::is_empty()`: Check if style has any properties
- `Style::has_visible_properties()`: Check if style has visible properties
- `Font::is_bold()`: Check if font is bold
- `Font::is_italic()`: Check if font is italic
- `Fill::is_visible()`: Check if fill is visible
- `Borders::has_visible_borders()`: Check if borders are visible

## Migration Guide

The style reading functionality is backward compatible. Existing code will continue to work without changes. To add style reading support:

1. Update your cell iteration to check for styles using `has_style()`
2. Use `get_style()` to access style information
3. Access specific style properties through the getter methods
4. Check for visibility using methods like `is_visible()` and `has_visible_borders()`

## Examples

See the `examples/style.rs` file for a complete working example of style extraction and usage. 