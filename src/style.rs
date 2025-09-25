// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use std::fmt;

/// Represents a color in ARGB format
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Color {
    /// Alpha channel (0-255)
    pub alpha: u8,
    /// Red channel (0-255)
    pub red: u8,
    /// Green channel (0-255)
    pub green: u8,
    /// Blue channel (0-255)
    pub blue: u8,
}

impl Color {
    /// Create a new color from ARGB values
    pub fn new(alpha: u8, red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha,
            red,
            green,
            blue,
        }
    }

    /// Create a color from RGB values (alpha = 255)
    pub fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::new(255, red, green, blue)
    }

    /// Create a color from ARGB integer
    pub fn from_argb(argb: u32) -> Self {
        Self {
            alpha: ((argb >> 24) & 0xFF) as u8,
            red: ((argb >> 16) & 0xFF) as u8,
            green: ((argb >> 8) & 0xFF) as u8,
            blue: (argb & 0xFF) as u8,
        }
    }

    /// Convert to ARGB integer
    pub fn to_argb(&self) -> u32 {
        ((self.alpha as u32) << 24)
            | ((self.red as u32) << 16)
            | ((self.green as u32) << 8)
            | (self.blue as u32)
    }

    /// Check if the color is transparent
    pub fn is_transparent(&self) -> bool {
        self.alpha == 0
    }

    /// Check if the color is black
    pub fn is_black(&self) -> bool {
        self.red == 0 && self.green == 0 && self.blue == 0
    }

    /// Check if the color is white
    pub fn is_white(&self) -> bool {
        self.red == 255 && self.green == 255 && self.blue == 255
    }
}

impl fmt::Display for Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "#{:02X}{:02X}{:02X}", self.red, self.green, self.blue)
    }
}

/// Border style enumeration
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BorderStyle {
    /// No border
    None,
    /// Thin border
    Thin,
    /// Medium border
    Medium,
    /// Thick border
    Thick,
    /// Double border
    Double,
    /// Hair border
    Hair,
    /// Dashed border
    Dashed,
    /// Dotted border
    Dotted,
    /// Medium dashed border
    MediumDashed,
    /// Dash dot border
    DashDot,
    /// Dash dot dot border
    DashDotDot,
    /// Slant dash dot border
    SlantDashDot,
}

impl Default for BorderStyle {
    fn default() -> Self {
        BorderStyle::None
    }
}

/// Border side
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Border {
    /// Border style
    pub style: BorderStyle,
    /// Border color
    pub color: Option<Color>,
}

impl Border {
    /// Create a new border with style
    pub fn new(style: BorderStyle) -> Self {
        Self { style, color: None }
    }

    /// Create a new border with style and color
    pub fn with_color(style: BorderStyle, color: Color) -> Self {
        Self {
            style,
            color: Some(color),
        }
    }

    /// Check if border is visible
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None
    }
}

/// All borders for a cell
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Borders {
    /// Left border
    pub left: Border,
    /// Right border
    pub right: Border,
    /// Top border
    pub top: Border,
    /// Bottom border
    pub bottom: Border,
    /// Diagonal down border
    pub diagonal_down: Border,
    /// Diagonal up border
    pub diagonal_up: Border,
}

impl Borders {
    /// Create new borders
    pub fn new() -> Self {
        Self::default()
    }

    /// Check if any border is visible
    pub fn has_visible_borders(&self) -> bool {
        self.left.is_visible()
            || self.right.is_visible()
            || self.top.is_visible()
            || self.bottom.is_visible()
            || self.diagonal_down.is_visible()
            || self.diagonal_up.is_visible()
    }
}

/// Font weight
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FontWeight {
    /// Normal weight
    Normal,
    /// Bold weight
    Bold,
}

impl Default for FontWeight {
    fn default() -> Self {
        FontWeight::Normal
    }
}

/// Font style
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FontStyle {
    /// Normal style
    Normal,
    /// Italic style
    Italic,
}

impl Default for FontStyle {
    fn default() -> Self {
        FontStyle::Normal
    }
}

/// Underline style
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnderlineStyle {
    /// No underline
    None,
    /// Single underline
    Single,
    /// Double underline
    Double,
    /// Single accounting underline
    SingleAccounting,
    /// Double accounting underline
    DoubleAccounting,
}

impl Default for UnderlineStyle {
    fn default() -> Self {
        UnderlineStyle::None
    }
}

/// Font properties
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Font {
    /// Font name
    pub name: Option<String>,
    /// Font size in points
    pub size: Option<f64>,
    /// Font weight
    pub weight: FontWeight,
    /// Font style
    pub style: FontStyle,
    /// Underline style
    pub underline: UnderlineStyle,
    /// Strikethrough
    pub strikethrough: bool,
    /// Font color
    pub color: Option<Color>,
    /// Font family
    pub family: Option<String>,
}

impl Font {
    /// Create a new font
    pub fn new() -> Self {
        Self::default()
    }

    /// Set font name
    pub fn with_name(mut self, name: String) -> Self {
        self.name = Some(name);
        self
    }

    /// Set font size
    pub fn with_size(mut self, size: f64) -> Self {
        self.size = Some(size);
        self
    }

    /// Set font weight
    pub fn with_weight(mut self, weight: FontWeight) -> Self {
        self.weight = weight;
        self
    }

    /// Set font style
    pub fn with_style(mut self, style: FontStyle) -> Self {
        self.style = style;
        self
    }

    /// Set underline
    pub fn with_underline(mut self, underline: UnderlineStyle) -> Self {
        self.underline = underline;
        self
    }

    /// Set strikethrough
    pub fn with_strikethrough(mut self, strikethrough: bool) -> Self {
        self.strikethrough = strikethrough;
        self
    }

    /// Set font color
    pub fn with_color(mut self, color: Color) -> Self {
        self.color = Some(color);
        self
    }

    /// Set font family
    pub fn with_family(mut self, family: String) -> Self {
        self.family = Some(family);
        self
    }

    /// Check if font is bold
    pub fn is_bold(&self) -> bool {
        self.weight == FontWeight::Bold
    }

    /// Check if font is italic
    pub fn is_italic(&self) -> bool {
        self.style == FontStyle::Italic
    }

    /// Check if font has underline
    pub fn has_underline(&self) -> bool {
        self.underline != UnderlineStyle::None
    }

    /// Check if font has strikethrough
    pub fn has_strikethrough(&self) -> bool {
        self.strikethrough
    }
}

/// Horizontal alignment
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum HorizontalAlignment {
    /// Left alignment
    Left,
    /// Center alignment
    Center,
    /// Right alignment
    Right,
    /// Justify alignment
    Justify,
    /// Distributed alignment
    Distributed,
    /// Fill alignment
    Fill,
    /// General alignment (default)
    General,
}

impl Default for HorizontalAlignment {
    fn default() -> Self {
        HorizontalAlignment::General
    }
}

/// Vertical alignment
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum VerticalAlignment {
    /// Top alignment
    Top,
    /// Center alignment
    Center,
    /// Bottom alignment
    Bottom,
    /// Justify alignment
    Justify,
    /// Distributed alignment
    Distributed,
}

impl Default for VerticalAlignment {
    fn default() -> Self {
        VerticalAlignment::Bottom
    }
}

/// Text rotation in degrees
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum TextRotation {
    /// No rotation
    None,
    /// Rotated by degrees (0-180)
    Degrees(u16),
    /// Stacked text
    Stacked,
}

impl Default for TextRotation {
    fn default() -> Self {
        TextRotation::None
    }
}

/// Cell alignment properties
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Alignment {
    /// Horizontal alignment
    pub horizontal: HorizontalAlignment,
    /// Vertical alignment
    pub vertical: VerticalAlignment,
    /// Text rotation
    pub text_rotation: TextRotation,
    /// Wrap text
    pub wrap_text: bool,
    /// Indent level
    pub indent: Option<u8>,
    /// Shrink to fit
    pub shrink_to_fit: bool,
}

impl Alignment {
    /// Create new alignment
    pub fn new() -> Self {
        Self::default()
    }

    /// Set horizontal alignment
    pub fn with_horizontal(mut self, horizontal: HorizontalAlignment) -> Self {
        self.horizontal = horizontal;
        self
    }

    /// Set vertical alignment
    pub fn with_vertical(mut self, vertical: VerticalAlignment) -> Self {
        self.vertical = vertical;
        self
    }

    /// Set text rotation
    pub fn with_text_rotation(mut self, rotation: TextRotation) -> Self {
        self.text_rotation = rotation;
        self
    }

    /// Set wrap text
    pub fn with_wrap_text(mut self, wrap: bool) -> Self {
        self.wrap_text = wrap;
        self
    }

    /// Set indent level
    pub fn with_indent(mut self, indent: u8) -> Self {
        self.indent = Some(indent);
        self
    }

    /// Set shrink to fit
    pub fn with_shrink_to_fit(mut self, shrink: bool) -> Self {
        self.shrink_to_fit = shrink;
        self
    }
}

/// Fill pattern type
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum FillPattern {
    /// No fill
    None,
    /// Solid fill
    Solid,
    /// Dark gray pattern
    DarkGray,
    /// Medium gray pattern
    MediumGray,
    /// Light gray pattern
    LightGray,
    /// Gray 125 pattern
    Gray125,
    /// Gray 0625 pattern
    Gray0625,
    /// Dark horizontal pattern
    DarkHorizontal,
    /// Dark vertical pattern
    DarkVertical,
    /// Dark down pattern
    DarkDown,
    /// Dark up pattern
    DarkUp,
    /// Dark grid pattern
    DarkGrid,
    /// Dark trellis pattern
    DarkTrellis,
    /// Light horizontal pattern
    LightHorizontal,
    /// Light vertical pattern
    LightVertical,
    /// Light down pattern
    LightDown,
    /// Light up pattern
    LightUp,
    /// Light grid pattern
    LightGrid,
    /// Light trellis pattern
    LightTrellis,
}

impl Default for FillPattern {
    fn default() -> Self {
        FillPattern::None
    }
}

/// Fill properties
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Fill {
    /// Fill pattern
    pub pattern: FillPattern,
    /// Foreground color
    pub foreground_color: Option<Color>,
    /// Background color
    pub background_color: Option<Color>,
}

impl Fill {
    /// Create new fill
    pub fn new() -> Self {
        Self::default()
    }

    /// Create solid fill with color
    pub fn solid(color: Color) -> Self {
        Self {
            pattern: FillPattern::Solid,
            foreground_color: Some(color),
            background_color: None,
        }
    }

    /// Set pattern
    pub fn with_pattern(mut self, pattern: FillPattern) -> Self {
        self.pattern = pattern;
        self
    }

    /// Set foreground color
    pub fn with_foreground_color(mut self, color: Color) -> Self {
        self.foreground_color = Some(color);
        self
    }

    /// Set background color
    pub fn with_background_color(mut self, color: Color) -> Self {
        self.background_color = Some(color);
        self
    }

    /// Check if fill is visible
    pub fn is_visible(&self) -> bool {
        self.pattern != FillPattern::None
    }

    /// Get the main fill color (foreground if available, otherwise background)
    pub fn get_color(&self) -> Option<Color> {
        self.foreground_color.or(self.background_color)
    }
}

/// Number format
#[derive(Debug, Clone, PartialEq)]
pub struct NumberFormat {
    /// Format code
    pub format_code: String,
    /// Format ID
    pub format_id: Option<u32>,
}

impl NumberFormat {
    /// Create new number format
    pub fn new(format_code: String) -> Self {
        Self {
            format_code,
            format_id: None,
        }
    }

    /// Create with format ID
    pub fn with_id(mut self, format_id: u32) -> Self {
        self.format_id = Some(format_id);
        self
    }
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self {
            format_code: "General".to_string(),
            format_id: None,
        }
    }
}

/// Cell protection properties
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Protection {
    /// Cell is locked
    pub locked: bool,
    /// Cell is hidden
    pub hidden: bool,
}

impl Protection {
    /// Create new protection
    pub fn new() -> Self {
        Self::default()
    }

    /// Set locked
    pub fn with_locked(mut self, locked: bool) -> Self {
        self.locked = locked;
        self
    }

    /// Set hidden
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }
}

/// Column width information
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnWidth {
    /// Column index (0-based)
    pub column: u32,
    /// Width in Excel units (characters)
    pub width: f64,
    /// Whether the width is custom (manually set)
    pub custom_width: bool,
    /// Whether the column is hidden
    pub hidden: bool,
    /// Best fit width
    pub best_fit: bool,
}

impl ColumnWidth {
    /// Create a new column width
    pub fn new(column: u32, width: f64) -> Self {
        Self {
            column,
            width,
            custom_width: false,
            hidden: false,
            best_fit: false,
        }
    }

    /// Set custom width flag
    pub fn with_custom_width(mut self, custom: bool) -> Self {
        self.custom_width = custom;
        self
    }

    /// Set hidden flag
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }

    /// Set best fit flag
    pub fn with_best_fit(mut self, best_fit: bool) -> Self {
        self.best_fit = best_fit;
        self
    }

    /// Check if column is visible
    pub fn is_visible(&self) -> bool {
        !self.hidden
    }
}

/// Row height information
#[derive(Debug, Clone, PartialEq)]
pub struct RowHeight {
    /// Row index (0-based)
    pub row: u32,
    /// Height in points
    pub height: f64,
    /// Whether the height is custom (manually set)
    pub custom_height: bool,
    /// Whether the row is hidden
    pub hidden: bool,
    /// Thick top border
    pub thick_top: bool,
    /// Thick bottom border
    pub thick_bottom: bool,
}

impl RowHeight {
    /// Create a new row height
    pub fn new(row: u32, height: f64) -> Self {
        Self {
            row,
            height,
            custom_height: false,
            hidden: false,
            thick_top: false,
            thick_bottom: false,
        }
    }

    /// Set custom height flag
    pub fn with_custom_height(mut self, custom: bool) -> Self {
        self.custom_height = custom;
        self
    }

    /// Set hidden flag
    pub fn with_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }

    /// Set thick top border
    pub fn with_thick_top(mut self, thick_top: bool) -> Self {
        self.thick_top = thick_top;
        self
    }

    /// Set thick bottom border
    pub fn with_thick_bottom(mut self, thick_bottom: bool) -> Self {
        self.thick_bottom = thick_bottom;
        self
    }

    /// Check if row is visible
    pub fn is_visible(&self) -> bool {
        !self.hidden
    }
}

/// Worksheet layout information
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WorksheetLayout {
    /// Column widths
    pub column_widths: Vec<ColumnWidth>,
    /// Row heights
    pub row_heights: Vec<RowHeight>,
    /// Default column width
    pub default_column_width: Option<f64>,
    /// Default row height
    pub default_row_height: Option<f64>,
}

impl WorksheetLayout {
    /// Create a new worksheet layout
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a column width
    pub fn add_column_width(mut self, column_width: ColumnWidth) -> Self {
        self.column_widths.push(column_width);
        self
    }

    /// Add a row height
    pub fn add_row_height(mut self, row_height: RowHeight) -> Self {
        self.row_heights.push(row_height);
        self
    }

    /// Set default column width
    pub fn with_default_column_width(mut self, width: f64) -> Self {
        self.default_column_width = Some(width);
        self
    }

    /// Set default row height
    pub fn with_default_row_height(mut self, height: f64) -> Self {
        self.default_row_height = Some(height);
        self
    }

    /// Get column width for a specific column
    pub fn get_column_width(&self, column: u32) -> Option<&ColumnWidth> {
        self.column_widths.iter().find(|cw| cw.column == column)
    }

    /// Get row height for a specific row
    pub fn get_row_height(&self, row: u32) -> Option<&RowHeight> {
        self.row_heights.iter().find(|rh| rh.row == row)
    }

    /// Get effective column width (custom or default)
    pub fn get_effective_column_width(&self, column: u32) -> f64 {
        self.get_column_width(column)
            .map(|cw| cw.width)
            .or(self.default_column_width)
            .unwrap_or(8.43) // Excel default column width
    }

    /// Get effective row height (custom or default)
    pub fn get_effective_row_height(&self, row: u32) -> f64 {
        self.get_row_height(row)
            .map(|rh| rh.height)
            .or(self.default_row_height)
            .unwrap_or(15.0) // Excel default row height
    }

    /// Check if layout has any custom dimensions
    pub fn has_custom_dimensions(&self) -> bool {
        !self.column_widths.is_empty()
            || !self.row_heights.is_empty()
            || self.default_column_width.is_some()
            || self.default_row_height.is_some()
    }
}

/// Complete cell style
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    /// Font properties
    pub font: Option<Font>,
    /// Fill properties
    pub fill: Option<Fill>,
    /// Border properties
    pub borders: Option<Borders>,
    /// Alignment properties
    pub alignment: Option<Alignment>,
    /// Number format
    pub number_format: Option<NumberFormat>,
    /// Protection properties
    pub protection: Option<Protection>,
    /// Style ID (for internal use)
    pub style_id: Option<u32>,
}

impl Style {
    /// Create new style
    pub fn new() -> Self {
        Self::default()
    }

    /// Set font
    pub fn with_font(mut self, font: Font) -> Self {
        self.font = Some(font);
        self
    }

    /// Set fill
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Set borders
    pub fn with_borders(mut self, borders: Borders) -> Self {
        self.borders = Some(borders);
        self
    }

    /// Set alignment
    pub fn with_alignment(mut self, alignment: Alignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    /// Set number format
    pub fn with_number_format(mut self, number_format: NumberFormat) -> Self {
        self.number_format = Some(number_format);
        self
    }

    /// Set protection
    pub fn with_protection(mut self, protection: Protection) -> Self {
        self.protection = Some(protection);
        self
    }

    /// Set style ID
    pub fn with_style_id(mut self, style_id: u32) -> Self {
        self.style_id = Some(style_id);
        self
    }

    /// Get font
    pub fn get_font(&self) -> Option<&Font> {
        self.font.as_ref()
    }

    /// Get fill
    pub fn get_fill(&self) -> Option<&Fill> {
        self.fill.as_ref()
    }

    /// Get borders
    pub fn get_borders(&self) -> Option<&Borders> {
        self.borders.as_ref()
    }

    /// Get alignment
    pub fn get_alignment(&self) -> Option<&Alignment> {
        self.alignment.as_ref()
    }

    /// Get number format
    pub fn get_number_format(&self) -> Option<&NumberFormat> {
        self.number_format.as_ref()
    }

    /// Get protection
    pub fn get_protection(&self) -> Option<&Protection> {
        self.protection.as_ref()
    }

    /// Check if style is empty (no properties set)
    pub fn is_empty(&self) -> bool {
        self.font.is_none()
            && self.fill.is_none()
            && self.borders.is_none()
            && self.alignment.is_none()
            && self.number_format.is_none()
            && self.protection.is_none()
    }

    /// Check if style has any visible properties
    pub fn has_visible_properties(&self) -> bool {
        (self
            .font
            .as_ref()
            .map_or(false, |f| f.color.is_some() || f.is_bold() || f.is_italic()))
            || (self.fill.as_ref().map_or(false, |f| f.is_visible()))
            || (self
                .borders
                .as_ref()
                .map_or(false, |b| b.has_visible_borders()))
            || (self.alignment.as_ref().map_or(false, |a| {
                a.horizontal != HorizontalAlignment::General
                    || a.vertical != VerticalAlignment::Bottom
                    || a.text_rotation != TextRotation::None
                    || a.wrap_text
                    || a.indent.is_some()
                    || a.shrink_to_fit
            }))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_color() {
        let color = Color::rgb(255, 0, 128);
        assert_eq!(color.red, 255);
        assert_eq!(color.green, 0);
        assert_eq!(color.blue, 128);
        assert_eq!(color.alpha, 255);
        assert_eq!(color.to_string(), "#FF0080");
    }

    #[test]
    fn test_font() {
        let font = Font::new()
            .with_name("Arial".to_string())
            .with_size(12.0)
            .with_weight(FontWeight::Bold)
            .with_color(Color::rgb(255, 0, 0));

        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.size, Some(12.0));
        assert!(font.is_bold());
        assert_eq!(font.color, Some(Color::rgb(255, 0, 0)));
    }

    #[test]
    fn test_style() {
        let style = Style::new()
            .with_font(Font::new().with_name("Arial".to_string()))
            .with_fill(Fill::solid(Color::rgb(255, 255, 0)));

        assert!(!style.is_empty());
        assert!(style.get_font().is_some());
        assert!(style.get_fill().is_some());
    }

    #[test]
    fn test_border_with_color() {
        let border = Border::with_color(BorderStyle::Thin, Color::rgb(255, 0, 0));
        assert_eq!(border.style, BorderStyle::Thin);
        assert_eq!(border.color, Some(Color::rgb(255, 0, 0)));
        assert!(border.is_visible());
    }

    #[test]
    fn test_border_without_color() {
        let border = Border::new(BorderStyle::Medium);
        assert_eq!(border.style, BorderStyle::Medium);
        assert_eq!(border.color, None);
        assert!(border.is_visible());
    }

    #[test]
    fn test_borders_with_mixed_colors() {
        let mut borders = Borders::new();
        borders.left = Border::with_color(BorderStyle::Thin, Color::rgb(255, 0, 0));
        borders.right = Border::new(BorderStyle::Medium);
        borders.top = Border::with_color(BorderStyle::Thick, Color::rgb(0, 255, 0));

        assert_eq!(borders.left.color, Some(Color::rgb(255, 0, 0)));
        assert_eq!(borders.right.color, None);
        assert_eq!(borders.top.color, Some(Color::rgb(0, 255, 0)));
        assert!(borders.has_visible_borders());
    }

    #[test]
    fn test_column_width() {
        let column_width = ColumnWidth::new(5, 12.5)
            .with_custom_width(true)
            .with_hidden(false)
            .with_best_fit(true);

        assert_eq!(column_width.column, 5);
        assert_eq!(column_width.width, 12.5);
        assert!(column_width.custom_width);
        assert!(!column_width.hidden);
        assert!(column_width.best_fit);
        assert!(column_width.is_visible());
    }

    #[test]
    fn test_row_height() {
        let row_height = RowHeight::new(10, 20.0)
            .with_custom_height(true)
            .with_hidden(false)
            .with_thick_top(true)
            .with_thick_bottom(false);

        assert_eq!(row_height.row, 10);
        assert_eq!(row_height.height, 20.0);
        assert!(row_height.custom_height);
        assert!(!row_height.hidden);
        assert!(row_height.thick_top);
        assert!(!row_height.thick_bottom);
        assert!(row_height.is_visible());
    }

    #[test]
    fn test_worksheet_layout() {
        let layout = WorksheetLayout::new()
            .add_column_width(ColumnWidth::new(0, 10.0))
            .add_column_width(ColumnWidth::new(1, 15.0))
            .add_row_height(RowHeight::new(0, 18.0))
            .add_row_height(RowHeight::new(1, 22.0))
            .with_default_column_width(8.43)
            .with_default_row_height(15.0);

        assert_eq!(layout.column_widths.len(), 2);
        assert_eq!(layout.row_heights.len(), 2);
        assert_eq!(layout.default_column_width, Some(8.43));
        assert_eq!(layout.default_row_height, Some(15.0));
        assert!(layout.has_custom_dimensions());

        // Test getting specific column width
        let col_width = layout.get_column_width(0).unwrap();
        assert_eq!(col_width.width, 10.0);

        // Test getting specific row height
        let row_height = layout.get_row_height(1).unwrap();
        assert_eq!(row_height.height, 22.0);

        // Test effective widths/heights
        assert_eq!(layout.get_effective_column_width(0), 10.0); // Custom width
        assert_eq!(layout.get_effective_column_width(5), 8.43); // Default width
        assert_eq!(layout.get_effective_row_height(0), 18.0); // Custom height
        assert_eq!(layout.get_effective_row_height(5), 15.0); // Default height
    }

    #[test]
    fn test_worksheet_layout_defaults() {
        let layout = WorksheetLayout::new();

        assert!(!layout.has_custom_dimensions());
        assert_eq!(layout.get_effective_column_width(0), 8.43); // Excel default
        assert_eq!(layout.get_effective_row_height(0), 15.0); // Excel default
    }
}
