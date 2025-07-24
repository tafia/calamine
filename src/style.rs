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
}
