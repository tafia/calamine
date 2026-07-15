// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Cell style types for spreadsheet formatting.
//!
//! This module contains the data model for cell formatting: colors, fonts,
//! fills, borders, alignment, number formats and protection, combined into a
//! [`Style`] or [`StyleRange`].

use std::fmt;

/// Represents a color in ARGB format.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Color {
    /// Alpha channel (0-255).
    pub alpha: u8,

    /// Red channel (0-255).
    pub red: u8,

    /// Green channel (0-255).
    pub green: u8,

    /// Blue channel (0-255).
    pub blue: u8,
}

impl Color {
    /// Create a new color from ARGB values.
    pub fn new(alpha: u8, red: u8, green: u8, blue: u8) -> Self {
        Self {
            alpha,
            red,
            green,
            blue,
        }
    }

    /// Create a color from RGB values (alpha = 255).
    pub fn rgb(red: u8, green: u8, blue: u8) -> Self {
        Self::new(255, red, green, blue)
    }

    /// Create a color from an ARGB integer.
    pub fn from_argb(argb: u32) -> Self {
        Self {
            alpha: ((argb >> 24) & 0xFF) as u8,
            red: ((argb >> 16) & 0xFF) as u8,
            green: ((argb >> 8) & 0xFF) as u8,
            blue: (argb & 0xFF) as u8,
        }
    }

    /// Convert to an ARGB integer.
    pub fn to_argb(&self) -> u32 {
        ((self.alpha as u32) << 24)
            | ((self.red as u32) << 16)
            | ((self.green as u32) << 8)
            | (self.blue as u32)
    }

    /// Check if the color is black.
    pub fn is_black(&self) -> bool {
        self.red == 0 && self.green == 0 && self.blue == 0
    }

    /// Check if the color is white.
    pub fn is_white(&self) -> bool {
        self.red == 255 && self.green == 255 && self.blue == 255
    }
}

impl fmt::Display for Color {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "#{:02X}{:02X}{:02X}", self.red, self.green, self.blue)
    }
}

/// Border style enumeration.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum BorderStyle {
    /// No border.
    #[default]
    None,

    /// Thin border.
    Thin,

    /// Medium border.
    Medium,

    /// Thick border.
    Thick,

    /// Double border.
    Double,

    /// Hair border.
    Hair,

    /// Dashed border.
    Dashed,

    /// Dotted border.
    Dotted,

    /// Medium dashed border.
    MediumDashed,

    /// Dash dot border.
    DashDot,

    /// Dash dot dot border.
    DashDotDot,

    /// Slant dash dot border.
    SlantDashDot,
}

/// Border side.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Border {
    /// Border style.
    pub style: BorderStyle,

    /// Border color.
    pub color: Option<Color>,
}

impl Border {
    // Create a new border with style.
    pub(crate) fn new(style: BorderStyle) -> Self {
        Self { style, color: None }
    }

    // Create a new border with style and color.
    pub(crate) fn with_color(style: BorderStyle, color: Color) -> Self {
        Self {
            style,
            color: Some(color),
        }
    }

    /// Check if border is visible.
    pub fn is_visible(&self) -> bool {
        self.style != BorderStyle::None
    }
}

/// All borders for a cell.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Borders {
    /// Left border.
    pub left: Border,

    /// Right border.
    pub right: Border,

    /// Top border.
    pub top: Border,

    /// Bottom border.
    pub bottom: Border,

    /// Diagonal down border.
    pub diagonal_down: Border,

    /// Diagonal up border.
    pub diagonal_up: Border,
}

impl Borders {
    // Create new borders.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    /// Check if any border is visible.
    pub fn has_visible_borders(&self) -> bool {
        self.left.is_visible()
            || self.right.is_visible()
            || self.top.is_visible()
            || self.bottom.is_visible()
            || self.diagonal_down.is_visible()
            || self.diagonal_up.is_visible()
    }
}

/// Font weight.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FontWeight {
    /// Normal weight.
    #[default]
    Normal,

    /// Bold weight.
    Bold,
}

/// Font style.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FontStyle {
    /// Normal style.
    #[default]
    Normal,

    /// Italic style.
    Italic,
}

/// Underline style.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum UnderlineStyle {
    /// No underline.
    #[default]
    None,

    /// Single underline.
    Single,

    /// Double underline.
    Double,

    /// Single accounting underline.
    SingleAccounting,

    /// Double accounting underline.
    DoubleAccounting,
}

/// Font properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Font {
    /// Font name.
    pub name: Option<String>,

    /// Font size in points.
    pub size: Option<f64>,

    /// Font weight.
    pub weight: FontWeight,

    /// Font style.
    pub style: FontStyle,

    /// Underline style.
    pub underline: UnderlineStyle,

    /// Strikethrough.
    pub strikethrough: bool,

    /// Font color.
    pub color: Option<Color>,

    /// Font family class as defined by OOXML (1 = Roman, 2 = Swiss, 3 =
    /// Modern, 4 = Script, 5 = Decorative).
    pub family: Option<u8>,
}

impl Font {
    // Create a new font.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    // Set font name.
    pub(crate) fn set_name(mut self, name: String) -> Self {
        self.name = Some(name);
        self
    }

    // Set font size.
    pub(crate) fn set_size(mut self, size: f64) -> Self {
        self.size = Some(size);
        self
    }

    // Set font weight.
    pub(crate) fn set_weight(mut self, weight: FontWeight) -> Self {
        self.weight = weight;
        self
    }

    // Set font style.
    pub(crate) fn set_style(mut self, style: FontStyle) -> Self {
        self.style = style;
        self
    }

    // Set underline.
    pub(crate) fn set_underline(mut self, underline: UnderlineStyle) -> Self {
        self.underline = underline;
        self
    }

    // Set strikethrough.
    pub(crate) fn set_strikethrough(mut self, strikethrough: bool) -> Self {
        self.strikethrough = strikethrough;
        self
    }

    // Set font color.
    pub(crate) fn set_color(mut self, color: Color) -> Self {
        self.color = Some(color);
        self
    }

    // Set the font family class.
    pub(crate) fn set_family(mut self, family: u8) -> Self {
        self.family = Some(family);
        self
    }

    /// Check if font is bold.
    pub fn is_bold(&self) -> bool {
        self.weight == FontWeight::Bold
    }

    /// Check if font is italic.
    pub fn is_italic(&self) -> bool {
        self.style == FontStyle::Italic
    }

    /// Check if font has underline.
    pub fn has_underline(&self) -> bool {
        self.underline != UnderlineStyle::None
    }

    /// Check if font has strikethrough.
    pub fn has_strikethrough(&self) -> bool {
        self.strikethrough
    }
}

/// Horizontal alignment.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum HorizontalAlignment {
    /// Left alignment.
    Left,

    /// Center alignment.
    Center,

    /// Right alignment.
    Right,

    /// Justify alignment.
    Justify,

    /// Distributed alignment.
    Distributed,

    /// Fill alignment.
    Fill,

    /// General alignment (default).
    #[default]
    General,
}

/// Vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum VerticalAlignment {
    /// Top alignment.
    Top,

    /// Center alignment.
    Center,

    /// Bottom alignment.
    #[default]
    Bottom,

    /// Justify alignment.
    Justify,

    /// Distributed alignment.
    Distributed,
}

/// Text rotation in degrees.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum TextRotation {
    /// No rotation.
    #[default]
    None,

    /// Rotated by degrees (0-180).
    Degrees(u16),

    /// Stacked text.
    Stacked,
}

/// Cell alignment properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Alignment {
    /// Horizontal alignment.
    pub horizontal: HorizontalAlignment,

    /// Vertical alignment.
    pub vertical: VerticalAlignment,

    /// Text rotation.
    pub text_rotation: TextRotation,

    /// Wrap text.
    pub wrap_text: bool,

    /// Indent level.
    pub indent: Option<u8>,

    /// Shrink to fit.
    pub shrink_to_fit: bool,
}

impl Alignment {
    // Create new alignment.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    // Set horizontal alignment.
    pub(crate) fn set_horizontal(mut self, horizontal: HorizontalAlignment) -> Self {
        self.horizontal = horizontal;
        self
    }

    // Set vertical alignment.
    pub(crate) fn set_vertical(mut self, vertical: VerticalAlignment) -> Self {
        self.vertical = vertical;
        self
    }

    // Set text rotation.
    pub(crate) fn set_text_rotation(mut self, rotation: TextRotation) -> Self {
        self.text_rotation = rotation;
        self
    }

    // Set wrap text.
    pub(crate) fn set_wrap_text(mut self, wrap: bool) -> Self {
        self.wrap_text = wrap;
        self
    }

    // Set indent level.
    pub(crate) fn set_indent(mut self, indent: u8) -> Self {
        self.indent = Some(indent);
        self
    }

    // Set shrink to fit.
    pub(crate) fn set_shrink_to_fit(mut self, shrink: bool) -> Self {
        self.shrink_to_fit = shrink;
        self
    }
}

/// Fill pattern type.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FillPattern {
    /// No fill.
    #[default]
    None,

    /// Solid fill.
    Solid,

    /// Dark gray pattern.
    DarkGray,

    /// Medium gray pattern.
    MediumGray,

    /// Light gray pattern.
    LightGray,

    /// Gray 125 pattern.
    Gray125,

    /// Gray 0625 pattern.
    Gray0625,

    /// Dark horizontal pattern.
    DarkHorizontal,

    /// Dark vertical pattern.
    DarkVertical,

    /// Dark down pattern.
    DarkDown,

    /// Dark up pattern.
    DarkUp,

    /// Dark grid pattern.
    DarkGrid,

    /// Dark trellis pattern.
    DarkTrellis,

    /// Light horizontal pattern.
    LightHorizontal,

    /// Light vertical pattern.
    LightVertical,

    /// Light down pattern.
    LightDown,

    /// Light up pattern.
    LightUp,

    /// Light grid pattern.
    LightGrid,

    /// Light trellis pattern.
    LightTrellis,
}

/// Fill properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Fill {
    /// Fill pattern.
    pub pattern: FillPattern,

    /// Foreground color.
    pub foreground_color: Option<Color>,

    /// Background color.
    pub background_color: Option<Color>,
}

impl Fill {
    // Create new fill.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    // Create solid fill with color.
    #[cfg(test)]
    pub(crate) fn solid(color: Color) -> Self {
        Self {
            pattern: FillPattern::Solid,
            foreground_color: Some(color),
            background_color: None,
        }
    }

    // Set pattern.
    pub(crate) fn set_pattern(mut self, pattern: FillPattern) -> Self {
        self.pattern = pattern;
        self
    }

    // Set foreground color.
    pub(crate) fn set_foreground_color(mut self, color: Color) -> Self {
        self.foreground_color = Some(color);
        self
    }

    // Set background color.
    pub(crate) fn set_background_color(mut self, color: Color) -> Self {
        self.background_color = Some(color);
        self
    }

    /// Check if fill is visible.
    pub fn is_visible(&self) -> bool {
        self.pattern != FillPattern::None
    }

    /// Get the main fill color (foreground if available, otherwise background).
    pub fn get_color(&self) -> Option<Color> {
        self.foreground_color.or(self.background_color)
    }
}

/// Number format.
#[derive(Debug, Clone, PartialEq)]
pub struct NumberFormat {
    /// Format code.
    pub format_code: String,

    /// Format ID.
    pub format_id: Option<u32>,
}

impl NumberFormat {
    // Create new number format.
    pub(crate) fn new(format_code: String) -> Self {
        Self {
            format_code,
            format_id: None,
        }
    }

    // Create with format ID.
    pub(crate) fn set_format_id(mut self, format_id: u32) -> Self {
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

/// Cell protection properties.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Protection {
    /// Cell is locked.
    pub locked: bool,

    /// Cell is hidden.
    pub hidden: bool,
}

impl Protection {
    // Create new protection.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    // Set locked.
    pub(crate) fn set_locked(mut self, locked: bool) -> Self {
        self.locked = locked;
        self
    }

    // Set hidden.
    pub(crate) fn set_hidden(mut self, hidden: bool) -> Self {
        self.hidden = hidden;
        self
    }
}

/// Complete cell style.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    /// Font properties.
    pub font: Option<Font>,

    /// Fill properties.
    pub fill: Option<Fill>,

    /// Border properties.
    pub borders: Option<Borders>,

    /// Alignment properties.
    pub alignment: Option<Alignment>,

    /// Number format.
    pub number_format: Option<NumberFormat>,

    /// Protection properties.
    pub protection: Option<Protection>,
}

impl Style {
    // Create new style.
    pub(crate) fn new() -> Self {
        Self::default()
    }

    // Set font.
    pub(crate) fn set_font(mut self, font: Font) -> Self {
        self.font = Some(font);
        self
    }

    // Set fill.
    pub(crate) fn set_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    // Set borders.
    pub(crate) fn set_borders(mut self, borders: Borders) -> Self {
        self.borders = Some(borders);
        self
    }

    // Set alignment.
    pub(crate) fn set_alignment(mut self, alignment: Alignment) -> Self {
        self.alignment = Some(alignment);
        self
    }

    // Set number format.
    pub(crate) fn set_number_format(mut self, number_format: NumberFormat) -> Self {
        self.number_format = Some(number_format);
        self
    }

    // Set protection.
    pub(crate) fn set_protection(mut self, protection: Protection) -> Self {
        self.protection = Some(protection);
        self
    }

    /// Check if style is empty (no properties set).
    pub fn is_empty(&self) -> bool {
        self.font.is_none()
            && self.fill.is_none()
            && self.borders.is_none()
            && self.alignment.is_none()
            && self.number_format.is_none()
            && self.protection.is_none()
    }

    /// Check if style has any visible properties.
    pub fn has_visible_properties(&self) -> bool {
        (self
            .font
            .as_ref()
            .is_some_and(|f| f.color.is_some() || f.is_bold() || f.is_italic()))
            || (self.fill.as_ref().is_some_and(|f| f.is_visible()))
            || (self
                .borders
                .as_ref()
                .is_some_and(|b| b.has_visible_borders()))
            || (self.alignment.as_ref().is_some_and(|a| {
                a.horizontal != HorizontalAlignment::General
                    || a.vertical != VerticalAlignment::Bottom
                    || a.text_rotation != TextRotation::None
                    || a.wrap_text
                    || a.indent.is_some()
                    || a.shrink_to_fit
            }))
    }
}

/// RLE-compressed style storage for a worksheet range.
///
/// Instead of storing one [`Style`] per cell (which wastes memory when many
/// cells share the same style), this stores:
///
/// - A palette of the workbook's cell formats (`xf` records, indexed by style
///   id, with index 0 being the default format).
/// - Runs of consecutive cells (in row-major order over the bounding box of
///   the styled cells) that share the same style id.
///
/// This dramatically reduces memory usage for large worksheets.
///
/// Runs are stored as two parallel vectors: a style id per run and the
/// cumulative (exclusive) end offset of each run. The cumulative offsets
/// double as the prefix-sum index used for `O(log runs)` random access in
/// [`StyleRange::get`].
#[derive(Debug, Clone, Default)]
pub struct StyleRange {
    start: (u32, u32),
    end: (u32, u32),

    // Palette of styles indexed by style id. Id 0 is the default format,
    // which is also used for cells without an explicit style.
    palette: Vec<Style>,

    // Style id for each run, in row-major order.
    run_ids: Vec<u32>,

    // Cumulative exclusive end offset (in cells) of each run.
    run_ends: Vec<u64>,

    // Total cell count of the bounding box.
    total_cells: u64,
}

impl StyleRange {
    /// Create an empty `StyleRange`.
    pub fn empty() -> Self {
        Self::default()
    }

    // Create a `StyleRange` from style ids and a palette.
    //
    // - `cells`: `(row, col, style_id)` triples where `style_id` indexes into
    //   `palette`. Entries with style id 0 (the default format) or an id
    //   outside the palette are ignored.
    // - `palette`: the palette of unique styles (takes ownership).
    //
    // The runs are built directly from the sparse input, so memory usage is
    // proportional to the number of styled cells, not to the bounding box
    // area.
    pub(crate) fn from_style_ids(mut cells: Vec<(u32, u32, usize)>, palette: Vec<Style>) -> Self {
        // Drop entries that cannot be resolved against the palette. Style id 0
        // is the default format and is represented implicitly by gap runs.
        cells.retain(|&(_, _, id)| id > 0 && id < palette.len());

        if cells.is_empty() {
            return Self::empty();
        }

        // Find the bounding box of the styled cells.
        let mut row_start = u32::MAX;
        let mut row_end = 0;
        let mut col_start = u32::MAX;
        let mut col_end = 0;
        for &(r, c, _) in &cells {
            row_start = row_start.min(r);
            row_end = row_end.max(r);
            col_start = col_start.min(c);
            col_end = col_end.max(c);
        }

        let width = (col_end - col_start + 1) as u64;
        let height = (row_end - row_start + 1) as u64;
        let total_cells = width * height;

        // The input normally arrives in XML stream order, which is already
        // row-major, so this sort is close to O(n) in practice.
        cells.sort_unstable_by_key(|&(r, c, _)| (r, c));

        let mut run_ids: Vec<u32> = Vec::new();
        let mut run_ends: Vec<u64> = Vec::new();
        let mut covered: u64 = 0; // Number of cells covered by runs so far.

        // Append a run, coalescing with the previous run when ids match.
        let push_run = |run_ids: &mut Vec<u32>, run_ends: &mut Vec<u64>, id: u32, count: u64| {
            if count == 0 {
                return;
            }
            match (run_ids.last(), run_ends.last_mut()) {
                (Some(&last_id), Some(last_end)) if last_id == id => *last_end += count,
                _ => {
                    let prev_end = run_ends.last().copied().unwrap_or(0);
                    run_ids.push(id);
                    run_ends.push(prev_end + count);
                }
            }
        };

        for &(r, c, id) in &cells {
            let linear = (r - row_start) as u64 * width + (c - col_start) as u64;
            if linear < covered {
                // Duplicate position (malformed input); first entry wins.
                continue;
            }

            // Gap run of default-styled cells before this cell.
            push_run(&mut run_ids, &mut run_ends, 0, linear - covered);

            // The styled cell itself.
            push_run(&mut run_ids, &mut run_ends, id as u32, 1);
            covered = linear + 1;
        }
        // Trailing gap to the end of the bounding box.
        push_run(&mut run_ids, &mut run_ends, 0, total_cells - covered);

        run_ids.shrink_to_fit();
        run_ends.shrink_to_fit();

        StyleRange {
            start: (row_start, col_start),
            end: (row_end, col_end),
            palette,
            run_ids,
            run_ends,
            total_cells,
        }
    }

    /// Get the start position of the range.
    pub fn start(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.start)
        }
    }

    /// Get the end position of the range.
    pub fn end(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.end)
        }
    }

    /// Check if the range is empty.
    pub fn is_empty(&self) -> bool {
        self.run_ids.is_empty()
    }

    /// Get width of the range.
    pub fn width(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.1 - self.start.1 + 1) as usize
        }
    }

    /// Get height of the range.
    pub fn height(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.0 - self.start.0 + 1) as usize
        }
    }

    /// Get style at a position (relative to the range start).
    ///
    /// Returns `None` if the position is out of bounds. Cells without an
    /// explicit style resolve to the default format (palette entry 0).
    pub fn get(&self, pos: (usize, usize)) -> Option<&Style> {
        if pos.0 >= self.height() || pos.1 >= self.width() {
            return None;
        }

        let linear = pos.0 as u64 * self.width() as u64 + pos.1 as u64;
        // Binary search over the cumulative run end offsets.
        let run = self.run_ends.partition_point(|&end| end <= linear);
        let style_id = *self.run_ids.get(run)?;
        self.palette.get(style_id as usize)
    }

    /// Iterate over all cells in the bounding box with their relative
    /// positions and styles.
    pub fn cells(&self) -> StyleRangeCells<'_> {
        StyleRangeCells {
            range: self,
            run_idx: 0,
            linear_idx: 0,
        }
    }

    // Get the number of RLE runs. Used in tests to verify run coalescing.
    #[cfg(test)]
    fn run_count(&self) -> usize {
        self.run_ids.len()
    }
}

/// Iterator over cells in a [`StyleRange`].
pub struct StyleRangeCells<'a> {
    range: &'a StyleRange,
    run_idx: usize,
    linear_idx: u64,
}

impl<'a> Iterator for StyleRangeCells<'a> {
    type Item = (usize, usize, &'a Style);

    fn next(&mut self) -> Option<Self::Item> {
        if self.linear_idx >= self.range.total_cells {
            return None;
        }

        let width = self.range.width() as u64;
        if width == 0 {
            return None;
        }

        // Advance past runs that have been fully consumed.
        while self.linear_idx >= *self.range.run_ends.get(self.run_idx)? {
            self.run_idx += 1;
        }

        let row = (self.linear_idx / width) as usize;
        let col = (self.linear_idx % width) as usize;
        let style_id = *self.range.run_ids.get(self.run_idx)?;
        let style = self.range.palette.get(style_id as usize)?;

        self.linear_idx += 1;

        Some((row, col, style))
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
    fn test_color_argb_roundtrip() {
        let color = Color::from_argb(0x80FF8040);
        assert_eq!(color.alpha, 0x80);
        assert_eq!(color.red, 0xFF);
        assert_eq!(color.green, 0x80);
        assert_eq!(color.blue, 0x40);
        assert_eq!(color.to_argb(), 0x80FF8040);
    }

    #[test]
    fn test_font() {
        let font = Font::new()
            .set_name("Arial".to_string())
            .set_size(12.0)
            .set_weight(FontWeight::Bold)
            .set_color(Color::rgb(255, 0, 0));

        assert_eq!(font.name, Some("Arial".to_string()));
        assert_eq!(font.size, Some(12.0));
        assert!(font.is_bold());
        assert_eq!(font.color, Some(Color::rgb(255, 0, 0)));
    }

    #[test]
    fn test_style() {
        let style = Style::new()
            .set_font(Font::new().set_name("Arial".to_string()))
            .set_fill(Fill::solid(Color::rgb(255, 255, 0)));

        assert!(!style.is_empty());
        assert!(style.font.is_some());
        assert!(style.fill.is_some());
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

    // A small palette where id 1 is bold and id 2 is a yellow fill.
    fn test_palette() -> Vec<Style> {
        vec![
            Style::new(),
            Style::new().set_font(Font::new().set_weight(FontWeight::Bold)),
            Style::new().set_fill(Fill::solid(Color::rgb(255, 255, 0))),
        ]
    }

    #[test]
    fn test_style_range_empty() {
        let range = StyleRange::empty();
        assert!(range.is_empty());
        assert_eq!(range.start(), None);
        assert_eq!(range.end(), None);
        assert_eq!(range.width(), 0);
        assert_eq!(range.height(), 0);
        assert_eq!(range.get((0, 0)), None);
        assert_eq!(range.cells().count(), 0);
    }

    #[test]
    fn test_style_range_from_style_ids() {
        // Styled cells at (1,1)=1, (1,3)=1, (3,1)=2. Bounding box is
        // rows 1-3, cols 1-3 => 3x3 grid.
        let cells = vec![(1, 1, 1), (1, 3, 1), (3, 1, 2)];
        let range = StyleRange::from_style_ids(cells, test_palette());

        assert!(!range.is_empty());
        assert_eq!(range.start(), Some((1, 1)));
        assert_eq!(range.end(), Some((3, 3)));
        assert_eq!(range.width(), 3);
        assert_eq!(range.height(), 3);

        // Styled positions (relative to start).
        assert!(range.get((0, 0)).unwrap().font.as_ref().unwrap().is_bold());
        assert!(range.get((0, 2)).unwrap().font.as_ref().unwrap().is_bold());
        assert!(range.get((2, 0)).unwrap().fill.is_some());

        // Gap cells resolve to the default (empty) style.
        assert!(range.get((0, 1)).unwrap().is_empty());
        assert!(range.get((1, 1)).unwrap().is_empty());

        // Out of bounds.
        assert_eq!(range.get((3, 0)), None);
        assert_eq!(range.get((0, 3)), None);
    }

    #[test]
    fn test_style_range_run_coalescing() {
        // A full row of the same style id must coalesce into a single run.
        let cells: Vec<_> = (0..10u32).map(|c| (0, c, 1)).collect();
        let range = StyleRange::from_style_ids(cells, test_palette());

        assert_eq!(range.run_count(), 1);
        assert_eq!(range.width(), 10);
        assert_eq!(range.height(), 1);
        for c in 0..10 {
            assert!(range.get((0, c)).unwrap().font.as_ref().unwrap().is_bold());
        }
    }

    #[test]
    fn test_style_range_ignores_unresolvable_ids() {
        // Id 0 (default) and out-of-palette ids are dropped.
        let cells = vec![(0, 0, 0), (0, 1, 99), (5, 5, 1)];
        let range = StyleRange::from_style_ids(cells, test_palette());

        // Only (5,5) survives, so the bounding box is a single cell.
        assert_eq!(range.start(), Some((5, 5)));
        assert_eq!(range.end(), Some((5, 5)));
        assert_eq!(range.run_count(), 1);
    }

    #[test]
    fn test_style_range_all_unresolvable_is_empty() {
        let cells = vec![(0, 0, 0), (0, 1, 99)];
        let range = StyleRange::from_style_ids(cells, test_palette());
        assert!(range.is_empty());
    }

    #[test]
    fn test_style_range_iterator_covers_bounding_box() {
        let cells = vec![(2, 2, 1), (4, 4, 2)];
        let range = StyleRange::from_style_ids(cells, test_palette());

        // 3x3 bounding box: iterator visits every cell once, row-major.
        let visited: Vec<(usize, usize, bool)> = range
            .cells()
            .map(|(r, c, s)| (r, c, !s.is_empty()))
            .collect();
        assert_eq!(visited.len(), 9);
        assert_eq!(visited[0], (0, 0, true)); // (2,2) styled
        assert_eq!(visited[4], (1, 1, false)); // gap
        assert_eq!(visited[8], (2, 2, true)); // (4,4) styled

        let styled = visited.iter().filter(|&&(_, _, s)| s).count();
        assert_eq!(styled, 2);
    }

    #[test]
    fn test_style_range_sparse_no_dense_allocation() {
        // A pathological sparse range: two styled cells at opposite corners
        // of a huge bounding box. Constructing the range must not allocate
        // per-cell storage for the bounding box (~16 billion cells here).
        let cells = vec![(0, 0, 1), (999_999, 16_383, 2)];
        let range = StyleRange::from_style_ids(cells, test_palette());

        assert_eq!(range.width(), 16_384);
        assert_eq!(range.height(), 1_000_000);
        // 2 styled cells + 1 gap run between them.
        assert_eq!(range.run_count(), 3);
        assert!(range.get((0, 0)).unwrap().font.as_ref().unwrap().is_bold());
        assert!(range.get((999_999, 16_383)).unwrap().fill.is_some());
        assert!(range.get((500_000, 8_000)).unwrap().is_empty());
    }

    #[test]
    fn test_style_range_unsorted_input() {
        // Input out of stream order must still produce a correct range.
        let cells = vec![(3, 1, 2), (1, 3, 1), (1, 1, 1)];
        let range = StyleRange::from_style_ids(cells, test_palette());

        assert_eq!(range.start(), Some((1, 1)));
        assert_eq!(range.end(), Some((3, 3)));
        assert!(range.get((0, 0)).unwrap().font.as_ref().unwrap().is_bold());
        assert!(range.get((2, 0)).unwrap().fill.is_some());
    }
}
