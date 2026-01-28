// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use std::collections::BTreeMap;
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
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum BorderStyle {
    /// No border
    #[default]
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
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FontWeight {
    /// Normal weight
    #[default]
    Normal,
    /// Bold weight
    Bold,
}

/// Font style
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FontStyle {
    /// Normal style
    #[default]
    Normal,
    /// Italic style
    Italic,
}

/// Underline style
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum UnderlineStyle {
    /// No underline
    #[default]
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
#[derive(Debug, Clone, Copy, PartialEq, Default)]
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
    #[default]
    General,
}

/// Vertical alignment
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum VerticalAlignment {
    /// Top alignment
    Top,
    /// Center alignment
    Center,
    /// Bottom alignment
    #[default]
    Bottom,
    /// Justify alignment
    Justify,
    /// Distributed alignment
    Distributed,
}

/// Text rotation in degrees
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum TextRotation {
    /// No rotation
    #[default]
    None,
    /// Rotated by degrees (0-180)
    Degrees(u16),
    /// Stacked text
    Stacked,
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
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum FillPattern {
    /// No fill
    #[default]
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
    /// Column widths (keyed by column index)
    pub column_widths: BTreeMap<u32, ColumnWidth>,
    /// Row heights (keyed by row index)
    pub row_heights: BTreeMap<u32, RowHeight>,
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
        self.column_widths.insert(column_width.column, column_width);
        self
    }

    /// Add a row height
    pub fn add_row_height(mut self, row_height: RowHeight) -> Self {
        self.row_heights.insert(row_height.row, row_height);
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

    /// Get column width for a specific column (O(log n) lookup)
    pub fn get_column_width(&self, column: u32) -> Option<&ColumnWidth> {
        self.column_widths.get(&column)
    }

    /// Get row height for a specific row (O(log n) lookup)
    pub fn get_row_height(&self, row: u32) -> Option<&RowHeight> {
        self.row_heights.get(&row)
    }

    /// Get effective column width (custom or default).
    ///
    /// Returns the column width in Excel's character-based units. If no custom
    /// width is set, returns the worksheet's default column width, or 8.43 if
    /// no default is specified.
    ///
    /// **Note:** Excel column widths are stored in character units relative to
    /// the workbook's default font, not pixels. Converting to pixels requires
    /// font metrics and is font-dependent. The value 8.43 is Excel's standard
    /// default for Calibri 11pt.
    pub fn get_effective_column_width(&self, column: u32) -> f64 {
        self.get_column_width(column)
            .map(|cw| cw.width)
            .or(self.default_column_width)
            .unwrap_or(8.43)
    }

    /// Get effective row height (custom or default).
    ///
    /// Returns the row height in points. If no custom height is set, returns
    /// the worksheet's default row height, or 15.0 if no default is specified.
    ///
    /// **Note:** Row heights in Excel are stored in points (1/72 inch), but
    /// the actual displayed height may vary slightly depending on the default
    /// font. The value 15.0 is Excel's standard default for Calibri 11pt.
    pub fn get_effective_row_height(&self, row: u32) -> f64 {
        self.get_row_height(row)
            .map(|rh| rh.height)
            .or(self.default_row_height)
            .unwrap_or(15.0)
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

/// A run of consecutive cells with the same style (row-major order)
#[derive(Debug, Clone)]
struct StyleRun {
    /// Index into the palette (0 = no style/default)
    style_id: u16,
    /// Number of consecutive cells with this style
    count: u32,
}

/// RLE-compressed style storage for a worksheet range.
///
/// Instead of storing one Style per cell (which wastes memory when many cells
/// share the same style), this stores:
/// - A palette of unique styles
/// - Runs of consecutive cells (row-major) that share the same style
///
/// This dramatically reduces memory usage and clone overhead for large worksheets.
#[derive(Debug, Clone, Default)]
pub struct StyleRange {
    start: (u32, u32),
    end: (u32, u32),
    /// Palette of unique styles. Index 0 is reserved for "no style" (empty).
    palette: Vec<Style>,
    /// RLE-encoded runs in row-major order
    runs: Vec<StyleRun>,
    /// Total cell count (for validation)
    total_cells: u64,
}

impl StyleRange {
    /// Create an empty StyleRange
    pub fn empty() -> Self {
        Self::default()
    }

    /// Create a StyleRange from style IDs and a palette (zero-copy).
    ///
    /// This is more efficient than `from_sparse` as it avoids cloning styles.
    ///
    /// - `cells`: Vec of (row, col, style_id) where style_id indexes into palette
    /// - `palette`: The shared palette of unique styles (taken ownership)
    pub fn from_style_ids(cells: Vec<(u32, u32, usize)>, palette: Vec<Style>) -> Self {
        if cells.is_empty() {
            return Self::empty();
        }

        // Find bounds
        let mut row_start = u32::MAX;
        let mut row_end = 0;
        let mut col_start = u32::MAX;
        let mut col_end = 0;
        for (r, c, _) in &cells {
            row_start = row_start.min(*r);
            row_end = row_end.max(*r);
            col_start = col_start.min(*c);
            col_end = col_end.max(*c);
        }

        let width = (col_end - col_start + 1) as usize;
        let height = (row_end - row_start + 1) as usize;
        let total_cells = (width * height) as u64;

        // Create dense style ID array (temporary)
        let mut style_ids = vec![0u16; width * height];

        for (r, c, style_id) in cells {
            let row = (r - row_start) as usize;
            let col = (c - col_start) as usize;
            let idx = row * width + col;
            // style_id is already an index, just need to fit in u16
            style_ids[idx] = style_id.min(u16::MAX as usize) as u16;
        }

        // Compress into RLE runs
        let mut runs = Vec::new();
        if !style_ids.is_empty() {
            let mut current_style = style_ids[0];
            let mut count = 1u32;

            for &style_id in &style_ids[1..] {
                if style_id == current_style {
                    count += 1;
                } else {
                    runs.push(StyleRun {
                        style_id: current_style,
                        count,
                    });
                    current_style = style_id;
                    count = 1;
                }
            }
            runs.push(StyleRun {
                style_id: current_style,
                count,
            });
        }

        runs.shrink_to_fit();

        StyleRange {
            start: (row_start, col_start),
            end: (row_end, col_end),
            palette,
            runs,
            total_cells,
        }
    }

    /// Create a StyleRange from sparse cell data
    ///
    /// Takes cells with positions and styles, compresses into RLE format.
    pub fn from_sparse(cells: Vec<(u32, u32, Style)>) -> Self {
        if cells.is_empty() {
            return Self::empty();
        }

        // Find bounds
        let mut row_start = u32::MAX;
        let mut row_end = 0;
        let mut col_start = u32::MAX;
        let mut col_end = 0;
        for (r, c, _) in &cells {
            row_start = row_start.min(*r);
            row_end = row_end.max(*r);
            col_start = col_start.min(*c);
            col_end = col_end.max(*c);
        }

        let width = (col_end - col_start + 1) as usize;
        let height = (row_end - row_start + 1) as usize;
        let total_cells = (width * height) as u64;

        // Build palette and map styles to IDs
        // Use style_id from Excel if available, otherwise assign sequential IDs
        let mut palette: Vec<Style> = vec![Style::default()]; // Index 0 = empty/default
        let mut style_to_id: std::collections::HashMap<u32, u16> = std::collections::HashMap::new();

        // Create dense style ID array (temporary)
        let mut style_ids = vec![0u16; width * height];

        for (r, c, style) in cells {
            let row = (r - row_start) as usize;
            let col = (c - col_start) as usize;
            let idx = row * width + col;

            if style.is_empty() {
                continue; // Leave as 0
            }

            // Use Excel's style_id if available for deduplication
            // This groups cells with the same formatting together
            let excel_style_id = style.style_id.unwrap_or_else(|| {
                // Fallback: use palette length as unique ID (no dedup for these)
                palette.len() as u32
            });

            let style_id = if let Some(&id) = style_to_id.get(&excel_style_id) {
                id
            } else {
                let id = palette.len() as u16;
                palette.push(style);
                style_to_id.insert(excel_style_id, id);
                id
            };

            style_ids[idx] = style_id;
        }

        // Compress into RLE runs
        let mut runs = Vec::new();
        if !style_ids.is_empty() {
            let mut current_style = style_ids[0];
            let mut count = 1u32;

            for &style_id in &style_ids[1..] {
                if style_id == current_style {
                    count += 1;
                } else {
                    runs.push(StyleRun {
                        style_id: current_style,
                        count,
                    });
                    current_style = style_id;
                    count = 1;
                }
            }
            // Push final run
            runs.push(StyleRun {
                style_id: current_style,
                count,
            });
        }

        runs.shrink_to_fit();
        palette.shrink_to_fit();

        StyleRange {
            start: (row_start, col_start),
            end: (row_end, col_end),
            palette,
            runs,
            total_cells,
        }
    }

    /// Get the start position of the range
    pub fn start(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.start)
        }
    }

    /// Get the end position of the range
    pub fn end(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.end)
        }
    }

    /// Check if the range is empty
    pub fn is_empty(&self) -> bool {
        self.runs.is_empty()
    }

    /// Get width of the range
    pub fn width(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.1 - self.start.1 + 1) as usize
        }
    }

    /// Get height of the range
    pub fn height(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.0 - self.start.0 + 1) as usize
        }
    }

    /// Get style at a position (relative to range start)
    ///
    /// Returns None if position is out of bounds, or reference to style.
    pub fn get(&self, pos: (usize, usize)) -> Option<&Style> {
        let width = self.width();
        let height = self.height();

        if pos.0 >= height || pos.1 >= width {
            return None;
        }

        let linear_idx = pos.0 * width + pos.1;
        let style_id = self.style_id_at(linear_idx)?;
        self.palette.get(style_id as usize)
    }

    /// Get style ID at a linear index using binary search on runs
    fn style_id_at(&self, linear_idx: usize) -> Option<u16> {
        let mut offset = 0usize;
        for run in &self.runs {
            let run_end = offset + run.count as usize;
            if linear_idx < run_end {
                return Some(run.style_id);
            }
            offset = run_end;
        }
        None
    }

    /// Iterate over all cells with their positions and styles
    pub fn cells(&self) -> StyleRangeCells<'_> {
        StyleRangeCells {
            range: self,
            run_idx: 0,
            run_offset: 0,
            linear_idx: 0,
        }
    }

    /// Get number of unique styles (excluding empty)
    pub fn unique_style_count(&self) -> usize {
        self.palette.len().saturating_sub(1)
    }

    /// Get number of RLE runs (for diagnostics)
    pub fn run_count(&self) -> usize {
        self.runs.len()
    }

    /// Get compression ratio (cells / runs)
    pub fn compression_ratio(&self) -> f64 {
        if self.runs.is_empty() {
            0.0
        } else {
            self.total_cells as f64 / self.runs.len() as f64
        }
    }
}

/// Iterator over cells in a StyleRange
pub struct StyleRangeCells<'a> {
    range: &'a StyleRange,
    run_idx: usize,
    run_offset: u32,
    linear_idx: u64,
}

impl<'a> Iterator for StyleRangeCells<'a> {
    type Item = (usize, usize, &'a Style);

    fn next(&mut self) -> Option<Self::Item> {
        if self.run_idx >= self.range.runs.len() {
            return None;
        }

        let width = self.range.width();
        if width == 0 {
            return None;
        }

        let row = (self.linear_idx / width as u64) as usize;
        let col = (self.linear_idx % width as u64) as usize;

        let run = &self.range.runs[self.run_idx];
        let style = self.range.palette.get(run.style_id as usize)?;

        self.linear_idx += 1;
        self.run_offset += 1;

        if self.run_offset >= run.count {
            self.run_idx += 1;
            self.run_offset = 0;
        }

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
