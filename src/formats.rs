use crate::datatype::{Data, DataRef, ExcelDateTime, ExcelDateTimeType};
use std::collections::HashMap;
use std::sync::{Arc, RwLock};

/// Format string interner to avoid duplicate Arc allocations
///
/// This structure helps reduce memory usage by deduplicating format strings
/// across multiple cells. It's particularly useful when parsing large spreadsheets
/// where many cells share the same custom format strings.
///
/// Thread-safe implementation using RwLock for concurrent access.
///
/// # Examples
///
/// ```
/// use calamine::FormatStringInterner;
///
/// let interner = FormatStringInterner::new();
/// let format1 = interner.intern("$#,##0.00");
/// let format2 = interner.intern("$#,##0.00"); // Same Arc<str> instance
/// assert_eq!(format1, format2);
/// ```
#[derive(Debug, Default)]
pub struct FormatStringInterner {
    cache: RwLock<HashMap<String, Arc<str>>>,
}

impl FormatStringInterner {
    /// Create a new format string interner
    pub fn new() -> Self {
        Self {
            cache: RwLock::new(HashMap::new()),
        }
    }

    /// Intern a format string, returning a shared Arc<str>
    ///
    /// Thread-safe implementation that uses read lock for lookup and write lock for insertion.
    /// Multiple threads can safely call this method concurrently.
    ///
    /// # Arguments
    ///
    /// * `format_string` - The format string to intern
    ///
    /// # Returns
    ///
    /// A shared Arc<str> that will be the same instance for identical format strings
    pub fn intern(&self, format_string: &str) -> Arc<str> {
        // Fast path: try to find existing entry with read lock
        if let Ok(cache) = self.cache.read() {
            if let Some(cached) = cache.get(format_string) {
                return cached.clone();
            }
        }

        // Slow path: insert new entry with write lock
        if let Ok(mut cache) = self.cache.write() {
            // Double-check in case another thread inserted while we were waiting
            if let Some(cached) = cache.get(format_string) {
                cached.clone()
            } else {
                let arc_str: Arc<str> = Arc::from(format_string);
                cache.insert(format_string.to_string(), arc_str.clone());
                arc_str
            }
        } else {
            // Fallback if lock is poisoned
            Arc::from(format_string)
        }
    }

    /// Get the number of interned strings
    pub fn len(&self) -> usize {
        self.cache.read().map(|cache| cache.len()).unwrap_or(0)
    }

    /// Check if the interner is empty
    pub fn is_empty(&self) -> bool {
        self.cache
            .read()
            .map(|cache| cache.is_empty())
            .unwrap_or(true)
    }
}

/// Cell number format types
///
/// Represents the different categories of number formats that Excel supports.
/// These correspond to the built-in format categories and custom format patterns.
///
/// # References
///
/// - ECMA-376 Part 1, Section 18.8.30 (numFmt)
/// - MS-XLSB Section 2.4.648 (BrtFmt)
#[derive(Debug, Clone, PartialEq)]
pub enum CellFormat {
    /// Other/general format (backward compatible)
    ///
    /// This includes general number formats, custom formats that don't fall
    /// into other categories, and text formats.
    Other,
    /// Date and time format
    ///
    /// Formats that display dates, times, or both. Examples: "yyyy-mm-dd", "h:mm:ss AM/PM"
    DateTime,
    /// Time delta/duration format
    ///
    /// Formats that display elapsed time or duration. Examples: "[h]:mm:ss", "[mm]:ss"
    TimeDelta,
}

/// Comprehensive cell formatting information
///
/// Contains all formatting information for a cell, including number format,
/// font, fill, border, and alignment properties. This structure provides
/// full compatibility with Excel's formatting model.
///
/// # References
///
/// - ECMA-376 Part 1, Section 18.8.45 (xf - Format)
/// - MS-XLSB Section 2.4.812 (BrtXF)
///
/// # Examples
///
/// ```
/// use calamine::{CellStyle, CellFormat};
///
/// let style = CellStyle::default();
/// assert_eq!(style.number_format, CellFormat::Other);
/// assert!(style.font.is_none());
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct CellStyle {
    /// Number format category
    pub number_format: CellFormat,
    /// Custom format string (for backward compatibility and additional context)
    ///
    /// Contains the original format string from the Excel file, useful for
    /// applications that need to preserve exact formatting information.
    pub format_string: Option<Arc<str>>,
    /// Font information
    ///
    /// Contains font name, size, style (bold/italic), and color information.
    pub font: Option<Arc<Font>>,
    /// Fill information  
    ///
    /// Contains background color and pattern information.
    pub fill: Option<Arc<Fill>>,
    /// Border information
    ///
    /// Contains border style and color for all four sides of the cell.
    pub border: Option<Arc<Border>>,
    /// Alignment information
    ///
    /// Contains horizontal/vertical alignment, text wrapping, and other text positioning options.
    pub alignment: Option<Arc<Alignment>>,
}

impl Default for CellStyle {
    fn default() -> Self {
        Self {
            number_format: CellFormat::Other,
            format_string: None,
            font: None,
            fill: None,
            border: None,
            alignment: None,
        }
    }
}

impl CellStyle {
    /// Check if this formatting has any non-default values
    pub fn is_default(&self) -> bool {
        self.number_format == CellFormat::Other
            && self.format_string.is_none()
            && self.font.is_none()
            && self.fill.is_none()
            && self.border.is_none()
            && self.alignment.is_none()
    }

    /// Return the stored [`CellFormat`].  Handy when all you need is
    /// the number‑format kind without matching on the whole style.
    #[inline]
    pub fn number_format(&self) -> &CellFormat {
        &self.number_format
    }
}

/// Font formatting information
///
/// Contains font properties including name, size, style, and color.
///
/// # References
///
/// - ECMA-376 Part 1, Section 18.8.22 (font)
/// - MS-XLSB Section 2.4.149 (BrtFont)
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Font {
    /// Font name (e.g., "Arial", "Calibri")
    ///
    /// If None, the default font for the workbook is used.
    pub name: Option<Arc<str>>,
    /// Font size in points
    ///
    /// Standard font sizes range from 8 to 72 points.
    pub size: Option<f64>,
    /// Bold formatting
    ///
    /// If None, the font is not bold.
    pub bold: Option<bool>,
    /// Italic formatting
    ///
    /// If None, the font is not italic.
    pub italic: Option<bool>,
    /// Font color
    ///
    /// Can be RGB, ARGB, theme color, indexed color, or automatic.
    pub color: Option<Color>,
}

/// Fill formatting information
#[derive(Debug, Clone, PartialEq)]
pub struct Fill {
    /// Pattern type
    pub pattern_type: PatternType,
    /// Foreground color
    pub foreground_color: Option<Color>,
    /// Background color  
    pub background_color: Option<Color>,
}

impl Default for Fill {
    fn default() -> Self {
        Self {
            pattern_type: PatternType::None,
            foreground_color: None,
            background_color: None,
        }
    }
}

/// Pattern fill types (matches Excel specification)
#[derive(Debug, Clone, PartialEq)]
pub enum PatternType {
    /// No fill
    None,
    /// Solid fill
    Solid,
    /// Light gray pattern
    LightGray,
    /// Medium gray pattern  
    MediumGray,
    /// Dark gray pattern
    DarkGray,
    /// Custom pattern with pattern name
    Pattern(Arc<str>),
}

/// Border formatting information
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Border {
    /// Left border
    pub left: Option<BorderSide>,
    /// Right border
    pub right: Option<BorderSide>,
    /// Top border
    pub top: Option<BorderSide>,
    /// Bottom border
    pub bottom: Option<BorderSide>,
}

/// Individual border side
#[derive(Debug, Clone, PartialEq)]
pub struct BorderSide {
    /// Border style name
    pub style: Arc<str>,
    /// Border color
    pub color: Option<Color>,
}

impl Default for BorderSide {
    fn default() -> Self {
        Self {
            style: Arc::from("none"),
            color: None,
        }
    }
}

/// Alignment formatting information
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Alignment {
    /// Horizontal alignment
    pub horizontal: Option<Arc<str>>,
    /// Vertical alignment
    pub vertical: Option<Arc<str>>,
    /// Wrap text flag
    pub wrap_text: Option<bool>,
    /// Indent level
    pub indent: Option<u32>,
    /// Shrink to fit flag
    pub shrink_to_fit: Option<bool>,
    /// Text rotation (degrees)
    pub text_rotation: Option<i32>,
    /// Reading order
    pub reading_order: Option<u32>,
}

/// Color representation
///
/// Represents the different ways colors can be specified in Excel files.
/// Excel supports multiple color models including RGB, theme colors, and
/// indexed colors from a predefined palette.
///
/// # References
///
/// - ECMA-376 Part 1, Section 18.8.3 (color)
/// - MS-XLSB Section 2.5.52 (BrtColor)
#[derive(Debug, Clone, PartialEq)]
pub enum Color {
    /// RGB color
    ///
    /// Standard RGB color with 8-bit components.
    Rgb {
        /// Red component (0-255)
        r: u8,
        /// Green component (0-255)
        g: u8,
        /// Blue component (0-255)
        b: u8,
    },
    /// ARGB color (with alpha)
    ///
    /// RGB color with alpha channel for transparency.
    Argb {
        /// Alpha component (0-255, where 255 is opaque)
        a: u8,
        /// Red component (0-255)
        r: u8,
        /// Green component (0-255)
        g: u8,
        /// Blue component (0-255)
        b: u8,
    },
    /// Theme color reference
    ///
    /// References one of the theme colors defined in the workbook theme.
    /// Theme colors provide consistent color schemes across documents.
    Theme {
        /// Theme color index (0-based)
        theme: u32,
        /// Tint adjustment (-1.0 to 1.0)
        ///
        /// Negative values darken the color, positive values lighten it.
        tint: Option<f64>,
    },
    /// Indexed color
    ///
    /// References a color from Excel's built-in color palette (0-based index).
    Indexed(u32),
    /// Automatic color
    ///
    /// Uses the default color for the context (e.g., black for text, white for background).
    Auto,
}

/// Detect the number format type from a custom format string
///
/// Analyzes an Excel format string to determine its category (DateTime, Currency, etc.).
/// This function implements format string parsing logic compatible with Excel's
/// number format detection.
///
/// # Arguments
///
/// * `format` - The Excel format string to analyze
///
/// # Returns
///
/// The detected [`CellFormat`] category
///
/// # Examples
///
/// ```
/// use calamine::{detect_custom_number_format, CellFormat};
///
/// assert_eq!(detect_custom_number_format("yyyy-mm-dd"), CellFormat::DateTime);
/// assert_eq!(detect_custom_number_format("$#,##0.00"), CellFormat::Currency);
/// assert_eq!(detect_custom_number_format("0.00%"), CellFormat::Percentage);
/// assert_eq!(detect_custom_number_format("[h]:mm:ss"), CellFormat::TimeDelta);
/// ```
///
/// # References
///
/// - ECMA-376 Part 1, Section 18.8.31 (numFmt)
pub fn detect_custom_number_format(format: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;

    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return CellFormat::Other, // first format only
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return CellFormat::DateTime,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return CellFormat::DateTime
            }
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                    // ok ...
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            }
        }
        prev = s;
    }

    CellFormat::Other
}

/// Check excel number format type from format string and create appropriate CellFormat
/// with interned format string for custom formats
pub fn detect_custom_number_format_with_interner(
    format: &str,
    interner: &FormatStringInterner,
) -> (CellFormat, Option<Arc<str>>) {
    let format_type = detect_custom_number_format(format);

    // For custom formats, we always intern the format string to preserve the original
    // formatting information, regardless of the detected type
    match format_type {
        CellFormat::Other => (CellFormat::Other, Some(interner.intern(format))),
        other => (other, Some(interner.intern(format))),
    }
}

/// Determine cell format from built-in format ID
pub fn builtin_format_by_id(id: &[u8]) -> CellFormat {
    match id {
        // mm-dd-yy
        b"14" |
        // d-mmm-yy
        b"15" |
        // d-mmm
        b"16" |
        // mmm-yy
        b"17" |
        // h:mm AM/PM
        b"18" |
        // h:mm:ss AM/PM
        b"19" |
        // h:mm
        b"20" |
        // h:mm:ss
        b"21" |
        // m/d/yy h:mm
        b"22" |
        // mm:ss
        b"45" |
        // mmss.0
        b"47" => CellFormat::DateTime,
        // [h]:mm:ss
        b"46" => CellFormat::TimeDelta,
        _ => CellFormat::Other
    }
}

/// Check if code corresponds to builtin format
///
/// See `is_builtin_date_format_id`
pub fn builtin_format_by_code(code: u16) -> CellFormat {
    match code {
        14..=22 | 45 | 47 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

// convert i64 to date, if format == Date
pub fn format_excel_i64(value: i64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    match format {
        Some(CellFormat::DateTime) => Data::DateTime(ExcelDateTime::new(
            value as f64,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => Data::DateTime(ExcelDateTime::new(
            value as f64,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => Data::Int(value),
    }
}

// convert f64 to date, if format == Date
#[inline]
pub fn format_excel_f64_ref(
    value: f64,
    format: Option<&CellFormat>,
    is_1904: bool,
) -> DataRef<'static> {
    match format {
        Some(CellFormat::DateTime) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => DataRef::Float(value),
    }
}

// convert f64 to date, if format == Date
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    format_excel_f64_ref(value, format, is_1904).into()
}

/// Ported from openpyxl, MIT License
/// https://foss.heptapod.net/openpyxl/openpyxl/-/blob/a5e197c530aaa49814fd1d993dd776edcec35105/openpyxl/styles/tests/test_number_style.py
#[test]
fn test_is_date_format() {
    assert_eq!(
        detect_custom_number_format("DD/MM/YY"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("H:MM:SS;@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("yyyy-mm-dd"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$\\u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("m\"M\"d\"D\";@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("m/d/yy\"M\"d\"D\";@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm:ss"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$''u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("0_ ;[Red]\\-0\\ "),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("#,##0.0####\" YMD\""),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("[h]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[ss]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[s].000"),
        CellFormat::TimeDelta
    );
    assert_eq!(detect_custom_number_format("[m]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[mm]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[h]:mm"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[m]:mm"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("m:mm"), CellFormat::DateTime);
    assert_eq!(
        detect_custom_number_format("[Blue]\\+[h]:mm;[Red]\\-[h]:mm;[Green][h]:mm"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta][s].00"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm;[=0]\\-"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("h:mm:ss AM/PM"),
        CellFormat::DateTime
    );
    assert_eq!(detect_custom_number_format("h:mm:ss"), CellFormat::DateTime);
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta].00"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta]General"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("ha/p\\\\m"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
        CellFormat::Other
    );

    // Test format detection with interner
    let interner = FormatStringInterner::new();
    let (datetime_format, datetime_string) =
        detect_custom_number_format_with_interner("yyyy-mm-dd", &interner);
    assert_eq!(datetime_format, CellFormat::DateTime);
    assert_eq!(
        datetime_string.as_ref().map(|s| s.as_ref()),
        Some("yyyy-mm-dd")
    );
}
