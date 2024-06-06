use std::borrow::Cow;

/// Rich (formatted) text cell.
#[derive(Debug, PartialEq, Clone, Default)]
pub struct RichText {
    /// Total text, where the different formats are saved in the `formatted` field.
    text: String,
    /// List of format annotations, with each having:
    /// - Length of string section
    /// - Format of the string section
    formatted: Vec<(usize, FontFormat)>,
}

impl RichText {
    /// Create new empty rich text.
    pub fn new() -> Self {
        Self::default()
    }

    /// Create an instance with only default plain text.
    pub fn plain(s: String) -> Self {
        let len = s.len();
        Self {
            text: s,
            formatted: vec![(len, FontFormat::default())],
        }
    }

    /// Is this empty?
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    /// Is this just plain text?
    #[inline]
    pub fn is_plain(&self) -> bool {
        self.formatted.iter().all(|(_len, f)| f.is_default())
    }

    /// Get the full text value without formatting.
    #[inline]
    pub fn text(&self) -> &String {
        &self.text
    }

    /// Deconstruct this into a plain string with the text.
    #[inline]
    pub fn into_text(self) -> String {
        self.text
    }

    /// Add a formatted text element to the rich text.
    #[inline]
    pub fn add_element(&mut self, element: RichTextPart<'_>) -> &mut Self {
        if !element.text.is_empty() {
            self.text.push_str(element.text);
            self.formatted
                .push((element.text.len(), element.format.into_owned()));
        }
        self
    }

    /// Iterate over the differently formatted text elements.
    pub fn elements(&self) -> impl Iterator<Item = RichTextPart<'_>> {
        let mut current = 0;
        self.formatted.iter().map(move |(len, f)| {
            let s = &self.text[current..(current + *len)];
            current += *len;
            RichTextPart {
                text: s,
                format: Cow::Borrowed(f),
            }
        })
    }
}

/// Part of a rich text formatted cell.
#[derive(Debug, PartialEq, Clone)]
pub struct RichTextPart<'a> {
    /// Text value.
    pub text: &'a str,
    /// Text format.
    pub format: Cow<'a, FontFormat>,
}

impl<'a> RichTextPart<'a> {
    /// Is this part plain text?
    #[inline]
    pub fn is_plain(&self) -> bool {
        self.format.is_default()
    }
}

/// Format of a font / text format.
#[derive(Debug, PartialEq, Clone)]
pub struct FontFormat {
    /// Bold?
    pub bold: bool,
    /// Italic?
    pub italic: bool,
    /// Underlined?
    pub underlined: bool,
    /// Striked?
    pub striked: bool,
    /// Font size.
    pub size: f64,
    /// Font color.
    pub color: Color,
    /// Font name (or default if none).
    pub name: Option<String>,
    /// Font family number.
    pub family_number: i32,
}

impl Default for FontFormat {
    fn default() -> Self {
        Self {
            bold: false,
            italic: false,
            underlined: false,
            striked: false,
            size: Self::DEFAULT_FONT_SIZE,
            color: Color::default(),
            name: None,
            family_number: Self::DEFAULT_FONT_FAMILY,
        }
    }
}

impl FontFormat {
    const DEFAULT_FONT_SIZE: f64 = 11.0;
    const DEFAULT_FONT_FAMILY: i32 = 2;

    /// Is this the default format?
    pub fn is_default(&self) -> bool {
        !self.bold
            && !self.italic
            && !self.underlined
            && !self.striked
            && (self.size - Self::DEFAULT_FONT_SIZE).abs() < f64::EPSILON
            && self.color == Color::default()
            && self.name.is_none() // TODO: Detect file default value.
            && self.family_number == Self::DEFAULT_FONT_FAMILY
    }
}

/// Spreadsheet color for background, text, etc.
#[derive(Debug, PartialEq, Clone, Copy, Default)]
pub enum Color {
    /// Default, no color.
    #[default]
    Default,
    /// Using the [colorIndex property](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12))
    /// `index_id`: color index id.
    Index(u8),
    /// Use the Theme color format, i.e. leave the color decision to the current theme of xlsx.
    /// You can determine the theme color id and tint by `theme_id` and `tint`.
    Theme(u8, f64),
    /// Using the ARGB color format.
    ARGB(u8, u8, u8, u8),
}
