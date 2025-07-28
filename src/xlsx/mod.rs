mod cells_reader;

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{BufReader, Read, Seek};
use std::str::FromStr;
use std::sync::Arc;

use log::warn;
use quick_xml::{
    events::{
        attributes::{Attribute, Attributes},
        BytesStart, Event,
    },
    name::QName,
    Reader as XmlReader,
};
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::conditional_formatting::{ConditionalFormatting, DifferentialFormat};
use crate::datatype::DataRef;
use crate::formats::{
    builtin_format_by_id, detect_custom_number_format_with_interner, Alignment, Border, BorderSide,
    CellFormat, CellStyle, Color, Fill, Font, FormatStringInterner,
};
use crate::vba::VbaProject;
use crate::{
    Cell, CellErrorType, Data, DataWithFormatting, Dimensions, HeaderRow, Metadata, Range, Reader, ReaderRef, Sheet,
    SheetType, SheetVisible, Table,
};
pub use cells_reader::XlsxCellReader;

pub(crate) type XlReader<'a, RS> = XmlReader<BufReader<ZipFile<'a, RS>>>;

/// Maximum number of rows allowed in an xlsx file
pub const MAX_ROWS: u32 = 1_048_576;

/// Maximum number of columns allowed in an xlsx file
pub const MAX_COLUMNS: u32 = 16_384;

/// An enum for Xlsx specific errors
#[derive(Debug)]
pub enum XlsxError {
    /// Io error
    Io(std::io::Error),
    /// Zip error
    Zip(zip::result::ZipError),
    /// Vba error
    Vba(crate::vba::VbaError),
    /// Xml error
    Xml(quick_xml::Error),
    /// Xml attribute error
    XmlAttr(quick_xml::events::attributes::AttrError),
    /// Parse error
    Parse(std::string::ParseError),
    /// Float error
    ParseFloat(std::num::ParseFloatError),
    /// `ParseInt` error
    ParseInt(std::num::ParseIntError),
    /// Unexpected end of xml
    XmlEof(&'static str),
    /// Unexpected node
    UnexpectedNode(&'static str),
    /// File not found
    FileNotFound(String),
    /// Relationship not found
    RelationshipNotFound,
    /// Expecting alphanumeric character
    Alphanumeric(u8),
    /// Numeric column
    NumericColumn(u8),
    /// Wrong dimension count
    DimensionCount(usize),
    /// Cell 't' attribute error
    CellTAttribute(String),
    /// There is no column component in the range string
    RangeWithoutColumnComponent,
    /// There is no row component in the range string
    RangeWithoutRowComponent,
    /// Unexpected error
    Unexpected(&'static str),
    /// Unrecognized data
    Unrecognized {
        /// data type
        typ: &'static str,
        /// value found
        val: String,
    },
    /// Cell error
    CellError(String),
    /// Workbook is password protected
    Password,
    /// Worksheet not found
    WorksheetNotFound(String),
    /// Table not found
    TableNotFound(String),
    /// The specified sheet is not a worksheet
    NotAWorksheet(String),
    /// XML Encoding error
    Encoding(quick_xml::encoding::EncodingError),
    /// XML attribute error
    XmlAttribute(quick_xml::events::attributes::AttrError),
}

from_err!(std::io::Error, XlsxError, Io);
from_err!(zip::result::ZipError, XlsxError, Zip);
from_err!(crate::vba::VbaError, XlsxError, Vba);
from_err!(quick_xml::Error, XlsxError, Xml);
from_err!(std::string::ParseError, XlsxError, Parse);
from_err!(std::num::ParseFloatError, XlsxError, ParseFloat);
from_err!(std::num::ParseIntError, XlsxError, ParseInt);
from_err!(quick_xml::encoding::EncodingError, XlsxError, Encoding);
from_err!(
    quick_xml::events::attributes::AttrError,
    XlsxError,
    XmlAttribute
);

impl std::fmt::Display for XlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsxError::Io(e) => write!(f, "I/O error: {e}"),
            XlsxError::Zip(e) => write!(f, "Zip error: {e}"),
            XlsxError::Xml(e) => write!(f, "Xml error: {e}"),
            XlsxError::XmlAttr(e) => write!(f, "Xml attribute error: {e}"),
            XlsxError::Vba(e) => write!(f, "Vba error: {e}"),
            XlsxError::Parse(e) => write!(f, "Parse string error: {e}"),
            XlsxError::ParseInt(e) => write!(f, "Parse integer error: {e}"),
            XlsxError::ParseFloat(e) => write!(f, "Parse float error: {e}"),

            XlsxError::XmlEof(e) => write!(f, "Unexpected end of xml, expecting '</{e}>'"),
            XlsxError::UnexpectedNode(e) => write!(f, "Expecting '{e}' node"),
            XlsxError::FileNotFound(e) => write!(f, "File not found '{e}'"),
            XlsxError::RelationshipNotFound => write!(f, "Relationship not found"),
            XlsxError::Alphanumeric(e) => {
                write!(f, "Expecting alphanumeric character, got {e:X}")
            }
            XlsxError::NumericColumn(e) => write!(
                f,
                "Numeric character is not allowed for column name, got {e}",
            ),
            XlsxError::DimensionCount(e) => {
                write!(f, "Range dimension must be lower than 2. Got {e}")
            }
            XlsxError::CellTAttribute(e) => write!(f, "Unknown cell 't' attribute: {e:?}"),
            XlsxError::RangeWithoutColumnComponent => {
                write!(f, "Range is missing the expected column component.")
            }
            XlsxError::RangeWithoutRowComponent => {
                write!(f, "Range is missing the expected row component.")
            }
            XlsxError::Unexpected(e) => write!(f, "{e}"),
            XlsxError::Unrecognized { typ, val } => write!(f, "Unrecognized {typ}: {val}"),
            XlsxError::CellError(e) => write!(f, "Unsupported cell error value '{e}'"),
            XlsxError::WorksheetNotFound(n) => write!(f, "Worksheet '{n}' not found"),
            XlsxError::Password => write!(f, "Workbook is password protected"),
            XlsxError::TableNotFound(n) => write!(f, "Table '{n}' not found"),
            XlsxError::NotAWorksheet(typ) => write!(f, "Expecting a worksheet, got {typ}"),
            XlsxError::Encoding(e) => write!(f, "XML encoding error: {e}"),
            XlsxError::XmlAttribute(e) => write!(f, "XML attribute error: {e}"),
        }
    }
}

impl std::error::Error for XlsxError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsxError::Io(e) => Some(e),
            XlsxError::Zip(e) => Some(e),
            XlsxError::Xml(e) => Some(e),
            XlsxError::Vba(e) => Some(e),
            XlsxError::Parse(e) => Some(e),
            XlsxError::ParseInt(e) => Some(e),
            XlsxError::ParseFloat(e) => Some(e),
            XlsxError::Encoding(e) => Some(e),
            XlsxError::XmlAttribute(e) => Some(e),
            _ => None,
        }
    }
}

impl FromStr for CellErrorType {
    type Err = XlsxError;
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "#DIV/0!" => Ok(CellErrorType::Div0),
            "#N/A" => Ok(CellErrorType::NA),
            "#NAME?" => Ok(CellErrorType::Name),
            "#NULL!" => Ok(CellErrorType::Null),
            "#NUM!" => Ok(CellErrorType::Num),
            "#REF!" => Ok(CellErrorType::Ref),
            "#VALUE!" => Ok(CellErrorType::Value),
            _ => Err(XlsxError::CellError(s.into())),
        }
    }
}

type Tables = Option<Vec<(String, String, Vec<String>, Dimensions)>>;

/// A struct representing xml zipped excel file
/// Xlsx, Xlsm, Xlam
pub struct Xlsx<RS> {
    zip: ZipArchive<RS>,
    /// Shared strings
    strings: Vec<String>,
    /// Sheets paths
    sheets: Vec<(String, String)>,
    /// Tables: Name, Sheet, Columns, Data dimensions
    tables: Tables,
    /// Cell formats (backward compatible)
    formats: Vec<CellFormat>,
    /// Cell formats (comprehensive formatting information)
    styles: Vec<CellStyle>,
    /// Format string interner for reuse across sheets
    format_interner: FormatStringInterner,
    /// 1904 datetime system
    is_1904: bool,
    /// Metadata
    metadata: Metadata,
    /// Pictures
    #[cfg(feature = "picture")]
    pictures: Option<Vec<(String, Vec<u8>)>>,
    /// Merged Regions: Name, Sheet, Merged Dimensions
    merged_regions: Option<Vec<(String, String, Dimensions)>>,
    /// Reader options
    options: XlsxOptions,
    /// Differential formats (for conditional formatting)
    dxf_formats: Vec<DifferentialFormat>,
    /// Conditional formatting rules by sheet name
    conditional_formats: BTreeMap<String, Vec<ConditionalFormatting>>,
}

/// Xlsx reader options
#[derive(Debug, Default)]
#[non_exhaustive]
struct XlsxOptions {
    pub header_row: HeaderRow,
}

impl<RS: Read + Seek> Xlsx<RS> {
    fn read_shared_strings(&mut self) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/sharedStrings.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => {
                    if let Some(s) = read_string(&mut xml, e.name())? {
                        self.strings.push(s);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sst" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sst")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    fn read_styles(&mut self) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/styles.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };

        let mut number_formats = BTreeMap::new();
        let format_interner = FormatStringInterner::new();

        let mut fonts: Vec<Arc<Font>> = Vec::new();
        let mut fills: Vec<Arc<Fill>> = Vec::new();
        let mut borders: Vec<Arc<Border>> = Vec::new();

        let mut buf = Vec::with_capacity(1024);
        let mut inner_buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => {
                    // Parse custom number formats
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                                let mut id = 0u32;
                                let mut format = String::new();
                                for a in e.attributes() {
                                    match a.map_err(XlsxError::XmlAttr)? {
                                        Attribute {
                                            key: QName(b"numFmtId"),
                                            value: v,
                                        } => {
                                            id = atoi_simd::parse::<u32>(&v).unwrap_or(0);
                                        }
                                        Attribute {
                                            key: QName(b"formatCode"),
                                            value: v,
                                        } => format = xml.decoder().decode(&v)?.into_owned(),
                                        _ => (),
                                    }
                                }
                                if !format.is_empty() {
                                    number_formats.insert(id, format);
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"numFmts" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("numFmts")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fonts" => {
                    // Parse fonts
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                                let font = Self::parse_font_element(&mut xml, &mut inner_buf)?;
                                fonts.push(Arc::new(font));
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fonts" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fonts")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fills" => {
                    // Parse fills
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                                let fill = Self::parse_fill_element(&mut xml, &mut inner_buf)?;
                                fills.push(Arc::new(fill));
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fills" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("fills")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"borders" => {
                    // Parse borders
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                                let border = Self::parse_border_element(&mut xml, &mut inner_buf)?;
                                borders.push(Arc::new(border));
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"borders" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("borders")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                    // Parse cell formats (comprehensive formatting)
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                                let mut cell_formatting = CellStyle::default();

                                // Parse attributes to get references to fonts, fills, borders, number formats
                                for attr in e.attributes() {
                                    match attr.map_err(XlsxError::XmlAttr)? {
                                        Attribute {
                                            key: QName(b"numFmtId"),
                                            value: v,
                                        } => {
                                            let num_fmt_id =
                                                atoi_simd::parse::<u32>(&v).unwrap_or(0);
                                            if let Some(fmt) = number_formats.get(&num_fmt_id) {
                                                let (detected_format, format_string) =
                                                    detect_custom_number_format_with_interner(
                                                        fmt,
                                                        &format_interner,
                                                    );
                                                cell_formatting.number_format = detected_format;
                                                cell_formatting.format_string = format_string;
                                            } else {
                                                cell_formatting.number_format =
                                                    builtin_format_by_id(
                                                        &num_fmt_id.to_string().into_bytes(),
                                                    );
                                                cell_formatting.format_string = None;
                                            }
                                        }
                                        Attribute {
                                            key: QName(b"fontId"),
                                            value: v,
                                        } => {
                                            let font_id =
                                                atoi_simd::parse::<usize>(&v).unwrap_or(0);
                                            cell_formatting.font = fonts.get(font_id).cloned();
                                        }
                                        Attribute {
                                            key: QName(b"fillId"),
                                            value: v,
                                        } => {
                                            let fill_id =
                                                atoi_simd::parse::<usize>(&v).unwrap_or(0);
                                            cell_formatting.fill = fills.get(fill_id).cloned();
                                        }
                                        Attribute {
                                            key: QName(b"borderId"),
                                            value: v,
                                        } => {
                                            let border_id =
                                                atoi_simd::parse::<usize>(&v).unwrap_or(0);
                                            cell_formatting.border =
                                                borders.get(border_id).cloned();
                                        }
                                        _ => (),
                                    }
                                }

                                // Parse alignment if present
                                cell_formatting.alignment =
                                    Self::parse_alignment_from_xf(&mut xml, &mut inner_buf)?
                                        .map(Arc::new);

                                // For backward compatibility, also push to the old formats field
                                self.formats.push(cell_formatting.number_format.clone());
                                self.styles.push(cell_formatting);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("cellXfs")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dxfs" => {
                    // Parse differential formats
                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dxf" => {
                                let dxf = Self::parse_dxf_element(&mut xml, &mut inner_buf, &number_formats, &format_interner)?;
                                self.dxf_formats.push(dxf);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxfs" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("dxfs")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("styleSheet")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(())
    }

    /// Parse a font element from XML
    fn parse_font_element(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
    ) -> Result<Font, XlsxError> {
        use crate::formats::Font;

        let mut font = Font {
            name: None,
            size: None,
            bold: None,
            italic: None,
            color: None,
        };

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"name" => {
                        if let Some(val) = get_attribute(e.attributes(), QName(b"val"))? {
                            font.name = Some(Arc::from(xml.decoder().decode(val)?.as_ref()));
                        }
                    }
                    b"sz" => {
                        if let Some(val) = get_attribute(e.attributes(), QName(b"val"))? {
                            if let Ok(size) = xml.decoder().decode(val)?.parse::<f64>() {
                                font.size = Some(size);
                            }
                        }
                    }
                    b"b" => font.bold = Some(true),
                    b"i" => font.italic = Some(true),
                    b"color" => {
                        font.color = Self::parse_color_from_attributes(e.attributes())?;
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(font)
    }

    /// Parse a fill element from XML
    fn parse_fill_element(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
    ) -> Result<Fill, XlsxError> {
        use crate::formats::{Fill, PatternType};

        let mut fill = Fill {
            pattern_type: PatternType::None,
            foreground_color: None,
            background_color: None,
        };

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"patternFill" => {
                        if let Some(pattern_type) =
                            get_attribute(e.attributes(), QName(b"patternType"))?
                        {
                            let pattern_str = xml.decoder().decode(pattern_type)?;
                            fill.pattern_type = match pattern_str.as_ref() {
                                "none" => PatternType::None,
                                "solid" => PatternType::Solid,
                                "lightGray" => PatternType::LightGray,
                                "mediumGray" => PatternType::MediumGray,
                                "darkGray" => PatternType::DarkGray,
                                other => PatternType::Pattern(Arc::from(other)),
                            };
                        }
                    }
                    b"fgColor" => {
                        fill.foreground_color = Self::parse_color_from_attributes(e.attributes())?;
                    }
                    b"bgColor" => {
                        fill.background_color = Self::parse_color_from_attributes(e.attributes())?;
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(fill)
    }

    /// Parse a border element from XML
    fn parse_border_element(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
    ) -> Result<Border, XlsxError> {
        use crate::formats::Border;

        let mut border = Border {
            left: None,
            right: None,
            top: None,
            bottom: None,
        };

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"left" => {
                        let mut temp_buf = Vec::new();
                        border.left = Self::parse_border_side(xml, e, &mut temp_buf)?;
                    }
                    b"right" => {
                        let mut temp_buf = Vec::new();
                        border.right = Self::parse_border_side(xml, e, &mut temp_buf)?;
                    }
                    b"top" => {
                        let mut temp_buf = Vec::new();
                        border.top = Self::parse_border_side(xml, e, &mut temp_buf)?;
                    }
                    b"bottom" => {
                        let mut temp_buf = Vec::new();
                        border.bottom = Self::parse_border_side(xml, e, &mut temp_buf)?;
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"border" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(border)
    }

    /// Parse border side information
    fn parse_border_side(
        xml: &mut XlReader<'_, RS>,
        element: &BytesStart<'_>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<BorderSide>, XlsxError> {
        use crate::formats::BorderSide;

        let style = match get_attribute(element.attributes(), QName(b"style"))? {
            Some(style_attr) => Arc::from(xml.decoder().decode(style_attr)?.as_ref()),
            None => return Ok(None),
        };

        let mut color = None;
        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                    color = Self::parse_color_from_attributes(e.attributes())?;
                }
                Ok(Event::End(ref e)) if e.local_name() == element.local_name() => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("border side")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(Some(BorderSide { style, color }))
    }

    /// Parse alignment information from cellXfs
    fn parse_alignment_from_xf(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
    ) -> Result<Option<Alignment>, XlsxError> {
        use crate::formats::Alignment;

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alignment" => {
                    let mut alignment = Alignment {
                        horizontal: None,
                        vertical: None,
                        wrap_text: None,
                        indent: None,
                        shrink_to_fit: None,
                        text_rotation: None,
                        reading_order: None,
                    };

                    for attr in e.attributes() {
                        match attr.map_err(XlsxError::XmlAttr)? {
                            Attribute {
                                key: QName(b"horizontal"),
                                value: v,
                            } => {
                                alignment.horizontal =
                                    Some(Arc::from(xml.decoder().decode(&v)?.as_ref()));
                            }
                            Attribute {
                                key: QName(b"vertical"),
                                value: v,
                            } => {
                                alignment.vertical =
                                    Some(Arc::from(xml.decoder().decode(&v)?.as_ref()));
                            }
                            Attribute {
                                key: QName(b"wrapText"),
                                value: v,
                            } => {
                                alignment.wrap_text = Some(&*v == b"1" || &*v == b"true");
                            }
                            Attribute {
                                key: QName(b"indent"),
                                value: v,
                            } => {
                                if let Ok(indent) = xml.decoder().decode(&v)?.parse::<u32>() {
                                    alignment.indent = Some(indent);
                                }
                            }
                            Attribute {
                                key: QName(b"shrinkToFit"),
                                value: v,
                            } => {
                                alignment.shrink_to_fit = Some(&*v == b"1" || &*v == b"true");
                            }
                            Attribute {
                                key: QName(b"textRotation"),
                                value: v,
                            } => {
                                if let Ok(rotation) = xml.decoder().decode(&v)?.parse::<i32>() {
                                    alignment.text_rotation = Some(rotation);
                                }
                            }
                            Attribute {
                                key: QName(b"readingOrder"),
                                value: v,
                            } => {
                                if let Ok(order) = xml.decoder().decode(&v)?.parse::<u32>() {
                                    alignment.reading_order = Some(order);
                                }
                            }
                            _ => (),
                        }
                    }

                    return Ok(Some(alignment));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xf" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("xf")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(None)
    }

    /// Parse color from element attributes
    /// Follows Excel precedence: rgb > theme > indexed > auto
    fn parse_color_from_attributes(attributes: Attributes<'_>) -> Result<Option<Color>, XlsxError> {
        use crate::formats::Color;

        let mut rgb_color = None;
        let mut theme_color = None;
        let mut indexed_color = None;
        let mut auto_color = None;
        let mut tint_value = None;

        // First pass: collect all attributes
        for attr in attributes {
            match attr.map_err(XlsxError::XmlAttr)? {
                Attribute {
                    key: QName(b"rgb"),
                    value: v,
                } => {
                    let color_str = std::str::from_utf8(&v)
                        .map_err(|_| XlsxError::Unexpected("Invalid UTF-8 in color RGB value"))?;
                    if color_str.len() == 8 {
                        // ARGB format
                        if let (Ok(a), Ok(r), Ok(g), Ok(b)) = (
                            u8::from_str_radix(&color_str[0..2], 16),
                            u8::from_str_radix(&color_str[2..4], 16),
                            u8::from_str_radix(&color_str[4..6], 16),
                            u8::from_str_radix(&color_str[6..8], 16),
                        ) {
                            rgb_color = Some(Color::Argb { a, r, g, b });
                        } else {
                            log::warn!("Invalid ARGB color format: {}", color_str);
                        }
                    } else if color_str.len() == 6 {
                        // RGB format
                        if let (Ok(r), Ok(g), Ok(b)) = (
                            u8::from_str_radix(&color_str[0..2], 16),
                            u8::from_str_radix(&color_str[2..4], 16),
                            u8::from_str_radix(&color_str[4..6], 16),
                        ) {
                            rgb_color = Some(Color::Rgb { r, g, b });
                        } else {
                            log::warn!("Invalid RGB color format: {}", color_str);
                        }
                    } else {
                        log::warn!("Invalid color format length: {}", color_str);
                    }
                }
                Attribute {
                    key: QName(b"theme"),
                    value: v,
                } => {
                    if let Ok(theme) = atoi_simd::parse::<u32>(&v) {
                        theme_color = Some(theme);
                    }
                }
                Attribute {
                    key: QName(b"tint"),
                    value: v,
                } => {
                    if let Ok(tint_str) = std::str::from_utf8(&v) {
                        if let Ok(tint) = tint_str.parse::<f64>() {
                            // Clamp tint to valid range [-1.0, 1.0]
                            tint_value = Some(tint.clamp(-1.0, 1.0));
                        }
                    }
                }
                Attribute {
                    key: QName(b"indexed"),
                    value: v,
                } => {
                    if let Ok(indexed) = atoi_simd::parse::<u32>(&v) {
                        indexed_color = Some(indexed);
                    }
                }
                Attribute {
                    key: QName(b"auto"),
                    value: v,
                } => {
                    if &*v == b"1" || &*v == b"true" {
                        auto_color = Some(());
                    }
                }
                _ => (),
            }
        }

        // Apply precedence: rgb > theme > indexed > auto
        if let Some(color) = rgb_color {
            // RGB colors can also have tint in some cases
            match color {
                Color::Rgb { r, g, b } if tint_value.is_some() => {
                    Ok(Some(Color::Argb { a: 255, r, g, b }))
                }
                other => Ok(Some(other)),
            }
        } else if let Some(theme) = theme_color {
            Ok(Some(Color::Theme {
                theme,
                tint: tint_value,
            }))
        } else if let Some(indexed) = indexed_color {
            Ok(Some(Color::Indexed(indexed)))
        } else if auto_color.is_some() {
            Ok(Some(Color::Auto))
        } else {
            Ok(None)
        }
    }

    /// Get conditional formatting for a worksheet
    pub fn worksheet_conditional_formatting(
        &mut self,
        name: &str,
    ) -> Result<&[ConditionalFormatting], XlsxError> {
        // Find the sheet path
        let sheet_path = match self.sheets.iter().find(|(n, _)| n == name) {
            Some((_, path)) => path.clone(),
            None => return Err(XlsxError::WorksheetNotFound(name.to_string())),
        };

        // Check if we've already loaded this sheet's conditional formatting
        if !self.conditional_formats.contains_key(name) {
            // Load the conditional formatting
            let formats = Self::parse_worksheet_conditional_formatting(&sheet_path, &mut self.zip)?;
            self.conditional_formats.insert(name.to_string(), formats);
        }

        Ok(self.conditional_formats.get(name).map(|v| v.as_slice()).unwrap_or(&[]))
    }

    /// Get differential formats
    pub fn dxf_formats(&self) -> &[DifferentialFormat] {
        &self.dxf_formats
    }

    /// Parse conditional formatting from a worksheet
    fn parse_worksheet_conditional_formatting(
        sheet_path: &str,
        zip: &mut ZipArchive<RS>,
    ) -> Result<Vec<ConditionalFormatting>, XlsxError> {
        use crate::conditional_formatting::ConditionalFormatting;

        let mut xml = match xml_reader(zip, sheet_path) {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };

        let mut conditional_formats = Vec::new();
        let mut buf = Vec::with_capacity(1024);

        // Skip to conditionalFormatting elements
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"conditionalFormatting" => {
                    let mut ranges = Vec::new();
                    let mut pivot = false;

                    // Parse attributes
                    for attr in e.attributes() {
                        match attr.map_err(XlsxError::XmlAttr)? {
                            Attribute {
                                key: QName(b"sqref"),
                                value: v,
                            } => {
                                let sqref = xml.decoder().decode(&v)?;
                                // Split by space and parse each range
                                for range_str in sqref.split_whitespace() {
                                    if let Ok(dims) = get_dimension(range_str.as_bytes()) {
                                        ranges.push(dims);
                                    }
                                }
                            }
                            Attribute {
                                key: QName(b"pivot"),
                                value: v,
                            } => {
                                pivot = &*v == b"1" || &*v == b"true";
                            }
                            _ => (),
                        }
                    }

                    // Parse rules
                    let mut rules = Vec::new();
                    let mut inner_buf = Vec::new();

                    loop {
                        inner_buf.clear();
                        match xml.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cfRule" => {
                                let mut rule_buf = Vec::new();
                                let rule = Self::parse_cf_rule(&mut xml, e, &mut rule_buf, pivot)?;
                                rules.push(rule);
                            }
                            Ok(Event::End(ref e))
                                if e.local_name().as_ref() == b"conditionalFormatting" =>
                            {
                                break
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("conditionalFormatting")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }

                    if !rules.is_empty() && !ranges.is_empty() {
                        conditional_formats.push(ConditionalFormatting { 
                            ranges, 
                            rules,
                            scope: None,
                            table: None,
                        });
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"worksheet" => break,
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(conditional_formats)
    }

    /// Parse a single cfRule element
    fn parse_cf_rule(
        xml: &mut XlReader<'_, RS>,
        rule_start: &BytesStart<'_>,
        buf: &mut Vec<u8>,
        pivot: bool,
    ) -> Result<crate::conditional_formatting::ConditionalFormatRule, XlsxError> {
        use crate::conditional_formatting::{
            CfvoType, ColorScale, ComparisonOperator, ConditionalFormatRule, ConditionalFormatType,
            ConditionalFormatValue, DataBar, IconSet, IconSetType, TimePeriod,
        };

        let mut rule_type = ConditionalFormatType::Expression;
        let mut priority = 0i32;
        let mut stop_if_true = false;
        let mut dxf_id = None;
        let mut formulas = Vec::new();
        let mut operator = None;
        let mut text = None;
        let mut time_period = None;
        let mut rank = None;
        let mut bottom = false;
        let mut percent = false;
        let mut above_average = true;
        let mut equal_average = false;
        let mut std_dev = None;

        // Parse attributes
        for attr in rule_start.attributes() {
            match attr.map_err(XlsxError::XmlAttr)? {
                Attribute {
                    key: QName(b"type"),
                    value: v,
                } => {
                    let type_str = xml.decoder().decode(&v)?;
                    rule_type = match type_str.as_ref() {
                        "cellIs" => ConditionalFormatType::CellIs {
                            operator: ComparisonOperator::Equal,
                        },
                        "expression" => ConditionalFormatType::Expression,
                        "top10" => ConditionalFormatType::Top10 {
                            bottom: false,
                            percent: false,
                            rank: 10,
                        },
                        "duplicateValues" => ConditionalFormatType::DuplicateValues,
                        "uniqueValues" => ConditionalFormatType::UniqueValues,
                        "containsText" => ConditionalFormatType::ContainsText {
                            text: String::new(),
                        },
                        "notContainsText" => ConditionalFormatType::NotContainsText {
                            text: String::new(),
                        },
                        "beginsWith" => ConditionalFormatType::BeginsWith {
                            text: String::new(),
                        },
                        "endsWith" => ConditionalFormatType::EndsWith {
                            text: String::new(),
                        },
                        "containsBlanks" => ConditionalFormatType::ContainsBlanks,
                        "notContainsBlanks" => ConditionalFormatType::NotContainsBlanks,
                        "containsErrors" => ConditionalFormatType::ContainsErrors,
                        "notContainsErrors" => ConditionalFormatType::NotContainsErrors,
                        "timePeriod" => ConditionalFormatType::TimePeriod {
                            period: TimePeriod::Today,
                        },
                        "aboveAverage" => ConditionalFormatType::AboveAverage {
                            below: false,
                            equal_average: false,
                            std_dev: None,
                        },
                        "dataBar" => ConditionalFormatType::DataBar(DataBar {
                            min_cfvo: ConditionalFormatValue {
                                value_type: CfvoType::Min,
                                value: None,
                                gte: false,
                            },
                            max_cfvo: ConditionalFormatValue {
                                value_type: CfvoType::Max,
                                value: None,
                                gte: false,
                            },
                            color: crate::formats::Color::Rgb { r: 0, g: 0, b: 255 },
                            negative_color: None,
                            show_value: true,
                            min_length: 10,
                            max_length: 90,
                            direction: None,
                            bar_only: false,
                            border_color: None,
                            negative_border_color: None,
                            gradient: true,
                            axis_position: None,
                            axis_color: None,
                        }),
                        "colorScale" => ConditionalFormatType::ColorScale(ColorScale {
                            cfvos: Vec::new(),
                            colors: Vec::new(),
                        }),
                        "iconSet" => ConditionalFormatType::IconSet(IconSet {
                            icon_set: IconSetType::Arrows3,
                            cfvos: Vec::new(),
                            show_value: true,
                            reverse: false,
                            custom_icons: Vec::new(),
                            percent: false,
                        }),
                        _ => ConditionalFormatType::Expression,
                    };
                }
                Attribute {
                    key: QName(b"dxfId"),
                    value: v,
                } => {
                    if let Ok(id) = atoi_simd::parse::<u32>(&v) {
                        dxf_id = Some(id);
                    }
                }
                Attribute {
                    key: QName(b"priority"),
                    value: v,
                } => {
                    if let Ok(p) = atoi_simd::parse::<i32>(&v) {
                        priority = p;
                    }
                }
                Attribute {
                    key: QName(b"stopIfTrue"),
                    value: v,
                } => {
                    stop_if_true = &*v == b"1" || &*v == b"true";
                }
                Attribute {
                    key: QName(b"operator"),
                    value: v,
                } => {
                    let op_str = xml.decoder().decode(&v)?;
                    operator = Some(match op_str.as_ref() {
                        "lessThan" => ComparisonOperator::LessThan,
                        "lessThanOrEqual" => ComparisonOperator::LessThanOrEqual,
                        "equal" => ComparisonOperator::Equal,
                        "notEqual" => ComparisonOperator::NotEqual,
                        "greaterThanOrEqual" => ComparisonOperator::GreaterThanOrEqual,
                        "greaterThan" => ComparisonOperator::GreaterThan,
                        "between" => ComparisonOperator::Between,
                        "notBetween" => ComparisonOperator::NotBetween,
                        "containsText" => ComparisonOperator::ContainsText,
                        "notContains" => ComparisonOperator::NotContains,
                        _ => ComparisonOperator::Equal,
                    });
                }
                Attribute {
                    key: QName(b"text"),
                    value: v,
                } => {
                    text = Some(xml.decoder().decode(&v)?.into_owned());
                }
                Attribute {
                    key: QName(b"timePeriod"),
                    value: v,
                } => {
                    let period_str = xml.decoder().decode(&v)?;
                    time_period = Some(match period_str.as_ref() {
                        "today" => TimePeriod::Today,
                        "yesterday" => TimePeriod::Yesterday,
                        "tomorrow" => TimePeriod::Tomorrow,
                        "last7Days" => TimePeriod::Last7Days,
                        "thisWeek" => TimePeriod::ThisWeek,
                        "lastWeek" => TimePeriod::LastWeek,
                        "nextWeek" => TimePeriod::NextWeek,
                        "thisMonth" => TimePeriod::ThisMonth,
                        "lastMonth" => TimePeriod::LastMonth,
                        "nextMonth" => TimePeriod::NextMonth,
                        "thisQuarter" => TimePeriod::ThisQuarter,
                        "lastQuarter" => TimePeriod::LastQuarter,
                        "nextQuarter" => TimePeriod::NextQuarter,
                        "thisYear" => TimePeriod::ThisYear,
                        "lastYear" => TimePeriod::LastYear,
                        "nextYear" => TimePeriod::NextYear,
                        "yearToDate" => TimePeriod::YearToDate,
                        "allDatesInPeriodJanuary" => TimePeriod::AllDatesInJanuary,
                        "allDatesInPeriodFebruary" => TimePeriod::AllDatesInFebruary,
                        "allDatesInPeriodMarch" => TimePeriod::AllDatesInMarch,
                        "allDatesInPeriodApril" => TimePeriod::AllDatesInApril,
                        "allDatesInPeriodMay" => TimePeriod::AllDatesInMay,
                        "allDatesInPeriodJune" => TimePeriod::AllDatesInJune,
                        "allDatesInPeriodJuly" => TimePeriod::AllDatesInJuly,
                        "allDatesInPeriodAugust" => TimePeriod::AllDatesInAugust,
                        "allDatesInPeriodSeptember" => TimePeriod::AllDatesInSeptember,
                        "allDatesInPeriodOctober" => TimePeriod::AllDatesInOctober,
                        "allDatesInPeriodNovember" => TimePeriod::AllDatesInNovember,
                        "allDatesInPeriodDecember" => TimePeriod::AllDatesInDecember,
                        "allDatesInPeriodQuarter1" => TimePeriod::AllDatesInQ1,
                        "allDatesInPeriodQuarter2" => TimePeriod::AllDatesInQ2,
                        "allDatesInPeriodQuarter3" => TimePeriod::AllDatesInQ3,
                        "allDatesInPeriodQuarter4" => TimePeriod::AllDatesInQ4,
                        _ => TimePeriod::Today,
                    });
                }
                Attribute {
                    key: QName(b"rank"),
                    value: v,
                } => {
                    if let Ok(r) = atoi_simd::parse::<u32>(&v) {
                        rank = Some(r);
                    }
                }
                Attribute {
                    key: QName(b"bottom"),
                    value: v,
                } => {
                    bottom = &*v == b"1" || &*v == b"true";
                }
                Attribute {
                    key: QName(b"percent"),
                    value: v,
                } => {
                    percent = &*v == b"1" || &*v == b"true";
                }
                Attribute {
                    key: QName(b"aboveAverage"),
                    value: v,
                } => {
                    above_average = &*v != b"0" && &*v != b"false";
                }
                Attribute {
                    key: QName(b"equalAverage"),
                    value: v,
                } => {
                    equal_average = &*v == b"1" || &*v == b"true";
                }
                Attribute {
                    key: QName(b"stdDev"),
                    value: v,
                } => {
                    if let Ok(dev) = atoi_simd::parse::<u32>(&v) {
                        std_dev = Some(dev);
                    }
                }
                _ => (),
            }
        }

        // Update rule type with parsed attributes
        rule_type = match rule_type {
            ConditionalFormatType::CellIs { .. } => ConditionalFormatType::CellIs {
                operator: operator.unwrap_or(ComparisonOperator::Equal),
            },
            ConditionalFormatType::Top10 { .. } => ConditionalFormatType::Top10 {
                bottom,
                percent,
                rank: rank.unwrap_or(10),
            },
            ConditionalFormatType::ContainsText { .. } => ConditionalFormatType::ContainsText {
                text: text.clone().unwrap_or_default(),
            },
            ConditionalFormatType::BeginsWith { .. } => ConditionalFormatType::BeginsWith {
                text: text.clone().unwrap_or_default(),
            },
            ConditionalFormatType::EndsWith { .. } => ConditionalFormatType::EndsWith {
                text: text.clone().unwrap_or_default(),
            },
            ConditionalFormatType::TimePeriod { .. } => ConditionalFormatType::TimePeriod {
                period: time_period.unwrap_or(TimePeriod::Today),
            },
            ConditionalFormatType::AboveAverage { .. } => ConditionalFormatType::AboveAverage {
                below: !above_average,
                equal_average,
                std_dev,
            },
            _ => rule_type,
        };

        // Parse child elements
        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"formula" => {
                        if let Ok(Event::Text(ref t)) = xml.read_event_into(buf) {
                            let formula_text = xml.decoder().decode(t)?.into_owned();
                            formulas.push(formula_text);
                        }
                    }
                    b"dataBar" => {
                        if let ConditionalFormatType::DataBar(ref mut data_bar) = rule_type {
                            Self::parse_data_bar(xml, buf, data_bar)?;
                        }
                    }
                    b"colorScale" => {
                        if let ConditionalFormatType::ColorScale(ref mut color_scale) = rule_type {
                            Self::parse_color_scale(xml, buf, color_scale)?;
                        }
                    }
                    b"iconSet" => {
                        if let ConditionalFormatType::IconSet(ref mut icon_set) = rule_type {
                            Self::parse_icon_set(xml, buf, icon_set)?;
                        }
                    }
                    b"extLst" => {
                        // Skip extensions for now
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cfRule" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("cfRule")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(ConditionalFormatRule {
            rule_type,
            priority,
            stop_if_true,
            dxf_id,
            formulas,
            pivot,
            text,
            operator: operator.map(|op| op.to_string()),
            bottom: if bottom { Some(true) } else { None },
            percent: if percent { Some(true) } else { None },
            rank: rank.map(|r| r as i32),
            above_average: if above_average { Some(true) } else { None },
            equal_average: if equal_average { Some(true) } else { None },
            std_dev: std_dev.map(|d| d as i32),
        })
    }

    /// Parse data bar element
    fn parse_data_bar(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
        data_bar: &mut crate::conditional_formatting::DataBar,
    ) -> Result<(), XlsxError> {
        use crate::conditional_formatting::{AxisPosition, BarDirection};

        let mut cfvo_count = 0;

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"dataBar" => {
                        // Parse dataBar attributes
                        for attr in e.attributes() {
                            match attr.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"showValue"),
                                    value: v,
                                } => {
                                    data_bar.show_value = &*v != b"0" && &*v != b"false";
                                }
                                Attribute {
                                    key: QName(b"minLength"),
                                    value: v,
                                } => {
                                    if let Ok(len) = atoi_simd::parse::<u32>(&v) {
                                        data_bar.min_length = len;
                                    }
                                }
                                Attribute {
                                    key: QName(b"maxLength"),
                                    value: v,
                                } => {
                                    if let Ok(len) = atoi_simd::parse::<u32>(&v) {
                                        data_bar.max_length = len;
                                    }
                                }
                                _ => (),
                            }
                        }
                    }
                    b"cfvo" => {
                        let cfvo = Self::parse_cfvo(e.attributes(), xml)?;
                        if cfvo_count == 0 {
                            data_bar.min_cfvo = cfvo;
                        } else if cfvo_count == 1 {
                            data_bar.max_cfvo = cfvo;
                        }
                        cfvo_count += 1;
                    }
                    b"color" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            data_bar.color = color;
                        }
                    }
                    b"negativeFillColor" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            data_bar.negative_color = Some(color);
                        }
                    }
                    b"borderColor" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            data_bar.border_color = Some(color);
                        }
                    }
                    b"negativeBorderColor" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            data_bar.negative_border_color = Some(color);
                        }
                    }
                    b"axisColor" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            data_bar.axis_color = Some(color);
                        }
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                    b"dataBar" => {
                        // Handle self-closing dataBar tag with attributes
                        for attr in e.attributes() {
                            match attr.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"direction"),
                                    value: v,
                                } => {
                                    let dir_str = xml.decoder().decode(&v)?;
                                    data_bar.direction = Some(match dir_str.as_ref() {
                                        "leftToRight" => BarDirection::LeftToRight,
                                        "rightToLeft" => BarDirection::RightToLeft,
                                        _ => BarDirection::LeftToRight,
                                    });
                                }
                                Attribute {
                                    key: QName(b"gradient"),
                                    value: v,
                                } => {
                                    data_bar.gradient = &*v != b"0" && &*v != b"false";
                                }
                                Attribute {
                                    key: QName(b"axisPosition"),
                                    value: v,
                                } => {
                                    let pos_str = xml.decoder().decode(&v)?;
                                    data_bar.axis_position = Some(match pos_str.as_ref() {
                                        "automatic" => AxisPosition::Automatic,
                                        "midpoint" => AxisPosition::Midpoint,
                                        "none" => AxisPosition::None,
                                        _ => AxisPosition::Automatic,
                                    });
                                }
                                _ => (),
                            }
                        }
                    }
                    _ => (),
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dataBar" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("dataBar")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(())
    }

    /// Parse color scale element
    fn parse_color_scale(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
        color_scale: &mut crate::conditional_formatting::ColorScale,
    ) -> Result<(), XlsxError> {
        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"cfvo" => {
                        let cfvo = Self::parse_cfvo(e.attributes(), xml)?;
                        color_scale.cfvos.push(cfvo);
                    }
                    b"color" => {
                        if let Some(color) = Self::parse_color_from_attributes(e.attributes())? {
                            color_scale.colors.push(color);
                        }
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"colorScale" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("colorScale")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(())
    }

    /// Parse icon set element
    fn parse_icon_set(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
        icon_set: &mut crate::conditional_formatting::IconSet,
    ) -> Result<(), XlsxError> {
        use crate::conditional_formatting::IconSetType;

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"iconSet" => {
                    // Parse attributes
                    for attr in e.attributes() {
                        match attr.map_err(XlsxError::XmlAttr)? {
                Attribute {
                    key: QName(b"iconSet"),
                    value: v,
                } => {
                    let icon_str = xml.decoder().decode(&v)?;
                    icon_set.icon_set = match icon_str.as_ref() {
                        "3Arrows" => IconSetType::Arrows3,
                        "3ArrowsGray" => IconSetType::Arrows3Gray,
                        "4Arrows" => IconSetType::Arrows4,
                        "4ArrowsGray" => IconSetType::Arrows4Gray,
                        "5Arrows" => IconSetType::Arrows5,
                        "5ArrowsGray" => IconSetType::Arrows5Gray,
                        "3Flags" => IconSetType::Flags3,
                        "3TrafficLights1" => IconSetType::TrafficLights3,
                        "3TrafficLights2" => IconSetType::TrafficLights3Rimmed,
                        "4TrafficLights" => IconSetType::TrafficLights4,
                        "3Signs" => IconSetType::Signs3,
                        "3Symbols" => IconSetType::Symbols3,
                        "3Symbols2" => IconSetType::Symbols3Uncircled,
                        "4Rating" => IconSetType::Rating4,
                        "5Rating" => IconSetType::Rating5,
                        "5Quarters" => IconSetType::Quarters5,
                        "3Stars" => IconSetType::Stars3,
                        "3Triangles" => IconSetType::Triangles3,
                        "5Boxes" => IconSetType::Boxes5,
                        "4RedToBlack" => IconSetType::RedToBlack4,
                        "4RatingBars" => IconSetType::RatingBars4,
                        "5RatingBars" => IconSetType::RatingBars5,
                        "3ColoredArrows" => IconSetType::ColoredArrows3,
                        "4ColoredArrows" => IconSetType::ColoredArrows4,
                        "5ColoredArrows" => IconSetType::ColoredArrows5,
                        "3WhiteArrows" => IconSetType::WhiteArrows3,
                        "4WhiteArrows" => IconSetType::WhiteArrows4,
                        "5WhiteArrows" => IconSetType::WhiteArrows5,
                        _ => IconSetType::Arrows3,
                    };
                }
                Attribute {
                    key: QName(b"showValue"),
                    value: v,
                } => {
                    icon_set.show_value = &*v != b"0" && &*v != b"false";
                }
                Attribute {
                    key: QName(b"reverse"),
                    value: v,
                } => {
                    icon_set.reverse = &*v == b"1" || &*v == b"true";
                }
                _ => (),
                        }
                    }
                }
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"cfvo" => {
                        let cfvo = Self::parse_cfvo(e.attributes(), xml)?;
                        icon_set.cfvos.push(cfvo);
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"iconSet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("iconSet")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(())
    }

    /// Parse conditional format value object (cfvo)
    fn parse_cfvo(
        attributes: quick_xml::events::attributes::Attributes<'_>,
        xml: &XlReader<'_, RS>,
    ) -> Result<crate::conditional_formatting::ConditionalFormatValue, XlsxError> {
        use crate::conditional_formatting::{CfvoType, ConditionalFormatValue};

        let mut cfvo = ConditionalFormatValue {
            value_type: CfvoType::Min,
            value: None,
            gte: false,
        };

        for attr in attributes {
            match attr.map_err(XlsxError::XmlAttr)? {
                Attribute {
                    key: QName(b"type"),
                    value: v,
                } => {
                    let type_str = xml.decoder().decode(&v)?;
                    cfvo.value_type = match type_str.as_ref() {
                        "min" => CfvoType::Min,
                        "max" => CfvoType::Max,
                        "num" => CfvoType::Number,
                        "percent" => CfvoType::Percent,
                        "percentile" => CfvoType::Percentile,
                        "formula" => CfvoType::Formula,
                        "autoMin" => CfvoType::AutoMin,
                        "autoMax" => CfvoType::AutoMax,
                        _ => CfvoType::Number,
                    };
                }
                Attribute {
                    key: QName(b"val"),
                    value: v,
                } => {
                    cfvo.value = Some(xml.decoder().decode(&v)?.into_owned());
                }
                Attribute {
                    key: QName(b"gte"),
                    value: v,
                } => {
                    cfvo.gte = &*v == b"1" || &*v == b"true";
                }
                _ => (),
            }
        }

        Ok(cfvo)
    }

    /// Parse a dxf (differential format) element
    fn parse_dxf_element(
        xml: &mut XlReader<'_, RS>,
        buf: &mut Vec<u8>,
        _number_formats: &BTreeMap<u32, String>,
        _format_interner: &FormatStringInterner,
    ) -> Result<DifferentialFormat, XlsxError> {
        use crate::conditional_formatting::{
            DifferentialAlignment, DifferentialBorder, DifferentialBorderSide, DifferentialFill,
            DifferentialFont, DifferentialFormat, DifferentialNumberFormat, PatternFill,
        };

        let mut dxf = DifferentialFormat::default();

        loop {
            buf.clear();
            match xml.read_event_into(buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"font" => {
                        let mut font = DifferentialFont::default();
                        let mut inner_buf = Vec::new();

                        loop {
                            inner_buf.clear();
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                                    b"name" => {
                                        if let Some(val) = get_attribute(e.attributes(), QName(b"val"))? {
                                            font.name = Some(xml.decoder().decode(val)?.into_owned());
                                        }
                                    }
                                    b"sz" => {
                                        if let Some(val) = get_attribute(e.attributes(), QName(b"val"))? {
                                            if let Ok(size) = xml.decoder().decode(val)?.parse::<f64>() {
                                                font.size = Some(size);
                                            }
                                        }
                                    }
                                    b"b" => font.bold = Some(true),
                                    b"i" => font.italic = Some(true),
                                    b"u" => font.underline = Some(true),
                                    b"strike" => font.strike = Some(true),
                                    b"color" => {
                                        font.color = Self::parse_color_from_attributes(e.attributes())?;
                                    }
                                    _ => {
                                        let mut temp_buf = Vec::new();
                                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                                    }
                                },
                                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
                                Ok(Event::Eof) => return Err(XlsxError::XmlEof("font")),
                                Err(e) => return Err(XlsxError::Xml(e)),
                                _ => (),
                            }
                        }
                        dxf.font = Some(font);
                    }
                    b"fill" => {
                        let mut pattern_fill = PatternFill {
                            pattern_type: None,
                            fg_color: None,
                            bg_color: None,
                        };
                        let mut inner_buf = Vec::new();

                        loop {
                            inner_buf.clear();
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"patternFill" => {
                                    for attr in e.attributes() {
                                        if let Attribute {
                                            key: QName(b"patternType"),
                                            value: v,
                                        } = attr.map_err(XlsxError::XmlAttr)?
                                        {
                                            pattern_fill.pattern_type = Some(xml.decoder().decode(&v)?.into_owned());
                                        }
                                    }

                                    let mut pattern_buf = Vec::new();
                                    loop {
                                        pattern_buf.clear();
                                        match xml.read_event_into(&mut pattern_buf) {
                                            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                                                b"fgColor" => {
                                                    pattern_fill.fg_color = Self::parse_color_from_attributes(e.attributes())?;
                                                }
                                                b"bgColor" => {
                                                    pattern_fill.bg_color = Self::parse_color_from_attributes(e.attributes())?;
                                                }
                                                _ => {
                                                    let mut temp_buf = Vec::new();
                                                    xml.read_to_end_into(e.name(), &mut temp_buf)?;
                                                }
                                            },
                                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"patternFill" => break,
                                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("patternFill")),
                                            Err(e) => return Err(XlsxError::Xml(e)),
                                            _ => (),
                                        }
                                    }
                                }
                                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
                                Ok(Event::Eof) => return Err(XlsxError::XmlEof("fill")),
                                Err(e) => return Err(XlsxError::Xml(e)),
                                _ => (),
                            }
                        }
                        dxf.fill = Some(DifferentialFill { pattern_fill });
                    }
                    b"border" => {
                        let mut border = DifferentialBorder::default();
                        let mut inner_buf = Vec::new();

                        // Parse border attributes
                        for attr in e.attributes() {
                            match attr.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"diagonalUp"),
                                    value: v,
                                } => {
                                    border.diagonal_up = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"diagonalDown"),
                                    value: v,
                                } => {
                                    border.diagonal_down = Some(&*v == b"1" || &*v == b"true");
                                }
                                _ => (),
                            }
                        }

                        loop {
                            inner_buf.clear();
                            match xml.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref e)) => {
                                    let side_name = e.local_name();
                                    let side = match side_name.as_ref() {
                                        b"left" => &mut border.left,
                                        b"right" => &mut border.right,
                                        b"top" => &mut border.top,
                                        b"bottom" => &mut border.bottom,
                                        b"diagonal" => &mut border.diagonal,
                                        _ => {
                                            let mut temp_buf = Vec::new();
                                            xml.read_to_end_into(e.name(), &mut temp_buf)?;
                                            continue;
                                        }
                                    };

                                    let mut border_side = DifferentialBorderSide {
                                        style: None,
                                        color: None,
                                    };

                                    // Parse style attribute
                                    for attr in e.attributes() {
                                        if let Attribute {
                                            key: QName(b"style"),
                                            value: v,
                                        } = attr.map_err(XlsxError::XmlAttr)?
                                        {
                                            border_side.style = Some(xml.decoder().decode(&v)?.into_owned());
                                        }
                                    }

                                    // Parse color element
                                    let mut side_buf = Vec::new();
                                    loop {
                                        side_buf.clear();
                                        match xml.read_event_into(&mut side_buf) {
                                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                                                border_side.color = Self::parse_color_from_attributes(e.attributes())?;
                                            }
                                            Ok(Event::End(ref e)) if e.local_name() == side_name => break,
                                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("border side")),
                                            Err(e) => return Err(XlsxError::Xml(e)),
                                            _ => (),
                                        }
                                    }

                                    *side = Some(border_side);
                                }
                                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"border" => break,
                                Ok(Event::Eof) => return Err(XlsxError::XmlEof("border")),
                                Err(e) => return Err(XlsxError::Xml(e)),
                                _ => (),
                            }
                        }
                        dxf.border = Some(border);
                    }
                    b"numFmt" => {
                        let mut format_code = String::new();
                        for attr in e.attributes() {
                            if let Attribute {
                                key: QName(b"formatCode"),
                                value: v,
                            } = attr.map_err(XlsxError::XmlAttr)?
                            {
                                format_code = xml.decoder().decode(&v)?.into_owned();
                            }
                        }
                        if !format_code.is_empty() {
                            dxf.number_format = Some(DifferentialNumberFormat { 
                                format_code,
                                num_fmt_id: None,
                            });
                        }
                    }
                    b"alignment" => {
                        let mut alignment = DifferentialAlignment::default();
                        for attr in e.attributes() {
                            match attr.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"horizontal"),
                                    value: v,
                                } => {
                                    alignment.horizontal = Some(xml.decoder().decode(&v)?.into_owned());
                                }
                                Attribute {
                                    key: QName(b"vertical"),
                                    value: v,
                                } => {
                                    alignment.vertical = Some(xml.decoder().decode(&v)?.into_owned());
                                }
                                Attribute {
                                    key: QName(b"wrapText"),
                                    value: v,
                                } => {
                                    alignment.wrap_text = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"shrinkToFit"),
                                    value: v,
                                } => {
                                    alignment.shrink_to_fit = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"textRotation"),
                                    value: v,
                                } => {
                                    if let Ok(rotation) = xml.decoder().decode(&v)?.parse::<i32>() {
                                        alignment.text_rotation = Some(rotation);
                                    }
                                }
                                Attribute {
                                    key: QName(b"indent"),
                                    value: v,
                                } => {
                                    if let Ok(indent) = xml.decoder().decode(&v)?.parse::<u32>() {
                                        alignment.indent = Some(indent);
                                    }
                                }
                                _ => (),
                            }
                        }
                        dxf.alignment = Some(alignment);
                    }
                    _ => {
                        let mut temp_buf = Vec::new();
                        xml.read_to_end_into(e.name(), &mut temp_buf)?;
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxf" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("dxf")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }

        Ok(dxf)
    }

    fn read_workbook(
        &mut self,
        relationships: &BTreeMap<Vec<u8>, String>,
    ) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/workbook.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut defined_names = Vec::new();
        let mut buf = Vec::with_capacity(1024);
        let mut val_buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheet" => {
                    let mut name = String::new();
                    let mut path = String::new();
                    let mut visible = SheetVisible::Visible;
                    for a in e.attributes() {
                        let a = a.map_err(XlsxError::XmlAttr)?;
                        match a {
                            Attribute {
                                key: QName(b"name"),
                                ..
                            } => {
                                name = a.decode_and_unescape_value(xml.decoder())?.to_string();
                            }
                            Attribute {
                                key: QName(b"state"),
                                ..
                            } => {
                                visible = match a.decode_and_unescape_value(xml.decoder())?.as_ref()
                                {
                                    "visible" => SheetVisible::Visible,
                                    "hidden" => SheetVisible::Hidden,
                                    "veryHidden" => SheetVisible::VeryHidden,
                                    v => {
                                        return Err(XlsxError::Unrecognized {
                                            typ: "sheet:state",
                                            val: v.to_string(),
                                        })
                                    }
                                }
                            }
                            Attribute {
                                key: QName(b"r:id"),
                                value: v,
                            }
                            | Attribute {
                                key: QName(b"relationships:id"),
                                value: v,
                            } => {
                                let r = &relationships
                                    .get(&*v)
                                    .ok_or(XlsxError::RelationshipNotFound)?[..];
                                // target may have pre-prended "/xl/" or "xl/" path;
                                // strip if present
                                path = if r.starts_with("/xl/") {
                                    r[1..].to_string()
                                } else if r.starts_with("xl/") {
                                    r.to_string()
                                } else {
                                    format!("xl/{r}")
                                };
                            }
                            _ => (),
                        }
                    }
                    let typ = match path.split('/').nth(1) {
                        Some("worksheets") => SheetType::WorkSheet,
                        Some("chartsheets") => SheetType::ChartSheet,
                        Some("dialogsheets") => SheetType::DialogSheet,
                        _ => {
                            return Err(XlsxError::Unrecognized {
                                typ: "sheet:type",
                                val: path.to_string(),
                            })
                        }
                    };
                    self.metadata.sheets.push(Sheet {
                        name: name.to_string(),
                        typ,
                        visible,
                    });
                    self.sheets.push((name, path));
                }
                Ok(Event::Start(ref e)) if e.name().as_ref() == b"workbookPr" => {
                    self.is_1904 = match e.try_get_attribute("date1904")? {
                        Some(c) => ["1", "true"].contains(
                            &c.decode_and_unescape_value(xml.decoder())
                                .map_err(XlsxError::Xml)?
                                .as_ref(),
                        ),
                        None => false,
                    };
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"definedName" => {
                    if let Some(a) = e
                        .attributes()
                        .filter_map(std::result::Result::ok)
                        .find(|a| a.key == QName(b"name"))
                    {
                        let name = a.decode_and_unescape_value(xml.decoder())?.to_string();
                        val_buf.clear();
                        let mut value = String::new();
                        loop {
                            match xml.read_event_into(&mut val_buf)? {
                                Event::Text(t) => value.push_str(&t.unescape()?),
                                Event::End(end) if end.name() == e.name() => break,
                                Event::Eof => return Err(XlsxError::XmlEof("workbook")),
                                _ => (),
                            }
                        }
                        defined_names.push((name, value));
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"workbook" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("workbook")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        self.metadata.names = defined_names;
        Ok(())
    }

    fn read_relationships(&mut self) -> Result<BTreeMap<Vec<u8>, String>, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            None => {
                return Err(XlsxError::FileNotFound(
                    "xl/_rels/workbook.xml.rels".to_string(),
                ));
            }
            Some(x) => x?,
        };
        let mut relationships = BTreeMap::new();
        let mut buf = Vec::with_capacity(64);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match a.map_err(XlsxError::XmlAttr)? {
                            Attribute {
                                key: QName(b"Id"),
                                value: v,
                            } => id.extend_from_slice(&v),
                            Attribute {
                                key: QName(b"Target"),
                                value: v,
                            } => target = xml.decoder().decode(&v)?.into_owned(),
                            _ => (),
                        }
                    }
                    relationships.insert(id, target);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"Relationships" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("Relationships")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(relationships)
    }

    // sheets must be added before this is called!!
    fn read_table_metadata(&mut self) -> Result<(), XlsxError> {
        let mut new_tables = Vec::new();
        for (sheet_name, sheet_path) in &self.sheets {
            let last_folder_index = sheet_path.rfind('/').expect("should be in a folder");
            let (base_folder, file_name) = sheet_path.split_at(last_folder_index);
            let rel_path = format!("{base_folder}/_rels{file_name}.rels");

            let mut table_locations = Vec::new();
            let mut buf = Vec::with_capacity(64);
            // we need another mutable borrow of self.zip later so we enclose this borrow within braces
            {
                let mut xml = match xml_reader(&mut self.zip, &rel_path) {
                    None => continue,
                    Some(x) => x?,
                };
                loop {
                    buf.clear();
                    match xml.read_event_into(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                            let mut id = Vec::new();
                            let mut target = String::new();
                            let mut table_type = false;
                            for a in e.attributes() {
                                match a.map_err(XlsxError::XmlAttr)? {
                                    Attribute {
                                        key: QName(b"Id"),
                                        value: v,
                                    } => id.extend_from_slice(&v),
                                    Attribute {
                                        key: QName(b"Target"),
                                        value: v,
                                    } => target = xml.decoder().decode(&v)?.into_owned(),
                                    Attribute {
                                        key: QName(b"Type"),
                                        value: v,
                                    } => table_type = *v == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"[..],
                                    _ => (),
                                }
                            }
                            if table_type {
                                if target.starts_with("../") {
                                    // this is an incomplete implementation, but should be good enough for excel
                                    let new_index =
                                        base_folder.rfind('/').expect("Must be a parent folder");
                                    let full_path =
                                        format!("{}{}", &base_folder[..new_index], &target[2..]);
                                    table_locations.push(full_path);
                                } else if target.is_empty() { // do nothing
                                } else {
                                    table_locations.push(target);
                                }
                            }
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"Relationships" => {
                            break
                        }
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("Relationships")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                }
            }
            for table_file in table_locations {
                let mut xml = match xml_reader(&mut self.zip, &table_file) {
                    None => continue,
                    Some(x) => x?,
                };
                let mut column_names = Vec::new();
                let mut table_meta = InnerTableMetadata::new();
                loop {
                    buf.clear();
                    match xml.read_event_into(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"table" => {
                            for a in e.attributes() {
                                match a.map_err(XlsxError::XmlAttr)? {
                                    Attribute {
                                        key: QName(b"displayName"),
                                        value: v,
                                    } => {
                                        table_meta.display_name =
                                            xml.decoder().decode(&v)?.into_owned();
                                    }
                                    Attribute {
                                        key: QName(b"ref"),
                                        value: v,
                                    } => {
                                        table_meta.ref_cells =
                                            xml.decoder().decode(&v)?.into_owned();
                                    }
                                    Attribute {
                                        key: QName(b"headerRowCount"),
                                        value: v,
                                    } => {
                                        table_meta.header_row_count =
                                            xml.decoder().decode(&v)?.parse()?;
                                    }
                                    Attribute {
                                        key: QName(b"insertRow"),
                                        value: v,
                                    } => table_meta.insert_row = *v != b"0"[..],
                                    Attribute {
                                        key: QName(b"totalsRowCount"),
                                        value: v,
                                    } => {
                                        table_meta.totals_row_count =
                                            xml.decoder().decode(&v)?.parse()?;
                                    }
                                    _ => (),
                                }
                            }
                        }
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableColumn" => {
                            for a in e.attributes().flatten() {
                                if let Attribute {
                                    key: QName(b"name"),
                                    value: v,
                                } = a
                                {
                                    column_names.push(xml.decoder().decode(&v)?.into_owned());
                                }
                            }
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"table" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("Table")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                }
                let mut dims = get_dimension(table_meta.ref_cells.as_bytes())?;
                if table_meta.header_row_count != 0 {
                    dims.start.0 += table_meta.header_row_count;
                }
                if table_meta.totals_row_count != 0 {
                    dims.end.0 -= table_meta.header_row_count;
                }
                if table_meta.insert_row {
                    dims.end.0 -= 1;
                }
                new_tables.push((
                    table_meta.display_name,
                    sheet_name.clone(),
                    column_names,
                    dims,
                ));
            }
        }
        self.tables = Some(new_tables);
        Ok(())
    }

    /// Read pictures
    #[cfg(feature = "picture")]
    fn read_pictures(&mut self) -> Result<(), XlsxError> {
        let mut pics = Vec::new();
        for i in 0..self.zip.len() {
            let mut zfile = self.zip.by_index(i)?;
            let zname = zfile.name();
            if zname.starts_with("xl/media") {
                if let Some(ext) = zname.split('.').next_back() {
                    if [
                        "emf", "wmf", "pict", "jpeg", "jpg", "png", "dib", "gif", "tiff", "eps",
                        "bmp", "wpg",
                    ]
                    .contains(&ext)
                    {
                        let ext = ext.to_string();
                        let mut buf: Vec<u8> = Vec::new();
                        zfile.read_to_end(&mut buf)?;
                        pics.push((ext, buf));
                    }
                }
            }
        }
        if !pics.is_empty() {
            self.pictures = Some(pics);
        }
        Ok(())
    }

    // sheets must be added before this is called!!
    fn read_merged_regions(&mut self) -> Result<(), XlsxError> {
        let mut regions = Vec::new();
        for (sheet_name, sheet_path) in &self.sheets {
            // we need another mutable borrow of self.zip later so we enclose this borrow within braces
            {
                let mut xml = match xml_reader(&mut self.zip, sheet_path) {
                    None => continue,
                    Some(x) => x?,
                };
                let mut buf = Vec::new();
                loop {
                    buf.clear();
                    match xml.read_event_into(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name() == QName(b"mergeCell").into() => {
                            if let Some(attr) = get_attribute(e.attributes(), QName(b"ref"))? {
                                let dismension = get_dimension(attr)?;
                                regions.push((
                                    sheet_name.to_string(),
                                    sheet_path.to_string(),
                                    dismension,
                                ));
                            }
                        }
                        Ok(Event::Eof) => break,
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                }
            }
        }
        self.merged_regions = Some(regions);
        Ok(())
    }

    #[inline]
    fn get_table_meta(&self, table_name: &str) -> Result<TableMetadata, XlsxError> {
        let match_table_meta = self
            .tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .find(|(table, ..)| table == table_name)
            .ok_or_else(|| XlsxError::TableNotFound(table_name.into()))?;

        let name = match_table_meta.0.to_owned();
        let sheet_name = match_table_meta.1.clone();
        let columns = match_table_meta.2.clone();
        let dimensions = Dimensions {
            start: match_table_meta.3.start,
            end: match_table_meta.3.end,
        };

        Ok(TableMetadata {
            name,
            sheet_name,
            columns,
            dimensions,
        })
    }

    /// Get comprehensive formatting information for a cell by its style index
    pub fn get_cell_formatting(&self, style_index: usize) -> Option<&CellStyle> {
        self.styles.get(style_index)
    }

    /// Get all available cell formats
    pub fn get_all_cell_formats(&self) -> &[CellStyle] {
        &self.styles
    }

    /// Get access to the format string interner for reuse across sheets
    /// The interner is thread-safe and can be shared across threads
    pub fn get_format_interner(&self) -> &FormatStringInterner {
        &self.format_interner
    }

    /// Load the merged regions
    pub fn load_merged_regions(&mut self) -> Result<(), XlsxError> {
        if self.merged_regions.is_none() {
            self.read_merged_regions()
        } else {
            Ok(())
        }
    }

    /// Get the merged regions of all the sheets
    pub fn merged_regions(&self) -> &Vec<(String, String, Dimensions)> {
        self.merged_regions
            .as_ref()
            .expect("Merged Regions must be loaded before the are referenced")
    }

    /// Get the merged regions by sheet name
    pub fn merged_regions_by_sheet(&self, name: &str) -> Vec<(&String, &String, &Dimensions)> {
        self.merged_regions()
            .iter()
            .filter(|s| s.0 == name)
            .map(|(name, sheet, region)| (name, sheet, region))
            .collect()
    }

    /// Load the tables from
    pub fn load_tables(&mut self) -> Result<(), XlsxError> {
        if self.tables.is_none() {
            self.read_table_metadata()
        } else {
            Ok(())
        }
    }

    /// Get the names of all the tables
    pub fn table_names(&self) -> Vec<&String> {
        self.tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .map(|(name, ..)| name)
            .collect()
    }

    /// Get the names of all the tables in a sheet
    pub fn table_names_in_sheet(&self, sheet_name: &str) -> Vec<&String> {
        self.tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .filter(|(_, sheet, ..)| sheet == sheet_name)
            .map(|(name, ..)| name)
            .collect()
    }

    /// Get the table by name (owned)
    // TODO: If retrieving multiple tables from a single sheet, get tables by sheet will be more efficient
    pub fn table_by_name(&mut self, table_name: &str) -> Result<Table<DataWithFormatting>, XlsxError> {
        let TableMetadata {
            name,
            sheet_name,
            columns,
            dimensions,
        } = self.get_table_meta(table_name)?;
        let Dimensions { start, end } = dimensions;
        let range = self.worksheet_range(&sheet_name)?;
        let tbl_rng = range.range(start, end);

        Ok(Table {
            name,
            sheet_name,
            columns,
            data: tbl_rng,
        })
    }

    /// Get the table by name (ref)
    pub fn table_by_name_ref(&mut self, table_name: &str) -> Result<Table<DataRef<'_>>, XlsxError> {
        let TableMetadata {
            name,
            sheet_name,
            columns,
            dimensions,
        } = self.get_table_meta(table_name)?;
        let Dimensions { start, end } = dimensions;
        let range = self.worksheet_range_ref(&sheet_name)?;
        let tbl_rng = range.range(start, end);

        Ok(Table {
            name,
            sheet_name,
            columns,
            data: tbl_rng,
        })
    }

    /// Gets the worksheet merge cell dimensions
    pub fn worksheet_merge_cells(
        &mut self,
        name: &str,
    ) -> Option<Result<Vec<Dimensions>, XlsxError>> {
        let (_, path) = self.sheets.iter().find(|(n, _)| n == name)?;
        let xml = xml_reader(&mut self.zip, path);

        xml.map(|xml| {
            let mut xml = xml?;
            let mut merge_cells = Vec::new();
            let mut buffer = Vec::new();

            loop {
                buffer.clear();

                match xml.read_event_into(&mut buffer) {
                    Ok(Event::Start(event)) if event.local_name().as_ref() == b"mergeCells" => {
                        if let Ok(cells) = read_merge_cells(&mut xml) {
                            merge_cells = cells;
                        }

                        break;
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => (),
                }
            }

            Ok(merge_cells)
        })
    }

    /// Get the nth worksheet. Shortcut for getting the nth
    /// sheet name, then the corresponding worksheet.
    pub fn worksheet_merge_cells_at(
        &mut self,
        n: usize,
    ) -> Option<Result<Vec<Dimensions>, XlsxError>> {
        let name = self
            .metadata()
            .sheets
            .get(n)
            .map(|sheet| sheet.name.clone())?;

        self.worksheet_merge_cells(&name)
    }

    /// Get a cell reader for the worksheet (with comprehensive formatting)
    pub fn worksheet_cells_reader_ext(
        &mut self,
        name: &str,
    ) -> Result<XlsxCellReader<'_, RS>, XlsxError> {
        let xml = xml_reader(&mut self.zip, &format!("xl/worksheets/{}.xml", name))
            .ok_or_else(|| XlsxError::FileNotFound(format!("xl/worksheets/{}.xml", name)))??;
        let is_1904 = self.is_1904;
        let strings = &self.strings;
        let formats = &self.styles;
        XlsxCellReader::new(xml, strings, formats, is_1904)
    }
}

struct TableMetadata {
    name: String,
    sheet_name: String,
    columns: Vec<String>,
    dimensions: Dimensions,
}

struct InnerTableMetadata {
    display_name: String,
    ref_cells: String,
    header_row_count: u32,
    insert_row: bool,
    totals_row_count: u32,
}

impl InnerTableMetadata {
    fn new() -> Self {
        Self {
            display_name: String::new(),
            ref_cells: String::new(),
            header_row_count: 1,
            insert_row: false,
            totals_row_count: 0,
        }
    }
}

impl<RS: Read + Seek> Xlsx<RS> {
    /// Get a reader over all used cells in the given worksheet cell reader
    pub fn worksheet_cells_reader<'a>(
        &'a mut self,
        name: &str,
    ) -> Result<XlsxCellReader<'a, RS>, XlsxError> {
        let (_, path) = self
            .sheets
            .iter()
            .find(|&(n, _)| n == name)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))?;
        let xml = xml_reader(&mut self.zip, path)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))??;
        let is_1904 = self.is_1904;
        let strings = &self.strings;
        let formats = &self.styles;
        XlsxCellReader::new(xml, strings, formats, is_1904)
    }
}

impl<RS: Read + Seek> Reader<RS> for Xlsx<RS> {
    type Error = XlsxError;

    fn new(mut reader: RS) -> Result<Self, XlsxError> {
        check_for_password_protected(&mut reader)?;

        let mut xlsx = Xlsx {
            zip: ZipArchive::new(reader)?,
            strings: Vec::new(),
            formats: Vec::new(),
            styles: Vec::new(),
            format_interner: FormatStringInterner::new(),
            is_1904: false,
            sheets: Vec::new(),
            tables: None,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
            merged_regions: None,
            options: XlsxOptions::default(),
            dxf_formats: Vec::new(),
            conditional_formats: BTreeMap::new(),
        };
        xlsx.read_shared_strings()?;
        xlsx.read_styles()?;
        let relationships = xlsx.read_relationships()?;
        xlsx.read_workbook(&relationships)?;
        #[cfg(feature = "picture")]
        xlsx.read_pictures()?;

        Ok(xlsx)
    }

    fn with_header_row(&mut self, header_row: HeaderRow) -> &mut Self {
        self.options.header_row = header_row;
        self
    }

    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, XlsxError>> {
        let mut f = self.zip.by_name("xl/vbaProject.bin").ok()?;
        let len = f.size() as usize;
        Some(
            VbaProject::new(&mut f, len)
                .map(Cow::Owned)
                .map_err(XlsxError::Vba),
        )
    }

    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    fn worksheet_range(&mut self, name: &str) -> Result<Range<DataWithFormatting>, XlsxError> {
        let header_row = self.options.header_row;
        let mut cell_reader = match self.worksheet_cells_reader(name) {
            Ok(reader) => reader,
            Err(XlsxError::NotAWorksheet(typ)) => {
                log::warn!("'{typ}' not a valid worksheet");
                return Ok(Range::default());
            }
            Err(e) => return Err(e),
        };
        let len = cell_reader.dimensions().len();
        let mut cells = Vec::new();
        if len < 100_000 {
            cells.reserve(len as usize);
        }

        match header_row {
            HeaderRow::FirstNonEmptyRow => {
                // the header row is the row of the first non-empty cell
                while let Some((cell, formatting)) = cell_reader.next_cell_with_formatting()? {
                    if matches!(cell.val, DataRef::Empty) {
                        continue;
                    }
                    let data_with_formatting = DataWithFormatting::new(
                        cell.val.into(),
                        formatting.cloned(),
                    );
                    cells.push(Cell::new(cell.pos, data_with_formatting));
                }
            }
            HeaderRow::Row(header_row_idx) => {
                // If `header_row` is a row index, we only add non-empty cells after this index.
                while let Some((cell, formatting)) = cell_reader.next_cell_with_formatting()? {
                    if matches!(cell.val, DataRef::Empty) {
                        continue;
                    }
                    if cell.pos.0 >= header_row_idx {
                        let data_with_formatting = DataWithFormatting::new(
                            cell.val.into(),
                            formatting.cloned(),
                        );
                        cells.push(Cell::new(cell.pos, data_with_formatting));
                    }
                }

                // If `header_row` is set and the first non-empty cell is not at the `header_row`, we add
                // an empty cell at the beginning with row `header_row` and same column as the first non-empty cell.
                if cells.first().is_some_and(|c| c.pos.0 != header_row_idx) {
                    cells.insert(
                        0,
                        Cell {
                            pos: (
                                header_row_idx,
                                cells.first().expect("cells should not be empty").pos.1,
                            ),
                            val: DataWithFormatting::default(),
                        },
                    );
                }
            }
        }

        Ok(Range::from_sparse(cells))
    }

    fn worksheet_formula(&mut self, name: &str) -> Result<Range<DataWithFormatting>, XlsxError> {
        let mut cell_reader = match self.worksheet_cells_reader(name) {
            Ok(reader) => reader,
            Err(XlsxError::NotAWorksheet(typ)) => {
                warn!("'{typ}' not a worksheet");
                return Ok(Range::default());
            }
            Err(e) => return Err(e),
        };
        let len = cell_reader.dimensions().len();
        let mut cells = Vec::new();
        if len < 100_000 {
            cells.reserve(len as usize);
        }
        while let Some((cell, formatting)) = cell_reader.next_formula_with_formatting()? {
            if !cell.val.is_empty() {
                let data_with_formatting = DataWithFormatting::new(
                    Data::String(cell.val),
                    formatting.cloned(),
                );
                cells.push(Cell::new(cell.pos, data_with_formatting));
            }
        }
        Ok(Range::from_sparse(cells))
    }

    fn worksheets(&mut self) -> Vec<(String, Range<DataWithFormatting>)> {
        let names = self
            .sheets
            .iter()
            .map(|(n, _)| n.clone())
            .collect::<Vec<_>>();
        names
            .into_iter()
            .filter_map(|n| {
                let rge = self.worksheet_range(&n).ok()?;
                Some((n, rge))
            })
            .collect()
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        self.pictures.to_owned()
    }
}

impl<RS: Read + Seek> ReaderRef<RS> for Xlsx<RS> {
    fn worksheet_range_ref<'a>(&'a mut self, name: &str) -> Result<Range<DataRef<'a>>, XlsxError> {
        let header_row = self.options.header_row;
        let mut cell_reader = match self.worksheet_cells_reader(name) {
            Ok(reader) => reader,
            Err(XlsxError::NotAWorksheet(typ)) => {
                log::warn!("'{typ}' not a valid worksheet");
                return Ok(Range::default());
            }
            Err(e) => return Err(e),
        };
        let len = cell_reader.dimensions().len();
        let mut cells = Vec::new();
        if len < 100_000 {
            cells.reserve(len as usize);
        }

        match header_row {
            HeaderRow::FirstNonEmptyRow => {
                // the header row is the row of the first non-empty cell
                loop {
                    match cell_reader.next_cell() {
                        Ok(Some(Cell {
                            val: DataRef::Empty,
                            ..
                        })) => (),
                        Ok(Some(cell)) => cells.push(cell),
                        Ok(None) => break,
                        Err(e) => return Err(e),
                    }
                }
            }
            HeaderRow::Row(header_row_idx) => {
                // If `header_row` is a row index, we only add non-empty cells after this index.
                loop {
                    match cell_reader.next_cell() {
                        Ok(Some(Cell {
                            val: DataRef::Empty,
                            ..
                        })) => (),
                        Ok(Some(cell)) => {
                            if cell.pos.0 >= header_row_idx {
                                cells.push(cell);
                            }
                        }
                        Ok(None) => break,
                        Err(e) => return Err(e),
                    }
                }

                // If `header_row` is set and the first non-empty cell is not at the `header_row`, we add
                // an empty cell at the beginning with row `header_row` and same column as the first non-empty cell.
                if cells.first().is_some_and(|c| c.pos.0 != header_row_idx) {
                    cells.insert(
                        0,
                        Cell {
                            pos: (
                                header_row_idx,
                                cells.first().expect("cells should not be empty").pos.1,
                            ),
                            val: DataRef::Empty,
                        },
                    );
                }
            }
        }

        Ok(Range::from_sparse(cells))
    }
}

fn xml_reader<'a, RS: Read + Seek>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
) -> Option<Result<XlReader<'a, RS>, XlsxError>> {
    let actual_path = zip
        .file_names()
        .find(|n| n.eq_ignore_ascii_case(path))?
        .to_owned();
    match zip.by_name(&actual_path) {
        Ok(f) => {
            let mut r = XmlReader::from_reader(BufReader::new(f));
            let config = r.config_mut();
            config.check_end_names = false;
            config.trim_text(false);
            config.check_comments = false;
            config.expand_empty_elements = true;
            Some(Ok(r))
        }
        Err(ZipError::FileNotFound) => None,
        Err(e) => Some(Err(e.into())),
    }
}

/// search through an Element's attributes for the named one
pub(crate) fn get_attribute<'a>(
    atts: Attributes<'a>,
    n: QName,
) -> Result<Option<&'a [u8]>, XlsxError> {
    for a in atts {
        match a {
            Ok(Attribute {
                key,
                value: Cow::Borrowed(value),
            }) if key == n => return Ok(Some(value)),
            Err(e) => return Err(XlsxError::XmlAttr(e)),
            _ => {} // ignore other attributes
        }
    }
    Ok(None)
}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column),
/// - bottom right (row, column)
pub(crate) fn get_dimension(dimension: &[u8]) -> Result<Dimensions, XlsxError> {
    let parts: Vec<_> = dimension
        .split(|c| *c == b':')
        .map(get_row_column)
        .collect::<Result<Vec<_>, XlsxError>>()?;

    match parts.len() {
        0 => Err(XlsxError::DimensionCount(0)),
        1 => Ok(Dimensions {
            start: parts[0],
            end: parts[0],
        }),
        2 => {
            let rows = parts[1].0 - parts[0].0;
            let columns = parts[1].1 - parts[0].1;
            if rows > MAX_ROWS {
                warn!("xlsx has more than maximum number of rows ({rows} > {MAX_ROWS})");
            }
            if columns > MAX_COLUMNS {
                warn!("xlsx has more than maximum number of columns ({columns} > {MAX_COLUMNS})");
            }
            Ok(Dimensions {
                start: parts[0],
                end: parts[1],
            })
        }
        len => Err(XlsxError::DimensionCount(len)),
    }
}

/// Converts a text range name into its position (row, column) (0 based index).
/// If the row or column component in the range is missing, an Error is returned.
pub(crate) fn get_row_column(range: &[u8]) -> Result<(u32, u32), XlsxError> {
    let (row, col) = get_row_and_optional_column(range)?;
    let col = col.ok_or(XlsxError::RangeWithoutColumnComponent)?;
    Ok((row, col))
}

/// Converts a text row name into its position (0 based index).
/// If the row component in the range is missing, an Error is returned.
/// If the text row name also contains a column component, it is ignored.
pub(crate) fn get_row(range: &[u8]) -> Result<u32, XlsxError> {
    get_row_and_optional_column(range).map(|(row, _)| row)
}

/// Converts a text range name into its position (row, column) (0 based index).
/// If the row component in the range is missing, an Error is returned.
/// If the column component in the range is missing, an None is returned for the column.
fn get_row_and_optional_column(range: &[u8]) -> Result<(u32, Option<u32>), XlsxError> {
    let (mut row, mut col) = (0, 0);
    let mut pow = 1;
    let mut readrow = true;
    for c in range.iter().rev() {
        match *c {
            c @ b'0'..=b'9' => {
                if readrow {
                    row += ((c - b'0') as u32) * pow;
                    pow *= 10;
                } else {
                    return Err(XlsxError::NumericColumn(c));
                }
            }
            c @ b'A'..=b'Z' => {
                if readrow {
                    if row == 0 {
                        return Err(XlsxError::RangeWithoutRowComponent);
                    }
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'A') as u32 + 1) * pow;
                pow *= 26;
            }
            c @ b'a'..=b'z' => {
                if readrow {
                    if row == 0 {
                        return Err(XlsxError::RangeWithoutRowComponent);
                    }
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'a') as u32 + 1) * pow;
                pow *= 26;
            }
            _ => return Err(XlsxError::Alphanumeric(*c)),
        }
    }
    let row = row
        .checked_sub(1)
        .ok_or(XlsxError::RangeWithoutRowComponent)?;
    Ok((row, col.checked_sub(1)))
}

/// attempts to read either a simple or richtext string
pub(crate) fn read_string<RS>(
    xml: &mut XlReader<'_, RS>,
    closing: QName,
) -> Result<Option<String>, XlsxError>
where
    RS: Read + Seek,
{
    let mut buf = Vec::with_capacity(1024);
    let mut val_buf = Vec::with_capacity(1024);
    let mut rich_buffer: Option<String> = None;
    let mut is_phonetic_text = false;
    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                if rich_buffer.is_none() {
                    // use a buffer since richtext has multiples <r> and <t> for the same cell
                    rich_buffer = Some(String::new());
                }
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPh" => {
                is_phonetic_text = true;
            }
            Ok(Event::End(ref e)) if e.name() == closing => {
                return Ok(rich_buffer);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rPh" => {
                is_phonetic_text = false;
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" && !is_phonetic_text => {
                val_buf.clear();
                let mut value = String::new();
                loop {
                    match xml.read_event_into(&mut val_buf)? {
                        Event::Text(t) => value.push_str(&t.unescape()?),
                        Event::End(end) if end.name() == e.name() => break,
                        Event::Eof => return Err(XlsxError::XmlEof("t")),
                        _ => (),
                    }
                }
                if let Some(ref mut s) = rich_buffer {
                    s.push_str(&value);
                } else {
                    // consume any remaining events up to expected closing tag
                    xml.read_to_end_into(closing, &mut val_buf)?;
                    return Ok(Some(value));
                }
            }
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }
}

fn check_for_password_protected<RS: Read + Seek>(reader: &mut RS) -> Result<(), XlsxError> {
    let offset_end = reader.seek(std::io::SeekFrom::End(0))? as usize;
    reader.seek(std::io::SeekFrom::Start(0))?;

    if let Ok(cfb) = crate::cfb::Cfb::new(reader, offset_end) {
        if cfb.has_directory("EncryptedPackage") {
            return Err(XlsxError::Password);
        }
    }

    Ok(())
}

fn read_merge_cells<RS>(xml: &mut XlReader<'_, RS>) -> Result<Vec<Dimensions>, XlsxError>
where
    RS: Read + Seek,
{
    let mut merge_cells = Vec::new();

    loop {
        let mut buffer = Vec::new();

        match xml.read_event_into(&mut buffer) {
            Ok(Event::Start(event)) if event.local_name().as_ref() == b"mergeCell" => {
                for attribute in event.attributes() {
                    let attribute = attribute.map_err(XlsxError::XmlAttr)?;

                    if attribute.key == QName(b"ref") {
                        let dimensions = get_dimension(&attribute.value)?;
                        merge_cells.push(dimensions);

                        break;
                    }
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == b"mergeCells" => {
                break;
            }
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }

    Ok(merge_cells)
}

/// advance the cell name by the offset
fn offset_cell_name(name: &[char], offset: (i64, i64)) -> Result<Vec<u8>, XlsxError> {
    let cell = get_row_column(name.iter().map(|c| *c as u8).collect::<Vec<_>>().as_slice())?;
    coordinate_to_name((
        (cell.0 as i64 + offset.0) as u32,
        (cell.1 as i64 + offset.1) as u32,
    ))
}

/// advance all valid cell names in the string by the offset
fn replace_cell_names(s: &str, offset: (i64, i64)) -> Result<String, XlsxError> {
    let mut res: Vec<u8> = Vec::new();
    let mut cell: Vec<char> = Vec::new();
    let mut is_cell_row = false;
    let mut in_quote = false;
    for c in s.chars() {
        if c == '"' {
            in_quote = !in_quote;
        }
        if in_quote {
            res.push(c as u8);
            continue;
        }
        if c.is_ascii_alphabetic() {
            if is_cell_row {
                // two cell not possible stick togather in formula
                res.extend(cell.iter().map(|c| *c as u8));
                cell.clear();
                is_cell_row = false;
            }
            cell.push(c);
        } else if c.is_ascii_digit() {
            is_cell_row = true;
            cell.push(c);
        } else {
            if let Ok(cell_name) = offset_cell_name(cell.as_ref(), offset) {
                res.extend(cell_name);
            } else {
                res.extend(cell.iter().map(|c| *c as u8));
            }
            cell.clear();
            is_cell_row = false;
            res.push(c as u8);
        }
    }
    if !cell.is_empty() {
        if let Ok(cell_name) = offset_cell_name(cell.as_ref(), offset) {
            res.extend(cell_name);
        } else {
            res.extend(cell.iter().map(|c| *c as u8));
        }
    }
    match String::from_utf8(res) {
        Ok(s) => Ok(s),
        Err(_) => Err(XlsxError::Unexpected("fail to convert cell name")),
    }
}

/// Convert the integer to Excelsheet column title.
/// If the column number not in 1~16384, an Error is returned.
pub(crate) fn column_number_to_name(num: u32) -> Result<Vec<u8>, XlsxError> {
    if num >= MAX_COLUMNS {
        return Err(XlsxError::Unexpected("column number overflow"));
    }
    let mut col: Vec<u8> = Vec::new();
    let mut num = num + 1;
    while num > 0 {
        let integer = ((num - 1) % 26 + 65) as u8;
        col.push(integer);
        num = (num - 1) / 26;
    }
    col.reverse();
    Ok(col)
}

/// Convert a cell coordinate to Excelsheet cell name.
/// If the column number not in 1~16384, an Error is returned.
pub(crate) fn coordinate_to_name(cell: (u32, u32)) -> Result<Vec<u8>, XlsxError> {
    let cell = &[
        column_number_to_name(cell.1)?,
        (cell.0 + 1).to_string().into_bytes(),
    ];
    Ok(cell.concat())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    #[test]
    fn test_dimensions() {
        assert_eq!(get_row_column(b"A1").unwrap(), (0, 0));
        assert_eq!(get_row_column(b"C107").unwrap(), (106, 2));
        assert_eq!(
            get_dimension(b"C2:D35").unwrap(),
            Dimensions {
                start: (1, 2),
                end: (34, 3)
            }
        );
        assert_eq!(
            get_dimension(b"A1:XFD1048576").unwrap(),
            Dimensions {
                start: (0, 0),
                end: (1_048_575, 16_383),
            }
        );
    }

    #[test]
    fn test_dimension_length() {
        assert_eq!(get_dimension(b"A1:Z99").unwrap().len(), 2_574);
        assert_eq!(
            get_dimension(b"A1:XFD1048576").unwrap().len(),
            17_179_869_184
        );
    }

    #[test]
    fn test_parse_error() {
        assert_eq!(
            CellErrorType::from_str("#DIV/0!").unwrap(),
            CellErrorType::Div0
        );
        assert_eq!(CellErrorType::from_str("#N/A").unwrap(), CellErrorType::NA);
        assert_eq!(
            CellErrorType::from_str("#NAME?").unwrap(),
            CellErrorType::Name
        );
        assert_eq!(
            CellErrorType::from_str("#NULL!").unwrap(),
            CellErrorType::Null
        );
        assert_eq!(
            CellErrorType::from_str("#NUM!").unwrap(),
            CellErrorType::Num
        );
        assert_eq!(
            CellErrorType::from_str("#REF!").unwrap(),
            CellErrorType::Ref
        );
        assert_eq!(
            CellErrorType::from_str("#VALUE!").unwrap(),
            CellErrorType::Value
        );
    }

    #[test]
    fn test_column_number_to_name() {
        assert_eq!(column_number_to_name(0).unwrap(), b"A");
        assert_eq!(column_number_to_name(25).unwrap(), b"Z");
        assert_eq!(column_number_to_name(26).unwrap(), b"AA");
        assert_eq!(column_number_to_name(27).unwrap(), b"AB");
        assert_eq!(column_number_to_name(MAX_COLUMNS - 1).unwrap(), b"XFD");
    }

    #[test]
    fn test_coordinate_to_name() {
        assert_eq!(coordinate_to_name((0, 0)).unwrap(), b"A1");
        assert_eq!(
            coordinate_to_name((MAX_ROWS - 1, MAX_COLUMNS - 1)).unwrap(),
            b"XFD1048576"
        );
    }

    #[test]
    fn test_replace_cell_names() {
        assert_eq!(replace_cell_names("A1", (1, 0)).unwrap(), "A2".to_owned());
        assert_eq!(
            replace_cell_names("CONCATENATE(A1, \"a\")", (1, 0)).unwrap(),
            "CONCATENATE(A2, \"a\")".to_owned()
        );
        assert_eq!(
            replace_cell_names(
                "A1 is a cell, B1 is another, also C107, but XFE123 is not and \"A3\" in quote wont change.",
                (1, 0)
            )
            .unwrap(),
            "A2 is a cell, B2 is another, also C108, but XFE123 is not and \"A3\" in quote wont change.".to_owned()
        );
    }

    #[test]
    fn test_read_shared_strings_with_namespaced_si_name() {
        let shared_strings_data = br#"<?xml version="1.0" encoding="utf-8"?>
<x:sst count="1187" uniqueCount="1187" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <x:si>
        <x:t>String 1</x:t>
    </x:si>
    <x:si>
        <x:r>
            <x:rPr>
                <x:sz val="11"/>
            </x:rPr>
            <x:t>String 2</x:t>
        </x:r>
    </x:si>
    <x:si>
        <x:r>
            <x:t>String 3</x:t>
        </x:r>
    </x:si>
</x:sst>"#;

        let mut buf = [0; 1000];
        let mut zip_writer = ZipWriter::new(std::io::Cursor::new(&mut buf[..]));
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        zip_writer
            .start_file("xl/sharedStrings.xml", options)
            .unwrap();
        zip_writer.write_all(shared_strings_data).unwrap();
        let zip_size = zip_writer.finish().unwrap().position() as usize;

        let zip = ZipArchive::new(std::io::Cursor::new(&buf[..zip_size])).unwrap();

        let mut xlsx = Xlsx {
            zip,
            strings: vec![],
            sheets: vec![],
            tables: None,
            formats: vec![],
            styles: vec![],
            format_interner: FormatStringInterner::new(),
            is_1904: false,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
            merged_regions: None,
            options: XlsxOptions::default(),
            dxf_formats: vec![],
            conditional_formats: BTreeMap::new(),
        };

        assert!(xlsx.read_shared_strings().is_ok());
        assert_eq!(3, xlsx.strings.len());
        assert_eq!("String 1", &xlsx.strings[0]);
        assert_eq!("String 2", &xlsx.strings[1]);
        assert_eq!("String 3", &xlsx.strings[2]);
    }
}

#[cfg(test)]
mod comprehensive_formatting_tests {
    use super::*;
    use crate::formats::{
        Alignment, Border, BorderSide, CellFormat, Color, Fill, Font, PatternType,
    };

    #[test]
    fn test_cell_formatting_structure() {
        // Test that we can create and access CellFormatting structures
        let formatting = CellStyle {
            number_format: CellFormat::Other,
            format_string: None,
            font: Some(Arc::new(Font {
                name: Some(Arc::from("Arial")),
                size: Some(12.0),
                bold: Some(true),
                italic: Some(false),
                color: Some(Color::Rgb { r: 255, g: 0, b: 0 }),
            })),
            fill: Some(Arc::new(Fill {
                pattern_type: PatternType::Solid,
                foreground_color: Some(Color::Rgb { r: 0, g: 255, b: 0 }),
                background_color: None,
            })),
            border: Some(Arc::new(Border {
                left: Some(BorderSide {
                    style: Arc::from("thin"),
                    color: Some(Color::Rgb { r: 0, g: 0, b: 255 }),
                }),
                right: None,
                top: None,
                bottom: None,
            })),
            alignment: Some(Arc::new(Alignment {
                horizontal: Some(Arc::from("center")),
                vertical: Some(Arc::from("middle")),
                wrap_text: Some(true),
                indent: Some(1),
                shrink_to_fit: None,
                text_rotation: None,
                reading_order: None,
            })),
        };

        // Verify the formatting was set correctly
        assert_eq!(formatting.number_format, CellFormat::Other);
        assert!(formatting.font.is_some());
        assert!(formatting.fill.is_some());
        assert!(formatting.border.is_some());
        assert!(formatting.alignment.is_some());

        if let Some(font) = &formatting.font {
            assert_eq!(font.name, Some(Arc::from("Arial")));
            assert_eq!(font.size, Some(12.0));
            assert_eq!(font.bold, Some(true));
        }

        if let Some(fill) = &formatting.fill {
            assert_eq!(fill.pattern_type, PatternType::Solid);
        }

        if let Some(border) = &formatting.border {
            assert!(border.left.is_some());
            if let Some(left_border) = &border.left {
                assert_eq!(left_border.style, Arc::from("thin"));
            }
        }

        if let Some(alignment) = &formatting.alignment {
            assert_eq!(alignment.horizontal, Some(Arc::from("center")));
            assert_eq!(alignment.wrap_text, Some(true));
        }
    }
}
