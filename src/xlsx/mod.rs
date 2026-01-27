// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

#![warn(missing_docs)]

mod cells_reader;
mod style_parser;

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::BufReader;
use std::io::{Read, Seek};
use std::str::FromStr;

use log::warn;
use quick_xml::events::attributes::{Attribute, Attributes};
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::datatype::DataRef;
use crate::formats::{builtin_format_by_id, detect_custom_number_format, CellFormat};
use crate::style::{ColumnWidth, RowHeight, WorksheetLayout};
use crate::utils::{unescape_entity_to_buffer, unescape_xml};
use crate::vba::VbaProject;
use crate::{
    Cell, CellErrorType, Data, Dimensions, HeaderRow, Metadata, Range, Reader, ReaderRef, Sheet,
    SheetType, SheetVisible, Style, Table,
};
pub use cells_reader::XlsxCellReader;

pub(crate) type XlReader<'a, RS> = XmlReader<BufReader<ZipFile<'a, RS>>>;

/// Maximum number of rows allowed in an XLSX file.
pub const MAX_ROWS: u32 = 1_048_576;

/// Maximum number of columns allowed in an XLSX file.
pub const MAX_COLUMNS: u32 = 16_384;

/// An enum for Xlsx specific errors.
#[derive(Debug)]
pub enum XlsxError {
    /// A wrapper for a variety of [`std::io::Error`] errors such as file
    /// permissions when reading an XLSX file. This can be caused by a
    /// non-existent file or parent directory or, commonly on Windows, if the
    /// file is already open in Excel.
    Io(std::io::Error),

    /// A wrapper for a variety of [`zip::result::ZipError`] errors from
    /// [`zip::ZipWriter`]. These relate to errors arising from reading the XLSX
    /// file zip container.
    Zip(zip::result::ZipError),

    /// A general error when reading a VBA project from an XLSX file.
    Vba(crate::vba::VbaError),

    /// A wrapper for a variety of [`quick_xml::Error`] XML parsing errors, but
    /// most commonly for missing data in the target file.
    Xml(quick_xml::Error),

    /// A wrapper for a variety of [`quick_xml::events::attributes::AttrError`]
    /// errors related to attributes in XML elements.
    XmlAttr(quick_xml::events::attributes::AttrError),

    /// A wrapper for a variety of [`std::string::ParseError`] errors when
    /// parsing strings.
    Parse(std::string::ParseError),

    /// A wrapper for a variety of [`std::num::ParseFloatError`] errors when
    /// parsing floats.
    ParseFloat(std::num::ParseFloatError),

    /// A wrapper for a variety of [`std::num::ParseIntError`] errors when
    /// parsing integers.
    ParseInt(std::num::ParseIntError),

    /// Unexpected end of XML file, usually when an end tag is missing.
    XmlEof(&'static str),

    /// Unexpected node in XML.
    UnexpectedNode(&'static str),

    /// XML file not found in XLSX container.
    FileNotFound(String),

    /// Relationship file not found in XLSX container.
    RelationshipNotFound,

    /// Non alphanumeric character found when parsing `A1` style range string.
    Alphanumeric(u8),

    /// Error when parsing the column name in a `A1` style range string.
    NumericColumn(u8),

    /// Missing column name when parsing an `A1` style range string.
    RangeWithoutColumnComponent,

    /// Missing row number when parsing an `A1` style range string.
    RangeWithoutRowComponent,

    /// Column number exceeds maximum allowed columns.
    ColumnNumberOverflow,

    /// Row number exceeds maximum allowed rows.
    RowNumberOverflow,

    /// Error when parsing dimensions of a worksheet.
    DimensionCount(usize),

    /// Unknown cell type (`t`) attribute.
    CellTAttribute(String),

    /// Unexpected XML element or attribute error.
    Unexpected(&'static str),

    /// Unrecognized worksheet type or state.
    Unrecognized {
        /// The data type.
        typ: &'static str,

        /// The value found.
        val: String,
    },

    /// Unrecognized cell error type.
    CellError(String),

    /// Workbook is password protected.
    Password,

    /// Specified worksheet was not found.
    WorksheetNotFound(String),

    /// Specified worksheet Table was not found.
    TableNotFound(String),

    /// The specified sheet is not a worksheet.
    NotAWorksheet(String),

    /// A wrapper for a variety of [`quick_xml::encoding::EncodingError`]
    /// encoding errors.
    Encoding(quick_xml::encoding::EncodingError),
}

from_err!(std::io::Error, XlsxError, Io);
from_err!(zip::result::ZipError, XlsxError, Zip);
from_err!(crate::vba::VbaError, XlsxError, Vba);
from_err!(quick_xml::Error, XlsxError, Xml);
from_err!(std::num::ParseFloatError, XlsxError, ParseFloat);
from_err!(std::num::ParseIntError, XlsxError, ParseInt);
from_err!(quick_xml::encoding::EncodingError, XlsxError, Encoding);
from_err!(quick_xml::events::attributes::AttrError, XlsxError, XmlAttr);

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
            XlsxError::ColumnNumberOverflow => write!(f, "column number overflow"),
            XlsxError::RowNumberOverflow => write!(f, "row number overflow"),
            XlsxError::Unexpected(e) => write!(f, "{e}"),
            XlsxError::Unrecognized { typ, val } => write!(f, "Unrecognized {typ}: {val}"),
            XlsxError::CellError(e) => write!(f, "Unsupported cell error value '{e}'"),
            XlsxError::WorksheetNotFound(n) => write!(f, "Worksheet '{n}' not found"),
            XlsxError::Password => write!(f, "Workbook is password protected"),
            XlsxError::TableNotFound(n) => write!(f, "Table '{n}' not found"),
            XlsxError::NotAWorksheet(typ) => write!(f, "Expecting a worksheet, got {typ}"),
            XlsxError::Encoding(e) => write!(f, "XML encoding error: {e}"),
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
    /// Cell (number) formats
    formats: Vec<CellFormat>,
    /// Cell styles
    pub styles: Vec<Style>,
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
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"si" => {
                    if let Some(s) = read_string(&mut xml, e.name())? {
                        self.strings.push(s);
                    }
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sst" => break,
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
        let mut fonts = Vec::new();
        let mut fills = Vec::new();
        let mut borders = Vec::new();

        let mut buf = Vec::with_capacity(1024);
        let mut inner_buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"numFmts" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(e)) if e.local_name().as_ref() == b"numFmt" => {
                            let mut id = Vec::new();
                            let mut format = String::new();
                            for a in e.attributes() {
                                let a = a?;
                                match a {
                                    Attribute {
                                        key: QName(b"numFmtId"),
                                        value: v,
                                    } => id.extend_from_slice(&v),
                                    Attribute {
                                        key: QName(b"formatCode"),
                                        ..
                                    } => {
                                        let format_code = a
                                            .decode_and_unescape_value(xml.decoder())?
                                            .into_owned();
                                        // Excel format codes use backslashes to escape special characters
                                        // Remove escape backslashes (backslash followed by any character becomes just the character)
                                        let mut unescaped = String::new();
                                        let mut chars = format_code.chars().peekable();
                                        while let Some(ch) = chars.next() {
                                            if ch == '\\' {
                                                // If there's a next character, use it without the backslash
                                                if let Some(next_ch) = chars.next() {
                                                    unescaped.push(next_ch);
                                                } else {
                                                    // Trailing backslash, keep it
                                                    unescaped.push(ch);
                                                }
                                            } else {
                                                unescaped.push(ch);
                                            }
                                        }
                                        format = unescaped;
                                    }
                                    _ => (),
                                }
                            }
                            if !format.is_empty() {
                                number_formats.insert(id, format);
                            }
                        }
                        Ok(Event::End(e)) if e.local_name().as_ref() == b"numFmts" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("numFmts")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fonts" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                            let font = style_parser::parse_font(&mut xml, e)?;
                            fonts.push(font);
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fonts" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("fonts")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fills" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                            let fill = style_parser::parse_fill(&mut xml, e)?;
                            fills.push(fill);
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fills" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("fills")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"borders" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                            let border = style_parser::parse_border(&mut xml, e)?;
                            borders.push(border);
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"borders" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("borders")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                            // Parse the style by building it from referenced components
                            let mut style = Style::new();
                            let mut num_fmt_id_bytes: Option<Vec<u8>> = None;

                            // Parse attributes to get references to fonts, fills, borders
                            for a in e.attributes() {
                                let a = a?;
                                match a.key.as_ref() {
                                    b"fontId" => {
                                        if let Ok(font_id) =
                                            xml.decoder().decode(&a.value)?.parse::<usize>()
                                        {
                                            if let Some(font) = fonts.get(font_id) {
                                                style = style.with_font(font.clone());
                                            }
                                        }
                                    }
                                    b"fillId" => {
                                        if let Ok(fill_id) =
                                            xml.decoder().decode(&a.value)?.parse::<usize>()
                                        {
                                            if let Some(fill) = fills.get(fill_id) {
                                                style = style.with_fill(fill.clone());
                                            }
                                        }
                                    }
                                    b"borderId" => {
                                        if let Ok(border_id) =
                                            xml.decoder().decode(&a.value)?.parse::<usize>()
                                        {
                                            if let Some(border) = borders.get(border_id) {
                                                style = style.with_borders(border.clone());
                                            }
                                        }
                                    }
                                    b"numFmtId" => {
                                        // Store for both Style and CellFormat
                                        num_fmt_id_bytes = Some(a.value.to_vec());

                                        if let Ok(num_fmt_id) =
                                            xml.decoder().decode(&a.value)?.parse::<u32>()
                                        {
                                            let mut fmt_id_bytes = Vec::new();
                                            fmt_id_bytes.extend_from_slice(&a.value);
                                            let format_code = match number_formats
                                                .get(&fmt_id_bytes)
                                            {
                                                Some(fmt) => fmt.clone(),
                                                None => {
                                                    // Use built-in format
                                                    match num_fmt_id {
                                                        0 => "General".to_string(),
                                                        1 => "0".to_string(),
                                                        2 => "0.00".to_string(),
                                                        3 => "#,##0".to_string(),
                                                        4 => "#,##0.00".to_string(),
                                                        9 => "0%".to_string(),
                                                        10 => "0.00%".to_string(),
                                                        11 => "0.00E+00".to_string(),
                                                        12 => "# ?/?".to_string(),
                                                        13 => "# ??/??".to_string(),
                                                        14 => "mm-dd-yy".to_string(),
                                                        15 => "d-mmm-yy".to_string(),
                                                        16 => "d-mmm".to_string(),
                                                        17 => "mmm-yy".to_string(),
                                                        18 => "h:mm AM/PM".to_string(),
                                                        19 => "h:mm:ss AM/PM".to_string(),
                                                        20 => "h:mm".to_string(),
                                                        21 => "h:mm:ss".to_string(),
                                                        22 => "m/d/yy h:mm".to_string(),
                                                        37 => "#,##0 ;(#,##0)".to_string(),
                                                        38 => "#,##0 ;[Red](#,##0)".to_string(),
                                                        39 => "#,##0.00;(#,##0.00)".to_string(),
                                                        40 => "#,##0.00;[Red](#,##0.00)".to_string(),
                                                        41 => "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)".to_string(),
                                                        42 => "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)".to_string(),
                                                        43 => "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)".to_string(),
                                                        44 => "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)".to_string(),
                                                        45 => "mm:ss".to_string(),
                                                        46 => "[h]:mm:ss".to_string(),
                                                        47 => "mmss.0".to_string(),
                                                        48 => "##0.0E+0".to_string(),
                                                        49 => "@".to_string(),
                                                        _ => "General".to_string(),
                                                    }
                                                }
                                            };

                                            use crate::style::NumberFormat;
                                            let number_format =
                                                NumberFormat::new(format_code).with_id(num_fmt_id);
                                            style = style.with_number_format(number_format);
                                        }
                                    }
                                    _ => {}
                                }
                            }

                            // Also parse any nested elements like alignment and protection
                            let mut nested_buf = Vec::with_capacity(512);
                            loop {
                                nested_buf.clear();
                                match xml.read_event_into(&mut nested_buf) {
                                    Ok(Event::Start(ref nested_e)) => match nested_e
                                        .local_name()
                                        .as_ref()
                                    {
                                        b"alignment" => {
                                            let alignment =
                                                style_parser::parse_alignment(&mut xml, nested_e)?;
                                            style = style.with_alignment(alignment);
                                        }
                                        b"protection" => {
                                            let protection =
                                                style_parser::parse_protection(&mut xml, nested_e)?;
                                            style = style.with_protection(protection);
                                        }
                                        _ => {
                                            // Skip unknown nested elements
                                            xml.read_to_end_into(nested_e.name(), &mut Vec::new())?;
                                        }
                                    },
                                    Ok(Event::End(ref end_e))
                                        if end_e.local_name().as_ref() == b"xf" =>
                                    {
                                        break
                                    }
                                    Ok(Event::Eof) => return Err(XlsxError::XmlEof("xf")),
                                    Err(e) => return Err(XlsxError::Xml(e)),
                                    _ => {}
                                }
                            }

                            self.styles.push(style);

                            // Also add format for backward compatibility
                            self.formats.push(
                                num_fmt_id_bytes
                                    .map_or(CellFormat::Other, |id_bytes| {
                                        match number_formats.get(&id_bytes) {
                                            Some(fmt) => detect_custom_number_format(fmt),
                                            None => builtin_format_by_id(&id_bytes),
                                        }
                                    }),
                            );
                        }
                        Ok(Event::End(e)) if e.local_name().as_ref() == b"cellXfs" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("cellXfs")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("styleSheet")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
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
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"sheet" => {
                    let mut name = String::new();
                    let mut path = String::new();
                    let mut visible = SheetVisible::Visible;
                    for a in e.attributes() {
                        let a = a?;
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
                                key: QName(b"r:id" | b"relationships:id"),
                                value: v,
                            } => {
                                let r = &relationships
                                    .get(&*v)
                                    .ok_or(XlsxError::RelationshipNotFound)?[..];
                                // target may have prepended "/xl/" or "xl/" path;
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
                Ok(Event::Start(e)) if e.name().as_ref() == b"workbookPr" => {
                    self.is_1904 = match e.try_get_attribute("date1904")? {
                        Some(c) => ["1", "true"].contains(
                            &c.decode_and_unescape_value(xml.decoder())
                                .map_err(XlsxError::Xml)?
                                .as_ref(),
                        ),
                        None => false,
                    };
                }
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"definedName" => {
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
                                Event::Text(t) => value.push_str(&t.xml10_content()?),
                                Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut value)?,
                                Event::End(end) if end.name() == e.name() => break,
                                Event::Eof => return Err(XlsxError::XmlEof("workbook")),
                                _ => (),
                            }
                        }
                        defined_names.push((name, value));
                    }
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"workbook" => break,
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
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match a? {
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
                Ok(Event::End(e)) if e.local_name().as_ref() == b"Relationships" => break,
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
                        Ok(Event::Start(e)) if e.local_name().as_ref() == b"Relationship" => {
                            let mut id = Vec::new();
                            let mut target = String::new();
                            let mut table_type = false;
                            for a in e.attributes() {
                                match a? {
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
                                    // Relative path.
                                    let new_index =
                                        base_folder.rfind('/').expect("Must be a parent folder");
                                    let full_path =
                                        format!("{}{}", &base_folder[..new_index], &target[2..]);
                                    table_locations.push(full_path);
                                } else if let Some(stripped) = target.strip_prefix('/') {
                                    // Absolute path.
                                    table_locations.push(stripped.to_string());
                                } else if !target.is_empty() {
                                    // Assume absolute path without leading slash.
                                    table_locations.push(target);
                                }
                            }
                        }
                        Ok(Event::End(e)) if e.local_name().as_ref() == b"Relationships" => break,
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
                        Ok(Event::Start(e)) if e.local_name().as_ref() == b"table" => {
                            for a in e.attributes() {
                                match a? {
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
                        Ok(Event::Start(e)) if e.local_name().as_ref() == b"tableColumn" => {
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
                        Ok(Event::End(e)) if e.local_name().as_ref() == b"table" => break,
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

    // Read pictures.
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
                        Ok(Event::Start(e)) if e.local_name() == QName(b"mergeCell").into() => {
                            if let Some(attr) = get_attribute(e.attributes(), QName(b"ref"))? {
                                let dimension = get_dimension(attr)?;
                                regions.push((
                                    sheet_name.to_string(),
                                    sheet_path.to_string(),
                                    dimension,
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

    /// Load the merged regions in the workbook.
    ///
    /// A merged region in Excel is a range of cells that have been merged to
    /// act as a single cell. It is often used to create headers or titles that
    /// span multiple columns or rows.
    ///
    /// This method must be called before accessing the merged regions using the
    /// methods:
    ///
    /// - [`Xlsx::merged_regions()`].
    /// - [`Xlsx::merged_regions_by_sheet()`].
    ///
    /// These methods are explained below.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::Xml`].
    ///
    pub fn load_merged_regions(&mut self) -> Result<(), XlsxError> {
        if self.merged_regions.is_none() {
            self.read_merged_regions()
        } else {
            Ok(())
        }
    }

    /// Get the merged regions for all the worksheets in a workbook.
    ///
    /// The function returns a ref to a vector of tuples containing the sheet
    /// name, the sheet path, and the [`Dimensions`] of the merged region. The
    /// middle element of the tuple can generally be ignored.
    ///
    /// The [`Xlsx::load_merged_regions()`] method must be called before calling
    /// this method.
    ///
    /// # Examples
    ///
    /// An example of getting all the merged regions in an Excel workbook.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/merged_range.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Load the merged regions in the workbook.
    ///     workbook.load_merged_regions()?;
    ///
    ///     // Get all the merged regions in the workbook.
    ///     let merged_regions = workbook.merged_regions();
    ///
    ///     // Print the sheet name and dimensions of each merged region.
    ///     for (sheet_name, _, dimensions) in merged_regions {
    ///         println!("{sheet_name}: {dimensions:?}");
    ///     }
    ///
    ///     Ok(())
    /// }
    ///
    /// ```
    ///
    /// Output:
    ///
    /// ```text
    /// Sheet1: Dimensions { start: (0, 7), end: (1, 7) }
    /// Sheet1: Dimensions { start: (0, 0), end: (1, 0) }
    /// Sheet1: Dimensions { start: (0, 1), end: (1, 1) }
    /// Sheet1: Dimensions { start: (0, 2), end: (1, 3) }
    /// Sheet1: Dimensions { start: (2, 2), end: (2, 3) }
    /// Sheet1: Dimensions { start: (3, 2), end: (3, 3) }
    /// Sheet1: Dimensions { start: (0, 4), end: (1, 4) }
    /// Sheet1: Dimensions { start: (0, 5), end: (1, 5) }
    /// Sheet1: Dimensions { start: (0, 6), end: (1, 6) }
    /// Sheet2: Dimensions { start: (0, 0), end: (3, 0) }
    /// Sheet2: Dimensions { start: (2, 2), end: (3, 3) }
    /// Sheet2: Dimensions { start: (0, 5), end: (3, 7) }
    /// Sheet2: Dimensions { start: (0, 1), end: (1, 1) }
    /// Sheet2: Dimensions { start: (0, 2), end: (1, 3) }
    /// Sheet2: Dimensions { start: (0, 4), end: (1, 4) }
    /// ```
    ///
    pub fn merged_regions(&self) -> &Vec<(String, String, Dimensions)> {
        self.merged_regions
            .as_ref()
            .expect("Merged Regions must be loaded before the are referenced")
    }

    /// Get the merged regions in a workbook by the sheet name.
    ///
    /// The function returns a vector of tuples containing the sheet name, the
    /// sheet path, and the [`Dimensions`] of the merged region. The first two
    /// elements of the tuple can generally be ignored.
    ///
    /// The [`Xlsx::load_merged_regions()`] method must be called before calling
    /// this method.
    ///
    /// # Parameters
    ///
    /// - `sheet_name`: The name of the worksheet to get the merged regions from.
    ///
    /// # Examples
    ///
    /// An example of getting the merged regions in an Excel workbook, by individual
    /// worksheet.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Reader, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/merged_range.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the names of all the sheets in the workbook.
    ///     let sheet_names = workbook.sheet_names();
    ///
    ///     // Load the merged regions in the workbook.
    ///     workbook.load_merged_regions()?;
    ///
    ///     for sheet_name in &sheet_names {
    ///         println!("{sheet_name}: ");
    ///
    ///         // Get the merged regions in the current sheet.
    ///         let merged_regions = workbook.merged_regions_by_sheet(sheet_name);
    ///
    ///         for (_, _, dimensions) in &merged_regions {
    ///             // Print the dimensions of each merged region.
    ///             println!("    {dimensions:?}");
    ///         }
    ///     }
    ///
    ///     Ok(())
    /// }
    ///
    /// ```
    ///
    /// Output:
    ///
    ///
    /// ```text
    /// Sheet1:
    ///     Dimensions { start: (0, 7), end: (1, 7) }
    ///     Dimensions { start: (0, 0), end: (1, 0) }
    ///     Dimensions { start: (0, 1), end: (1, 1) }
    ///     Dimensions { start: (0, 2), end: (1, 3) }
    ///     Dimensions { start: (2, 2), end: (2, 3) }
    ///     Dimensions { start: (3, 2), end: (3, 3) }
    ///     Dimensions { start: (0, 4), end: (1, 4) }
    ///     Dimensions { start: (0, 5), end: (1, 5) }
    ///     Dimensions { start: (0, 6), end: (1, 6) }
    /// Sheet2:
    ///     Dimensions { start: (0, 0), end: (3, 0) }
    ///     Dimensions { start: (2, 2), end: (3, 3) }
    ///     Dimensions { start: (0, 5), end: (3, 7) }
    ///     Dimensions { start: (0, 1), end: (1, 1) }
    ///     Dimensions { start: (0, 2), end: (1, 3) }
    ///     Dimensions { start: (0, 4), end: (1, 4) }
    /// ```
    ///
    pub fn merged_regions_by_sheet(&self, name: &str) -> Vec<(&String, &String, &Dimensions)> {
        self.merged_regions()
            .iter()
            .filter(|s| s.0 == name)
            .map(|(name, sheet, region)| (name, sheet, region))
            .collect()
    }

    /// Load the worksheet tables from the XLSX file.
    ///
    /// Tables in Excel are a way of grouping a range of cells into a single
    /// entity that has common formatting or that can be referenced in formulas.
    /// In `calamine`, tables can be read as a [`Table`] object and converted to
    /// a data [`Range`] for further processing.
    ///
    /// Calamine does not automatically load table data from a workbook to avoid
    /// unnecessary overhead. Instead you must explicitly load the table data
    /// using the `Xlsx::load_tables()` method. Once the tables have been loaded
    /// the following methods can be used to extract and work with individual
    /// tables:
    ///
    /// - [`Xlsx::table_by_name()`].
    /// - [`Xlsx::table_by_name_ref()`].
    /// - [`Xlsx::table_names()`].
    /// - [`Xlsx::table_names_in_sheet()`].
    ///
    /// These methods are explained below. See also the [`Table`] struct for
    /// additional methods that can be used when working with tables.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::XmlAttr`].
    /// - [`XlsxError::XmlEof`].
    /// - [`XlsxError::Xml`].
    ///
    ///
    pub fn load_tables(&mut self) -> Result<(), XlsxError> {
        if self.tables.is_none() {
            self.read_table_metadata()
        } else {
            Ok(())
        }
    }

    /// Get the names of all the tables in the workbook.
    ///
    /// Read all the table names in the workbook. This can be used in
    /// conjunction with [`Xlsx::table_by_name()`] to iterate over the tables in
    /// the workbook.
    ///
    /// # Panics
    ///
    /// Panics if tables have not been loaded via [`Xlsx::load_tables()`].
    ///
    /// # Examples
    ///
    /// An example of getting the names of all the tables in an Excel workbook.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/table-multiple.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Load the tables in the workbook.
    ///     workbook.load_tables()?;
    ///
    ///     // Get all the table names in the workbook.
    ///     let table_names = workbook.table_names();
    ///
    ///     // Check the table names.
    ///     assert_eq!(
    ///         table_names,
    ///         vec!["Inventory", "Pricing", "Sales_Bob", "Sales_Alice"]
    ///     );
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn table_names(&self) -> Vec<&String> {
        self.tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .map(|(name, ..)| name)
            .collect()
    }

    /// Get the names of all the tables in a worksheet.
    ///
    /// Read all the table names in a worksheet. This can be used in conjunction
    /// with [`Xlsx::table_by_name()`] to iterate over the tables in the
    /// worksheet.
    ///
    /// # Parameters
    ///
    /// - `sheet_name`: The name of the worksheet to get the table names from.
    ///
    /// # Panics
    ///
    /// Panics if tables have not been loaded via [`Xlsx::load_tables()`].
    ///
    /// # Examples
    ///
    /// An example of getting the names of all the tables in an Excel workbook,
    /// sheet by sheet.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Reader, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/table-multiple.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the names of all the sheets in the workbook.
    ///     let sheet_names = workbook.sheet_names();
    ///
    ///     // Load the tables in the workbook.
    ///     workbook.load_tables()?;
    ///
    ///     for sheet_name in &sheet_names {
    ///         // Get the table names in the current sheet.
    ///         let table_names = workbook.table_names_in_sheet(sheet_name);
    ///
    ///         // Print the associated table names.
    ///         println!("{sheet_name} contains tables: {table_names:?}");
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output:
    ///
    /// ```text
    /// Sheet1 contains tables: ["Inventory"]
    /// Sheet2 contains tables: ["Pricing"]
    /// Sheet3 contains tables: ["Sales_Bob", "Sales_Alice"]
    /// ```
    ///
    pub fn table_names_in_sheet(&self, sheet_name: &str) -> Vec<&String> {
        self.tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .filter(|(_, sheet, ..)| sheet == sheet_name)
            .map(|(name, ..)| name)
            .collect()
    }

    /// Get a worksheet table by name.
    ///
    /// This method retrieves a [`Table`] from the workbook by its name. The
    /// table will contain an owned copy of the worksheet data in the table
    /// range.
    ///
    /// # Parameters
    ///
    /// - `table_name`: The name of the table to retrieve.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::TableNotFound`].
    /// - [`XlsxError::NotAWorksheet`].
    ///
    /// # Panics
    ///
    /// Panics if tables have not been loaded via [`Xlsx::load_tables()`].
    ///
    /// # Examples
    ///
    /// An example of getting an Excel worksheet table by its name. The file in
    /// this example contains 4 tables spread across 3 worksheets. This example
    /// gets an owned copy of the worksheet data in the table area.
    ///
    /// ```
    /// use calamine::{open_workbook, Data, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/table-multiple.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Load the tables in the workbook.
    ///     workbook.load_tables()?;
    ///
    ///     // Get the table by name.
    ///     let table = workbook.table_by_name("Inventory")?;
    ///
    ///     // Get the data range of the table. The data type is `&Range<Data>`.
    ///     let data_range = table.data();
    ///
    ///     // Do something with the data using the `Range` APIs. In this case
    ///     // we will just check for a cell value.
    ///     assert_eq!(
    ///         data_range.get((0, 1)),
    ///         Some(&Data::String("Apple".to_string()))
    ///     );
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    pub fn table_by_name(&mut self, table_name: &str) -> Result<Table<Data>, XlsxError> {
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

    /// Get a worksheet table by name, with referenced data.
    ///
    /// This method retrieves a [`Table`] from the workbook by its name. The
    /// table will contain an borrowed/referenced copy of the worksheet data in
    /// the table range. This is more efficient than [`Xlsx::table_by_name()`]
    /// for large tables.
    ///
    /// # Parameters
    ///
    /// - `table_name`: The name of the table to retrieve.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::TableNotFound`].
    /// - [`XlsxError::NotAWorksheet`].
    ///
    /// # Panics
    ///
    /// Panics if tables have not been loaded via [`Xlsx::load_tables()`].
    ///
    /// # Examples
    ///
    /// An example of getting an Excel worksheet table by its name. The file in
    /// this example contains 4 tables spread across 3 worksheets. This example
    /// gets a borrowed/referenced copy of the worksheet data in the table area.
    ///
    /// ```
    /// use calamine::{open_workbook, DataRef, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/table-multiple.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Load the tables in the workbook.
    ///     workbook.load_tables()?;
    ///
    ///     // Get the table by name.
    ///     let table = workbook.table_by_name_ref("Inventory")?;
    ///
    ///     // Get the data range of the table. The data type is `&Range<DataRef<'_>>`.
    ///     let data_range = table.data();
    ///
    ///     // Do something with the data using the `Range` APIs. In this case
    ///     // we will just check for a cell value.
    ///     assert_eq!(
    ///         data_range.get((0, 1)),
    ///         Some(&DataRef::SharedString("Apple"))
    ///     );
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
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

    /// Get the merged cells/regions in a workbook by the sheet name.
    ///
    /// Merged cells in Excel are a range of cells that have been merged to act
    /// as a single cell. It is often used to create headers or titles that span
    /// multiple columns or rows.
    ///
    /// The function returns a vector of [`Dimensions`] of the merged region.
    /// This is wrapped in a [`Result`] and an [`Option`].
    ///
    /// # Parameters
    ///
    /// - `sheet_name`: The name of the worksheet to get the merged regions
    ///   from.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::Xml`].
    ///
    /// # Examples
    ///
    /// An example of getting the merged regions/cells in an Excel workbook, by
    /// individual worksheet.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Reader, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/merged_range.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the names of all the sheets in the workbook.
    ///     let sheet_names = workbook.sheet_names();
    ///
    ///     for sheet_name in &sheet_names {
    ///         println!("{sheet_name}: ");
    ///
    ///         // Get the merged cells in the current sheet.
    ///         let merge_cells = workbook.worksheet_merge_cells(sheet_name);
    ///
    ///         if let Some(dimensions) = merge_cells {
    ///             let dimensions = dimensions?;
    ///
    ///             // Print the dimensions of each merged region.
    ///             for dimension in &dimensions {
    ///                 println!("    {dimension:?}");
    ///             }
    ///         }
    ///     }
    ///
    ///     Ok(())
    /// }
    ///
    /// ```
    ///
    /// Output:
    ///
    /// ```text
    /// Sheet1:
    ///     Dimensions { start: (0, 7), end: (1, 7) }
    ///     Dimensions { start: (0, 0), end: (1, 0) }
    ///     Dimensions { start: (0, 1), end: (1, 1) }
    ///     Dimensions { start: (0, 2), end: (1, 3) }
    ///     Dimensions { start: (2, 2), end: (2, 3) }
    ///     Dimensions { start: (3, 2), end: (3, 3) }
    ///     Dimensions { start: (0, 4), end: (1, 4) }
    ///     Dimensions { start: (0, 5), end: (1, 5) }
    ///     Dimensions { start: (0, 6), end: (1, 6) }
    /// Sheet2:
    ///     Dimensions { start: (0, 0), end: (3, 0) }
    ///     Dimensions { start: (2, 2), end: (3, 3) }
    ///     Dimensions { start: (0, 5), end: (3, 7) }
    ///     Dimensions { start: (0, 1), end: (1, 1) }
    ///     Dimensions { start: (0, 2), end: (1, 3) }
    ///     Dimensions { start: (0, 4), end: (1, 4) }
    /// ```
    ///
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

    /// Get the merged cells/regions in a workbook by the sheet index.
    ///
    /// Merged cells in Excel are a range of cells that have been merged to act
    /// as a single cell. It is often used to create headers or titles that span
    /// multiple columns or rows.
    ///
    /// The function returns a vector of [`Dimensions`] of the merged region.
    /// This is wrapped in a [`Result`] and an [`Option`].
    ///
    /// # Parameters
    ///
    /// - `sheet_index`: The zero index of the worksheet to get the merged
    ///   regions from.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::Xml`].
    ///
    /// # Examples
    ///
    /// An example of getting the merged regions/cells in an Excel workbook, by
    /// worksheet index.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/merged_range.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the merged cells in the first worksheet.
    ///     let merge_cells = workbook.worksheet_merge_cells_at(0);
    ///
    ///     if let Some(dimensions) = merge_cells {
    ///         let dimensions = dimensions?;
    ///
    ///         // Print the dimensions of each merged region.
    ///         for dimension in &dimensions {
    ///             println!("{dimension:?}");
    ///         }
    ///     }
    ///
    ///     Ok(())
    /// }
    ///
    /// ```
    ///
    /// Output:
    ///
    /// ```text
    /// Dimensions { start: (0, 7), end: (1, 7) }
    /// Dimensions { start: (0, 0), end: (1, 0) }
    /// Dimensions { start: (0, 1), end: (1, 1) }
    /// Dimensions { start: (0, 2), end: (1, 3) }
    /// Dimensions { start: (2, 2), end: (2, 3) }
    /// Dimensions { start: (3, 2), end: (3, 3) }
    /// Dimensions { start: (0, 4), end: (1, 4) }
    /// Dimensions { start: (0, 5), end: (1, 5) }
    /// Dimensions { start: (0, 6), end: (1, 6) }
    /// ```
    ///
    pub fn worksheet_merge_cells_at(
        &mut self,
        sheet_index: usize,
    ) -> Option<Result<Vec<Dimensions>, XlsxError>> {
        let name = self
            .metadata()
            .sheets
            .get(sheet_index)
            .map(|sheet| sheet.name.clone())?;

        self.worksheet_merge_cells(&name)
    }

    /// Get the cells reader for a worksheet.
    ///
    /// This function returns a [`XlsxCellReader`] for the specified worksheet.
    /// The reader can be used to iterate over the cells in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the worksheet to get the cells reader for.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::WorksheetNotFound`].
    ///
    /// # Examples
    ///
    /// An example of getting the cells reader for a worksheet.
    ///
    /// ```
    /// use calamine::{open_workbook, Error, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = "tests/merged_range.xlsx";
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the cells reader for the first worksheet.
    ///     let reader = workbook.worksheet_cells_reader("Sheet1")?;
    ///
    ///     // Iterate over the cells in the worksheet.
    ///     while let Some(cell) = reader.next_cell()? {
    ///         println!("{:?}", cell);
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output:
    ///
    /// ```text
    /// Cell {
    ///     row: 0,
    ///     col: 0,
    ///     val: "Hello, world!".to_string(),
    ///     err: None,
    ///     typ: DataType::String,
    /// }
    /// ```
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
        let formats = &self.formats;
        let styles = &self.styles;
        XlsxCellReader::new(xml, strings, formats, styles, is_1904)
    }

    /// Get the styles for a worksheet.
    ///
    /// This function returns a [`Range<Style>`] for the specified worksheet.
    /// The range contains the styles for the cells in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the worksheet to get the styles for.
    ///
    pub fn worksheet_style(&mut self, name: &str) -> Result<Range<Style>, XlsxError> {
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
        while let Some(cell) = cell_reader.next_style()? {
            if !cell.val.is_empty() {
                cells.push(cell);
            }
        }
        Ok(Range::from_sparse(cells))
    }

    /// Get the layout for a worksheet.
    ///
    /// This function returns a [`WorksheetLayout`] for the specified worksheet.
    /// The layout contains the column widths and row heights for the cells in the worksheet.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the worksheet to get the layout for.
    ///
    pub fn worksheet_layout(&mut self, name: &str) -> Result<WorksheetLayout, XlsxError> {
        let (_, path) = self
            .sheets
            .iter()
            .find(|&(n, _)| n == name)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))?;

        let mut xml = xml_reader(&mut self.zip, path)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))??;

        let mut layout = WorksheetLayout::new();
        let mut buf = Vec::with_capacity(1024);

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetFormatPr" => {
                    // Parse default column width and row height
                    for attr in e.attributes() {
                        let attr = attr.map_err(XlsxError::XmlAttr)?;
                        match attr.key.as_ref() {
                            b"defaultColWidth" => {
                                if let Ok(width_str) = xml.decoder().decode(&attr.value) {
                                    if let Ok(width) = width_str.parse::<f64>() {
                                        layout = layout.with_default_column_width(width);
                                    }
                                }
                            }
                            b"defaultRowHeight" => {
                                if let Ok(height_str) = xml.decoder().decode(&attr.value) {
                                    if let Ok(height) = height_str.parse::<f64>() {
                                        layout = layout.with_default_row_height(height);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cols" => {
                    // Parse column definitions
                    loop {
                        buf.clear();
                        match xml.read_event_into(&mut buf) {
                            Ok(Event::Start(ref col_e))
                                if col_e.local_name().as_ref() == b"col" =>
                            {
                                let mut col_info = None;
                                let mut width = 0.0;
                                let mut custom_width = false;
                                let mut hidden = false;
                                let mut best_fit = false;

                                for attr in col_e.attributes() {
                                    let attr = attr.map_err(XlsxError::XmlAttr)?;
                                    match attr.key.as_ref() {
                                        b"min" => {
                                            if let Ok(min_str) = xml.decoder().decode(&attr.value) {
                                                if let Ok(min_col) = min_str.parse::<u32>() {
                                                    col_info = Some(min_col - 1);
                                                    // Convert to 0-based
                                                }
                                            }
                                        }
                                        b"width" => {
                                            if let Ok(width_str) = xml.decoder().decode(&attr.value)
                                            {
                                                if let Ok(w) = width_str.parse::<f64>() {
                                                    width = w;
                                                }
                                            }
                                        }
                                        b"customWidth" => {
                                            custom_width = attr.value.as_ref() != b"0";
                                        }
                                        b"hidden" => {
                                            hidden = attr.value.as_ref() != b"0";
                                        }
                                        b"bestFit" => {
                                            best_fit = attr.value.as_ref() != b"0";
                                        }
                                        _ => {}
                                    }
                                }

                                if let Some(col) = col_info {
                                    let column_width = ColumnWidth::new(col, width)
                                        .with_custom_width(custom_width)
                                        .with_hidden(hidden)
                                        .with_best_fit(best_fit);
                                    layout = layout.add_column_width(column_width);
                                }
                            }
                            Ok(Event::End(ref end_e)) if end_e.local_name().as_ref() == b"cols" => {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("cols")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => {}
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    // Parse row definitions
                    loop {
                        buf.clear();
                        match xml.read_event_into(&mut buf) {
                            Ok(Event::Start(ref row_e))
                                if row_e.local_name().as_ref() == b"row" =>
                            {
                                let mut row_num = None;
                                let mut height = 0.0;
                                let mut custom_height = false;
                                let mut hidden = false;
                                let mut thick_top = false;
                                let mut thick_bottom = false;

                                for attr in row_e.attributes() {
                                    let attr = attr.map_err(XlsxError::XmlAttr)?;
                                    match attr.key.as_ref() {
                                        b"r" => {
                                            if let Ok(row_str) = xml.decoder().decode(&attr.value) {
                                                if let Ok(r) = row_str.parse::<u32>() {
                                                    row_num = Some(r - 1); // Convert to 0-based
                                                }
                                            }
                                        }
                                        b"ht" => {
                                            if let Ok(height_str) =
                                                xml.decoder().decode(&attr.value)
                                            {
                                                if let Ok(h) = height_str.parse::<f64>() {
                                                    height = h;
                                                }
                                            }
                                        }
                                        b"customHeight" => {
                                            custom_height = attr.value.as_ref() != b"0";
                                        }
                                        b"hidden" => {
                                            hidden = attr.value.as_ref() != b"0";
                                        }
                                        b"thickTop" => {
                                            thick_top = attr.value.as_ref() != b"0";
                                        }
                                        b"thickBot" => {
                                            thick_bottom = attr.value.as_ref() != b"0";
                                        }
                                        _ => {}
                                    }
                                }

                                // Only add row height if it's custom or has special properties
                                if let Some(row) = row_num {
                                    if custom_height
                                        || hidden
                                        || thick_top
                                        || thick_bottom
                                        || height > 0.0
                                    {
                                        let row_height = RowHeight::new(row, height)
                                            .with_custom_height(custom_height)
                                            .with_hidden(hidden)
                                            .with_thick_top(thick_top)
                                            .with_thick_bottom(thick_bottom);
                                        layout = layout.add_row_height(row_height);
                                    }
                                }

                                // Skip to the end of this row element
                                xml.read_to_end_into(row_e.name(), &mut Vec::new())?;
                            }
                            Ok(Event::End(ref end_e))
                                if end_e.local_name().as_ref() == b"sheetData" =>
                            {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => {}
                        }
                    }
                    break; // We're done after processing sheetData
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }

        Ok(layout)
    }

    /// Get all worksheets in the workbook.
    ///
    /// This function returns a vector of tuples, where each tuple contains the name of a worksheet and the range of cells in the worksheet.
    ///
    /// # Returns
    ///
    /// A vector of tuples, where each tuple contains the name of a worksheet and the range of cells in the worksheet.
    ///
    pub fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
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
    totals_row_count: u32,
}

impl InnerTableMetadata {
    fn new() -> Self {
        Self {
            display_name: String::new(),
            ref_cells: String::new(),
            header_row_count: 1,
            totals_row_count: 0,
        }
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
            is_1904: false,
            sheets: Vec::new(),
            tables: None,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
            merged_regions: None,
            options: XlsxOptions::default(),
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

    fn vba_project(&mut self) -> Result<Option<VbaProject>, XlsxError> {
        let Some(mut f) = self.zip.by_name("xl/vbaProject.bin").ok() else {
            return Ok(None);
        };
        let len = f.size() as usize;
        let vba = VbaProject::new(&mut f, len)?;
        Ok(Some(vba))
    }

    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    fn worksheet_range(&mut self, name: &str) -> Result<Range<Data>, XlsxError> {
        let rge = self.worksheet_range_ref(name)?;
        let inner = rge.inner.into_iter().map(|v| v.into()).collect();
        Ok(Range {
            start: rge.start,
            end: rge.end,
            inner,
        })
    }

    fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, XlsxError> {
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
        while let Some(cell) = cell_reader.next_formula()? {
            if !cell.val.is_empty() {
                cells.push(cell);
            }
        }
        Ok(Range::from_sparse(cells))
    }

    fn worksheet_style(&mut self, name: &str) -> Result<Range<Style>, XlsxError> {
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
        while let Some(cell) = cell_reader.next_style()? {
            if !cell.val.is_empty() {
                cells.push(cell);
            }
        }
        Ok(Range::from_sparse(cells))
    }

    fn worksheet_layout(&mut self, name: &str) -> Result<WorksheetLayout, XlsxError> {
        let (_, path) = self
            .sheets
            .iter()
            .find(|&(n, _)| n == name)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))?;

        let mut xml = xml_reader(&mut self.zip, path)
            .ok_or_else(|| XlsxError::WorksheetNotFound(name.into()))??;

        let mut layout = WorksheetLayout::new();
        let mut buf = Vec::with_capacity(1024);

        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetFormatPr" => {
                    // Parse default column width and row height
                    for attr in e.attributes() {
                        let attr = attr.map_err(XlsxError::XmlAttr)?;
                        match attr.key.as_ref() {
                            b"defaultColWidth" => {
                                if let Ok(width_str) = xml.decoder().decode(&attr.value) {
                                    if let Ok(width) = width_str.parse::<f64>() {
                                        layout = layout.with_default_column_width(width);
                                    }
                                }
                            }
                            b"defaultRowHeight" => {
                                if let Ok(height_str) = xml.decoder().decode(&attr.value) {
                                    if let Ok(height) = height_str.parse::<f64>() {
                                        layout = layout.with_default_row_height(height);
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cols" => {
                    // Parse column definitions
                    loop {
                        buf.clear();
                        match xml.read_event_into(&mut buf) {
                            Ok(Event::Start(ref col_e))
                                if col_e.local_name().as_ref() == b"col" =>
                            {
                                let mut col_info = None;
                                let mut width = 0.0;
                                let mut custom_width = false;
                                let mut hidden = false;
                                let mut best_fit = false;

                                for attr in col_e.attributes() {
                                    let attr = attr.map_err(XlsxError::XmlAttr)?;
                                    match attr.key.as_ref() {
                                        b"min" => {
                                            if let Ok(min_str) = xml.decoder().decode(&attr.value) {
                                                if let Ok(min_col) = min_str.parse::<u32>() {
                                                    col_info = Some(min_col - 1);
                                                    // Convert to 0-based
                                                }
                                            }
                                        }
                                        b"width" => {
                                            if let Ok(width_str) = xml.decoder().decode(&attr.value)
                                            {
                                                if let Ok(w) = width_str.parse::<f64>() {
                                                    width = w;
                                                }
                                            }
                                        }
                                        b"customWidth" => {
                                            custom_width = attr.value.as_ref() != b"0";
                                        }
                                        b"hidden" => {
                                            hidden = attr.value.as_ref() != b"0";
                                        }
                                        b"bestFit" => {
                                            best_fit = attr.value.as_ref() != b"0";
                                        }
                                        _ => {}
                                    }
                                }

                                if let Some(col) = col_info {
                                    let column_width = ColumnWidth::new(col, width)
                                        .with_custom_width(custom_width)
                                        .with_hidden(hidden)
                                        .with_best_fit(best_fit);
                                    layout = layout.add_column_width(column_width);
                                }
                            }
                            Ok(Event::End(ref end_e)) if end_e.local_name().as_ref() == b"cols" => {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("cols")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => {}
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    // Parse row definitions
                    loop {
                        buf.clear();
                        match xml.read_event_into(&mut buf) {
                            Ok(Event::Start(ref row_e))
                                if row_e.local_name().as_ref() == b"row" =>
                            {
                                let mut row_num = None;
                                let mut height = 0.0;
                                let mut custom_height = false;
                                let mut hidden = false;
                                let mut thick_top = false;
                                let mut thick_bottom = false;

                                for attr in row_e.attributes() {
                                    let attr = attr.map_err(XlsxError::XmlAttr)?;
                                    match attr.key.as_ref() {
                                        b"r" => {
                                            if let Ok(row_str) = xml.decoder().decode(&attr.value) {
                                                if let Ok(r) = row_str.parse::<u32>() {
                                                    row_num = Some(r - 1); // Convert to 0-based
                                                }
                                            }
                                        }
                                        b"ht" => {
                                            if let Ok(height_str) =
                                                xml.decoder().decode(&attr.value)
                                            {
                                                if let Ok(h) = height_str.parse::<f64>() {
                                                    height = h;
                                                }
                                            }
                                        }
                                        b"customHeight" => {
                                            custom_height = attr.value.as_ref() != b"0";
                                        }
                                        b"hidden" => {
                                            hidden = attr.value.as_ref() != b"0";
                                        }
                                        b"thickTop" => {
                                            thick_top = attr.value.as_ref() != b"0";
                                        }
                                        b"thickBot" => {
                                            thick_bottom = attr.value.as_ref() != b"0";
                                        }
                                        _ => {}
                                    }
                                }

                                // Only add row height if it's custom or has special properties
                                if let Some(row) = row_num {
                                    if custom_height
                                        || hidden
                                        || thick_top
                                        || thick_bottom
                                        || height > 0.0
                                    {
                                        let row_height = RowHeight::new(row, height)
                                            .with_custom_height(custom_height)
                                            .with_hidden(hidden)
                                            .with_thick_top(thick_top)
                                            .with_thick_bottom(thick_bottom);
                                        layout = layout.add_row_height(row_height);
                                    }
                                }

                                // Skip to the end of this row element
                                xml.read_to_end_into(row_e.name(), &mut Vec::new())?;
                            }
                            Ok(Event::End(ref end_e))
                                if end_e.local_name().as_ref() == b"sheetData" =>
                            {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => {}
                        }
                    }
                    break; // We're done after processing sheetData
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }

        Ok(layout)
    }

    fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
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
                            style: None,
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
    let zip_path = path_to_zip_path(zip, path);

    match zip.by_name(&zip_path) {
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
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"r" => {
                if rich_buffer.is_none() {
                    // use a buffer since richtext has multiples <r> and <t> for the same cell
                    rich_buffer = Some(String::new());
                }
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"rPh" => {
                is_phonetic_text = true;
            }
            Ok(Event::End(e)) if e.name() == closing => {
                if rich_buffer.is_none() {
                    // An empty <s></s> element, without <t> or other
                    // subelements, is treated as a valid empty string in Excel.
                    rich_buffer = Some(String::new());
                }

                return Ok(rich_buffer);
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == b"rPh" => {
                is_phonetic_text = false;
            }
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"t" && !is_phonetic_text => {
                val_buf.clear();
                let mut value = String::new();
                loop {
                    match xml.read_event_into(&mut val_buf)? {
                        Event::Text(t) => value.push_str(&unescape_xml(&t.xml10_content()?)),
                        Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut value)?,
                        Event::End(end) if end.name() == e.name() => break,
                        Event::Eof => return Err(XlsxError::XmlEof("t")),
                        _ => (),
                    }
                }
                if let Some(s) = &mut rich_buffer {
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
                    let attribute = attribute?;

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

/// Advance the cell name by the offset
///
/// This function advances the cell name by the offset.
///
/// # Parameters
///
/// - `name`: The cell name to advance.
/// - `offset`: The offset to advance the cell name by.
///
/// # Returns
///
/// A vector of bytes representing the advanced cell name.
///
pub fn offset_cell_name(name: &[u8], offset: (i64, i64)) -> Result<Vec<u8>, XlsxError> {
    let reference = Reference::parse(name)?;
    let offset_ref = reference.offset(offset)?;
    let mut buf = Vec::new();
    offset_ref.format(&mut buf)?;
    Ok(buf)
}

#[derive(Debug, Copy, Clone)]
enum Reference {
    Cell {
        row: u32,
        col: u32,
        absolute_row: bool,
        absolute_col: bool,
    },
    Row {
        row: u32,
        absolute: bool,
    },
    Column {
        col: u32,
        absolute: bool,
    },
}

impl Reference {
    // Create a cell reference with validation.
    fn cell(row: u32, col: u32, absolute_row: bool, absolute_col: bool) -> Result<Self, XlsxError> {
        let reference = Reference::Cell {
            row,
            col,
            absolute_row,
            absolute_col,
        };
        reference.validate()?;
        Ok(reference)
    }

    // Create a column reference with validation.
    fn column(col: u32, absolute: bool) -> Result<Self, XlsxError> {
        let reference = Reference::Column { col, absolute };
        reference.validate()?;
        Ok(reference)
    }

    // Create a row reference with validation.
    fn row(row: u32, absolute: bool) -> Result<Self, XlsxError> {
        let reference = Reference::Row { row, absolute };
        reference.validate()?;
        Ok(reference)
    }

    // Parse a reference (e.g., "A1", "$A1", "A$1", "$A$1", "E", "$E", "5", "$5").
    fn parse(name: &[u8]) -> Result<Self, XlsxError> {
        let mut iter = name.iter().peekable();
        let mut col: u32 = 0;
        let mut row: u32 = 0;
        let mut absolute_col = false;
        let mut absolute_row = false;

        while let Some(&c) = iter.next() {
            match (c, iter.peek()) {
                (b'$', Some(b'A'..=b'Z' | b'a'..=b'z')) => {
                    if row > 0 || col > 0 {
                        return Err(XlsxError::Alphanumeric(c));
                    }
                    absolute_col = true;
                }
                (b'$', Some(b'0'..=b'9')) => {
                    if row > 0 {
                        return Err(XlsxError::Alphanumeric(c));
                    }
                    absolute_row = true;
                }
                (b'$', _) => return Err(XlsxError::Alphanumeric(c)),
                (c @ (b'A'..=b'Z' | b'a'..=b'z'), _) => {
                    if row > 0 {
                        return Err(XlsxError::Alphanumeric(c));
                    }
                    col = col
                        .wrapping_mul(26)
                        .wrapping_add((c.to_ascii_uppercase() - b'A') as u32 + 1);
                }
                (c @ b'0'..=b'9', _) => {
                    row = row.wrapping_mul(10).wrapping_add((c - b'0') as u32);
                }
                _ => return Err(XlsxError::Alphanumeric(c)),
            }
        }

        match (col.checked_sub(1), row.checked_sub(1)) {
            (Some(col), Some(row)) => Reference::cell(row, col, absolute_row, absolute_col),
            (Some(col), None) => Reference::column(col, absolute_col),
            (None, Some(row)) => Reference::row(row, absolute_row),
            (None, None) => Err(XlsxError::Unexpected("Empty reference")),
        }
    }

    // Apply offset to create a new reference with validation.
    fn offset(self, offset: (i64, i64)) -> Result<Self, XlsxError> {
        let result = match self {
            Reference::Cell {
                row,
                col,
                absolute_row,
                absolute_col,
            } => {
                let new_col = if absolute_col {
                    col
                } else {
                    (col as i64 + offset.1) as u32
                };
                let new_row = if absolute_row {
                    row
                } else {
                    (row as i64 + offset.0) as u32
                };

                Reference::Cell {
                    row: new_row,
                    col: new_col,
                    absolute_row,
                    absolute_col,
                }
            }
            Reference::Column { col, absolute } => {
                let new_col = if absolute {
                    col
                } else {
                    (col as i64 + offset.1) as u32
                };

                Reference::Column {
                    col: new_col,
                    absolute,
                }
            }
            Reference::Row { row, absolute } => {
                let new_row = if absolute {
                    row
                } else {
                    (row as i64 + offset.0) as u32
                };

                Reference::Row {
                    row: new_row,
                    absolute,
                }
            }
        };

        result.validate()?;
        Ok(result)
    }

    // Validate that row/column values are in bounds.
    fn validate(&self) -> Result<(), XlsxError> {
        match self {
            Reference::Cell { row, col, .. } => {
                if *col >= MAX_COLUMNS {
                    return Err(XlsxError::ColumnNumberOverflow);
                }
                if *row >= MAX_ROWS {
                    return Err(XlsxError::RowNumberOverflow);
                }
                Ok(())
            }
            Reference::Column { col, .. } => {
                if *col >= MAX_COLUMNS {
                    return Err(XlsxError::ColumnNumberOverflow);
                }
                Ok(())
            }
            Reference::Row { row, .. } => {
                if *row >= MAX_ROWS {
                    return Err(XlsxError::RowNumberOverflow);
                }
                Ok(())
            }
        }
    }

    // Format a reference to bytes.
    fn format(&self, buf: &mut Vec<u8>) -> Result<(), XlsxError> {
        match self {
            Reference::Cell {
                row,
                col,
                absolute_row,
                absolute_col,
            } => {
                if *absolute_col {
                    buf.push(b'$');
                }
                column_number_to_name(*col, buf)?;
                if *absolute_row {
                    buf.push(b'$');
                }
                buf.extend((row + 1).to_string().into_bytes());
                Ok(())
            }
            Reference::Column { col, absolute } => {
                if *absolute {
                    buf.push(b'$');
                }
                column_number_to_name(*col, buf)
            }
            Reference::Row { row, absolute } => {
                if *absolute {
                    buf.push(b'$');
                }
                buf.extend((row + 1).to_string().into_bytes());
                Ok(())
            }
        }
    }
}

// Advance a reference by the offset (e.g., "A1", "E:F", "5:6", "A1:B5").
fn offset_range(range: &[u8], offset: (i64, i64), buf: &mut Vec<u8>) -> Result<(), XlsxError> {
    let colon_pos = range.iter().position(|&b| b == b':');

    match colon_pos {
        None => {
            let reference = Reference::parse(range)?;
            if !matches!(reference, Reference::Cell { .. }) {
                return Err(XlsxError::Unexpected("Single reference type must be cell"));
            }
            let offset_ref = reference.offset(offset)?;
            offset_ref.format(buf)
        }
        Some(idx) => {
            let start = &range[..idx];
            let end = &range[idx + 1..];

            let start_ref = Reference::parse(start)?;
            let end_ref = Reference::parse(end)?;

            if std::mem::discriminant(&start_ref) != std::mem::discriminant(&end_ref) {
                return Err(XlsxError::Unexpected("Range type mismatch"));
            }

            let start_offset = start_ref.offset(offset)?;
            let end_offset = end_ref.offset(offset)?;

            start_offset.format(buf)?;
            buf.push(b':');
            end_offset.format(buf)
        }
    }
}

// Advance all valid cell names in the string by the offset.
fn replace_cell_names(s: &str, offset: (i64, i64)) -> Result<String, XlsxError> {
    let bytes = s.as_bytes();
    let mut res: Vec<u8> = Vec::new();
    let mut in_quote = false;

    let mut token_start = 0;
    let mut token_end = 0;

    for (i, &c) in bytes.iter().enumerate() {
        if !in_quote && (c.is_ascii_alphanumeric() || c == b'$' || c == b':') {
            token_end = i + 1;
        } else {
            if token_start < token_end
                && offset_range(&bytes[token_start..token_end], offset, &mut res).is_err()
            {
                res.extend(&bytes[token_start..token_end]);
            }
            res.push(c);
            token_start = i + 1;
            token_end = i + 1;

            if c == b'"' {
                in_quote = !in_quote;
            }
        }
    }

    if token_start < token_end
        && offset_range(&bytes[token_start..token_end], offset, &mut res).is_err()
    {
        res.extend(&bytes[token_start..token_end]);
    }

    match String::from_utf8(res) {
        Ok(s) => Ok(s),
        Err(_) => Err(XlsxError::Unexpected("fail to convert cell name")),
    }
}

/// Convert the integer to Excelsheet column title.
/// If the column number not in 1~16384, an Error is returned.
pub(crate) fn column_number_to_name(num: u32, buf: &mut Vec<u8>) -> Result<(), XlsxError> {
    if num >= MAX_COLUMNS {
        return Err(XlsxError::ColumnNumberOverflow);
    }
    let start = buf.len();
    let mut num = num + 1;
    while num > 0 {
        let integer = ((num - 1) % 26 + 65) as u8;
        buf.push(integer);
        num = (num - 1) / 26;
    }
    buf[start..].reverse();
    Ok(())
}

// Convert an Excel Open Packaging "Part" path like "xl/sharedStrings.xml" to
// the equivalent path/filename in the zip file. The file name in the zip file
// may be a case-insensitive version of the target path and may use backslashes.
pub(crate) fn path_to_zip_path<RS: Read + Seek>(zip: &ZipArchive<RS>, path: &str) -> String {
    for zip_path in zip.file_names() {
        let normalized_path = zip_path.replace('\\', "/");

        if path.eq_ignore_ascii_case(&normalized_path) {
            return zip_path.to_string();
        }
    }

    path.to_string()
}

// -----------------------------------------------------------------------
// Unit tests for Xlsx.
// -----------------------------------------------------------------------
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
        let check = |num, expected: &[u8]| {
            let mut buf = Vec::new();
            column_number_to_name(num, &mut buf).unwrap();
            assert_eq!(buf, expected);
        };

        check(0, b"A");
        check(25, b"Z");
        check(26, b"AA");
        check(27, b"AB");
        check(MAX_COLUMNS - 1, b"XFD");
    }

    #[test]
    fn test_parse_reference() {
        let check_cell =
            |input: &[u8], row, col, abs_row, abs_col| match Reference::parse(input).unwrap() {
                Reference::Cell {
                    row: r,
                    col: c,
                    absolute_row: ar,
                    absolute_col: ac,
                } => {
                    assert_eq!((r, c, ar, ac), (row, col, abs_row, abs_col));
                }
                _ => panic!("Expected Cell reference"),
            };

        let check_column = |input: &[u8], col, abs| match Reference::parse(input).unwrap() {
            Reference::Column {
                col: c,
                absolute: a,
            } => {
                assert_eq!((c, a), (col, abs));
            }
            _ => panic!("Expected Column reference"),
        };

        let check_row = |input: &[u8], row, abs| match Reference::parse(input).unwrap() {
            Reference::Row {
                row: r,
                absolute: a,
            } => {
                assert_eq!((r, a), (row, abs));
            }
            _ => panic!("Expected Row reference"),
        };

        // Cell references
        check_cell(b"A1", 0, 0, false, false);
        check_cell(b"$A1", 0, 0, false, true);
        check_cell(b"A$1", 0, 0, true, false);
        check_cell(b"$A$1", 0, 0, true, true);
        check_cell(b"XFD1048576", MAX_ROWS - 1, MAX_COLUMNS - 1, false, false);

        // Column references
        check_column(b"A", 0, false);
        check_column(b"$A", 0, true);
        check_column(b"XFD", MAX_COLUMNS - 1, false);

        // Row references
        check_row(b"1", 0, false);
        check_row(b"$1", 0, true);
        check_row(b"1048576", MAX_ROWS - 1, false);
    }

    #[test]
    fn test_format_reference() {
        let check_cell = |row, col, abs_row, abs_col, expected: &[u8]| {
            let mut buf = Vec::new();
            Reference::Cell {
                row,
                col,
                absolute_row: abs_row,
                absolute_col: abs_col,
            }
            .format(&mut buf)
            .unwrap();
            assert_eq!(buf, expected);
        };

        let check_column = |col, absolute, expected: &[u8]| {
            let mut buf = Vec::new();
            Reference::Column { col, absolute }
                .format(&mut buf)
                .unwrap();
            assert_eq!(buf, expected);
        };

        let check_row = |row, absolute, expected: &[u8]| {
            let mut buf = Vec::new();
            Reference::Row { row, absolute }.format(&mut buf).unwrap();
            assert_eq!(buf, expected);
        };

        // Cell references
        check_cell(0, 0, false, false, b"A1");
        check_cell(0, 0, false, true, b"$A1");
        check_cell(0, 0, true, false, b"A$1");
        check_cell(0, 0, true, true, b"$A$1");
        check_cell(MAX_ROWS - 1, MAX_COLUMNS - 1, false, false, b"XFD1048576");

        // Column references
        check_column(0, false, b"A");
        check_column(0, true, b"$A");
        check_column(MAX_COLUMNS - 1, false, b"XFD");

        // Row references
        check_row(0, false, b"1");
        check_row(0, true, b"$1");
        check_row(MAX_ROWS - 1, false, b"1048576");
    }

    #[test]
    fn test_format_reference_overflow() {
        let check_err = |reference: Reference, offset| {
            let result = reference.offset(offset);
            assert!(
                matches!(
                    result,
                    Err(XlsxError::ColumnNumberOverflow) | Err(XlsxError::RowNumberOverflow)
                ),
                "expected overflow error, got {:?}",
                result
            );
        };

        // Cell reference offset pushes column out of bounds
        check_err(
            Reference::Cell {
                row: 0,
                col: MAX_COLUMNS - 1,
                absolute_row: false,
                absolute_col: false,
            },
            (0, 1),
        );

        // Cell reference offset pushes row out of bounds
        check_err(
            Reference::Cell {
                row: MAX_ROWS - 1,
                col: 0,
                absolute_row: false,
                absolute_col: false,
            },
            (1, 0),
        );

        // Column reference offset pushes out of bounds
        check_err(
            Reference::Column {
                col: MAX_COLUMNS - 1,
                absolute: false,
            },
            (0, 1),
        );

        // Row reference offset pushes out of bounds
        check_err(
            Reference::Row {
                row: MAX_ROWS - 1,
                absolute: false,
            },
            (1, 0),
        );
    }

    #[test]
    fn test_offset_range() {
        let check = |input: &[u8], offset, expected: &[u8]| {
            let mut buf = Vec::new();
            offset_range(input, offset, &mut buf).unwrap();
            assert_eq!(buf, expected);
        };

        let check_err = |input: &[u8], offset| {
            let mut buf = Vec::new();
            let res = offset_range(input, offset, &mut buf);
            assert!(res.is_err());
            assert_eq!(buf.len(), 0)
        };

        // Cell references
        check(b"A1", (1, 1), b"B2");
        check(b"$A1", (1, 1), b"$A2");
        check(b"A$1", (1, 1), b"B$1");
        check(b"$A$1", (1, 1), b"$A$1");

        // Column references
        check_err(b"E", (0, 1));
        check_err(b"$E", (0, 1));

        // Row references
        check_err(b"5", (1, 0));
        check_err(b"$5", (1, 0));

        // Cell ranges
        check(b"A1:B2", (1, 1), b"B2:C3");
        check(b"$A$1:$B$2", (1, 1), b"$A$1:$B$2");

        // Column ranges
        check(b"E:F", (0, 1), b"F:G");
        check(b"$E:$F", (0, 1), b"$E:$F");
        check(b"E:F", (1, 0), b"E:F");

        // Row ranges
        check(b"5:6", (1, 0), b"6:7");
        check(b"$5:$6", (1, 0), b"$5:$6");
        check(b"5:6", (0, 1), b"5:6");
    }

    #[test]
    fn test_parse_reference_overflow() {
        let check_col_err = |input: &[u8]| {
            assert!(matches!(
                Reference::parse(input),
                Err(XlsxError::ColumnNumberOverflow)
            ));
        };
        let check_row_err = |input: &[u8]| {
            assert!(matches!(
                Reference::parse(input),
                Err(XlsxError::RowNumberOverflow)
            ));
        };
        let check_syntax_err = |input: &[u8]| {
            assert!(matches!(
                Reference::parse(input),
                Err(XlsxError::Alphanumeric(_))
            ));
        };

        // Invalid syntax
        check_syntax_err(b"A$A1");
        check_syntax_err(b"A1$2");
        check_syntax_err(b"$$A1");
        check_syntax_err(b"$A$$1");
        check_syntax_err(b"A$$1");
        check_syntax_err(b"1A");
        check_syntax_err(b"1A1");
        check_syntax_err(b"A1B2");

        // Cell references
        check_col_err(b"XFE1");
        check_col_err(b"AAAA1");
        check_row_err(b"A1048577");
        check_row_err(b"A99999999999999999999");
        check_col_err(b"$XFE$1");

        // Column references
        check_col_err(b"XFE");
        check_col_err(b"$XFE");

        // Row references
        check_row_err(b"1048577");
        check_row_err(b"$1048577");
    }

    #[test]
    fn test_offset_range_overflow() {
        let check_col_err = |input: &[u8], offset| {
            let mut buf = Vec::new();
            assert!(matches!(
                offset_range(input, offset, &mut buf),
                Err(XlsxError::ColumnNumberOverflow)
            ));
        };
        let check_row_err = |input: &[u8], offset| {
            let mut buf = Vec::new();
            assert!(matches!(
                offset_range(input, offset, &mut buf),
                Err(XlsxError::RowNumberOverflow)
            ));
        };

        // Original reference is out of bounds
        check_col_err(b"XFE1", (0, 0));
        check_col_err(b"$XFE$1", (0, 0));
        check_row_err(b"A1048577", (0, 0));
        check_row_err(b"$A$1048577", (0, 0));
        check_col_err(b"XFE:XFE", (0, 0));
        check_row_err(b"1048577:1048577", (0, 0));

        // Offset pushes valid cell out of bounds
        check_col_err(b"XFD1", (0, 1));
        check_row_err(b"A1048576", (1, 0));
        check_row_err(b"XFD1048576", (1, 0));
        check_col_err(b"XFD1048576", (0, 1));

        // Offset pushes valid range out of bounds
        check_col_err(b"XFD:XFD", (0, 1));
        check_row_err(b"1048576:1048576", (1, 0));
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
        assert_eq!(
            replace_cell_names("한글 A1 テスト", (0, 1)).unwrap(),
            "한글 B1 テスト".to_owned()
        );

        assert_eq!(
            replace_cell_names("ABC\"asd\"123", (1, 0)).unwrap(),
            "ABC\"asd\"123".to_owned()
        );

        // Column ranges
        assert_eq!(
            replace_cell_names("SUM(E:F)", (0, 1)).unwrap(),
            "SUM(F:G)".to_owned()
        );
        assert_eq!(
            replace_cell_names("SUM($E:$F)", (0, 1)).unwrap(),
            "SUM($E:$F)".to_owned()
        );
        assert_eq!(
            replace_cell_names("SUM($E:F)", (0, 1)).unwrap(),
            "SUM($E:G)".to_owned()
        );

        // Row ranges
        assert_eq!(
            replace_cell_names("SUM(5:6)", (1, 0)).unwrap(),
            "SUM(6:7)".to_owned()
        );
        assert_eq!(
            replace_cell_names("SUM($5:$6)", (1, 0)).unwrap(),
            "SUM($5:$6)".to_owned()
        );
        assert_eq!(
            replace_cell_names("SUM($5:6)", (1, 0)).unwrap(),
            "SUM($5:7)".to_owned()
        );

        // Mixed with cell references
        assert_eq!(
            replace_cell_names("SUM(A1:A5,E:F)", (0, 1)).unwrap(),
            "SUM(B1:B5,F:G)".to_owned()
        );

        // Invalid syntax
        assert_eq!(
            replace_cell_names(
                "Valid: A1 Invalid: A1B1 A1$ $$A1 $A$$1 A$$1 A:1 1:A 1 A A1:1 A1:B A$A1 A1$2 $1 $A Valid: C1:D1",
                (1, 1)
            )
            .unwrap(),
            "Valid: B2 Invalid: A1B1 A1$ $$A1 $A$$1 A$$1 A:1 1:A 1 A A1:1 A1:B A$A1 A1$2 $1 $A Valid: D2:E2"
                .to_owned()
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
            styles: Vec::new(),
            is_1904: false,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
            merged_regions: None,
            options: XlsxOptions::default(),
        };

        assert!(xlsx.read_shared_strings().is_ok());
        assert_eq!(3, xlsx.strings.len());
        assert_eq!("String 1", &xlsx.strings[0]);
        assert_eq!("String 2", &xlsx.strings[1]);
        assert_eq!("String 3", &xlsx.strings[2]);
    }
}
