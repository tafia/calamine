// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

#![warn(missing_docs)]

mod cells_reader;

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
use crate::utils::unescape_entity_to_buffer;
use crate::vba::VbaProject;
use crate::{
    Cell, CellErrorType, Data, Dimensions, HeaderRow, Metadata, Range, Reader, ReaderRef, Sheet,
    SheetType, SheetVisible, Table,
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
    /// errors related to missing attributes in XML elements.
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

    /// A wrapper for a variety of [`quick_xml::events::attributes::AttrError`]
    /// errors related to XML attributes.
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
    /// Cell (number) formats
    formats: Vec<CellFormat>,
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

        let mut buf = Vec::with_capacity(1024);
        let mut inner_buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                            let mut id = Vec::new();
                            let mut format = String::new();
                            for a in e.attributes() {
                                match a.map_err(XlsxError::XmlAttr)? {
                                    Attribute {
                                        key: QName(b"numFmtId"),
                                        value: v,
                                    } => id.extend_from_slice(&v),
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
                },
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => loop {
                    inner_buf.clear();
                    match xml.read_event_into(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                            self.formats.push(
                                e.attributes()
                                    .filter_map(|a| a.ok())
                                    .find(|a| a.key == QName(b"numFmtId"))
                                    .map_or(CellFormat::Other, |a| {
                                        match number_formats.get(&*a.value) {
                                            Some(fmt) => detect_custom_number_format(fmt),
                                            None => builtin_format_by_id(&a.value),
                                        }
                                    }),
                            );
                        }
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("cellXfs")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
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
        let formats = &self.formats;
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
                        Event::Text(t) => value.push_str(&t.xml10_content()?),
                        Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut value)?,
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
fn offset_cell_name(name: &[u8], offset: (i64, i64)) -> Result<Vec<u8>, XlsxError> {
    let cell = get_row_column(name.to_vec().as_slice())?;
    coordinate_to_name((
        (cell.0 as i64 + offset.0) as u32,
        (cell.1 as i64 + offset.1) as u32,
    ))
}

/// advance all valid cell names in the string by the offset
fn replace_cell_names(s: &str, offset: (i64, i64)) -> Result<String, XlsxError> {
    let mut res: Vec<u8> = Vec::new();
    let mut cell: Vec<u8> = Vec::new();
    let mut is_cell_row = false;
    let mut in_quote = false;
    for c in s.bytes() {
        if c == b'"' {
            in_quote = !in_quote;
        }
        if in_quote {
            res.push(c);
            continue;
        }
        if c.is_ascii_alphabetic() {
            if is_cell_row {
                // two cell not possible stick together in formula
                res.extend(cell.iter().copied());
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
                res.extend(cell.iter().copied());
            }
            cell.clear();
            is_cell_row = false;
            res.push(c);
        }
    }
    if !cell.is_empty() {
        if let Ok(cell_name) = offset_cell_name(cell.as_ref(), offset) {
            res.extend(cell_name);
        } else {
            res.extend(cell.iter().copied());
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
        assert_eq!(
            replace_cell_names("한글 A1 テスト", (0, 1)).unwrap(),
            "한글 B1 テスト".to_owned()
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
