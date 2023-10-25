use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::BufReader;
use std::io::{Read, Seek};
use std::str::FromStr;

use log::warn;
use nested::Nested;
use quick_xml::events::attributes::{Attribute, Attributes};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::formats::{
    builtin_format_by_id, detect_custom_number_format, format_excel_f64, CellFormat,
};
use crate::vba::VbaProject;
use crate::{
    Cell, CellErrorType, CellType, DataType, Metadata, Range, Reader, Sheet, SheetType,
    SheetVisible, Table,
};

type XlsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

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
    /// ParseInt error
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
}

from_err!(std::io::Error, XlsxError, Io);
from_err!(zip::result::ZipError, XlsxError, Zip);
from_err!(crate::vba::VbaError, XlsxError, Vba);
from_err!(quick_xml::Error, XlsxError, Xml);
from_err!(std::string::ParseError, XlsxError, Parse);
from_err!(std::num::ParseFloatError, XlsxError, ParseFloat);
from_err!(std::num::ParseIntError, XlsxError, ParseInt);

impl std::fmt::Display for XlsxError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsxError::Io(e) => write!(f, "I/O error: {}", e),
            XlsxError::Zip(e) => write!(f, "Zip error: {}", e),
            XlsxError::Xml(e) => write!(f, "Xml error: {}", e),
            XlsxError::XmlAttr(e) => write!(f, "Xml attribute error: {}", e),
            XlsxError::Vba(e) => write!(f, "Vba error: {}", e),
            XlsxError::Parse(e) => write!(f, "Parse string error: {}", e),
            XlsxError::ParseInt(e) => write!(f, "Parse integer error: {}", e),
            XlsxError::ParseFloat(e) => write!(f, "Parse float error: {}", e),

            XlsxError::XmlEof(e) => write!(f, "Unexpected end of xml, expecting '</{}>'", e),
            XlsxError::UnexpectedNode(e) => write!(f, "Expecting '{}' node", e),
            XlsxError::FileNotFound(e) => write!(f, "File not found '{}'", e),
            XlsxError::RelationshipNotFound => write!(f, "Relationship not found"),
            XlsxError::Alphanumeric(e) => {
                write!(f, "Expecting alphanumeric character, got {:X}", e)
            }
            XlsxError::NumericColumn(e) => write!(
                f,
                "Numeric character is not allowed for column name, got {}",
                e
            ),
            XlsxError::DimensionCount(e) => {
                write!(f, "Range dimension must be lower than 2. Got {}", e)
            }
            XlsxError::CellTAttribute(e) => write!(f, "Unknown cell 't' attribute: {:?}", e),
            XlsxError::RangeWithoutColumnComponent => {
                write!(f, "Range is missing the expected column component.")
            }
            XlsxError::RangeWithoutRowComponent => {
                write!(f, "Range is missing the expected row component.")
            }
            XlsxError::Unexpected(e) => write!(f, "{}", e),
            XlsxError::Unrecognized { typ, val } => write!(f, "Unrecognized {}: {}", typ, val),
            XlsxError::CellError(e) => write!(f, "Unsupported cell error value '{}'", e),
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
    strings: Nested<String>,
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
                                name = a.decode_and_unescape_value(&xml)?.to_string();
                            }
                            Attribute {
                                key: QName(b"state"),
                                ..
                            } => {
                                visible = match a.decode_and_unescape_value(&xml)?.as_ref() {
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
                                    format!("xl/{}", r)
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
                            &c.decode_and_unescape_value(&xml)
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
                        let name = a.decode_and_unescape_value(&xml)?.to_string();
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
        for (sheet_name, sheet_path) in &self.sheets {
            let last_folder_index = sheet_path.rfind('/').expect("should be in a folder");
            let (base_folder, file_name) = sheet_path.split_at(last_folder_index);
            let rel_path = format!("{}/_rels{}.rels", base_folder, file_name);

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
                                    let full_path = format!(
                                        "{}{}",
                                        base_folder[..new_index].to_owned(),
                                        target[2..].to_owned()
                                    );
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
            let mut new_tables = Vec::new();
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
                                            xml.decoder().decode(&v)?.into_owned()
                                    }
                                    Attribute {
                                        key: QName(b"ref"),
                                        value: v,
                                    } => {
                                        table_meta.ref_cells =
                                            xml.decoder().decode(&v)?.into_owned()
                                    }
                                    Attribute {
                                        key: QName(b"headerRowCount"),
                                        value: v,
                                    } => {
                                        table_meta.header_row_count =
                                            xml.decoder().decode(&v)?.parse()?
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
                                            xml.decoder().decode(&v)?.parse()?
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
                                    column_names.push(xml.decoder().decode(&v)?.into_owned())
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
            if let Some(tables) = &mut self.tables {
                tables.append(&mut new_tables);
            } else {
                self.tables = Some(new_tables);
            }
        }
        Ok(())
    }

    /// Read pictures
    #[cfg(feature = "picture")]
    fn read_pictures(&mut self) -> Result<(), XlsxError> {
        let mut pics = Vec::new();
        for i in 0..self.zip.len() {
            let mut zfile = self.zip.by_index(i)?;
            let zname = zfile.name().to_owned();
            if zname.starts_with("xl/media") {
                let name_ext: Vec<&str> = zname.split(".").collect();
                if let Some(ext) = name_ext.last() {
                    if [
                        "emf", "wmf", "pict", "jpeg", "jpg", "png", "dib", "gif", "tiff", "eps",
                        "bmp", "wpg",
                    ]
                    .contains(ext)
                    {
                        let mut buf: Vec<u8> = Vec::new();
                        zfile.read_to_end(&mut buf)?;
                        pics.push((ext.to_string(), buf));
                    }
                }
            }
        }
        if !pics.is_empty() {
            self.pictures = Some(pics);
        }
        Ok(())
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

    /// Get the table by name
    // TODO: If retrieving multiple tables from a single sheet, get tables by sheet will be more efficient
    pub fn table_by_name(
        &mut self,
        table_name: &str,
    ) -> Option<Result<Table<DataType>, XlsxError>> {
        let match_table_meta = self
            .tables
            .as_ref()
            .expect("Tables must be loaded before they are referenced")
            .iter()
            .find(|(table, ..)| table == table_name)?;
        let name = match_table_meta.0.to_owned();
        let sheet_name = match_table_meta.1.clone();
        let columns = match_table_meta.2.clone();
        let start_dim = match_table_meta.3.start;
        let end_dim = match_table_meta.3.end;
        let r_range = self.worksheet_range(&sheet_name)?;
        match r_range {
            Ok(range) => {
                let tbl_rng = range.range(start_dim, end_dim);
                Some(Ok(Table {
                    name,
                    sheet_name,
                    columns,
                    data: tbl_rng,
                }))
            }
            Err(e) => Some(Err(e)),
        }
    }
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

fn worksheet<T, F>(
    strings: &Nested<String>,
    formats: &[CellFormat],
    mut xml: XlsReader<'_>,
    read_data: &mut F,
) -> Result<Range<T>, XlsxError>
where
    T: CellType,
    F: FnMut(
        &Nested<String>,
        &[CellFormat],
        &mut XlsReader<'_>,
        &mut Vec<Cell<T>>,
    ) -> Result<(), XlsxError>,
{
    let mut cells = Vec::with_capacity(1024);
    let mut buf = Vec::with_capacity(1024);
    'xml: loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: QName(b"ref"),
                                value: rdim,
                            } = a.map_err(XlsxError::XmlAttr)?
                            {
                                let len = get_dimension(&rdim)?.len();
                                if len < 1_000_000 {
                                    // it is unlikely to have more than that
                                    // there may be of empty cells
                                    cells.reserve(len as usize);
                                }
                                continue 'xml;
                            }
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
                    b"sheetData" => {
                        read_data(strings, formats, &mut xml, &mut cells)?;
                        break;
                    }
                    _ => (),
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }
    Ok(Range::from_sparse(cells))
}

impl<RS: Read + Seek> Reader<RS> for Xlsx<RS> {
    type Error = XlsxError;

    fn new(reader: RS) -> Result<Self, XlsxError> {
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(reader)?,
            strings: Nested::new(),
            formats: Vec::new(),
            is_1904: false,
            sheets: Vec::new(),
            tables: None,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
        };
        xlsx.read_shared_strings()?;
        xlsx.read_styles()?;
        let relationships = xlsx.read_relationships()?;
        xlsx.read_workbook(&relationships)?;
        #[cfg(feature = "picture")]
        xlsx.read_pictures()?;

        Ok(xlsx)
    }

    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, XlsxError>> {
        self.zip.by_name("xl/vbaProject.bin").ok().map(|mut f| {
            let len = f.size() as usize;
            VbaProject::new(&mut f, len)
                .map(Cow::Owned)
                .map_err(XlsxError::Vba)
        })
    }

    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, XlsxError>> {
        let xml = match self.sheets.iter().find(|&(n, _)| n == name) {
            Some((_, path)) => xml_reader(&mut self.zip, path),
            None => return None,
        };
        let is_1904 = self.is_1904;
        let strings = &self.strings;
        let formats = &self.formats;
        xml.map(|xml| {
            worksheet(strings, formats, xml?, &mut |s, f, xml, cells| {
                read_sheet_data(xml, s, f, cells, is_1904)
            })
        })
    }

    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, XlsxError>> {
        let xml = match self.sheets.iter().find(|&(n, _)| n == name) {
            Some((_, path)) => xml_reader(&mut self.zip, path),
            None => return None,
        };

        let strings = &self.strings;
        let formats = &self.formats;
        xml.map(|xml| {
            worksheet(strings, formats, xml?, &mut |_, _, xml, cells| {
                read_sheet(xml, cells, &mut |cells, xml, e, pos, _| {
                    match e.local_name().as_ref() {
                        b"is" | b"v" => {
                            xml.read_to_end_into(e.name(), &mut Vec::new())?;
                        }
                        b"f" => {
                            let mut f_buf = Vec::with_capacity(512);
                            let mut f = String::new();
                            loop {
                                match xml.read_event_into(&mut f_buf)? {
                                    Event::Text(t) => f.push_str(&t.unescape()?),
                                    Event::End(end) if end.name() == e.name() => break,
                                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                                    _ => (),
                                }
                                f_buf.clear();
                            }
                            if !f.is_empty() {
                                cells.push(Cell::new(pos, f));
                            }
                        }
                        _ => return Err(XlsxError::UnexpectedNode("v, f, or is")),
                    }
                    Ok(())
                })
            })
        })
    }

    fn worksheets(&mut self) -> Vec<(String, Range<DataType>)> {
        let is_1904 = self.is_1904;
        self.sheets
            .clone()
            .into_iter()
            .filter_map(|(name, path)| {
                let xml = xml_reader(&mut self.zip, &path)?.ok()?;
                let range = worksheet(
                    &self.strings,
                    &self.formats,
                    xml,
                    &mut |s, f, xml, cells| read_sheet_data(xml, s, f, cells, is_1904),
                )
                .ok()?;
                Some((name, range))
            })
            .collect()
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        self.pictures.to_owned()
    }
}

fn xml_reader<'a, RS: Read + Seek>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
) -> Option<Result<XlsReader<'a>, XlsxError>> {
    match zip.by_name(path) {
        Ok(f) => {
            let mut r = XmlReader::from_reader(BufReader::new(f));
            r.check_end_names(false)
                .trim_text(false)
                .check_comments(false)
                .expand_empty_elements(true);
            Some(Ok(r))
        }
        Err(ZipError::FileNotFound) => None,
        Err(e) => Some(Err(e.into())),
    }
}

/// search through an Element's attributes for the named one
fn get_attribute<'a>(atts: Attributes<'a>, n: QName) -> Result<Option<&'a [u8]>, XlsxError> {
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

fn read_sheet<T, F>(
    xml: &mut XlsReader<'_>,
    cells: &mut Vec<Cell<T>>,
    push_cell: &mut F,
) -> Result<(), XlsxError>
where
    T: CellType,
    F: FnMut(
        &mut Vec<Cell<T>>,
        &mut XlsReader<'_>,
        &BytesStart<'_>,
        (u32, u32),
        &BytesStart<'_>,
    ) -> Result<(), XlsxError>,
{
    let mut buf = Vec::with_capacity(1024);
    let mut cell_buf = Vec::with_capacity(1024);

    let mut row_index = 0;
    let mut col_index = 0;

    loop {
        buf.clear();
        match xml.read_event_into(&mut buf) {
            Ok(Event::Start(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                if let Some(range) = attribute {
                    let row = get_row(range)?;
                    row_index = row;
                }
            }
            Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                row_index += 1;
                col_index = 0;
            }
            Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;

                let pos = if let Some(range) = attribute {
                    let (row, col) = get_row_column(range)?;
                    col_index = col;
                    (row, col)
                } else {
                    (row_index, col_index)
                };

                loop {
                    cell_buf.clear();
                    match xml.read_event_into(&mut cell_buf) {
                        Ok(Event::Start(ref e)) => push_cell(cells, xml, e, pos, c_element)?,
                        Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                }
                col_index += 1;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => return Ok(()),
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }
}

/// read sheetData node
fn read_sheet_data(
    xml: &mut XlsReader<'_>,
    strings: &Nested<String>,
    formats: &[CellFormat],
    cells: &mut Vec<Cell<DataType>>,
    is_1904: bool,
) -> Result<(), XlsxError> {
    /// read the contents of a <v> cell
    fn read_value(
        v: String,
        strings: &Nested<String>,
        formats: &[CellFormat],
        c_element: &BytesStart<'_>,
        is_1904: bool,
    ) -> Result<DataType, XlsxError> {
        let cell_format = match get_attribute(c_element.attributes(), QName(b"s")) {
            Ok(Some(style)) => {
                let id: usize = std::str::from_utf8(style).unwrap_or("0").parse()?;
                formats.get(id)
            }
            _ => Some(&CellFormat::Other),
        };

        match get_attribute(c_element.attributes(), QName(b"t"))? {
            Some(b"s") => {
                // shared string
                let idx: usize = v.parse()?;
                Ok(DataType::String(strings[idx].to_owned()))
            }
            Some(b"b") => {
                // boolean
                Ok(DataType::Bool(v != "0"))
            }
            Some(b"e") => {
                // error
                Ok(DataType::Error(v.parse()?))
            }
            Some(b"d") => {
                // date
                Ok(DataType::DateTimeIso(v))
            }
            Some(b"str") => {
                // see http://officeopenxml.com/SScontentOverview.php
                // str - refers to formula cells
                // * <c .. t='v' .. > indicates calculated value (this case)
                // * <c .. t='f' .. > to the formula string (ignored case
                // TODO: Fully support a DataType::Formula representing both Formula string &
                // last calculated value?
                //
                // NB: the result of a formula may not be a numeric value (=A3&" "&A4).
                // We do try an initial parse as Float for utility, but fall back to a string
                // representation if that fails
                v.parse().map(DataType::Float).or(Ok(DataType::String(v)))
            }
            Some(b"n") => {
                // n - number
                if v.is_empty() {
                    Ok(DataType::Empty)
                } else {
                    v.parse()
                        .map(|n| format_excel_f64(n, cell_format, is_1904))
                        .map_err(XlsxError::ParseFloat)
                }
            }
            None => {
                // If type is not known, we try to parse as Float for utility, but fall back to
                // String if this fails.
                v.parse()
                    .map(|n| format_excel_f64(n, cell_format, is_1904))
                    .or(Ok(DataType::String(v)))
            }
            Some(b"is") => {
                // this case should be handled in outer loop over cell elements, in which
                // case read_inline_str is called instead. Case included here for completeness.
                Err(XlsxError::Unexpected(
                    "called read_value on a cell of type inlineStr",
                ))
            }
            Some(t) => {
                let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
                Err(XlsxError::CellTAttribute(t))
            }
        }
    }

    read_sheet(xml, cells, &mut |cells, xml, e, pos, c_element| {
        match e.local_name().as_ref() {
            b"is" => {
                // inlineStr
                if let Some(s) = read_string(xml, e.name())? {
                    cells.push(Cell::new(pos, DataType::String(s)));
                }
            }
            b"v" => {
                // value
                let mut v = String::new();
                let mut v_buf = Vec::new();
                loop {
                    v_buf.clear();
                    match xml.read_event_into(&mut v_buf)? {
                        Event::Text(t) => v.push_str(&t.unescape()?),
                        Event::End(end) if end.name() == e.name() => break,
                        Event::Eof => return Err(XlsxError::XmlEof("v")),
                        _ => (),
                    }
                }
                match read_value(v, strings, formats, c_element, is_1904)? {
                    DataType::Empty => (),
                    v => cells.push(Cell::new(pos, v)),
                }
            }
            b"f" => {
                xml.read_to_end_into(e.name(), &mut Vec::new())?;
            }
            _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
        }
        Ok(())
    })
}

#[derive(Debug, PartialEq)]
struct Dimensions {
    start: (u32, u32),
    end: (u32, u32),
}

impl Dimensions {
    fn len(&self) -> u64 {
        (self.end.0 - self.start.0 + 1) as u64 * (self.end.1 - self.start.1 + 1) as u64
    }
}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column),
/// - bottom right (row, column)
fn get_dimension(dimension: &[u8]) -> Result<Dimensions, XlsxError> {
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
                warn!(
                    "xlsx has more than maximum number of rows ({} > {})",
                    rows, MAX_ROWS
                );
            }
            if columns > MAX_COLUMNS {
                warn!(
                    "xlsx has more than maximum number of columns ({} > {})",
                    columns, MAX_COLUMNS
                );
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
fn get_row_column(range: &[u8]) -> Result<(u32, u32), XlsxError> {
    let (row, col) = get_row_and_optional_column(range)?;
    let col = col.ok_or(XlsxError::RangeWithoutColumnComponent)?;
    Ok((row, col))
}

/// Converts a text row name into its position (0 based index).
/// If the row component in the range is missing, an Error is returned.
/// If the text row name also contains a column component, it is ignored.
fn get_row(range: &[u8]) -> Result<u32, XlsxError> {
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
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'A') as u32 + 1) * pow;
                pow *= 26;
            }
            c @ b'a'..=b'z' => {
                if readrow {
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
fn read_string(
    xml: &mut XlsReader<'_>,
    QName(closing): QName,
) -> Result<Option<String>, XlsxError> {
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
            Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => {
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
                    xml.read_to_end_into(QName(closing), &mut val_buf)?;
                    return Ok(Some(value));
                }
            }
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }
}

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
