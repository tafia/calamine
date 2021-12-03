use std::borrow::Cow;
use std::collections::HashMap;
use std::io::BufReader;
use std::io::{Read, Seek};
use std::str::FromStr;

use log::warn;
use quick_xml::events::attributes::{Attribute, Attributes};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::vba::VbaProject;
use crate::{Cell, CellErrorType, DataType, Metadata, Range, Reader, Table};

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
    /// Expecting alphanumeri character
    Alphanumeric(u8),
    /// Numeric column
    NumericColumn(u8),
    /// Wrong dimension count
    DimensionCount(usize),
    /// Cell 't' attribute error
    CellTAttribute(String),
    /// Cell 'r' attribute error
    CellRAttribute,
    /// Unexpected error
    Unexpected(&'static str),
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
            XlsxError::Vba(e) => write!(f, "Vba error: {}", e),
            XlsxError::Parse(e) => write!(f, "Parse string error: {}", e),
            XlsxError::ParseInt(e) => write!(f, "Parse integer error: {}", e),
            XlsxError::ParseFloat(e) => write!(f, "Parse float error: {}", e),

            XlsxError::XmlEof(e) => write!(f, "Unexpected end of xml, expecting '</{}>'", e),
            XlsxError::UnexpectedNode(e) => write!(f, "Expecting '{}' node", e),
            XlsxError::FileNotFound(e) => write!(f, "File not found '{}'", e),
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
            XlsxError::CellRAttribute => write!(f, "Cell missing 'r' attribute"),
            XlsxError::Unexpected(e) => write!(f, "{}", e),
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

#[derive(Debug)]
enum CellFormat {
    Other,
    Date,
}

/// A struct representing xml zipped excel file
/// Xlsx, Xlsm, Xlam
pub struct Xlsx<RS>
where
    RS: Read + Seek,
{
    zip: ZipArchive<RS>,
    /// Shared strings
    strings: Vec<String>,
    /// Sheets paths
    sheets: Vec<(String, String)>,
    /// Tables: Name, Sheet, Columns, Data dimensions
    tables: Option<Vec<(String, String, Vec<String>, Dimensions)>>,
    /// Cell (number) formats
    formats: Vec<CellFormat>,
    /// Metadata
    metadata: Metadata,
}

impl<RS: Read + Seek> Xlsx<RS> {
    fn read_shared_strings(&mut self) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/sharedStrings.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"si" => {
                    if let Some(s) = read_string(&mut xml, e.name())? {
                        self.strings.push(s);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name() == b"sst" => break,
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

        let mut number_formats = HashMap::new();

        let mut buf = Vec::new();
        let mut inner_buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"numFmts" => loop {
                    inner_buf.clear();
                    match xml.read_event(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name() == b"numFmt" => {
                            let mut id = Vec::new();
                            let mut format = String::new();
                            for a in e.attributes() {
                                match a? {
                                    Attribute {
                                        key: b"numFmtId",
                                        value: v,
                                    } => id.extend_from_slice(&v),
                                    Attribute {
                                        key: b"formatCode",
                                        value: v,
                                    } => format = xml.decode(&v).into_owned(),
                                    _ => (),
                                }
                            }
                            number_formats.insert(id, format);
                        }
                        Ok(Event::End(ref e)) if e.local_name() == b"numFmts" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("numFmts")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.local_name() == b"cellXfs" => loop {
                    inner_buf.clear();
                    match xml.read_event(&mut inner_buf) {
                        Ok(Event::Start(ref e)) if e.local_name() == b"xf" => {
                            self.formats.push(
                                e.attributes()
                                    .filter_map(|a| a.ok())
                                    .find(|a| a.key == b"numFmtId")
                                    .map_or(CellFormat::Other, |a| {
                                        match number_formats.get(&*a.value) {
                                            Some(fmt) if is_custom_date_format(fmt) => {
                                                CellFormat::Date
                                            }
                                            None if is_builtin_date_format_id(&a.value) => {
                                                CellFormat::Date
                                            }
                                            _ => CellFormat::Other,
                                        }
                                    }),
                            );
                        }
                        Ok(Event::End(ref e)) if e.local_name() == b"cellXfs" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("cellXfs")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::End(ref e)) if e.local_name() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("styleSheet")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    fn read_workbook(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/workbook.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut defined_names = Vec::new();
        let mut buf = Vec::new();
        let mut val_buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"sheet" => {
                    let mut name = String::new();
                    let mut path = String::new();
                    for a in e.attributes() {
                        let a = a?;
                        match a {
                            Attribute { key: b"name", .. } => {
                                name = a.unescape_and_decode_value(&xml)?;
                            }
                            Attribute {
                                key: b"r:id",
                                value: v,
                            }
                            | Attribute {
                                key: b"relationships:id",
                                value: v,
                            } => {
                                let r = &relationships[&*v][..];
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
                    self.metadata.sheets.push(name.to_string());
                    self.sheets.push((name, path));
                }
                Ok(Event::Start(ref e)) if e.local_name() == b"definedName" => {
                    if let Some(a) = e
                        .attributes()
                        .filter_map(|a| a.ok())
                        .find(|a| a.key == b"name")
                    {
                        let name = a.unescape_and_decode_value(&xml)?;
                        val_buf.clear();
                        let value = xml.read_text(e.name(), &mut val_buf)?;
                        defined_names.push((name, value));
                    }
                }
                Ok(Event::End(ref e)) if e.local_name() == b"workbook" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("workbook")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        self.metadata.names = defined_names;
        Ok(())
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            None => {
                return Err(XlsxError::FileNotFound(
                    "xl/_rels/workbook.xml.rels".to_string(),
                ));
            }
            Some(x) => x?,
        };
        let mut relationships = HashMap::new();
        let mut buf = Vec::new();
        loop {
            buf.clear();
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match a? {
                            Attribute {
                                key: b"Id",
                                value: v,
                            } => id.extend_from_slice(&v),
                            Attribute {
                                key: b"Target",
                                value: v,
                            } => target = xml.decode(&v).into_owned(),
                            _ => (),
                        }
                    }
                    relationships.insert(id, target);
                }
                Ok(Event::End(ref e)) if e.local_name() == b"Relationships" => break,
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
            let last_folder_index = sheet_path.rfind("/").expect("should be in a folder");
            let (base_folder, file_name) = sheet_path.split_at(last_folder_index);
            let rel_path = format!("{}/_rels{}.rels", base_folder, file_name);

            let mut table_locations = Vec::new();
            let mut buf = Vec::new();
            // we need another mutable borrow of self.zip later so we enclose this borrow within braces
            {
                let mut xml = match xml_reader(&mut self.zip, &rel_path) {
                    None => continue,
                    Some(x) => x?,
                };
                loop {
                    buf.clear();
                    match xml.read_event(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name() == b"Relationship" => {
                            let mut id = Vec::new();
                            let mut target = String::new();
                            let mut table_type = false;
                            for a in e.attributes() {
                                match a? {
                                Attribute {
                                    key: b"Id",
                                    value: v,
                                } => id.extend_from_slice(&v),
                                Attribute {
                                    key: b"Target",
                                    value: v,
                                } => target = xml.decode(&v).into_owned(),
                                Attribute {
                                    key: b"Type",
                                    value: v,
                                } => table_type = *v == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"[..],
                                _ => (),
                            }
                            }
                            if table_type {
                                if target.starts_with("../") {
                                    // this is an incomplete implementation, but should be good enough for excel
                                    let new_index =
                                        base_folder.rfind("/").expect("Must be a parent folder");
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
                        Ok(Event::End(ref e)) if e.local_name() == b"Relationships" => break,
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
                    match xml.read_event(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name() == b"table" => {
                            for a in e.attributes() {
                                match a? {
                                    Attribute {
                                        key: b"displayName",
                                        value: v,
                                    } => table_meta.display_name = xml.decode(&v).into_owned(),
                                    Attribute {
                                        key: b"ref",
                                        value: v,
                                    } => table_meta.ref_cells = xml.decode(&v).into_owned(),
                                    Attribute {
                                        key: b"headerRowCount",
                                        value: v,
                                    } => table_meta.header_row_count = xml.decode(&v).parse()?,
                                    Attribute {
                                        key: b"insertRow",
                                        value: v,
                                    } => table_meta.insert_row = *v != b"0"[..],
                                    Attribute {
                                        key: b"totalsRowCount",
                                        value: v,
                                    } => table_meta.totals_row_count = xml.decode(&v).parse()?,
                                    _ => (),
                                }
                            }
                        }
                        Ok(Event::Start(ref e)) if e.local_name() == b"tableColumn" => {
                            for a in e.attributes() {
                                match a? {
                                    Attribute {
                                        key: b"name",
                                        value: v,
                                    } => column_names.push(xml.decode(&v).into_owned()),
                                    _ => (),
                                }
                            }
                        }
                        Ok(Event::End(ref e)) if e.local_name() == b"table" => break,
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
    strings: &[String],
    formats: &[CellFormat],
    mut xml: XlsReader<'_>,
    read_data: &mut F,
) -> Result<Range<T>, XlsxError>
where
    T: Default + Clone + PartialEq,
    F: FnMut(
        &[String],
        &[CellFormat],
        &mut XlsReader<'_>,
        &mut Vec<Cell<T>>,
    ) -> Result<(), XlsxError>,
{
    let mut cells = Vec::new();
    let mut buf = Vec::new();
    'xml: loop {
        buf.clear();
        match xml.read_event(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: b"ref",
                                value: rdim,
                            } = a?
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
                        read_data(&strings, &formats, &mut xml, &mut cells)?;
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

impl<RS: Read + Seek> Reader for Xlsx<RS> {
    type RS = RS;
    type Error = XlsxError;

    fn new(reader: RS) -> Result<Self, XlsxError>
    where
        RS: Read + Seek,
    {
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(reader)?,
            strings: Vec::new(),
            formats: Vec::new(),
            sheets: Vec::new(),
            tables: None,
            metadata: Metadata::default(),
        };
        xlsx.read_shared_strings()?;
        xlsx.read_styles()?;
        let relationships = xlsx.read_relationships()?;
        xlsx.read_workbook(&relationships)?;
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
        let xml = match self.sheets.iter().find(|&&(ref n, _)| n == name) {
            Some(&(_, ref path)) => xml_reader(&mut self.zip, path),
            None => return None,
        };
        let strings = &self.strings;
        let formats = &self.formats;
        xml.map(|xml| {
            worksheet(strings, formats, xml?, &mut |s, f, xml, cells| {
                read_sheet_data(xml, s, f, cells)
            })
        })
    }

    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, XlsxError>> {
        let xml = match self.sheets.iter().find(|&&(ref n, _)| n == name) {
            Some(&(_, ref path)) => xml_reader(&mut self.zip, path),
            None => return None,
        };

        let strings = &self.strings;
        let formats = &self.formats;
        xml.map(|xml| {
            worksheet(strings, formats, xml?, &mut |_, _, xml, cells| {
                read_sheet(xml, cells, &mut |cells, xml, e, pos, _| {
                    match e.local_name() {
                        b"is" | b"v" => xml.read_to_end(e.name(), &mut Vec::new())?,
                        b"f" => {
                            let f = xml.read_text(e.name(), &mut Vec::new())?;
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
        self.sheets
            .clone()
            .into_iter()
            .filter_map(|(name, path)| {
                let xml = xml_reader(&mut self.zip, &path)?.ok()?;
                let range = worksheet(
                    &self.strings,
                    &self.formats,
                    xml,
                    &mut |s, f, xml, cells| read_sheet_data(xml, s, f, cells),
                )
                .ok()?;
                Some((name, range))
            })
            .collect()
    }
}

fn xml_reader<'a, RS>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
) -> Option<Result<XlsReader<'a>, XlsxError>>
where
    RS: Read + Seek,
{
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
fn get_attribute<'a>(atts: Attributes<'a>, n: &[u8]) -> Result<Option<&'a [u8]>, XlsxError> {
    for a in atts {
        match a {
            Ok(Attribute {
                key,
                value: Cow::Borrowed(value),
            }) if key == n => return Ok(Some(value)),
            Err(e) => return Err(XlsxError::Xml(e)),
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
    T: Clone + Default + PartialEq,
    F: FnMut(
        &mut Vec<Cell<T>>,
        &mut XlsReader<'_>,
        &BytesStart<'_>,
        (u32, u32),
        &BytesStart<'_>,
    ) -> Result<(), XlsxError>,
{
    let mut buf = Vec::new();
    let mut cell_buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event(&mut buf) {
            Ok(Event::Start(ref c_element)) if c_element.local_name() == b"c" => {
                let pos = get_attribute(c_element.attributes(), b"r")
                    .and_then(|o| o.ok_or(XlsxError::CellRAttribute))
                    .and_then(get_row_column)?;
                loop {
                    cell_buf.clear();
                    match xml.read_event(&mut cell_buf) {
                        Ok(Event::Start(ref e)) => push_cell(cells, xml, e, pos, c_element)?,
                        Ok(Event::End(ref e)) if e.local_name() == b"c" => break,
                        Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => (),
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name() == b"sheetData" => return Ok(()),
            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => (),
        }
    }
}

/// read sheetData node
fn read_sheet_data(
    xml: &mut XlsReader<'_>,
    strings: &[String],
    formats: &[CellFormat],
    cells: &mut Vec<Cell<DataType>>,
) -> Result<(), XlsxError> {
    /// read the contents of a <v> cell
    fn read_value<'a>(
        v: String,
        strings: &[String],
        formats: &[CellFormat],
        c_element: &BytesStart<'a>,
    ) -> Result<DataType, XlsxError> {
        let is_date_time = match get_attribute(c_element.attributes(), b"s") {
            Ok(Some(style)) => {
                let id: usize = std::str::from_utf8(style).unwrap_or("0").parse()?;
                match formats.get(id) {
                    Some(CellFormat::Date) => true,
                    _ => false,
                }
            }
            _ => false,
        };

        match get_attribute(c_element.attributes(), b"t")? {
            Some(b"s") => {
                // shared string
                let idx: usize = v.parse()?;
                Ok(DataType::String(strings[idx].clone()))
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
                // TODO: create a DataType::Date
                // currently just return as string (ISO 8601)
                Ok(DataType::String(v))
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
                        .map(|n| {
                            if is_date_time {
                                DataType::DateTime(n)
                            } else {
                                DataType::Float(n)
                            }
                        })
                        .map_err(XlsxError::ParseFloat)
                }
            }
            None => {
                // If type is not known, we try to parse as Float for utility, but fall back to
                // String if this fails.
                v.parse()
                    .map(|n| {
                        if is_date_time {
                            DataType::DateTime(n)
                        } else {
                            DataType::Float(n)
                        }
                    })
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
        match e.local_name() {
            b"is" => {
                // inlineStr
                if let Some(s) = read_string(xml, e.name())? {
                    cells.push(Cell::new(pos, DataType::String(s)));
                }
            }
            b"v" => {
                // value
                let v = xml.read_text(e.name(), &mut Vec::new())?;
                match read_value(v, strings, formats, c_element)? {
                    DataType::Empty => (),
                    v => cells.push(Cell::new(pos, v)),
                }
            }
            b"f" => xml.read_to_end(e.name(), &mut Vec::new())?,
            _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
        }
        Ok(())
    })
}

// This tries to detect number formats that are definitely date/time formats.
// This is definitely not perfect!
fn is_custom_date_format(format: &str) -> bool {
    format.bytes().all(|c| b"mdyMDYhsHS-/.: \\".contains(&c))
}

fn is_builtin_date_format_id(id: &[u8]) -> bool {
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
    // [h]:mm:ss
    b"46" |
    // mmss.0
    b"47" => true,
    _ => false
    }
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
        .map(|s| get_row_column(s))
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

/// converts a text range name into its position (row, column) (0 based index)
fn get_row_column(range: &[u8]) -> Result<(u32, u32), XlsxError> {
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
    Ok((row - 1, col - 1))
}

/// attempts to read either a simple or richtext string
fn read_string(xml: &mut XlsReader<'_>, closing: &[u8]) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut val_buf = Vec::new();
    let mut rich_buffer: Option<String> = None;
    let mut is_phonetic_text = false;
    loop {
        buf.clear();
        match xml.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name() == b"r" => {
                if rich_buffer.is_none() {
                    // use a buffer since richtext has multiples <r> and <t> for the same cell
                    rich_buffer = Some(String::new());
                }
            }
            Ok(Event::Start(ref e)) if e.local_name() == b"rPh" => {
                is_phonetic_text = true;
            }
            Ok(Event::End(ref e)) if e.local_name() == closing => {
                return Ok(rich_buffer);
            }
            Ok(Event::End(ref e)) if e.local_name() == b"rPh" => {
                is_phonetic_text = false;
            }
            Ok(Event::Start(ref e)) if e.local_name() == b"t" && !is_phonetic_text => {
                val_buf.clear();
                let value = xml.read_text(e.name(), &mut val_buf)?;
                if let Some(ref mut s) = rich_buffer {
                    s.push_str(&value);
                } else {
                    // consume any remaining events up to expected closing tag
                    xml.read_to_end(closing, &mut val_buf)?;
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
