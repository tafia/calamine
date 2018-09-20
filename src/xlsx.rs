use std::borrow::Cow;
use std::collections::HashMap;
use std::io::BufReader;
use std::io::{Read, Seek};
use std::str::FromStr;

use quick_xml::events::attributes::{Attribute, Attributes};
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use vba::VbaProject;
use {Cell, CellErrorType, DataType, Metadata, Range, Reader};

type XlsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

/// An enum for Xlsx specific errors
#[derive(Debug, Fail)]
pub enum XlsxError {
    /// Io error
    #[fail(display = "{}", _0)]
    Io(#[cause] ::std::io::Error),
    /// Zip error
    #[fail(display = "{}", _0)]
    Zip(#[cause] ::zip::result::ZipError),
    /// Vba error
    #[fail(display = "{}", _0)]
    Vba(#[cause] ::vba::VbaError),
    /// Xml error
    #[fail(display = "{}", _0)]
    Xml(#[cause] ::quick_xml::Error),
    /// Parse error
    #[fail(display = "{}", _0)]
    Parse(#[cause] ::std::string::ParseError),
    /// Float error
    #[fail(display = "{}", _0)]
    ParseFloat(#[cause] ::std::num::ParseFloatError),
    /// ParseInt error
    #[fail(display = "{}", _0)]
    ParseInt(#[cause] ::std::num::ParseIntError),

    /// Unexpected end of xml
    #[fail(display = "Unexpected end of xml, expecting '</{}>'", _0)]
    XmlEof(&'static str),
    /// Unexpected node
    #[fail(display = "Expecting '{}' node", _0)]
    UnexpectedNode(&'static str),
    /// File not found
    #[fail(display = "File not found '{}'", _0)]
    FileNotFound(String),
    /// Expecting alphanumeri character
    #[fail(display = "Expecting alphanumeric character, got {:X}", _0)]
    Alphanumeric(u8),
    /// Numeric column
    #[fail(
        display = "Numeric character is not allowed for column name, got {}",
        _0
    )]
    NumericColumn(u8),
    /// Wrong dimension count
    #[fail(display = "Range dimension must be lower than 2. Got {}", _0)]
    DimensionCount(usize),
    /// Cell 't' attribute error
    #[fail(display = "Unknown cell 't' attribute: {:?}", _0)]
    CellTAttribute(String),
    /// Cell 'r' attribute error
    #[fail(display = "Cell missing 'r' attribute")]
    CellRAttribute,
    /// Unexpected error
    #[fail(display = "{}", _0)]
    Unexpected(&'static str),
    /// Cell error
    #[fail(display = "Unsupported cell error value '{}'", _0)]
    CellError(String),
}

from_err!(::std::io::Error, XlsxError, Io);
from_err!(::zip::result::ZipError, XlsxError, Zip);
from_err!(::vba::VbaError, XlsxError, Vba);
from_err!(::quick_xml::Error, XlsxError, Xml);
from_err!(::std::string::ParseError, XlsxError, Parse);
from_err!(::std::num::ParseFloatError, XlsxError, ParseFloat);
from_err!(::std::num::ParseIntError, XlsxError, ParseInt);

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
            _ => return Err(XlsxError::CellError(s.into())),
        }
    }
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
                    self.sheets.push((name, path));
                }
                Ok(Event::Start(ref e)) if e.local_name() == b"definedName" => if let Some(a) = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .find(|a| a.key == b"name")
                {
                    let name = a.unescape_and_decode_value(&xml)?;
                    val_buf.clear();
                    let value = xml.read_text(b"definedName", &mut val_buf)?;
                    defined_names.push((name, value));
                },
                Ok(Event::End(ref e)) if e.local_name() == b"workbook" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("workbook")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        self.metadata.names = defined_names;
        self.metadata.sheets = self.sheets.iter().map(|&(ref s, _)| s.clone()).collect();
        Ok(())
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>, XlsxError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            None => {
                return Err(XlsxError::FileNotFound(
                    "xl/_rels/workbook.xml.rels".to_string(),
                ))
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
}

fn worksheet<T, F>(
    strings: &[String],
    mut xml: XlsReader,
    read_data: &mut F,
) -> Result<Range<T>, XlsxError>
where
    T: Default + Clone + PartialEq,
    F: FnMut(&[String], &mut XlsReader, &mut Vec<Cell<T>>) -> Result<(), XlsxError>,
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
                                let (start, end) = get_dimension(&rdim)?;
                                let len = (end.0 - start.0 + 1) * (end.1 - start.1 + 1);
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
                        read_data(&strings, &mut xml, &mut cells)?;
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
            sheets: Vec::new(),
            metadata: Metadata::default(),
        };
        xlsx.read_shared_strings()?;
        let relationships = xlsx.read_relationships()?;
        xlsx.read_workbook(&relationships)?;
        Ok(xlsx)
    }

    fn vba_project(&mut self) -> Option<Result<Cow<VbaProject>, XlsxError>> {
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
        xml.map(|xml| {
            worksheet(strings, xml?, &mut |s, xml, cells| {
                read_sheet_data(xml, s, cells)
            })
        })
    }

    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, XlsxError>> {
        let xml = match self.sheets.iter().find(|&&(ref n, _)| n == name) {
            Some(&(_, ref path)) => xml_reader(&mut self.zip, path),
            None => return None,
        };

        let strings = &self.strings;
        xml.map(|xml| {
            worksheet(strings, xml?, &mut |_, xml, cells| {
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
    use std::borrow::Cow;
    for a in atts {
        match a {
            Ok(Attribute {
                key,
                value: Cow::Borrowed(value),
            })
                if key == n =>
            {
                return Ok(Some(value))
            }
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {} // ignore other attributes
        }
    }
    Ok(None)
}

fn read_sheet<T, F>(
    xml: &mut XlsReader,
    cells: &mut Vec<Cell<T>>,
    push_cell: &mut F,
) -> Result<(), XlsxError>
where
    T: Clone + Default + PartialEq,
    F: FnMut(&mut Vec<Cell<T>>, &mut XlsReader, &BytesStart, (u32, u32), &BytesStart)
        -> Result<(), XlsxError>,
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
    xml: &mut XlsReader,
    strings: &[String],
    cells: &mut Vec<Cell<DataType>>,
) -> Result<(), XlsxError> {
    /// read the contents of a <v> cell
    fn read_value<'a>(
        v: String,
        strings: &[String],
        atts: Attributes<'a>,
    ) -> Result<DataType, XlsxError> {
        match get_attribute(atts, b"t")? {
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
                v.parse()
                    .map(DataType::Float)
                    .map_err(XlsxError::ParseFloat)
                    .or_else::<XlsxError, _>(|_| Ok(DataType::String(v)))
            }
            Some(b"n") => {
                // n - number
                v.parse()
                    .map(DataType::Float)
                    .map_err(XlsxError::ParseFloat)
            }
            None => {
                // If type is not known, we try to parse as Float for utility, but fall back to
                // String if this fails.
                v.parse()
                    .map(DataType::Float)
                    .map_err(XlsxError::ParseFloat)
                    .or_else::<XlsxError, _>(|_| Ok(DataType::String(v)))
            }
            Some(b"is") => {
                // this case should be handled in outer loop over cell elements, in which
                // case read_inline_str is called instead. Case included here for completeness.
                return Err(XlsxError::Unexpected(
                    "called read_value on a cell of type inlineStr",
                ));
            }
            Some(t) => {
                let t = ::std::str::from_utf8(t)
                    .unwrap_or("<utf8 error>")
                    .to_string();
                return Err(XlsxError::CellTAttribute(t));
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
                match read_value(v, strings, c_element.attributes())? {
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

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column),
/// - bottom right (row, column)
fn get_dimension(dimension: &[u8]) -> Result<((u32, u32), (u32, u32)), XlsxError> {
    let parts: Vec<_> = dimension
        .split(|c| *c == b':')
        .map(|s| get_row_column(s))
        .collect::<Result<Vec<_>, XlsxError>>()?;

    match parts.len() {
        0 => Err(XlsxError::DimensionCount(0)),
        1 => Ok((parts[0], parts[0])),
        2 => Ok((parts[0], parts[1])),
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
            c @ b'0'...b'9' => if readrow {
                row += ((c - b'0') as u32) * pow;
                pow *= 10;
            } else {
                return Err(XlsxError::NumericColumn(c));
            },
            c @ b'A'...b'Z' => {
                if readrow {
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'A') as u32 + 1) * pow;
                pow *= 26;
            }
            c @ b'a'...b'z' => {
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
fn read_string(xml: &mut XlsReader, closing: &[u8]) -> Result<Option<String>, XlsxError> {
    let mut buf = Vec::new();
    let mut val_buf = Vec::new();
    let mut rich_buffer: Option<String> = None;
    loop {
        buf.clear();
        match xml.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name() == b"r" => {
                if rich_buffer.is_none() {
                    // use a buffer since richtext has multiples <r> and <t> for the same cell
                    rich_buffer = Some(String::new());
                }
            }
            Ok(Event::End(ref e)) if e.local_name() == closing => {
                return Ok(rich_buffer);
            }
            Ok(Event::Start(ref e)) if e.local_name() == b"t" => {
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
    assert_eq!(get_dimension(b"C2:D35").unwrap(), ((1, 2), (34, 3)));
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
