use std::fs::File;
use std::io::BufReader;
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::reader::Reader;
use quick_xml::events::Event;
use quick_xml::events::attributes::{Attribute, Attributes};

use {DataType, ExcelReader, Range, Cell};
use vba::VbaProject;
use errors::*;

/// A struct representing xml zipped excel file
/// Xlsx, Xlsm, Xlam
pub struct Xlsx {
    zip: ZipArchive<File>,
}

impl Xlsx {
    fn xml_reader<'a>(&'a mut self, path: &str) -> Option<Result<Reader<BufReader<ZipFile<'a>>>>> {
        match self.zip.by_name(path) {
            Ok(f) => {
                let mut r = Reader::from_reader(BufReader::new(f));
                r.check_end_names(false)
                    .trim_text(false)
                    .check_comments(false)
                    .expand_empty_elements(true);
                Some(Ok(r))
            }
            Err(ZipError::FileNotFound) => None,
            Err(e) => return Some(Err(e.into())),
        }
    }
}

impl ExcelReader for Xlsx {
    fn new(f: File) -> Result<Self> {
        Ok(Xlsx { zip: ZipArchive::new(f)? })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        let mut f = self.zip.by_name("xl/vbaProject.bin")?;
        let len = f.size() as usize;
        VbaProject::new(&mut f, len).map(|v| Cow::Owned(v))
    }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        let mut xml = match self.xml_reader("xl/sharedStrings.xml") {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };
        let mut strings = Vec::new();
        let mut buf = Vec::new();
        loop {
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"si" => {
                    if let Some(s) = read_string(&mut xml, e.name())? {
                        strings.push(s);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name() == b"sst" => break,
                Ok(Event::Eof) => return Err("unexpected end of xml (no </sst>)".into()),
                _ => (),
            }
            buf.clear();
        }
        Ok(strings)
    }

    fn read_sheets_names(&mut self,
                         relationships: &HashMap<Vec<u8>, String>)
                         -> Result<Vec<(String, String)>> {
        let mut xml = match self.xml_reader("xl/workbook.xml") {
            None => return Ok(Vec::new()),
            Some(x) => x?,
        };
        let mut sheets = Vec::new();
        let mut buf = Vec::new();
        loop {
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
                    sheets.push((name, path));
                }
                Ok(Event::End(ref e)) if e.local_name() == b"workbook" => break,
                Ok(Event::Eof) => return Err("unexpected end of xml (no </workbook>)".into()),
                Err(e) => return Err(e.into()),
                _ => (),
            }
            buf.clear();
        }
        Ok(sheets)
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let mut xml = match self.xml_reader("xl/_rels/workbook.xml.rels") {
            None => return Err("Cannot find relationships file".into()),
            Some(x) => x?,
        };
        let mut relationships = HashMap::new();
        let mut buf = Vec::new();
        loop {
            match xml.read_event(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match a? {
                            Attribute {
                                key: b"Id",
                                value: v,
                            } => id.extend_from_slice(v),
                            Attribute {
                                key: b"Target",
                                value: v,
                            } => target = xml.decode(v).into_owned(),
                            _ => (),
                        }
                    }
                    relationships.insert(id, target);
                }
                Ok(Event::End(ref e)) if e.local_name() == b"Relationships" => break,
                Ok(Event::Eof) => return Err("unexpected end of xml (no </Relationships>)".into()),
                Err(e) => return Err(e.into()),
                _ => (),
            }
            buf.clear();
        }
        Ok(relationships)
    }

    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        let mut xml = match self.xml_reader(path) {
            None => return Err(format!("Cannot find {} path", path).into()),
            Some(x) => x?,
        };
        let mut cells = Vec::new();
        let mut buf = Vec::new();
        'xml: loop {
            match xml.read_event(&mut buf) {
                Err(e) => return Err(e.into()),
                Ok(Event::Start(ref e)) => {
                    match e.local_name() {
                        b"dimension" => {
                            for a in e.attributes() {
                                if let Attribute {
                                           key: b"ref",
                                           value: rdim,
                                       } = a? {
                                    let (start, end) = get_dimension(rdim)?;
                                    cells.reserve(((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as
                                                  usize);
                                    continue 'xml;
                                }
                            }
                            return Err(format!("Expecting dimension, got {:?}", e).into());
                        }
                        b"sheetData" => read_sheet_data(&mut xml, strings, &mut cells)?,
                        _ => (),
                    }
                }
                Ok(Event::End(ref e)) if e.local_name() == b"worksheet" => break,
                Ok(Event::Eof) => return Err("unexpected end of xml (no </worksheet>)".into()),
                _ => (),
            }
            buf.clear();
        }
        Ok(Range::from_sparse(cells))
    }
}

/// read sheetData node
fn read_sheet_data(xml: &mut Reader<BufReader<ZipFile>>,
                   strings: &[String],
                   cells: &mut Vec<Cell>)
                   -> Result<()> {
    /// read the contents of a <v> cell
    fn read_value<'a>(v: String, strings: &[String], atts: Attributes<'a>) -> Result<DataType> {
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
                    .map_err(Error::from)
                    .or_else::<Error, _>(|_| Ok(DataType::String(v)))
            }
            Some(b"n") => {
                // n - number
                v.parse().map(DataType::Float).map_err(Error::from)
            }
            None => {
                // If type is not known, we try to parse as Float for utility, but fall back to
                // String if this fails.
                v.parse()
                    .map(DataType::Float)
                    .map_err(Error::from)
                    .or_else::<Error, _>(|_| Ok(DataType::String(v)))
            }
            Some(b"is") => {
                // this case should be handled in outer loop over cell elements, in which
                // case read_inline_str is called instead. Case included here for completeness.
                return Err("called read_value on a cell of type inlineStr".into());
            }
            Some(t) => return Err(format!("unknown cell 't' attribute={:?}", t).into()),
        }
    }

    /// search through an Element's attributes for the named one
    fn get_attribute<'a>(atts: Attributes<'a>, n: &'a [u8]) -> Result<Option<&'a [u8]>> {
        for a in atts {
            match a {
                Ok(Attribute { key: k, value: v }) if k == n => return Ok(Some(v)),
                Err(qe) => return Err(qe.into()),
                _ => {} // ignore other attributes
            }
        }
        Ok(None)
    }

    let mut buf = Vec::new();
    /// main content of read_sheet_data
    loop {
        match xml.read_event(&mut buf) {
            Err(e) => return Err(e.into()),
            Ok(Event::Start(ref c_element)) if c_element.local_name() == b"c" => {
                let pos = get_attribute(c_element.attributes(), b"r")
                    .and_then(|o| o.ok_or_else(|| "Cell missing 'r' attribute tag".into()))
                    .and_then(get_row_column)?;

                loop {
                    let mut buf = Vec::new();
                    match xml.read_event(&mut buf) {
                        Err(e) => return Err(e.into()),
                        Ok(Event::Start(ref e)) => {
                            debug!("e: {:?}", e);
                            match e.local_name() {
                                b"is" => {
                                    // inlineStr
                                    if let Some(s) = read_string(xml, e.name())? {
                                        cells.push(Cell::new(pos, DataType::String(s)));
                                    }
                                    break;
                                }
                                b"v" => {
                                    // value
                                    let v = xml.read_text(e.name(), &mut Vec::new())?;
                                    cells.push(Cell::new(pos,
                                                         read_value(v,
                                                                    strings,
                                                                    c_element.attributes())?));
                                    break;
                                }
                                b"f" => {} // ignore f nodes
                                n => {
                                    return Err(format!("not a 'v', 'f', or 'is' node: {:?}", n)
                                                   .into())
                                }
                            }
                        }
                        Ok(Event::End(ref e)) if e.local_name() == b"c" => break,
                        Ok(Event::Eof) => return Err("unexpected end of xml (no </c>)".into()),
                        o => debug!("ignored Event: {:?}", o),
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name() == b"sheetData" => return Ok(()),
            Ok(Event::Eof) => return Err("unexpected end of xml (no </sheetData>)".into()),
            _ => (),
        }
        buf.clear();
    }
}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column),
/// - bottom right (row, column)
fn get_dimension(dimension: &[u8]) -> Result<((u32, u32), (u32, u32))> {
    let parts: Vec<_> = dimension
        .split(|c| *c == b':')
        .map(|s| get_row_column(s))
        .collect::<Result<Vec<_>>>()?;

    match parts.len() {
        0 => Err("dimension cannot be empty".into()),
        1 => Ok((parts[0], parts[0])),
        2 => Ok((parts[0], parts[1])),
        len => Err(format!("range dimension has 0 or 1 ':', got {}", len).into()),
    }
}

/// converts a text range name into its position (row, column) (0 based index)
fn get_row_column(range: &[u8]) -> Result<(u32, u32)> {
    let (mut row, mut col) = (0, 0);
    let mut pow = 1;
    let mut readrow = true;
    for c in range.iter().rev() {
        match *c {
            c @ b'0'...b'9' => {
                if readrow {
                    row += ((c - b'0') as u32) * pow;
                    pow *= 10;
                } else {
                    return Err(format!("Numeric character are only allowed \
                        at the end of the range: {:x}",
                                       c)
                                       .into());
                }
            }
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
            _ => return Err(format!("Expecting alphanumeric character, got {:x}", c).into()),
        }
    }
    Ok((row - 1, col - 1))
}

/// attempts to read either a simple or richtext string
fn read_string(xml: &mut Reader<BufReader<ZipFile>>, closing: &[u8]) -> Result<Option<String>> {
    let mut buf = Vec::new();
    let mut rich_buffer: Option<String> = None;
    loop {
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
                let value = xml.read_text(e.name(), &mut Vec::new())?;
                if let Some(ref mut s) = rich_buffer {
                    s.push_str(&value);
                } else {
                    // consume any remaining events up to expected closing tag
                    xml.read_to_end(closing, &mut Vec::new())?;
                    return Ok(Some(value));
                }
            }
            Ok(Event::Eof) => return Err("unexpected end of xml".into()),
            Err(e) => return Err(e.into()),
            _ => (),
        }
        buf.clear();
    }
}

#[test]
fn test_dimensions() {
    assert_eq!(get_row_column(b"A1").unwrap(), (0, 0));
    assert_eq!(get_row_column(b"C107").unwrap(), (106, 2));
    assert_eq!(get_dimension(b"C2:D35").unwrap(), ((1, 2), (34, 3)));
}
