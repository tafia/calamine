use std::io::BufReader;
use std::io::{Read, Seek};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;
use quick_xml::reader::Reader as XmlReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::events::attributes::{Attribute, Attributes};

use {Cell, DataType, Metadata, Range, Reader};
use vba::VbaProject;
use errors::*;

type XlsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

/// A struct representing xml zipped excel file
/// Xlsx, Xlsm, Xlam
pub struct Xlsx<RS> where RS: Read + Seek {
    zip: ZipArchive<RS>,
    /// Shared strings
    strings: Vec<String>,
    /// Sheets paths
    sheets: Vec<(String, String)>,
}

impl<RS> Xlsx<RS> where RS: Read + Seek {
    fn read_shared_strings(&mut self) -> Result<()> {
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
                Ok(Event::Eof) => bail!("unexpected end of xml (no </sst>)"),
                _ => (),
            }
        }
        Ok(())
    }

    fn read_workbook(
        &mut self,
        relationships: &HashMap<Vec<u8>, String>,
    ) -> Result<Vec<(String, String)>> {
        let mut xml = match xml_reader(&mut self.zip, "xl/workbook.xml") {
            None => return Ok(Vec::new()),
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
                Ok(Event::Start(ref e)) if e.local_name() == b"definedName" => if let Some(a) =
                    e.attributes()
                        .filter_map(|a| a.ok())
                        .find(|a| a.key == b"name")
                {
                    let name = a.unescape_and_decode_value(&xml)?;
                    val_buf.clear();
                    let value = xml.read_text(b"definedName", &mut val_buf)?;
                    defined_names.push((name, value));
                },
                Ok(Event::End(ref e)) if e.local_name() == b"workbook" => break,
                Ok(Event::Eof) => bail!("unexpected end of xml (no </workbook>)"),
                Err(e) => bail!(e),
                _ => (),
            }
        }
        Ok(defined_names)
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let mut xml = match xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            None => bail!("Cannot find relationships file"),
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
                Ok(Event::Eof) => bail!("unexpected end of xml (no </Relationships>)"),
                Err(e) => bail!(e),
                _ => (),
            }
        }
        Ok(relationships)
    }

    fn read_worksheet<T, F>(&mut self, name: &str, read_data: &mut F) -> Result<Range<T>>
    where
        T: Default + Clone + PartialEq,
        F: FnMut(&[String], &mut XlsReader, &mut Vec<Cell<T>>) -> Result<()>,
    {
        let &(_, ref path) = self.sheets
            .iter()
            .find(|&&(ref n, _)| n == name)
            .ok_or_else(|| ErrorKind::WorksheetName(name.to_string()))?;
        let mut xml = match xml_reader(&mut self.zip, path) {
            None => bail!("Cannot find {} path", path),
            Some(x) => x?,
        };
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
                            bail!("Expecting dimension, got {:?}", e);
                        }
                        b"sheetData" => {
                            read_data(&self.strings, &mut xml, &mut cells)?;
                            break;
                        }
                        _ => (),
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => bail!(e),
                _ => (),
            }
        }
        Ok(Range::from_sparse(cells))
    }
}

impl<RS> Reader<RS> for Xlsx<RS> where RS: Read + Seek {
    fn new(reader: RS) -> Result<Self> where RS: Read + Seek {
        Ok(Xlsx {
            zip: ZipArchive::new(reader)?,
            strings: Vec::new(),
            sheets: Vec::new(),
        })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        let mut f = self.zip.by_name("xl/vbaProject.bin")?;
        let len = f.size() as usize;
        VbaProject::new(&mut f, len).map(Cow::Owned)
    }

    fn initialize(&mut self) -> Result<Metadata> {
        self.read_shared_strings()?;
        let relationships = self.read_relationships()?;
        let defined_names = self.read_workbook(&relationships)?;
        Ok(Metadata {
            sheets: self.sheets.iter().map(|&(ref s, _)| s.clone()).collect(),
            defined_names: defined_names,
        })
    }

    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>> {
        self.read_worksheet(name, &mut |s, xml, cells| read_sheet_data(xml, s, cells))
    }

    fn read_worksheet_formula(&mut self, name: &str) -> Result<Range<String>> {
        self.read_worksheet(name, &mut |_, xml, cells| {
            read_sheet(xml, cells, &mut |cells, xml, e, pos, _| {
                match e.local_name() {
                    b"is" | b"v" => xml.read_to_end(e.name(), &mut Vec::new())?,
                    b"f" => {
                        let f = xml.read_text(e.name(), &mut Vec::new())?;
                        if !f.is_empty() {
                            cells.push(Cell::new(pos, f));
                        }
                    }
                    n => bail!("not a 'v', 'f', or 'is' node: {:?}", n),
                }
                Ok(())
            })
        })
    }
}

fn xml_reader<'a, RS>(zip: &'a mut ZipArchive<RS>, path: &str) -> Option<Result<XlsReader<'a>>> where RS: Read + Seek {
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
fn get_attribute<'a>(atts: Attributes<'a>, n: &[u8]) -> Result<Option<&'a [u8]>> {
    use std::borrow::Cow;
    for a in atts {
        match a {
            Ok(Attribute { key, value: Cow::Borrowed(value) }) if key == n => return Ok(Some(value)),
            Err(qe) => bail!(qe),
            _ => {} // ignore other attributes
        }
    }
    Ok(None)
}

fn read_sheet<T, F>(xml: &mut XlsReader, cells: &mut Vec<Cell<T>>, push_cell: &mut F) -> Result<()>
where
    T: Clone + Default + PartialEq,
    F: FnMut(&mut Vec<Cell<T>>, &mut XlsReader, &BytesStart, (u32, u32), &BytesStart)
        -> Result<()>,
{
    let mut buf = Vec::new();
    let mut cell_buf = Vec::new();
    loop {
        buf.clear();
        match xml.read_event(&mut buf) {
            Ok(Event::Start(ref c_element)) if c_element.local_name() == b"c" => {
                let pos = get_attribute(c_element.attributes(), b"r")
                    .and_then(|o| o.ok_or_else(|| "Cell missing 'r' attribute tag".into()))
                    .and_then(get_row_column)?;
                loop {
                    cell_buf.clear();
                    match xml.read_event(&mut cell_buf) {
                        Ok(Event::Start(ref e)) => push_cell(cells, xml, e, pos, c_element)?,
                        Ok(Event::End(ref e)) if e.local_name() == b"c" => break,
                        Ok(Event::Eof) => bail!("unexpected end of xml (no </c>)"),
                        Err(e) => bail!(e),
                        _ => (),
                    }
                }
            }
            Ok(Event::End(ref e)) if e.local_name() == b"sheetData" => return Ok(()),
            Ok(Event::Eof) => bail!("unexpected end of xml (no </sheetData>)"),
            Err(e) => bail!(e),
            _ => (),
        }
    }
}

/// read sheetData node
fn read_sheet_data(
    xml: &mut XlsReader,
    strings: &[String],
    cells: &mut Vec<Cell<DataType>>,
) -> Result<()> {
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
                bail!("called read_value on a cell of type inlineStr");
            }
            Some(t) => bail!("unknown cell 't' attribute={:?}", t),
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
            n => bail!("not a 'v', 'f', or 'is' node: {:?}", n),
        }
        Ok(())
    })
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
        len => Err(
            format!("range dimension has 0 or 1 ':', got {}", len).into(),
        ),
    }
}

/// converts a text range name into its position (row, column) (0 based index)
fn get_row_column(range: &[u8]) -> Result<(u32, u32)> {
    let (mut row, mut col) = (0, 0);
    let mut pow = 1;
    let mut readrow = true;
    for c in range.iter().rev() {
        match *c {
            c @ b'0'...b'9' => if readrow {
                row += ((c - b'0') as u32) * pow;
                pow *= 10;
            } else {
                bail!(
                    "Numeric character are only allowed at the end of the range: {:x}",
                    c
                );
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
            _ => bail!("Expecting alphanumeric character, got {:x}", c),
        }
    }
    Ok((row - 1, col - 1))
}

/// attempts to read either a simple or richtext string
fn read_string(xml: &mut XlsReader, closing: &[u8]) -> Result<Option<String>> {
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
            Ok(Event::Eof) => bail!("unexpected end of xml"),
            Err(e) => bail!(e),
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
