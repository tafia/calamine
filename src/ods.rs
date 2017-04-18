//! A module to parse Open Document Spreadsheets
//!
//! # Reference
//! OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
//! http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf

use std::fs::File;
use std::io::{BufReader, Read};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::reader::Reader as XmlReader;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attributes;

use {Metadata, DataType, Reader, Range};
use vba::VbaProject;
use errors::*;

const MIMETYPE: &'static [u8] = b"application/vnd.oasis.opendocument.spreadsheet";

type OdsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

enum Content {
    Zip(ZipArchive<File>),
    Sheets(HashMap<String, Range<DataType>>),
}

/// An OpenDocument Spreadsheet document parser
///
/// # Reference
/// OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
/// http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf
pub struct Ods {
    /// A zip package or an already parsed xml content
    content: Content,
}

impl Reader for Ods {
    /// Creates a new instance based on the actual file
    fn new(f: File) -> Result<Self> {
        let mut zip = ZipArchive::new(f)?;

        // check mimetype
        match zip.by_name("mimetype") {
            Ok(mut f) => {
                let mut buf = [0u8; 46];
                f.read_exact(&mut buf)?;
                if &buf[..] != MIMETYPE {
                    bail!("Invalid mimetype, expecting {:?}, found {:?}",
                          MIMETYPE,
                          &buf[..]);
                }
            }
            Err(ZipError::FileNotFound) => bail!("Cannot find 'mimetype' file"),
            Err(e) => bail!(e),
        }

        Ok(Ods { content: Content::Zip(zip) })
    }

    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool {
        // TODO: implement code parsing
        false
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        unimplemented!();
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn initialize(&mut self) -> Result<Metadata> {
        let defined_names = self.parse_content()?;
        let sheets = if let Content::Sheets(ref s) = self.content {
            s.keys().map(|k| k.to_string()).collect()
        } else {
            Vec::new()
        };
        Ok(Metadata {
               sheets: sheets,
               defined_names: defined_names,
           })
    }

    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>> {
        self.parse_content()?;
        if let Content::Sheets(ref s) = self.content {
            if let Some(r) = s.get(name) {
                return Ok(r.to_owned());
            }
        }
        bail!("Cannot find '{}' sheet", name);
    }
}

impl Ods {
    /// Parses content.xml and store the result in `self.content`
    fn parse_content(&mut self) -> Result<Vec<(String, String)>> {
        let (sheets, defined_names) = if let Content::Zip(ref mut zip) = self.content {
            let mut reader = match zip.by_name("content.xml") {
                Ok(f) => {
                    let mut r = XmlReader::from_reader(BufReader::new(f));
                    r.check_end_names(false)
                        .trim_text(true)
                        .check_comments(false)
                        .expand_empty_elements(true);
                    r
                }
                Err(ZipError::FileNotFound) => bail!("Cannot find 'content.xml' file"),
                Err(e) => bail!(e),
            };
            let mut buf = Vec::new();
            let mut sheets = HashMap::new();
            let mut defined_names = Vec::new();
            loop {
                match reader.read_event(&mut buf) {
                    Ok(Event::Start(ref e)) if e.name() == b"table:table" => {
                        if let Some(ref a) = e.attributes()
                               .filter_map(|a| a.ok())
                               .find(|ref a| a.key == b"table:name") {
                            let name = a.unescape_and_decode_value(&mut reader)?;
                            let range = read_table(&mut reader)?;
                            sheets.insert(name, range);
                        }
                    }
                    Ok(Event::Start(ref e)) if e.name() == b"table:named-expressions" => {
                        defined_names = read_named_expressions(&mut reader)?;
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => bail!(e),
                    _ => (),
                }
                buf.clear();
            }
            (sheets, defined_names)
        } else {
            return Ok(Vec::new());
        };
        self.content = Content::Sheets(sheets);
        Ok(defined_names)
    }
}

fn read_table(reader: &mut OdsReader) -> Result<Range<DataType>> {
    let mut cells = Vec::new();
    let mut cols = Vec::new();
    let mut buf = Vec::new();
    let mut row_buf = Vec::new();
    let mut cell_buf = Vec::new();
    cols.push(0);
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-row" => {
                read_row(reader, &mut row_buf, &mut cell_buf, &mut cells)?;
                cols.push(cells.len());
            }
            Ok(Event::End(ref e)) if e.name() == b"table:table" => break,
            Err(e) => bail!(e),
            Ok(_) => (),
        }
        buf.clear();
    }
    Ok(get_range(cells, &cols))
}

fn get_range(mut cells: Vec<DataType>, cols: &[usize]) -> Range<DataType> {

    // find smallest area with non empty Cells
    let mut row_min = None;
    let mut row_max = 0;
    let mut col_min = ::std::usize::MAX;
    let mut col_max = 0;
    {
        let not_empty = |c| if let &DataType::Empty = c {
            false
        } else {
            true
        };
        for (i, w) in cols.windows(2).enumerate() {
            let row = &cells[w[0]..w[1]];
            if let Some(p) = row.iter().position(|c| not_empty(c)) {
                if row_min.is_none() {
                    row_min = Some(i);
                }
                row_max = i;
                if p < col_min {
                    col_min = p;
                }
                if let Some(p) = row.iter().rposition(|c| not_empty(c)) {
                    if p > col_max {
                        col_max = p;
                    }
                }
            }
        }
    }
    let row_min = match row_min {
        Some(min) => min,
        _ => return Range::default(),
    };

    // rebuild cells into its smallest non empty area
    let cells_len = (row_max + 1 - row_min) * (col_max + 1 - col_min);
    if cells.len() != cells_len {
        let mut new_cells = Vec::with_capacity(cells_len);
        let empty_cells = vec![DataType::Empty; col_max + 1];
        for w in cols.windows(2).skip(row_min).take(row_max + 1) {
            let row = &cells[w[0]..w[1]];
            if row.len() < col_max + 1 {
                new_cells.extend_from_slice(&row[col_min..]);
                new_cells.extend_from_slice(&empty_cells[row.len()..]);
            } else if row.len() == col_max + 1 {
                new_cells.extend_from_slice(&row[col_min..]);
            } else {
                new_cells.extend_from_slice(&row[col_min..col_max + 1]);
            }
        }
        cells = new_cells;
    }
    Range {
        start: (row_min as u32, col_min as u32),
        end: (row_max as u32, col_max as u32),
        inner: cells,
    }
}

fn read_row(reader: &mut OdsReader,
            row_buf: &mut Vec<u8>,
            cell_buf: &mut Vec<u8>,
            cells: &mut Vec<DataType>)
            -> Result<()> {
    loop {
        row_buf.clear();
        match reader.read_event(row_buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-cell" => {
                let (value, is_closed) = get_datatype(reader, e.attributes(), cell_buf)?;
                cells.push(value);
                if !is_closed {
                    reader.read_to_end(b"table:table-cell", cell_buf)?;
                }
            }
            Ok(Event::End(ref e)) if e.name() == b"table:table-row" => break,
            Err(e) => bail!(e),
            Ok(e) => bail!("Expecting 'table-cell' event, found {:?}", e),
        }
    }
    Ok(())
}

/// Converts table-cell element into a DataType
///
/// ODF 1.2-19.385
fn get_datatype(reader: &mut OdsReader,
                atts: Attributes,
                buf: &mut Vec<u8>)
                -> Result<(DataType, bool)> {
    let mut is_string = false;
    for a in atts {
        let a = a?;
        match a.key {
            b"office:value" => {
                let v = reader.decode(a.value);
                return v.parse()
                           .map(|f| (DataType::Float(f), false))
                           .map_err(|e| e.into());
            }
            b"office:string-value" |
            b"office:date-value" |
            b"office:time-value" => {
                return Ok((DataType::String(a.unescape_and_decode_value(reader)?), false));
            }
            b"office:boolean-value" => return Ok((DataType::Bool(a.value == b"TRUE"), false)),
            b"office:value-type" => is_string = a.value == b"string",
            _ => (),
        }
    }
    if is_string {
        // If the value type is string and the office:string-value attribute
        // is not present, the element content defines the value.
        loop {
            buf.clear();
            match reader.read_event(buf) {
                Ok(Event::Text(ref e)) => {
                    return Ok((DataType::String(e.unescape_and_decode(reader)?), false));
                }
                Ok(Event::End(ref e)) if e.name() == b"table:table-cell" => {
                    return Ok((DataType::String("".to_string()), true));
                }
                Err(e) => bail!(e),
                Ok(Event::Eof) => bail!("Expecting 'table:table-cell' end element, found EOF"),
                _ => (),
            }
        }
    } else {
        Ok((DataType::Empty, false))
    }
}

fn read_named_expressions(reader: &mut OdsReader) -> Result<Vec<(String, String)>> {
    let mut defined_names = Vec::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:named-range" ||
                                       e.name() == b"table:named-expression" => {
                let mut name = String::new();
                let mut formula = String::new();
                for a in e.attributes() {
                    let a = a?;
                    match a.key {
                        b"table:name" => name = a.unescape_and_decode_value(reader)?,
                        b"table:cell-range-address" |
                        b"table:expression" => formula = a.unescape_and_decode_value(reader)?,
                        _ => (),
                    }
                }
                defined_names.push((name, formula));
            }
            Ok(Event::End(ref e)) if e.name() == b"table:named-range" ||
                                     e.name() == b"table:named-expression" => (),
            Ok(Event::End(ref e)) if e.name() == b"table:named-expressions" => break,
            Err(e) => bail!(e),
            Ok(e) => bail!("Expecting 'table:named-expressions' event, found {:?}", e),
        }
    }
    Ok(defined_names)
}
