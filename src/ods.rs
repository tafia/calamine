//! A module to parse Open Document Spreasheets
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
use quick_xml::reader::Reader;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attributes;

use {DataType, ExcelReader, Range};
use vba::VbaProject;
use errors::*;

const MIMETYPE: &'static [u8] = b"application/vnd.oasis.opendocument.spreadsheet";

type OdsReader<'a> = Reader<BufReader<ZipFile<'a>>>;

enum Content {
    Zip(ZipArchive<File>),
    Sheets(HashMap<String, Range>),
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

impl ExcelReader for Ods {
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
                          MIMETYPE, &buf[..]);
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
    fn vba_project(&mut self) -> Result<Cow<VbaProject>>{
        unimplemented!();
    }

    /// Gets vba references
    fn read_shared_strings(&mut self) -> Result<Vec<String>>{
        Ok(Vec::new())
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn read_sheets_names(&mut self,
                         _: &HashMap<Vec<u8>, String>)
                         -> Result<Vec<(String, String)>> {
        self.parse_content()?;
        if let Content::Sheets(ref s) = self.content {
            Ok(s.keys().map(|k| (k.to_string(), k.to_string())).collect())
        } else {
            Ok(Vec::new())
        }
    }

    /// Read workbook relationships
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>>{
        Ok(HashMap::new())
    }

    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, path: &str, _: &[String]) -> Result<Range> {
        self.parse_content()?;
        if let Content::Sheets(ref s) = self.content {
            if let Some(r) = s.get(path) {
                return Ok(r.to_owned());
            }
        }
        bail!("Cannot find '{}' sheet", path);
    }
}

impl Ods {
    /// Parses content.xml and store the result in `self.content`
    fn parse_content(&mut self) -> Result<()> {
        let sheets = if let Content::Zip(ref mut zip) = self.content {
            let mut reader = match zip.by_name("content.xml") {
                Ok(f) => {
                    let mut r = Reader::from_reader(BufReader::new(f));
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
            loop {
                match reader.read_event(&mut buf) {
                    Ok(Event::Start(ref e)) if e.name() == b"table:table" => {
                        if let Some(ref a) = e.attributes().filter_map(|a| a.ok())
                            .find(|ref a| a.key == b"table:name") {
                            let name = a.unescape_and_decode_value(&mut reader)?;
                            let range = read_table(&mut reader)?;
                            sheets.insert(name, range);
                        }
                    },
                    Ok(Event::Eof) => break,
                    Err(e) => bail!(e),
                    _ => (),
                }
                buf.clear();
            }
            sheets
        } else {
            return Ok(());
        };
        self.content = Content::Sheets(sheets);
        Ok(())
    }
}

fn read_table(reader: &mut OdsReader) -> Result<Range> {
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

    // prune cells so it doesn't necessarily starts at 'A1'
    let mut row_min = None;
    let mut row_max = 0;
    let mut col_min = ::std::usize::MAX;
    let mut col_max = 0;
    {
        let not_empty = |c| if let &DataType::Empty = c { false } else { true };
        for (i, w) in cols.windows(2).enumerate() {
            let row = &cells[w[0]..w[1]];
            if let Some(p) = row.iter().position(|c| not_empty(c)) {
                if p < col_min { col_min = p; }
                if row_min.is_none() { row_min = Some(i); }
                row_max = i;
            }
            if let Some(p) = row.iter().rposition(|c| not_empty(c)) {
                if p > col_max { col_max = p; }
            }
        }
    }
    let row_min = match row_min {
        Some(min) => min,
        _ => return Ok(Range::default()),
    };

    // rebuild cells to it is rectangular
    let cells_len = (row_max + 1 - row_min) * (col_max + 1 - col_min);
    if cells.len() != cells_len {
        let mut new_cells = Vec::with_capacity(cells_len);
        for w in cols.windows(2).skip(row_min).take(row_max + 1) {
            let row = &cells[w[0]..w[1]];
            if row.len() < col_max + 1 {
                new_cells.extend_from_slice(&row[col_min..]);
                new_cells.extend(::std::iter::repeat(DataType::Empty).take(col_max + 1 - row.len()));
            } else if row.len() == col_max + 1 {
                new_cells.extend_from_slice(&row[col_min..]);
            } else {
                new_cells.extend_from_slice(&row[col_min..col_max + 1]);
            }
        }
        cells = new_cells;
    }
    Ok(Range {
        start: (row_min as u32, col_min as u32),
        end: (row_max as u32, col_max as u32),
        inner: cells,
    })
}

fn read_row(reader: &mut OdsReader,
            row_buf: &mut Vec<u8>,
            cell_buf: &mut Vec<u8>,
            cells: &mut Vec<DataType>) -> Result<()> {
    loop {
        row_buf.clear();
        match reader.read_event(row_buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-cell" => {
                let (cell_value, cell_closed) = attributes_to_datatype(reader, 
                                                                       e.attributes(),
                                                                       cell_buf)?;
                cells.push(cell_value);
                if !cell_closed {
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
fn attributes_to_datatype(reader: &mut OdsReader,
                          atts: Attributes,
                          buf: &mut Vec<u8>) -> Result<(DataType, bool)> {
    let mut is_string = false;
    for a in atts {
        let a = a?;
        match a.key {
            b"office:value" => {
                let v = reader.decode(a.value);
                return v.parse()
                    .map(|f| (DataType::Float(f), false))
                    .map_err(|e| e.into());
            },
            b"office:string-value" | b"office:date-value" | b"office:time-value" => {
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
                },
                Ok(Event::End(ref e)) if e.name() == b"table:table-cell" => {
                    return Ok((DataType::String("".to_string()), true));
                },
                Err(e) => bail!(e),
                Ok(Event::Eof) => bail!("Expecting 'table:table-cell' end element, found EOF"),
                _ => (),
            }
        }
    } else {
        Ok((DataType::Empty, false))
    }
}
