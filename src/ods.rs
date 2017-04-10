//! A module to parse Open Document Spreasheets

use std::fs::File;
use std::io::{BufReader, Read};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::reader::Reader;
use quick_xml::events::{Event, BytesText};
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
            let mut reader = get_zip_reader(zip, "content.xml")?;
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
                    Ok(_) => (),
                    Err(e) => bail!(e),
                }
                buf.clear();
            }
            Some(sheets)
        } else {
            None
        };

        if let Some(sheets) = sheets {
            self.content = Content::Sheets(sheets);
        }

        Ok(())
    }
}

fn get_zip_reader<'a, 'b>(zip: &'a mut ZipArchive<File>, path: &'b str) -> Result<OdsReader<'a>> {
    match zip.by_name(path) {
        Ok(f) => {
            let mut r = Reader::from_reader(BufReader::new(f));
            r.check_end_names(false)
                .trim_text(true)
                .check_comments(false)
                .expand_empty_elements(true);
            Ok(r)
        }
        Err(ZipError::FileNotFound) => bail!("Cannot find '{}' file", path),
        Err(e) => bail!(e),
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
                if p < col_min {
                    col_min = p;
                }
                if row_min.is_none() {
                    row_min = Some(i);
                }
                row_max = i;
            }
            if let Some(p) = row.iter().rposition(|c| not_empty(c)) {
                if p > col_max {
                    col_max = p;
                }
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
        Ok(Range {
            start: (row_min as u32, col_min as u32),
            end: (row_max as u32, col_max as u32),
            inner: new_cells,
        })
    } else {
        Ok(Range {
            start: (row_min as u32, col_min as u32),
            end: (row_max as u32, col_max as u32),
            inner: cells,
        })
    }
}

fn read_row(reader: &mut OdsReader,
            row_buf: &mut Vec<u8>,
            cell_buf: &mut Vec<u8>,
            cells: &mut Vec<DataType>) -> Result<()> {
    let mut close_cell = false;
    loop {
        row_buf.clear();
        match reader.read_event(row_buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-cell" => {
                close_cell = true;
                if let Some(a) = e.attributes().filter_map(|a| a.ok())
                    .find(|ref a| a.key == b"office:value-type") {
                    cell_buf.clear();
                    match reader.read_event(cell_buf) {
                        Ok(Event::Start(ref c)) if c.name() == b"text:p" => (),
                        Ok(Event::End(ref end)) if end.name() == b"table:table-cell" => {
                            close_cell = false;
                            cells.push(attributes_to_datatype(reader, e.attributes())?);
                        },
                        Err(e) => bail!(e),
                        Ok(e) => bail!("Expecting 'text:p' event, found {:?}", e),
                    }
                    if close_cell {
                    cell_buf.clear();
                        match reader.read_event(cell_buf) {
                            Ok(Event::Text(ref c)) => {
                                cells.push(cell_value_to_datatype(reader, a.value, c)?);
                            },
                            Err(e) => bail!(e),
                            Ok(e) => bail!("Expecting Text event, found {:?}", e),
                        }
                        match reader.read_event(cell_buf) {
                            Ok(Event::End(ref c)) if c.name() == b"text:p" => (),
                            Err(e) => bail!(e),
                            Ok(e) => bail!("Expecting 'text:p' event, found {:?}", e),
                        }
                    }
                } else {
                    cells.push(DataType::Empty);
                }
            }
            Ok(Event::End(ref e)) if close_cell && e.name() == b"table:table-cell" => close_cell = false,
            Ok(Event::End(ref e)) if e.name() == b"table:table-row" => break,
            Err(e) => bail!(e),
            Ok(e) => bail!("Expecting 'table-cell' event, found {:?}", e),
        }
    }
    Ok(())
}

fn cell_value_to_datatype(reader: &OdsReader,
                          cell_type: &[u8],
                          cell_value: &BytesText) -> Result<DataType> {
    match cell_type {
        b"boolean" => Ok(DataType::Bool(&**cell_value == b"TRUE")),
        b"string" | b"date" | b"time" => {
            Ok(DataType::String(cell_value.unescape_and_decode(reader)?))
        }
        b"float" | b"percentage" => {
            let v = reader.decode(cell_value);
            v.parse().map(DataType::Float).map_err(|e| e.into())
        },
        b"void" => Ok(DataType::Empty),
        b"currency" => {
            let v = reader.decode(cell_value);
            v.parse()
                .map(DataType::Float)
                .or_else(|_| Ok(DataType::String(v.to_string())))
        }
        t => bail!("Unrecognized cell type: {:?}", t),
    }
}

fn attributes_to_datatype(reader: &OdsReader, atts: Attributes) -> Result<DataType> {
    for a in atts {
        let a = a?;
        match a.key {
            b"office:boolean-value" => return Ok(DataType::Bool(a.value == b"TRUE")),
            b"office:value" => {
                let v = reader.decode(a.value);
                return v.parse().map(DataType::Float).map_err(|e| e.into());
            },
            b"office:string-value" | b"office:date-value" | b"office:time-value" => {
                return Ok(DataType::String(a.unescape_and_decode_value(reader)?));
            }
            _ => (),
        }
    }
    Ok(DataType::Empty)
}
