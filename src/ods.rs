//! A module to parse Open Document Spreadsheets
//!
//! # Reference
//! OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
//! http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf

use std::io::{BufReader, Read, Seek};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;
use quick_xml::reader::Reader as XmlReader;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attributes;

use {DataType, Metadata, Range, Reader};
use vba::VbaProject;
use errors::CalError;

const MIMETYPE: &'static [u8] = b"application/vnd.oasis.opendocument.spreadsheet";

type OdsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

/// An enum for ods specific errors
#[derive(Debug, Fail)]
pub enum OdsError {
    #[fail(display = "{}", _0)] Io(#[cause] ::std::io::Error),
    #[fail(display = "{}", _0)] Zip(#[cause] ::zip::result::ZipError),
    #[fail(display = "{}", _0)] Xml(#[cause] ::quick_xml::errors::Error),
    #[fail(display = "{}", _0)] Parse(#[cause] ::std::string::ParseError),
    #[fail(display = "{}", _0)] ParseFloat(#[cause] ::std::num::ParseFloatError),

    #[fail(display = "Invalid MIME type: {:?}", _0)] InvalidMime(Vec<u8>),
    #[fail(display = "Cannot find '{}' file", _0)] FileNotFound(&'static str),
    #[fail(display = "Expecting {} found EOF", _0)] Eof(&'static str),
    #[fail(display = "Expecting {} found {}", expected, found)]
    Mismatch {
        expected: &'static str,
        found: String,
    },
}

impl_error!(::std::io::Error, OdsError, Io);
impl_error!(::zip::result::ZipError, OdsError, Zip);
impl_error!(::quick_xml::errors::Error, OdsError, Xml);
impl_error!(::std::string::ParseError, OdsError, Parse);
impl_error!(::std::num::ParseFloatError, OdsError, ParseFloat);

enum Content<RS>
where
    RS: Read + Seek,
{
    Zip(ZipArchive<RS>),
    Sheets(HashMap<String, (Range<DataType>, Range<String>)>),
}

/// An OpenDocument Spreadsheet document parser
///
/// # Reference
/// OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
/// http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf
pub struct Ods<RS>
where
    RS: Read + Seek,
{
    /// A zip package or an already parsed xml content
    content: Content<RS>,
}

impl<RS> Reader<RS> for Ods<RS>
where
    RS: Read + Seek,
{
    fn new(reader: RS) -> Result<Self, CalError>
    where
        RS: Read + Seek,
    {
        let mut zip = ZipArchive::new(reader).map_err(OdsError::Zip)?;

        // check mimetype
        match zip.by_name("mimetype") {
            Ok(mut f) => {
                let mut buf = [0u8; 46];
                f.read_exact(&mut buf)?;
                if &buf[..] != MIMETYPE {
                    return Err(CalError::Ods(OdsError::InvalidMime(buf.to_vec())));
                }
            }
            Err(ZipError::FileNotFound) => {
                return Err(CalError::Ods(OdsError::FileNotFound("mimetype")))
            }
            Err(e) => return Err(CalError::Ods(OdsError::Zip(e))),
        }

        Ok(Ods {
            content: Content::Zip(zip),
        })
    }

    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool {
        // TODO: implement code parsing
        false
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Result<Cow<VbaProject>, CalError> {
        unimplemented!();
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn initialize(&mut self) -> Result<Metadata, CalError> {
        let defined_names = self.parse_content().map_err(CalError::Ods)?;
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
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>, CalError> {
        self.parse_content().map_err(CalError::Ods)?;
        if let Content::Sheets(ref s) = self.content {
            if let Some(r) = s.get(name) {
                return Ok(r.0.to_owned());
            }
        }
        Err(CalError::WorksheetName(name.into()))
    }

    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_formula(&mut self, name: &str) -> Result<Range<String>, CalError> {
        self.parse_content().map_err(CalError::Ods)?;
        if let Content::Sheets(ref s) = self.content {
            if let Some(r) = s.get(name) {
                return Ok(r.1.to_owned());
            }
        }
        Err(CalError::WorksheetName(name.into()))
    }
}

impl<RS> Ods<RS>
where
    RS: Read + Seek,
{
    /// Parses content.xml and store the result in `self.content`
    fn parse_content(&mut self) -> Result<Vec<(String, String)>, OdsError> {
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
                Err(ZipError::FileNotFound) => return Err(OdsError::FileNotFound("content.xml")),
                Err(e) => return Err(OdsError::Zip(e)),
            };
            let mut buf = Vec::new();
            let mut sheets = HashMap::new();
            let mut defined_names = Vec::new();
            loop {
                match reader.read_event(&mut buf) {
                    Ok(Event::Start(ref e)) if e.name() == b"table:table" => if let Some(ref a) =
                        e.attributes()
                            .filter_map(|a| a.ok())
                            .find(|a| a.key == b"table:name")
                    {
                        let name = a.unescape_and_decode_value(&reader).map_err(OdsError::Xml)?;
                        let (range, formulas) = read_table(&mut reader)?;
                        sheets.insert(name, (range, formulas));
                    },
                    Ok(Event::Start(ref e)) if e.name() == b"table:named-expressions" => {
                        defined_names = read_named_expressions(&mut reader)?;
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(OdsError::Xml(e)),
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

fn read_table(reader: &mut OdsReader) -> Result<(Range<DataType>, Range<String>), OdsError> {
    let mut cells = Vec::new();
    let mut formulas = Vec::new();
    let mut cols = Vec::new();
    let mut buf = Vec::new();
    let mut row_buf = Vec::new();
    let mut cell_buf = Vec::new();
    cols.push(0);
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-row" => {
                read_row(
                    reader,
                    &mut row_buf,
                    &mut cell_buf,
                    &mut cells,
                    &mut formulas,
                )?;
                cols.push(cells.len());
            }
            Ok(Event::End(ref e)) if e.name() == b"table:table" => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(_) => (),
        }
        buf.clear();
    }
    Ok((get_range(cells, &cols), get_range(formulas, &cols)))
}

fn get_range<T: Default + Clone + PartialEq>(mut cells: Vec<T>, cols: &[usize]) -> Range<T> {
    // find smallest area with non empty Cells
    let mut row_min = None;
    let mut row_max = 0;
    let mut col_min = ::std::usize::MAX;
    let mut col_max = 0;
    {
        for (i, w) in cols.windows(2).enumerate() {
            let row = &cells[w[0]..w[1]];
            if let Some(p) = row.iter().position(|c| c != &T::default()) {
                if row_min.is_none() {
                    row_min = Some(i);
                }
                row_max = i;
                if p < col_min {
                    col_min = p;
                }
                if let Some(p) = row.iter().rposition(|c| c != &T::default()) {
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
        let empty_cells = vec![T::default(); col_max + 1];
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

fn read_row(
    reader: &mut OdsReader,
    row_buf: &mut Vec<u8>,
    cell_buf: &mut Vec<u8>,
    cells: &mut Vec<DataType>,
    formulas: &mut Vec<String>,
) -> Result<(), OdsError> {
    loop {
        row_buf.clear();
        match reader.read_event(row_buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table-cell" => {
                let (value, formula, is_closed) = get_datatype(reader, e.attributes(), cell_buf)?;
                cells.push(value);
                formulas.push(formula);
                if !is_closed {
                    reader.read_to_end(b"table:table-cell", cell_buf)?;
                }
            }
            Ok(Event::End(ref e)) if e.name() == b"table:table-row" => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(e) => {
                return Err(OdsError::Mismatch {
                    expected: "table-cell",
                    found: format!("{:?}", e),
                })
            }
        }
    }
    Ok(())
}

/// Converts table-cell element into a `DataType`
///
/// ODF 1.2-19.385
fn get_datatype(
    reader: &mut OdsReader,
    atts: Attributes,
    buf: &mut Vec<u8>,
) -> Result<(DataType, String, bool), OdsError> {
    let mut is_string = false;
    let mut is_value_set = false;
    let mut val = DataType::Empty;
    let mut formula = String::new();
    for a in atts {
        let a = a?;
        match a.key {
            b"office:value" if !is_value_set => {
                let v = reader.decode(&a.value);
                val = DataType::Float(v.parse().map_err(OdsError::ParseFloat)?);
                is_value_set = true;
            }
            b"office:string-value" | b"office:date-value" | b"office:time-value"
                if !is_value_set =>
            {
                val = DataType::String(a.unescape_and_decode_value(reader).map_err(OdsError::Xml)?);
                is_value_set = true;
            }
            b"office:boolean-value" if !is_value_set => {
                val = DataType::Bool(&*a.value == b"TRUE");
                is_value_set = true;
            }
            b"office:value-type" if !is_value_set => is_string = &*a.value == b"string",
            b"table:formula" => {
                formula = a.unescape_and_decode_value(reader).map_err(OdsError::Xml)?;
            }
            _ => (),
        }
    }
    if !is_value_set && is_string {
        // If the value type is string and the office:string-value attribute
        // is not present, the element content defines the value.
        loop {
            buf.clear();
            match reader.read_event(buf) {
                Ok(Event::Text(ref e)) => {
                    return Ok((
                        DataType::String(e.unescape_and_decode(reader).map_err(OdsError::Xml)?),
                        formula,
                        false,
                    ));
                }
                Ok(Event::End(ref e)) if e.name() == b"table:table-cell" => {
                    return Ok((DataType::String("".to_string()), formula, true));
                }
                Err(e) => return Err(OdsError::Xml(e)),
                Ok(Event::Eof) => return Err(OdsError::Eof("table:table-cell")),
                _ => (),
            }
        }
    } else {
        Ok((val, formula, false))
    }
}

fn read_named_expressions(reader: &mut OdsReader) -> Result<Vec<(String, String)>, OdsError> {
    let mut defined_names = Vec::new();
    let mut buf = Vec::new();
    loop {
        buf.clear();
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e))
                if e.name() == b"table:named-range" || e.name() == b"table:named-expression" =>
            {
                let mut name = String::new();
                let mut formula = String::new();
                for a in e.attributes() {
                    let a = a.map_err(OdsError::Xml)?;
                    match a.key {
                        b"table:name" => {
                            name = a.unescape_and_decode_value(reader).map_err(OdsError::Xml)?
                        }
                        b"table:cell-range-address" | b"table:expression" => {
                            formula = a.unescape_and_decode_value(reader).map_err(OdsError::Xml)?
                        }
                        _ => (),
                    }
                }
                defined_names.push((name, formula));
            }
            Ok(Event::End(ref e))
                if e.name() == b"table:named-range" || e.name() == b"table:named-expression" => {}
            Ok(Event::End(ref e)) if e.name() == b"table:named-expressions" => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(e) => {
                return Err(OdsError::Mismatch {
                    expected: "table:named-expressions",
                    found: format!("{:?}", e),
                })
            }
        }
    }
    Ok(defined_names)
}
