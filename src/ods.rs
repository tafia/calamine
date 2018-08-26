//! A module to parse Open Document Spreadsheets
//!
//! # Reference
//! OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
//! http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf

use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};

use quick_xml::events::attributes::Attributes;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use std::marker::PhantomData;
use vba::VbaProject;
use {DataType, Metadata, Range, Reader};

const MIMETYPE: &'static [u8] = b"application/vnd.oasis.opendocument.spreadsheet";

type OdsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

/// An enum for ods specific errors
#[derive(Debug, Fail)]
pub enum OdsError {
    /// Io error
    #[fail(display = "{}", _0)]
    Io(#[cause] ::std::io::Error),
    /// Zip error
    #[fail(display = "{}", _0)]
    Zip(#[cause] ::zip::result::ZipError),
    /// Xml error
    #[fail(display = "{}", _0)]
    Xml(#[cause] ::quick_xml::Error),
    /// Error while parsing string
    #[fail(display = "{}", _0)]
    Parse(#[cause] ::std::string::ParseError),
    /// Error while parsing integer
    #[fail(display = "{}", _0)]
    ParseInt(#[cause] ::std::num::ParseIntError),
    /// Error while parsing float
    #[fail(display = "{}", _0)]
    ParseFloat(#[cause] ::std::num::ParseFloatError),

    /// Invalid MIME
    #[fail(display = "Invalid MIME type: {:?}", _0)]
    InvalidMime(Vec<u8>),
    /// File not found
    #[fail(display = "Cannot find '{}' file", _0)]
    FileNotFound(&'static str),
    /// Unexpected end of file
    #[fail(display = "Expecting {} found EOF", _0)]
    Eof(&'static str),
    /// Unexpexted error
    #[fail(display = "Expecting {} found {}", expected, found)]
    Mismatch {
        /// Expected
        expected: &'static str,
        /// Found
        found: String,
    },
}

from_err!(::std::io::Error, OdsError, Io);
from_err!(::zip::result::ZipError, OdsError, Zip);
from_err!(::quick_xml::Error, OdsError, Xml);
from_err!(::std::string::ParseError, OdsError, Parse);
from_err!(::std::num::ParseFloatError, OdsError, ParseFloat);

/// An OpenDocument Spreadsheet document parser
///
/// # Reference
/// OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
/// http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf
pub struct Ods<RS>
where
    RS: Read + Seek,
{
    sheets: HashMap<String, (Range<DataType>, Range<String>)>,
    metadata: Metadata,
    marker: PhantomData<RS>,
}

impl<RS: Read + Seek> Reader for Ods<RS> {
    type RS = RS;
    type Error = OdsError;

    fn new(reader: RS) -> Result<Self, OdsError>
    where
        RS: Read + Seek,
    {
        let mut zip = ZipArchive::new(reader)?;

        // check mimetype
        match zip.by_name("mimetype") {
            Ok(mut f) => {
                let mut buf = [0u8; 46];
                f.read_exact(&mut buf)?;
                if &buf[..] != MIMETYPE {
                    return Err(OdsError::InvalidMime(buf.to_vec()));
                }
            }
            Err(ZipError::FileNotFound) => return Err(OdsError::FileNotFound("mimetype")),
            Err(e) => return Err(OdsError::Zip(e)),
        }

        let (sheets, sheet_names, defined_names) = parse_content(zip)?;
        let metadata = Metadata {
            sheets: sheet_names,
            names: defined_names,
        };

        Ok(Ods {
            sheets: sheets,
            metadata: metadata,
            marker: PhantomData,
        })
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<VbaProject>, OdsError>> {
        None
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, OdsError>> {
        self.sheets.get(name).map(|r| Ok(r.0.to_owned()))
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, OdsError>> {
        self.sheets.get(name).map(|r| Ok(r.1.to_owned()))
    }
}

/// Parses content.xml and store the result in `self.content`
fn parse_content<RS: Read + Seek>(
    mut zip: ZipArchive<RS>,
) -> Result<
    (
        HashMap<String, (Range<DataType>, Range<String>)>,
        Vec<String>,
        Vec<(String, String)>,
    ),
    OdsError,
> {
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
    let mut sheet_names = Vec::new();
    loop {
        match reader.read_event(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == b"table:table" => if let Some(ref a) = e
                .attributes()
                .filter_map(|a| a.ok())
                .find(|a| a.key == b"table:name")
            {
                let name = a.unescape_and_decode_value(&reader).map_err(OdsError::Xml)?;
                let (range, formulas) = read_table(&mut reader)?;
                sheet_names.push(name.clone());
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
    Ok((sheets, sheet_names, defined_names))
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
            Ok(Event::Start(ref e)) if e.name() == b"table:table-cell" || e.name() == b"table:covered-table-cell" => {
                let mut repeats = 1;
                for a in e.attributes() {
                    let a = a?;
                    if a.key == b"table:number-columns-repeated" {
                        repeats = reader.decode(&a.value).parse().map_err(OdsError::ParseInt)?;
                        break;
                    }
                }

                let (value, formula, is_closed) = get_datatype(reader, e.attributes(), cell_buf)?;
                for _ in 0..repeats {
                    cells.push(value.clone());
                    formulas.push(formula.clone());
                }
                if !is_closed {
                    reader.read_to_end(e.name(), cell_buf)?;
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
                let b = &*a.value == b"TRUE" || &*a.value == b"true";
                val = DataType::Bool(b);
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
