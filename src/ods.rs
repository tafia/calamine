//! A module to parse Open Document Spreadsheets
//!
//! # Reference
//! OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
//! http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf

use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap};
use std::io::{BufReader, Read, Seek};

use quick_xml::events::attributes::Attributes;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::vba::VbaProject;
use crate::{Data, DataType, HeaderRow, Metadata, Range, Reader, Sheet, SheetType, SheetVisible};
use std::marker::PhantomData;

const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet";

type OdsReader<'a> = XmlReader<BufReader<ZipFile<'a>>>;

/// An enum for ods specific errors
#[derive(Debug)]
pub enum OdsError {
    /// Io error
    Io(std::io::Error),
    /// Zip error
    Zip(zip::result::ZipError),
    /// Xml error
    Xml(quick_xml::Error),
    /// Xml attribute error
    XmlAttr(quick_xml::events::attributes::AttrError),
    /// Error while parsing string
    Parse(std::string::ParseError),
    /// Error while parsing integer
    ParseInt(std::num::ParseIntError),
    /// Error while parsing float
    ParseFloat(std::num::ParseFloatError),
    /// Error while parsing bool
    ParseBool(std::str::ParseBoolError),

    /// Invalid MIME
    InvalidMime(Vec<u8>),
    /// File not found
    FileNotFound(&'static str),
    /// Unexpected end of file
    Eof(&'static str),
    /// Unexpected error
    Mismatch {
        /// Expected
        expected: &'static str,
        /// Found
        found: String,
    },
    /// Workbook is password protected
    Password,
    /// Worksheet not found
    WorksheetNotFound(String),
}

/// Ods reader options
#[derive(Debug, Default)]
#[non_exhaustive]
struct OdsOptions {
    pub header_row: HeaderRow,
}

from_err!(std::io::Error, OdsError, Io);
from_err!(zip::result::ZipError, OdsError, Zip);
from_err!(quick_xml::Error, OdsError, Xml);
from_err!(std::string::ParseError, OdsError, Parse);
from_err!(std::num::ParseFloatError, OdsError, ParseFloat);

impl std::fmt::Display for OdsError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OdsError::Io(e) => write!(f, "I/O error: {e}"),
            OdsError::Zip(e) => write!(f, "Zip error: {e:?}"),
            OdsError::Xml(e) => write!(f, "Xml error: {e}"),
            OdsError::XmlAttr(e) => write!(f, "Xml attribute error: {e}"),
            OdsError::Parse(e) => write!(f, "Parse string error: {e}"),
            OdsError::ParseInt(e) => write!(f, "Parse integer error: {e}"),
            OdsError::ParseFloat(e) => write!(f, "Parse float error: {e}"),
            OdsError::ParseBool(e) => write!(f, "Parse bool error: {e}"),
            OdsError::InvalidMime(mime) => write!(f, "Invalid MIME type: {mime:?}"),
            OdsError::FileNotFound(file) => write!(f, "'{file}' file not found in archive"),
            OdsError::Eof(node) => write!(f, "Expecting '{node}' node, found end of xml file"),
            OdsError::Mismatch { expected, found } => {
                write!(f, "Expecting '{expected}', found '{found}'")
            }
            OdsError::Password => write!(f, "Workbook is password protected"),
            OdsError::WorksheetNotFound(name) => write!(f, "Worksheet '{name}' not found"),
        }
    }
}

impl std::error::Error for OdsError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            OdsError::Io(e) => Some(e),
            OdsError::Zip(e) => Some(e),
            OdsError::Xml(e) => Some(e),
            OdsError::Parse(e) => Some(e),
            OdsError::ParseInt(e) => Some(e),
            OdsError::ParseFloat(e) => Some(e),
            _ => None,
        }
    }
}

/// An OpenDocument Spreadsheet document parser
///
/// # Reference
/// OASIS Open Document Format for Office Application 1.2 (ODF 1.2)
/// http://docs.oasis-open.org/office/v1.2/OpenDocument-v1.2.pdf
pub struct Ods<RS> {
    sheets: BTreeMap<String, (Range<Data>, Range<String>)>,
    metadata: Metadata,
    marker: PhantomData<RS>,
    #[cfg(feature = "picture")]
    pictures: Option<Vec<(String, Vec<u8>)>>,
    /// Reader options
    options: OdsOptions,
}

impl<RS> Reader<RS> for Ods<RS>
where
    RS: Read + Seek,
{
    type Error = OdsError;

    fn new(reader: RS) -> Result<Self, OdsError> {
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

        check_for_password_protected(&mut zip)?;

        #[cfg(feature = "picture")]
        let pictures = read_pictures(&mut zip)?;

        let Content {
            sheets,
            sheets_metadata,
            defined_names,
        } = parse_content(zip)?;
        let metadata = Metadata {
            sheets: sheets_metadata,
            names: defined_names,
        };

        Ok(Ods {
            marker: PhantomData,
            metadata,
            sheets,
            #[cfg(feature = "picture")]
            pictures,
            options: OdsOptions::default(),
        })
    }

    fn with_header_row(&mut self, header_row: HeaderRow) -> &mut Self {
        self.options.header_row = header_row;
        self
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, OdsError>> {
        None
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Result<Range<Data>, OdsError> {
        let sheet = self
            .sheets
            .get(name)
            .ok_or_else(|| OdsError::WorksheetNotFound(name.into()))?
            .0
            .to_owned();

        match self.options.header_row {
            HeaderRow::FirstNonEmptyRow => Ok(sheet),
            HeaderRow::Row(header_row_idx) => {
                // If `header_row` is a row index, adjust the range
                if let (Some(start), Some(end)) = (sheet.start(), sheet.end()) {
                    Ok(sheet.range((header_row_idx, start.1), end))
                } else {
                    Ok(sheet)
                }
            }
        }
    }

    fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
        self.sheets
            .iter()
            .map(|(name, (range, _formula))| (name.to_owned(), range.clone()))
            .collect()
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, OdsError> {
        self.sheets
            .get(name)
            .ok_or_else(|| OdsError::WorksheetNotFound(name.into()))
            .map(|r| r.1.to_owned())
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        self.pictures.to_owned()
    }
}

struct Content {
    sheets: BTreeMap<String, (Range<Data>, Range<String>)>,
    sheets_metadata: Vec<Sheet>,
    defined_names: Vec<(String, String)>,
}

/// Check password protection
fn check_for_password_protected<RS: Read + Seek>(zip: &mut ZipArchive<RS>) -> Result<(), OdsError> {
    let mut reader = match zip.by_name("META-INF/manifest.xml") {
        Ok(f) => {
            let mut r = XmlReader::from_reader(BufReader::new(f));
            r.check_end_names(false)
                .trim_text(false)
                .check_comments(false)
                .expand_empty_elements(true);
            r
        }
        Err(ZipError::FileNotFound) => return Err(OdsError::FileNotFound("META-INF/manifest.xml")),
        Err(e) => return Err(OdsError::Zip(e)),
    };

    let mut buf = Vec::new();
    let mut inner = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == QName(b"manifest:file-entry") => {
                loop {
                    match reader.read_event_into(&mut inner) {
                        Ok(Event::Start(ref e))
                            if e.name() == QName(b"manifest:encryption-data") =>
                        {
                            return Err(OdsError::Password)
                        }
                        Ok(Event::Eof) => break,
                        Err(e) => return Err(OdsError::Xml(e)),
                        _ => (),
                    }
                }
                inner.clear()
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(OdsError::Xml(e)),
            _ => (),
        }
        buf.clear()
    }

    Ok(())
}

/// Parses content.xml and store the result in `self.content`
fn parse_content<RS: Read + Seek>(mut zip: ZipArchive<RS>) -> Result<Content, OdsError> {
    let mut reader = match zip.by_name("content.xml") {
        Ok(f) => {
            let mut r = XmlReader::from_reader(BufReader::new(f));
            r.check_end_names(false)
                .trim_text(false)
                .check_comments(false)
                .expand_empty_elements(true);
            r
        }
        Err(ZipError::FileNotFound) => return Err(OdsError::FileNotFound("content.xml")),
        Err(e) => return Err(OdsError::Zip(e)),
    };
    let mut buf = Vec::with_capacity(1024);
    let mut sheets = BTreeMap::new();
    let mut defined_names = Vec::new();
    let mut sheets_metadata = Vec::new();
    let mut styles = HashMap::new();
    let mut style_name: Option<String> = None;
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == QName(b"style:style") => {
                style_name = e
                    .try_get_attribute(b"style:name")?
                    .map(|a| a.decode_and_unescape_value(&reader))
                    .transpose()
                    .map_err(OdsError::Xml)?
                    .map(|x| x.to_string())
            }
            Ok(Event::Start(ref e))
                if style_name.is_some() && e.name() == QName(b"style:table-properties") =>
            {
                let visible = match e.try_get_attribute(b"table:display")? {
                    Some(a) => match a
                        .decode_and_unescape_value(&reader)
                        .map_err(OdsError::Xml)?
                        .parse()
                        .map_err(OdsError::ParseBool)?
                    {
                        true => SheetVisible::Visible,
                        false => SheetVisible::Hidden,
                    },
                    None => SheetVisible::Visible,
                };
                styles.insert(style_name.clone(), visible);
            }
            Ok(Event::Start(ref e)) if e.name() == QName(b"table:table") => {
                let visible = styles
                    .get(
                        &e.try_get_attribute(b"table:style-name")?
                            .map(|a| a.decode_and_unescape_value(&reader))
                            .transpose()
                            .map_err(OdsError::Xml)?
                            .map(|x| x.to_string()),
                    )
                    .cloned()
                    .unwrap_or(SheetVisible::Visible);
                if let Some(ref a) = e
                    .attributes()
                    .filter_map(|a| a.ok())
                    .find(|a| a.key == QName(b"table:name"))
                {
                    let name = a
                        .decode_and_unescape_value(&reader)
                        .map_err(OdsError::Xml)?
                        .to_string();
                    let (range, formulas) = read_table(&mut reader)?;
                    sheets_metadata.push(Sheet {
                        name: name.clone(),
                        typ: SheetType::WorkSheet,
                        visible,
                    });
                    sheets.insert(name, (range, formulas));
                }
            }
            Ok(Event::Start(ref e)) if e.name() == QName(b"table:named-expressions") => {
                defined_names = read_named_expressions(&mut reader)?;
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(OdsError::Xml(e)),
            _ => (),
        }
        buf.clear();
    }
    Ok(Content {
        sheets,
        sheets_metadata,
        defined_names,
    })
}

fn read_table(reader: &mut OdsReader<'_>) -> Result<(Range<Data>, Range<String>), OdsError> {
    let mut cells = Vec::new();
    let mut rows_repeats = Vec::new();
    let mut formulas = Vec::new();
    let mut cols = Vec::new();
    let mut buf = Vec::with_capacity(1024);
    let mut row_buf = Vec::with_capacity(1024);
    let mut cell_buf = Vec::with_capacity(1024);
    cols.push(0);
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.name() == QName(b"table:table-row") => {
                let row_repeats = match e.try_get_attribute(b"table:number-rows-repeated")? {
                    Some(c) => c
                        .decode_and_unescape_value(reader)
                        .map_err(OdsError::Xml)?
                        .parse()
                        .map_err(OdsError::ParseInt)?,
                    None => 1,
                };
                read_row(
                    reader,
                    &mut row_buf,
                    &mut cell_buf,
                    &mut cells,
                    &mut formulas,
                )?;
                cols.push(cells.len());
                rows_repeats.push(row_repeats);
            }
            Ok(Event::End(ref e)) if e.name() == QName(b"table:table") => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(_) => (),
        }
        buf.clear();
    }
    Ok((
        get_range(cells, &cols, &rows_repeats),
        get_range(formulas, &cols, &rows_repeats),
    ))
}

fn is_empty_row<T: Default + Clone + PartialEq>(row: &[T]) -> bool {
    row.iter().all(|x| x == &T::default())
}

fn get_range<T: Default + Clone + PartialEq>(
    mut cells: Vec<T>,
    cols: &[usize],
    rows_repeats: &[usize],
) -> Range<T> {
    // find smallest area with non empty Cells
    let mut row_min = None;
    let mut row_max = 0;
    let mut col_min = usize::MAX;
    let mut col_max = 0;
    let mut first_empty_rows_repeated = 0;
    {
        for (i, w) in cols.windows(2).enumerate() {
            let row = &cells[w[0]..w[1]];
            if let Some(p) = row.iter().position(|c| c != &T::default()) {
                if row_min.is_none() {
                    row_min = Some(i);
                    first_empty_rows_repeated =
                        rows_repeats.iter().take(i).sum::<usize>().saturating_sub(i);
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
    {
        let mut new_cells = Vec::with_capacity(cells_len);
        let empty_cells = vec![T::default(); col_max + 1];
        let mut empty_row_repeats = 0;
        let mut consecutive_empty_rows = 0;
        for (w, row_repeats) in cols
            .windows(2)
            .skip(row_min)
            .take(row_max + 1)
            .zip(rows_repeats.iter().skip(row_min).take(row_max + 1))
        {
            let row = &cells[w[0]..w[1]];
            let row_repeats = *row_repeats;

            if is_empty_row(row) {
                empty_row_repeats += row_repeats;
                consecutive_empty_rows += 1;
                continue;
            }

            if empty_row_repeats > 0 {
                row_max = row_max + empty_row_repeats - consecutive_empty_rows;
                for _ in 0..empty_row_repeats {
                    new_cells.extend_from_slice(&empty_cells);
                }
                empty_row_repeats = 0;
                consecutive_empty_rows = 0;
            };

            if row_repeats > 1 {
                row_max = row_max + row_repeats - 1;
            };

            for _ in 0..row_repeats {
                match row.len().cmp(&(col_max + 1)) {
                    std::cmp::Ordering::Less => {
                        new_cells.extend_from_slice(&row[col_min..]);
                        new_cells.extend_from_slice(&empty_cells[row.len()..]);
                    }
                    std::cmp::Ordering::Equal => {
                        new_cells.extend_from_slice(&row[col_min..]);
                    }
                    std::cmp::Ordering::Greater => {
                        new_cells.extend_from_slice(&row[col_min..=col_max]);
                    }
                }
            }
        }
        cells = new_cells;
    }
    let row_min = row_min + first_empty_rows_repeated;
    let row_max = row_max + first_empty_rows_repeated;
    Range {
        start: (row_min as u32, col_min as u32),
        end: (row_max as u32, col_max as u32),
        inner: cells,
    }
}

fn read_row(
    reader: &mut OdsReader<'_>,
    row_buf: &mut Vec<u8>,
    cell_buf: &mut Vec<u8>,
    cells: &mut Vec<Data>,
    formulas: &mut Vec<String>,
) -> Result<(), OdsError> {
    let mut empty_col_repeats = 0;
    loop {
        row_buf.clear();
        match reader.read_event_into(row_buf) {
            Ok(Event::Start(ref e))
                if e.name() == QName(b"table:table-cell")
                    || e.name() == QName(b"table:covered-table-cell") =>
            {
                let mut repeats = 1;
                for a in e.attributes() {
                    let a = a.map_err(OdsError::XmlAttr)?;
                    if a.key == QName(b"table:number-columns-repeated") {
                        repeats = reader
                            .decoder()
                            .decode(&a.value)?
                            .parse()
                            .map_err(OdsError::ParseInt)?;
                        break;
                    }
                }

                let (value, formula, is_closed) = get_datatype(reader, e.attributes(), cell_buf)?;

                for _ in 0..empty_col_repeats {
                    cells.push(Data::Empty);
                    formulas.push("".to_string());
                }
                empty_col_repeats = 0;

                if value.is_empty() && formula.is_empty() {
                    empty_col_repeats = repeats;
                } else {
                    for _ in 0..repeats {
                        cells.push(value.clone());
                        formulas.push(formula.clone());
                    }
                }
                if !is_closed {
                    reader.read_to_end_into(e.name(), cell_buf)?;
                }
            }
            Ok(Event::End(ref e)) if e.name() == QName(b"table:table-row") => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(e) => {
                return Err(OdsError::Mismatch {
                    expected: "table-cell",
                    found: format!("{:?}", e),
                });
            }
        }
    }
    Ok(())
}

/// Converts table-cell element into a `Data`
///
/// ODF 1.2-19.385
fn get_datatype(
    reader: &mut OdsReader<'_>,
    atts: Attributes<'_>,
    buf: &mut Vec<u8>,
) -> Result<(Data, String, bool), OdsError> {
    let mut is_string = false;
    let mut is_value_set = false;
    let mut val = Data::Empty;
    let mut formula = String::new();
    for a in atts {
        let a = a.map_err(OdsError::XmlAttr)?;
        match a.key {
            QName(b"office:value") if !is_value_set => {
                let v = reader.decoder().decode(&a.value)?;
                val = Data::Float(v.parse().map_err(OdsError::ParseFloat)?);
                is_value_set = true;
            }
            QName(b"office:string-value" | b"office:date-value" | b"office:time-value")
                if !is_value_set =>
            {
                let attr = a
                    .decode_and_unescape_value(reader)
                    .map_err(OdsError::Xml)?
                    .to_string();
                val = match a.key {
                    QName(b"office:date-value") => Data::DateTimeIso(attr),
                    QName(b"office:time-value") => Data::DurationIso(attr),
                    _ => Data::String(attr),
                };
                is_value_set = true;
            }
            QName(b"office:boolean-value") if !is_value_set => {
                let b = &*a.value == b"TRUE" || &*a.value == b"true";
                val = Data::Bool(b);
                is_value_set = true;
            }
            QName(b"office:value-type") if !is_value_set => is_string = &*a.value == b"string",
            QName(b"table:formula") => {
                formula = a
                    .decode_and_unescape_value(reader)
                    .map_err(OdsError::Xml)?
                    .to_string();
            }
            _ => (),
        }
    }
    if !is_value_set && is_string {
        // If the value type is string and the office:string-value attribute
        // is not present, the element content defines the value.
        let mut s = String::new();
        let mut first_paragraph = true;
        loop {
            buf.clear();
            match reader.read_event_into(buf) {
                Ok(Event::Text(ref e)) => {
                    s.push_str(&e.unescape()?);
                }
                Ok(Event::End(ref e))
                    if e.name() == QName(b"table:table-cell")
                        || e.name() == QName(b"table:covered-table-cell") =>
                {
                    return Ok((Data::String(s), formula, true));
                }
                Ok(Event::Start(ref e)) if e.name() == QName(b"office:annotation") => loop {
                    match reader.read_event_into(buf) {
                        Ok(Event::End(ref e)) if e.name() == QName(b"office:annotation") => {
                            break;
                        }
                        Err(e) => return Err(OdsError::Xml(e)),
                        _ => (),
                    }
                },
                Ok(Event::Start(ref e)) if e.name() == QName(b"text:p") => {
                    if first_paragraph {
                        first_paragraph = false;
                    } else {
                        s.push('\n');
                    }
                }
                Ok(Event::Start(ref e)) if e.name() == QName(b"text:s") => {
                    let count = match e.try_get_attribute("text:c")? {
                        Some(c) => c
                            .decode_and_unescape_value(reader)
                            .map_err(OdsError::Xml)?
                            .parse()
                            .map_err(OdsError::ParseInt)?,
                        None => 1,
                    };
                    for _ in 0..count {
                        s.push(' ');
                    }
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

fn read_named_expressions(reader: &mut OdsReader<'_>) -> Result<Vec<(String, String)>, OdsError> {
    let mut defined_names = Vec::new();
    let mut buf = Vec::with_capacity(512);
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e))
                if e.name() == QName(b"table:named-range")
                    || e.name() == QName(b"table:named-expression") =>
            {
                let mut name = String::new();
                let mut formula = String::new();
                for a in e.attributes() {
                    let a = a.map_err(OdsError::XmlAttr)?;
                    match a.key {
                        QName(b"table:name") => {
                            name = a
                                .decode_and_unescape_value(reader)
                                .map_err(OdsError::Xml)?
                                .to_string();
                        }
                        QName(b"table:cell-range-address" | b"table:expression") => {
                            formula = a
                                .decode_and_unescape_value(reader)
                                .map_err(OdsError::Xml)?
                                .to_string();
                        }
                        _ => (),
                    }
                }
                defined_names.push((name, formula));
            }
            Ok(Event::End(ref e))
                if e.name() == QName(b"table:named-range")
                    || e.name() == QName(b"table:named-expression") => {}
            Ok(Event::End(ref e)) if e.name() == QName(b"table:named-expressions") => break,
            Err(e) => return Err(OdsError::Xml(e)),
            Ok(e) => {
                return Err(OdsError::Mismatch {
                    expected: "table:named-expressions",
                    found: format!("{:?}", e),
                });
            }
        }
    }
    Ok(defined_names)
}

/// Read pictures
#[cfg(feature = "picture")]
fn read_pictures<RS: Read + Seek>(
    zip: &mut ZipArchive<RS>,
) -> Result<Option<Vec<(String, Vec<u8>)>>, OdsError> {
    let mut pics = Vec::new();
    for i in 0..zip.len() {
        let mut zfile = zip.by_index(i)?;
        let zname = zfile.name();
        // no Thumbnails
        if zname.starts_with("Pictures") {
            if let Some(ext) = zname.split('.').last() {
                if [
                    "emf", "wmf", "pict", "jpeg", "jpg", "png", "dib", "gif", "tiff", "eps", "bmp",
                    "wpg",
                ]
                .contains(&ext)
                {
                    let ext = ext.to_string();
                    let mut buf: Vec<u8> = Vec::new();
                    zfile.read_to_end(&mut buf)?;
                    pics.push((ext, buf));
                }
            }
        }
    }
    if pics.is_empty() {
        Ok(None)
    } else {
        Ok(Some(pics))
    }
}
