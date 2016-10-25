//! Rust Excel reader
//!
//! # Status
//!
//! Reads excel workbooks and vba project. This mainly works except for 
//! binary file format (xls and xlsb) where only VBA is currently supported.
//!
#![deny(missing_docs)]

extern crate zip;
extern crate quick_xml;
extern crate encoding;
extern crate byteorder;
#[macro_use]
extern crate error_chain;

#[macro_use]
extern crate log;

mod errors;
pub mod vba;

use std::path::Path;
use std::fs::File;
use std::io::BufReader;
use std::collections::HashMap;
use std::slice::Chunks;

pub use errors::*;
use vba::VbaProject;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::{XmlReader, Event, AsStr};

// https://msdn.microsoft.com/en-us/library/office/ff839168.aspx
/// An enum to represent all different excel errors that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq)]
pub enum CellErrorType {
    /// Division by 0 error
    Div0,
    /// Unavailable value error
    NA,
    /// Invalid name error
    Name,
    /// Null value error
    Null,
    /// Number error
    Num,
    /// Invalid cell reference error
    Ref,
    /// Value error
    Value,
}

/// An enum to represent all different excel data types that can appear as 
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq)]
pub enum DataType {
    /// Unsigned integer
    Int(i64),
    /// Float
    Float(f64),
    /// String
    String(String),
    /// Boolean
    Bool(bool),
    /// Error
    Error(CellErrorType),
    /// Empty cell
    Empty,
}

/// Excel file types
enum FileType {
    /// Compound File Binary Format [MS-CFB] (xls and xlsb)
    CFB(File),
    /// Regular file (xlsx, xlsm, xlam)
    Zip(ZipArchive<File>),
}

/// A wrapper struct over the Excel file
pub struct Excel {
    zip: FileType,
    strings: Vec<String>,
    relationships: HashMap<Vec<u8>, String>,
    /// Map of sheet names/sheet path within zip archive
    sheets: HashMap<String, String>,
}

/// A struct which represents a squared selection of cells 
#[derive(Debug, Default)]
pub struct Range {
    position: (u32, u32),
    size: (usize, usize),
    inner: Vec<DataType>,
}

/// An iterator to read `Range` struct row by row
pub struct Rows<'a> {
    inner: Chunks<'a, DataType>,
}

impl Excel {

    /// Opens a new workbook
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Excel> {
        let f = try!(File::open(&path));
        let zip = match path.as_ref().extension().and_then(|s| s.to_str()) {
            Some("xls") | Some("xla") => FileType::CFB(f),
            Some("xlsx") | Some("xlsb") | Some("xlsm") | 
                Some("xlam") => FileType::Zip(try!(ZipArchive::new(f))),
            Some(e) => return Err(format!("unrecognized extension: {:?}", e).into()),
            None => return Err("expecting a file with an extension".into()),
        };
        Ok(Excel { 
            zip: zip, 
            strings: vec![], 
            relationships: HashMap::new(),
            sheets: HashMap::new(),
        })
    }

    /// Does the workbook contain a vba project
    pub fn has_vba(&mut self) -> bool {
        match self.zip {
            FileType::CFB(_) => true,
            FileType::Zip(ref mut z) => z.by_name("xl/vbaProject.bin").is_ok()
        }
    }

    /// Gets vba project
    pub fn vba_project(&mut self) -> Result<VbaProject> {
        match self.zip {
            FileType::CFB(ref mut f) => {
                let len = try!(f.metadata()).len() as usize;
                VbaProject::new(f, len)
            },
            FileType::Zip(ref mut z) => {
                let f = try!(z.by_name("xl/vbaProject.bin"));
                let len = f.size() as usize;
                VbaProject::new(f, len)
            }
        }
    }

    /// Get all data from `Worksheet`
    pub fn worksheet_range(&mut self, name: &str) -> Result<Range> {
        try!(self.read_shared_strings());
        try!(self.read_relationships());
        try!(self.read_sheets_names());
        let strings = &self.strings;
        let z = match self.zip {
            FileType::CFB(_) => return Err("worksheet_range not implemented for CFB files".into()),
            FileType::Zip(ref mut z) => z
        };
        let ws = match self.sheets.get(name) {
            Some(p) => try!(z.by_name(p)),
            None => return Err(format!("Sheet '{}' does not exist", name).into()),
        };
        Range::from_worksheet(ws, strings)
    }

    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn read_sheets_names(&mut self) -> Result<()> {
        if self.sheets.is_empty() {
            let z = match self.zip {
                FileType::CFB(_) => return Err("read_sheet_names not implemented for CFB files".into()),
                FileType::Zip(ref mut z) => z
            };

            match z.by_name("xl/workbook.xml") {
                Ok(f) => {
                    let xml = XmlReader::from_reader(BufReader::new(f))
                        .with_check(false)
                        .trim_text(false);

                    for res_event in xml {
                        match res_event {
                            Ok(Event::Start(ref e)) if e.name() == b"sheet" => {
                                let mut name = String::new();
                                let mut path = String::new();
                                for a in e.attributes() {
                                    match try!(a) {
                                        (b"name", v) => name = try!(v.as_str()).to_string(),
                                        (b"r:id", v) => path = format!("xl/{}", self.relationships[v]),
                                        _ => (),
                                    }
                                }
                                self.sheets.insert(name, path);
                            }
                            Err(e) => return Err(e.into()),
                            _ => (),
                        }
                    }
                },
                Err(ZipError::FileNotFound) => (),
                Err(e) => return Err(e.into()),
            }
        }
        Ok(())
    }

    /// Read shared string list
    fn read_shared_strings(&mut self) -> Result<()> {
        if self.strings.is_empty() {
            let z = match self.zip {
                FileType::CFB(_) => return Err("read_shared_strings not implemented for CFB files".into()),
                FileType::Zip(ref mut z) => z
            };
            match z.by_name("xl/sharedStrings.xml") {
                Ok(f) => {
                    let mut xml = XmlReader::from_reader(BufReader::new(f))
                        .with_check(false)
                        .trim_text(false);

                    let mut rich_buffer: Option<String> = None;
                    let mut strings = Vec::new();
                    while let Some(res_event) = xml.next() {
                        match res_event {
                            Ok(Event::Start(ref e)) if e.name() == b"r" => {
                                if let None = rich_buffer {
                                    // use a buffer since richtext has multiples <r> and <t> for the same cell
                                    rich_buffer = Some(String::new());
                                }
                            },
                            Ok(Event::End(ref e)) if e.name() == b"si" => {
                                if let Some(s) = rich_buffer {
                                    strings.push(s);
                                    rich_buffer = None;
                                }
                            },
                            Ok(Event::Start(ref e)) if e.name() == b"t" => {
                                let value = try!(xml.read_text(b"t"));
                                if let Some(ref mut s) = rich_buffer {
                                    s.push_str(&value);
                                } else {
                                    strings.push(value);
                                }
                            }
                            Err(e) => return Err(e.into()),
                            _ => (),
                        }
                    }
                    self.strings = strings;
                },
                Err(ZipError::FileNotFound) => (),
                Err(e) => return Err(e.into()),
            }
        }

        Ok(())
    }

    /// Read workbook relationships
    fn read_relationships(&mut self) -> Result<()> {
        if self.relationships.is_empty() {
            let z = match self.zip {
                FileType::CFB(_) => return Err("read_relationships not implemented for CFB files".into()),
                FileType::Zip(ref mut z) => z
            };
            match z.by_name("xl/_rels/workbook.xml.rels") {
                Ok(f) => {
                    let xml = XmlReader::from_reader(BufReader::new(f))
                        .with_check(false)
                        .trim_text(false);

                    for res_event in xml {
                        match res_event {
                            Ok(Event::Start(ref e)) if e.name() == b"Relationship" => {
                                let mut id = Vec::new();
                                let mut target = String::new();
                                for a in e.attributes() {
                                    match try!(a) {
                                        (b"Id", v) => id.extend_from_slice(v),
                                        (b"Target", v) => target = try!(v.as_str()).to_string(),
                                        _ => (),
                                    }
                                }
                                self.relationships.insert(id, target);
                            }
                            Err(e) => return Err(e.into()),
                            _ => (),
                        }
                    }
                },
                Err(ZipError::FileNotFound) => (),
                Err(e) => return Err(e.into()),
            }
        }

        Ok(())
    }

}

impl Range {

    /// open a xml `ZipFile` reader and read content of *sheetData* and *dimension* node
    fn from_worksheet(xml: ZipFile, strings: &[String]) -> Result<Range> {
        let mut xml = XmlReader::from_reader(BufReader::new(xml))
            .with_check(false)
            .trim_text(false);
        let mut data = Range::default();
        while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(e.into()),
                Ok(Event::Start(ref e)) => {
                    match e.name() {
                        b"dimension" => {
                            let mut dim = None;
                            for a in e.attributes() {
                                if let (b"ref", rdim) = try!(a) {
                                    dim = Some(rdim);
                                    break;
                                }
                            }
                            match dim {
                                None => return Err(format!("Expecting dimension, got {:?}", e).into()),
                                Some(dim) => {
                                    let (position, size) = try!(get_dimension(dim));
                                    data.position = position;
                                    data.size = (size.0 as usize, size.1 as usize);
                                    data.inner.reserve_exact(data.size.0 * data.size.1);
                                }
                            }
                        },
                        b"sheetData" => try!(data.read_sheet_data(&mut xml, strings)),
                        _ => (),
                    }
                },
                _ => (),
            }
        }
        data.inner.shrink_to_fit();
        Ok(data)
    }
    
    /// get worksheet position (row, column)
    pub fn get_position(&self) -> (u32, u32) {
        self.position
    }

    /// get size
    pub fn get_size(&self) -> (usize, usize) {
        self.size
    }

    /// get cell value
    pub fn get_value(&self, i: usize, j: usize) -> &DataType {
        assert!((i, j) < self.size);
        let idx = i * self.size.1 + j;
        &self.inner[idx]
    }

    /// get an iterator over inner rows
    pub fn rows(&self) -> Rows {
        let width = self.size.1;
        Rows { inner: self.inner.chunks(width) }
    }

    /// read sheetData node
    fn read_sheet_data(&mut self, xml: &mut XmlReader<BufReader<ZipFile>>, strings: &[String]) 
        -> Result<()> 
    {
        while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(e.into()),
                Ok(Event::Start(ref c_element)) => {
                    if c_element.name() == b"c" {
                        loop {
                            match xml.next() {
                                Some(Err(e)) => return Err(e.into()),
                                Some(Ok(Event::Start(ref e))) => match e.name() {
                                    b"v" => {
                                        // value
                                        let v = try!(xml.read_text(b"v"));
                                        let value = match c_element.attributes()
                                            .filter_map(|a| a.ok())
                                            .find(|&(k, _)| k == b"t") {
                                                Some((_, b"s")) => {
                                                    // shared string
                                                    let idx: usize = try!(v.parse());
                                                    DataType::String(strings[idx].clone())
                                                },
                                                Some((_, b"str")) => {
                                                    // regular string
                                                    DataType::String(v)
                                                },
                                                Some((_, b"b")) => {
                                                    // boolean
                                                    DataType::Bool(v != "0")
                                                },
                                                Some((_, b"e")) => {
                                                    // error
                                                    DataType::Error(parse_error(v.as_ref()))
                                                },
                                                _ => match v.parse() {
                                                    // TODO: check in styles to know which type is
                                                    // supposed to be used
                                                    Ok(i) => DataType::Int(i),
                                                    Err(_) => try!(v.parse().map(DataType::Float)),
                                                },
                                            };
                                        self.inner.push(value);
                                        break;
                                    },
                                    b"f" => (), // formula, ignore
                                    _name => return Err("not v or f node".into()),
                                },
                                Some(Ok(Event::End(ref e))) => {
                                    if e.name() == b"c" {
                                        self.inner.push(DataType::Empty);
                                        break;
                                    }
                                }
                                None => return Err("End of xml".into()),
                                _ => (),
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.name() == b"sheetData" => return Ok(()),
                _ => (),
            }
        }
        Err("Could not find </sheetData>".into())
    }

}

impl<'a> Iterator for Rows<'a> {
    type Item = &'a [DataType];
    fn next(&mut self) -> Option<&'a [DataType]> {
        self.inner.next()
    }
}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column), 
/// - size (width, height)
fn get_dimension(dimension: &[u8]) -> Result<((u32, u32), (u32, u32))> {
    let parts: Vec<_> = try!(dimension.split(|c| *c == b':')
        .map(|s| get_row_column(s))
        .collect::<Result<Vec<_>>>());

    match parts.len() {
        0 => Err("dimension cannot be empty".into()),
        1 => Ok((parts[0], (1, 1))),
        2 => Ok((parts[0], (parts[1].0 - parts[0].0 + 1, parts[1].1 - parts[0].1 + 1))),
        len => Err(format!("range dimension has 0 or 1 ':', got {}", len).into()),
    }
}

/// converts a text range name into its position (row, column)
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
                        at the end of the range: {:x}", c).into());
                }
            }
            c @ b'A'...b'Z' => {
                if readrow { 
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'A') as u32 + 1) * pow;
                pow *= 26;
            },
            c @ b'a'...b'z' => {
                if readrow { 
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'a') as u32 + 1) * pow;
                pow *= 26;
            },
            _ => return Err(format!("Expecting alphanumeric character, got {:x}", c).into()),
        }
    }
    Ok((row, col))
}

/// converts a string into an `CellErrorType`
fn parse_error(v: &str) -> CellErrorType {
    match v {
        "#DIV/0!" => CellErrorType::Div0,
        "#N/A" => CellErrorType::NA,
        "#NAME?" => CellErrorType::Name,
        "#NULL!" => CellErrorType::Null,
        "#NUM!" => CellErrorType::Num,
        "#REF!" => CellErrorType::Ref,
        "#VALUE!" => CellErrorType::Value,
        _ => unimplemented!(),
    }
}

#[test]
fn test_parse_error() {
    assert_eq!(parse_error("#DIV/0!"), CellErrorType::Div0);
    assert_eq!(parse_error("#N/A"), CellErrorType::NA);
    assert_eq!(parse_error("#NAME?"), CellErrorType::Name);
    assert_eq!(parse_error("#NULL!"), CellErrorType::Null);
    assert_eq!(parse_error("#NUM!"), CellErrorType::Num);
    assert_eq!(parse_error("#REF!"), CellErrorType::Ref);
    assert_eq!(parse_error("#VALUE!"), CellErrorType::Value);
}

#[test]
fn test_dimensions() {
    assert_eq!(get_row_column(b"A1").unwrap(), (1, 1));
    assert_eq!(get_row_column(b"C107").unwrap(), (107, 3));
    assert_eq!(get_dimension(b"C2:D35").unwrap(), ((2, 3), (34, 2)));
}
