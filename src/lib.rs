//! Rust Excel reader
//!
//! # Status
//!
//! Reads excel workbooks and vba project. This mainly works except for 
//! binary file format (xls and xlsb) where only VBA is currently supported.
//!
#![deny(missing_docs)]
#![feature(conservative_impl_trait)]

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

macro_rules! unexp {
    ($pat: expr) => {
        {
            return Err($pat.into());
        }
    };
    ($pat: expr, $($args: expr)* ) => {
        {
            return Err(format!($pat, $($args)*).into());
        }
    };
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
            None => unexp!("Sheet '{}' does not exist", name),
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
                    let mut xml = XmlReader::from_reader(BufReader::new(f))
                        .with_check(false)
                        .trim_text(false);

                    while let Some(res_event) = xml.next() {
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

                    let mut strings = Vec::new();
                    while let Some(res_event) = xml.next() {
                        match res_event {
                            Ok(Event::Start(ref e)) if e.name() == b"t" => {
                                strings.push(try!(xml.read_text(b"t")));
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
                    let mut xml = XmlReader::from_reader(BufReader::new(f))
                        .with_check(false)
                        .trim_text(false);

                    while let Some(res_event) = xml.next() {
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
                        b"dimension" => match e.attributes().filter_map(|a| a.ok())
                                .find(|&(key, _)| key == b"ref") {
                            Some((_, dim)) => {
                                let (position, size) = try!(get_dimension(try!(dim.as_str())));
                                data.position = position;
                                data.size = (size.0 as usize, size.1 as usize);
                                data.inner.reserve_exact(data.size.0 * data.size.1);
                            },
                            None => unexp!("Expecting dimension, got {:?}", e),
                        },
                        b"sheetData" => {
                            let _ = try!(data.read_sheet_data(&mut xml, strings));
                        }
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
                                                Some((_, b"s")) => { // shared string
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
                                    _name => unexp!("not v or f node"),
                                },
                                Some(Ok(Event::End(ref e))) => {
                                    if e.name() == b"c" {
                                        self.inner.push(DataType::Empty);
                                        break;
                                    }
                                }
                                None => unexp!("End of xml"),
                                _ => (),
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.name() == b"sheetData" => return Ok(()),
                _ => (),
            }
        }
        unexp!("Could not find </sheetData>")
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
fn get_dimension(dimension: &str) -> Result<((u32, u32), (u32, u32))> {
    match dimension.chars().position(|c| c == ':') {
        None => {
            get_row_column(dimension).map(|position| (position, (1, 1)))
        }, 
        Some(p) => {
            let top_left = try!(get_row_column(&dimension[..p]));
            let bottom_right = try!(get_row_column(&dimension[p + 1..]));
            Ok((top_left, (bottom_right.0 - top_left.0 + 1, bottom_right.1 - top_left.1 + 1)))
        }
    }
}

/// converts a text range name into its position (row, column)
fn get_row_column(range: &str) -> Result<(u32, u32)> {
    let mut col = 0;
    let mut pow = 1;
    let mut rowpos = range.len();
    let mut readrow = true;
    for c in range.chars().rev() {
        match c {
            '0'...'9' => {
                if readrow {
                    rowpos -= 1;
                } else {
                    unexp!("Numeric character are only allowed at the end of the range: {}", c);
                }
            }
            c @ 'A'...'Z' => {
                readrow = false;
                col += ((c as u8 - b'A') as u32 + 1) * pow;
                pow *= 26;
            },
            c @ 'a'...'z' => {
                readrow = false;
                col += ((c as u8 - b'a') as u32 + 1) * pow;
                pow *= 26;
            },
            _ => unexp!("Expecting alphanumeric character, got {:?}", c),
        }
    }
    let row = try!(range[rowpos..].parse());
    Ok((row, col))
}

