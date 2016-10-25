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
mod utils;
mod xlsb;
mod xlsx;
mod xls;
pub mod vba;

use std::path::Path;
use std::collections::HashMap;
use std::fs::File;
use std::slice::Chunks;
use std::str::FromStr;

pub use errors::*;
use vba::VbaProject;

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

impl FromStr for CellErrorType {
    type Err = errors::Error;
    fn from_str(s: &str) -> Result<Self> {
        match s {
            "#DIV/0!" => Ok(CellErrorType::Div0),
            "#N/A" => Ok(CellErrorType::NA),
            "#NAME?" => Ok(CellErrorType::Name),
            "#NULL!" => Ok(CellErrorType::Null),
            "#NUM!" => Ok(CellErrorType::Num),
            "#REF!" => Ok(CellErrorType::Ref),
            "#VALUE!" => Ok(CellErrorType::Value),
            _ => Err(format!("{} is not an excel error", s).into()),
        }
    }
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
    /// Compound File Binary Format [MS-CFB] (xls, xla)
    Xls(xls::Xls),
    /// Regular xml zipped file (xlsx, xlsm, xlam)
    Xlsx(xlsx::Xlsx),
    /// Binary zipped file (xlsb)
    Xlsb(xlsb::Xlsb),
}

/// A wrapper struct over the Excel file
pub struct Excel {
    file: FileType,
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

macro_rules! inner {
    ($s:expr, $func:ident()) => {{
        match $s.file {
            FileType::Xls(ref mut f) => f.$func(),
            FileType::Xlsx(ref mut f) => f.$func(),
            FileType::Xlsb(ref mut f) => f.$func(),
        }
    }};
    ($s:expr, $func:ident($first_arg:expr $(, $args:expr)*)) => {{
        match $s.file {
            FileType::Xls(ref mut f) => f.$func($first_arg $(, $args)*),
            FileType::Xlsx(ref mut f) => f.$func($first_arg $(, $args)*),
            FileType::Xlsb(ref mut f) => f.$func($first_arg $(, $args)*),
        }
    }};
}

impl Excel {
    /// Opens a new workbook
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Excel> {
        let f = try!(File::open(&path));
        let file = match path.as_ref().extension().and_then(|s| s.to_str()) {
            Some("xls") | Some("xla") => FileType::Xls(try!(xls::Xls::new(f))),
            Some("xlsx") | Some("xlsm") | Some("xlam") => FileType::Xlsx(try!(xlsx::Xlsx::new(f))),
            Some("xlsb") => FileType::Xlsb(try!(xlsb::Xlsb::new(f))),
            Some(e) => return Err(format!("unrecognized extension: {:?}", e).into()),
            None => return Err("expecting a file with an extension".into()),
        };
        Ok(Excel { 
            file: file, 
            strings: vec![], 
            relationships: HashMap::new(),
            sheets: HashMap::new(),
        })
    }

    /// Get all data from `Worksheet`
    pub fn worksheet_range(&mut self, name: &str) -> Result<Range> {
        if self.strings.is_empty() {
            let strings = try!(inner!(self, read_shared_strings()));
            self.strings = strings;
        }

        if self.relationships.is_empty() {
            let rels = try!(inner!(self, read_relationships()));
            self.relationships = rels;
        }

        if self.sheets.is_empty() {
            let sheets = try!(inner!(self, read_sheets_names(&self.relationships)));
            self.sheets = sheets;
        }

        match self.sheets.get(name) {
            Some(ref p) => inner!(self, read_worksheet_range(p, &self.strings)), 
            None => Err(format!("Sheet '{}' does not exist", name).into()),
        }
    }

    /// Does the workbook contain a vba project
    pub fn has_vba(&mut self) -> bool {
        inner!(self, has_vba())
    }

    /// Gets vba project
    pub fn vba_project(&mut self) -> Result<VbaProject> {
        inner!(self, vba_project())
    }
}

/// A trait to share excel reader functions accross different `FileType`s
pub trait ExcelReader: Sized {
    /// Creates a new instance based on the actual file
    fn new(f: File) -> Result<Self>;
    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool;
    /// Gets vba project
    fn vba_project(&mut self) -> Result<VbaProject>;
    /// Read shared string list
    fn read_shared_strings(&mut self) -> Result<Vec<String>>;
    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<HashMap<String, String>>;
    /// Read workbook relationships
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>>;
    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range>;
}

impl Range {
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
}

impl<'a> Iterator for Rows<'a> {
    type Item = &'a [DataType];
    fn next(&mut self) -> Option<&'a [DataType]> {
        self.inner.next()
    }
}

#[test]
fn test_parse_error() {
    assert_eq!(CellErrorType::from_str("#DIV/0!").unwrap(), CellErrorType::Div0);
    assert_eq!(CellErrorType::from_str("#N/A").unwrap(), CellErrorType::NA);
    assert_eq!(CellErrorType::from_str("#NAME?").unwrap(), CellErrorType::Name);
    assert_eq!(CellErrorType::from_str("#NULL!").unwrap(), CellErrorType::Null);
    assert_eq!(CellErrorType::from_str("#NUM!").unwrap(), CellErrorType::Num);
    assert_eq!(CellErrorType::from_str("#REF!").unwrap(), CellErrorType::Ref);
    assert_eq!(CellErrorType::from_str("#VALUE!").unwrap(), CellErrorType::Value);
}

#[test]
fn test_dimensions() {
    assert_eq!(utils::get_row_column(b"A1").unwrap(), (1, 1));
    assert_eq!(utils::get_row_column(b"C107").unwrap(), (107, 3));
    assert_eq!(utils::get_dimension(b"C2:D35").unwrap(), ((2, 3), (34, 2)));
}
