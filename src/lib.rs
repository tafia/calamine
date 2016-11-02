//! Rust Excel reader
//!
//! # Status
//!
//! Reads excel workbooks and vba project. This mainly works except for 
//! binary file format (xls and xlsb) where only VBA is currently supported.
//!
//! # Examples
//! ```
//! use office::{Excel, Range, DataType};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook = Excel::open(path).unwrap();
//!
//! // Check if the workbook has a vba project
//! if workbook.has_vba() {
//!     let mut vba = workbook.vba_project().unwrap();
//!     let module1 = vba.get_module("Module 1").unwrap();
//!     println!("Module 1 code:");
//!     println!("{}", module1);
//!     for r in vba.get_references() {
//!         if r.is_missing() {
//!             println!("Reference {} is broken or not accessible", r.name);
//!         }
//!     }
//! }
//!
//! // Read whole worksheet data and provide some statistics
//! if let Ok(range) = workbook.worksheet_range("Sheet1") {
//!     let total_cells = range.get_size().0 * range.get_size().1;
//!     let non_empty_cells: usize = range.rows().map(|r| {
//!         r.iter().filter(|cell| cell != &&DataType::Empty).count()
//!     }).sum();
//!     println!("Found {} cells in 'Sheet1', including {} non empty cells",
//!              total_cells, non_empty_cells);
//! }
//! ```
//!
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
mod cfb;
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
    /// Getting data
    GettingData,
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
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    /// assert!(Excel::open(path).is_ok());
    /// ```
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
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).unwrap();
    /// let range = workbook.worksheet_range("Sheet1").unwrap();
    /// ```
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
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).unwrap();
    /// if workbook.has_vba() {
    ///     let mut vba = workbook.vba_project().unwrap();
    ///     println!("References: {:?}", vba.get_references());
    ///     println!("Modules: {:?}", vba.get_module_names());
    /// }
    /// ```
    pub fn vba_project(&mut self) -> Result<VbaProject> {
        inner!(self, vba_project())
    }

    /// Get all sheet names of this workbook
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).unwrap();
    /// println!("Sheets: {:#?}", workbook.sheet_names());
    /// ```
    pub fn sheet_names(&mut self) -> Result<Vec<String>> {

        if self.relationships.is_empty() {
            let rels = try!(inner!(self, read_relationships()));
            self.relationships = rels;
        }

        if self.sheets.is_empty() {
            let sheets = try!(inner!(self, read_sheets_names(&self.relationships)));
            self.sheets = sheets;
        }

        Ok(self.sheets.keys().map(|k| k.to_string()).collect())
    }
}

/// A trait to share excel reader functions accross different `FileType`s
pub trait ExcelReader: Sized {
    /// Creates a new instance based on the actual file
    fn new(f: File) -> Result<Self>;
    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Result<VbaProject>;
    /// Gets vba references
    fn read_shared_strings(&mut self) -> Result<Vec<String>>;
    /// Read sheets from workbook.xml and get their corresponding path from relationships
    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<HashMap<String, String>>;
    /// Read workbook relationships
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>>;
    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range>;
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

impl Range {

    /// Creates a new range
    pub fn new(position: (u32, u32), size: (usize, usize)) -> Range {
        Range {
            position: position,
            size: size,
            inner: vec![DataType::Empty; size.0 * size.1],
        }
    }

    /// Get top left cell position (row, column)
    pub fn get_position(&self) -> (u32, u32) {
        self.position
    }

    /// Get size
    pub fn get_size(&self) -> (usize, usize) {
        self.size
    }

    /// Set inner value
    ///
    /// Panics if indexes are out of range bounds
    ///
    /// # Examples
    /// ```
    /// use office::{Range, DataType};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    /// assert_eq!(range.get_value(2, 1), &DataType::Empty);
    /// range.set_value((2, 1), DataType::Float(1.0));
    /// assert_eq!(range.get_value(2, 1), &DataType::Float(1.0));
    /// ```
    pub fn set_value(&mut self, pos: (u32, u32), value: DataType) {
        assert!(self.position <= pos);
        let idx = (pos.0 - self.position.0) * self.size.1 as u32 + pos.1 - self.position.1;
        self.inner[idx as usize] = value;
    }

    /// Get cell value
    ///
    /// Panics if indexes are out of range bounds
    pub fn get_value(&self, i: usize, j: usize) -> &DataType {
        assert!((i, j) < self.size);
        let idx = i * self.size.1 + j;
        &self.inner[idx]
    }

    /// Get an iterator over inner rows
    ///
    /// # Examples
    /// ```
    /// use office::{Range, DataType};
    ///
    /// let range = Range::new((0, 0), (5, 2));
    /// // with rows item row: &[DataType]
    /// assert_eq!(range.rows().flat_map(|row| row).count(), 10);
    /// ```
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
