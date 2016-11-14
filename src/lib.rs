//! Rust Excel reader
//!
//! # Status
//!
//! **calamine** is a pure Rust library to read any excel file (`xls`, `xlsx`, `xlsm`, `xlsb`). 
//! 
//! Read both cell values and vba project.
//!
//! # Examples
//! ```
//! use calamine::{Excel, DataType};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook = Excel::open(path).unwrap();
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
//!
//! // Check if the workbook has a vba project
//! if workbook.has_vba() {
//!     let mut vba = workbook.vba_project().expect("Cannot find VbaProject");
//!     let vba = vba.to_mut();
//!     let module1 = vba.get_module("Module 1").unwrap();
//!     println!("Module 1 code:");
//!     println!("{}", module1);
//!     for r in vba.get_references() {
//!         if r.is_missing() {
//!             println!("Reference {} is broken or not accessible", r.name);
//!         }
//!     }
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
use std::borrow::Cow;

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
    /// use calamine::Excel;
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
    /// use calamine::Excel;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range("Sheet1").expect("Cannot find Sheet1");
    /// println!("Used range size: {:?}", range.get_size());
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
    /// use calamine::Excel;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).unwrap();
    /// if workbook.has_vba() {
    ///     let vba = workbook.vba_project().expect("Cannot find vba project");
    ///     println!("References: {:?}", vba.get_references());
    ///     println!("Modules: {:?}", vba.get_module_names());
    /// }
    /// ```
    pub fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        inner!(self, vba_project())
    }

    /// Get all sheet names of this workbook
    ///
    /// # Examples
    /// ```
    /// use calamine::Excel;
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
    fn vba_project(&mut self) -> Result<Cow<VbaProject>>;
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
#[derive(Debug, Default, Clone)]
pub struct Range {
    start: (u32, u32),
    end: (u32, u32),
    inner: Vec<DataType>,
}

/// An iterator to read `Range` struct row by row
pub struct Rows<'a> {
    inner: Option<Chunks<'a, DataType>>,
}

impl Range {

    /// Creates a new range
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range {
        Range {
            start: start,
            end: end,
            inner: vec![DataType::Empty; ((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize],
        }
    }

    /// Get top left cell position (row, column)
    pub fn get_position(&self) -> (u32, u32) {
        self.start
    }

    /// Get column width
    pub fn width(&self) -> usize {
        (self.end.1 - self.start.1 + 1) as usize
    }

    /// Get column width
    pub fn height(&self) -> usize {
        (self.end.0 - self.start.0 + 1) as usize
    }

    /// Get size
    pub fn get_size(&self) -> (usize, usize) {
        (self.height(), self.width())
    }

    /// Is range empty
    pub fn is_empty(&self) -> bool {
        self.start.0 > self.end.0 || self.start.1 > self.end.1
    }

    /// Set inner value
    ///
    /// Panics if indexes are out of range bounds
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    /// assert_eq!(range.get_value(2, 1), &DataType::Empty);
    /// range.set_value((2, 1), DataType::Float(1.0)).expect("Could not set value");
    /// assert_eq!(range.get_value(2, 1), &DataType::Float(1.0));
    /// ```
    pub fn set_value(&mut self, pos: (u32, u32), value: DataType) -> Result<()> {
        if self.start > pos {
            return Err(format!("invalid position, range start {:?} > position {:?}", 
                               self.start, pos).into());
        }

        // check if we need to change range dimension (strangely happens sometimes ...)
        match (self.end.0 < pos.0 , self.end.1 < pos.1) {
            (false, false) => (), // regular case, position within bounds
            (true, false) => {
                let len = (pos.0 - self.end.0 + 1) as usize * self.width();
                self.inner.extend_from_slice(&vec![DataType::Empty; len]);
                self.end.0 = pos.0;
            }, // missing some rows
            (e, true) => {
                let height = if e { 
                    (pos.0 - self.start.0 + 1) as usize 
                } else {
                    self.height() 
                };
                let width = (pos.1 - self.start.1 + 1) as usize;
                let old_width = self.width();
                let mut data = Vec::with_capacity(width * height);
                for sce in self.inner.chunks(old_width) {
                    data.extend_from_slice(sce);
                    data.extend_from_slice(&vec![DataType::Empty; width - old_width]);
                }
                data.extend_from_slice(&vec![DataType::Empty; width * (height - self.height())]);
                if e { self.end = pos } else { self.end.1 = pos.1 }
                self.inner = data;
            }, // missing some columns
        }

        let pos = (pos.0 - self.start.0, pos.1 - self.start.1);
        let idx = pos.0 as usize * self.width() + pos.1 as usize;
        self.inner[idx] = value;
        Ok(())
    }

    /// Get cell value
    ///
    /// Panics if indexes are out of range bounds
    pub fn get_value(&self, i: u32, j: u32) -> &DataType {
        assert!((i, j) < self.end);
        let idx = i as usize * self.width() + j as usize;
        &self.inner[idx]
    }

    /// Get an iterator over inner rows
    ///
    /// # Examples
    /// ```
    /// use calamine::Range;
    ///
    /// let range = Range::new((0, 0), (5, 2));
    /// // with rows item row: &[DataType]
    /// assert_eq!(range.rows().flat_map(|row| row).count(), 18);
    /// ```
    pub fn rows(&self) -> Rows {
        if self.inner.is_empty() {
            Rows { inner: None }
        } else {
            let width = self.width();
            Rows { inner: Some(self.inner.chunks(width)) }
        }
    }
}

impl<'a> Iterator for Rows<'a> {
    type Item = &'a [DataType];
    fn next(&mut self) -> Option<&'a [DataType]> {
        self.inner.as_mut().and_then(|c| c.next())
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
