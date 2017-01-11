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
//! let mut workbook = Excel::open(path).expect("Cannot open file");
//!
//! // Read whole worksheet data and provide some statistics
//! if let Ok(range) = workbook.worksheet_range("Sheet1") {
//!     let total_cells = range.get_size().0 * range.get_size().1;
//!     let non_empty_cells: usize = range.used_cells().count();
//!     println!("Found {} cells in 'Sheet1', including {} non empty cells",
//!              total_cells, non_empty_cells);
//!     // alternatively, we can manually filter rows
//!     assert_eq!(non_empty_cells, range.rows()
//!         .flat_map(|r| r.iter().filter(|&c| c != &DataType::Empty)).count());
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
    sheets: Vec<(String, String)>,
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
        let f = File::open(&path)?;
        let file = match path.as_ref().extension().and_then(|s| s.to_str()) {
            Some("xls") | Some("xla") => FileType::Xls(xls::Xls::new(f)?),
            Some("xlsx") | Some("xlsm") | Some("xlam") => FileType::Xlsx(xlsx::Xlsx::new(f)?),
            Some("xlsb") => FileType::Xlsb(xlsb::Xlsb::new(f)?),
            Some(e) => return Err(ErrorKind::InvalidExtension(e.to_string()).into()),
            None => return Err(ErrorKind::InvalidExtension("".to_string()).into()),
        };
        Ok(Excel { 
            file: file, 
            strings: vec![], 
            relationships: HashMap::new(),
            sheets: Vec::new(),
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
        self.initialize()?;
        let &(_, ref p) = self.sheets.iter().find(|&&(ref n, _)| n == name)
            .ok_or_else(|| ErrorKind::WorksheetName(name.to_string()))?;
        inner!(self, read_worksheet_range(p, &self.strings))
    }


    /// Get all data from `Worksheet` at index `idx` (0 based)
    ///
    /// # Examples
    /// ```
    /// use calamine::Excel;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Excel::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range_by_index(0).expect("Cannot find first sheet");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_range_by_index(&mut self, idx: usize) -> Result<Range> {
        self.initialize()?;
        let &(_, ref p) = self.sheets.get(idx).ok_or(ErrorKind::WorksheetIndex(idx))?;
        inner!(self, read_worksheet_range(p, &self.strings))
    }

    fn initialize(&mut self) -> Result<()> {
        if self.strings.is_empty() {
            self.strings = inner!(self, read_shared_strings())?;
        }
        if self.relationships.is_empty() {
            self.relationships = inner!(self, read_relationships())?;
        }
        if self.sheets.is_empty() {
            self.sheets = inner!(self, read_sheets_names(&self.relationships))?;
        }
        Ok(())
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
    pub fn sheet_names(&mut self) -> Result<Vec<&str>> {
        self.initialize()?;
        Ok(self.sheets.iter().map(|&(ref k, _)| &**k).collect())
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
    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<Vec<(String, String)>>;
    /// Read workbook relationships
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>>;
    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range>;
}

/// A struct to hold cell position and value
#[derive(Debug, Clone)]
pub struct Cell {
    /// Position for the cell (row, column)
    pos: (u32, u32),
    /// Value for the cell
    val: DataType,
}

impl Cell {

    /// Creates a new `Cell`
    pub fn new(position: (u32, u32), value: DataType) -> Cell {
        Cell {
            pos: position,
            val: value,
        }
    }

    /// Gets `Cell` position
    pub fn get_position(&self) -> (u32, u32) {
        self.pos
    }

    /// Gets `Cell` value
    pub fn get_value(&self) -> &DataType {
        &self.val
    }
}

/// A struct which represents a squared selection of cells 
#[derive(Debug, Default, Clone)]
pub struct Range {
    start: (u32, u32),
    end: (u32, u32),
    inner: Vec<DataType>,
}

impl Range {

    /// Creates a new `Range`
    ///
    /// When possible, prefer the more efficient `Range::from_sparse`
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range {
        Range {
            start: start,
            end: end,
            inner: vec![DataType::Empty; ((end.0 - start.0 + 1) 
                                          * (end.1 - start.1 + 1)) as usize],
        }
    }

    /// Creates a `Range` from a coo sparse vector of `Cell`s.
    ///
    /// Coordinate list (COO) is the natural way cells are stored in excel files 
    /// Inner size is defined only by non empty.
    ///
    /// cells: `Vec` of non empty `Cell`s, sorted by row
    /// 
    /// # Panics
    ///
    /// panics when a `Cell` row is lower than the first `Cell` row or 
    /// bigger than the last `Cell` row.
    ///
    /// # Examples
    /// 
    /// ```
    /// use calamine::{Range, DataType, Cell};
    /// 
    /// let v = vec![Cell::new((1, 200), DataType::Float(1.)),
    ///              Cell::new((55, 2),  DataType::String("a".to_string()))];
    /// let range = Range::from_sparse(v);
    /// 
    /// assert_eq!(range.get_size(), (55, 199));
    /// ```
    pub fn from_sparse(cells: Vec<Cell>) -> Range {
        if cells.is_empty() {
            Range { start: (0, 0), end: (0, 0), inner: Vec::new() }
        } else {
            // search bounds
            let row_start = cells.first().unwrap().pos.0;
            let row_end = cells.last().unwrap().pos.0;
            let mut col_start = ::std::u32::MAX;
            let mut col_end = 0;
            for c in cells.iter().map(|c| c.pos.1) {
                if c < col_start {
                    col_start = c;
                }
                if c > col_end {
                    col_end = c
                }
            }
            let width = col_end - col_start + 1;
            let len = ((row_end - row_start + 1) * width) as usize;
            let mut v = vec![DataType::Empty; len];
            v.shrink_to_fit();
            for c in cells {
                let idx = ((c.pos.0 - row_start) * width + (c.pos.1 - col_start)) as usize;
                v[idx] = c.val;
            }
            Range { start: (row_start, col_start), end: (row_end, col_end), inner: v }
        }
    }

    /// Get top left cell position (row, column)
    pub fn start(&self) -> (u32, u32) {
        self.start
    }

    /// Get bottom right cell position (row, column)
    pub fn end(&self) -> (u32, u32) {
        self.end
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
    /// Will try to resize inner structure if the value is out of bounds.
    ///
    /// Try to avoid this method as much as possible and prefer initializing
    /// the `Range` with `from_sparce` constructor.
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    /// assert_eq!(range.get_value((2, 1)), &DataType::Empty);
    /// range.set_value((2, 1), DataType::Float(1.0))
    ///     .expect("Cannot set value at position (2, 1)");
    /// assert_eq!(range.get_value((2, 1)), &DataType::Float(1.0));
    /// ```
    pub fn set_value(&mut self, pos: (u32, u32), value: DataType) -> Result<()> {
        if self.start > pos {
            return Err(ErrorKind::CellOutOfRange(pos, self.start).into());
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
    pub fn get_value(&self, pos: (u32, u32)) -> &DataType {
        assert!(pos <= self.end);
        let idx = (pos.0 - self.start.0) as usize * self.width() + (pos.1 - self.start.1) as usize;
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
    /// assert_eq!(range.rows().map(|r| r.len()).sum::<usize>(), 18);
    /// ```
    pub fn rows(&self) -> Rows {
        if self.inner.is_empty() {
            Rows { inner: None }
        } else {
            let width = self.width();
            Rows { inner: Some(self.inner.chunks(width)) }
        }
    }

    /// Get an iterator over used cells only
    ///
    /// This can be much faster than iterating rows as `Range` is saved as a sparce matrix
    pub fn used_cells(&self) -> UsedCells {
        UsedCells { width: self.width(), inner: self.inner.iter().enumerate() }
    }

}

/// A struct to iterate over used cells
#[derive(Debug)]
pub struct UsedCells<'a> {
    width: usize,
    inner: ::std::iter::Enumerate<::std::slice::Iter<'a, DataType>>,
}

impl<'a> Iterator for UsedCells<'a> {
    type Item = (usize, usize, &'a DataType);
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.by_ref().find(|&(_, v)| v != &DataType::Empty)
            .map(|(i, v)| {
                let row = i / self.width;
                let col = i % self.width;
                (row, col, v)
            })
    }
}

/// An iterator to read `Range` struct row by row
#[derive(Debug)]
pub struct Rows<'a> {
    inner: Option<::std::slice::Chunks<'a, DataType>>
}

impl<'a> Iterator for Rows<'a> {
    type Item = &'a[DataType];
    fn next(&mut self) -> Option<Self::Item> {
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
