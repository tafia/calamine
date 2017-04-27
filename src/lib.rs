//! Rust Excel/OpenDocument reader
//!
//! # Status
//!
//! **calamine** is a pure Rust library to read Excel and OpenDocument Spreasheet files.
//!
//! Read both cell values and vba project.
//!
//! # Examples
//! ```
//! use calamine::{Sheets, DataType};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook = Sheets::open(path).expect("Cannot open file");
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
//!
//! // You can also get defined names definition (string representation only)
//! for &(ref name, ref formula) in workbook.defined_names().expect("Cannot get defined names!") {
//!     println!("name: {}, formula: {}", name, formula);
//! }
//! ```

#![deny(missing_docs)]

extern crate zip;
extern crate quick_xml;
extern crate encoding_rs;
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
mod ods;
pub mod vba;

use std::borrow::Cow;
use std::fs::File;
use std::ops::{Index, IndexMut};
use std::path::Path;
use std::str::FromStr;

pub use errors::*;
use vba::VbaProject;

// https://msdn.microsoft.com/en-us/library/office/ff839168.aspx
/// An enum to represent all different errors that can appear as
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
            _ => Err(format!("Unsupported error '{}'", s).into()),
        }
    }
}

/// An enum to represent all different data types that can appear as
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

impl Default for DataType {
    fn default() -> DataType {
        DataType::Empty
    }
}

/// File types
enum FileType {
    /// Compound File Binary Format [MS-CFB] (xls, xla)
    Xls(xls::Xls),
    /// Regular xml zipped file (xlsx, xlsm, xlam)
    Xlsx(xlsx::Xlsx),
    /// Binary zipped file (xlsb)
    Xlsb(xlsb::Xlsb),
    /// OpenDocument Spreadsheet Document
    Ods(ods::Ods),
}

/// Common file metadata
///
/// Depending on file type, some extra information may be stored
/// in the Reader implementations
#[derive(Debug, Default)]
struct Metadata {
    sheets: Vec<String>,
    /// Map of sheet names/sheet path within zip archive
    defined_names: Vec<(String, String)>,
}

/// A wrapper struct over the spreadsheet file
pub struct Sheets {
    file: FileType,
    metadata: Metadata,
}

macro_rules! inner {
    ($s:expr, $func:ident()) => {{
        match $s.file {
            FileType::Xls(ref mut f) => f.$func(),
            FileType::Xlsx(ref mut f) => f.$func(),
            FileType::Xlsb(ref mut f) => f.$func(),
            FileType::Ods(ref mut f) => f.$func(),
        }
    }};
    ($s:expr, $func:ident($first_arg:expr $(, $args:expr)*)) => {{
        match $s.file {
            FileType::Xls(ref mut f) => f.$func($first_arg $(, $args)*),
            FileType::Xlsx(ref mut f) => f.$func($first_arg $(, $args)*),
            FileType::Xlsb(ref mut f) => f.$func($first_arg $(, $args)*),
            FileType::Ods(ref mut f) => f.$func($first_arg $(, $args)*),
        }
    }};
}

impl Sheets {
    /// Opens a new workbook
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    /// assert!(Sheets::open(path).is_ok());
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Sheets> {
        let f = File::open(&path)?;
        let file = match path.as_ref().extension().and_then(|s| s.to_str()) {
            Some("xls") | Some("xla") => FileType::Xls(xls::Xls::new(f)?),
            Some("xlsx") | Some("xlsm") | Some("xlam") => FileType::Xlsx(xlsx::Xlsx::new(f)?),
            Some("xlsb") => FileType::Xlsb(xlsb::Xlsb::new(f)?),
            Some("ods") => FileType::Ods(ods::Ods::new(f)?),
            Some(e) => return Err(ErrorKind::InvalidExtension(e.to_string()).into()),
            None => return Err(ErrorKind::InvalidExtension("".to_string()).into()),
        };
        Ok(Sheets {
               file: file,
               metadata: Metadata::default(),
           })
    }

    /// Get all data from worksheet
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range("Sheet1").expect("Cannot find Sheet1");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_range(&mut self, name: &str) -> Result<Range<DataType>> {
        self.initialize()?;
        inner!(self, read_worksheet_range(name))
    }

    /// Get all formula from worksheet
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_formula("Sheet1").expect("Cannot find Sheet1");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>> {
        self.initialize()?;
        inner!(self, read_worksheet_formula(name))
    }

    /// Get all data from `Worksheet` at index `idx` (0 based)
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range_by_index(0).expect("Cannot find first sheet");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_range_by_index(&mut self, idx: usize) -> Result<Range<DataType>> {
        self.initialize()?;
        let name = self.metadata
            .sheets
            .get(idx)
            .ok_or(ErrorKind::WorksheetIndex(idx))?;
        inner!(self, read_worksheet_range(name))
    }

    fn initialize(&mut self) -> Result<()> {
        if self.metadata.sheets.is_empty() {
            self.metadata = inner!(self, initialize())?;
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
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::open(path).unwrap();
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
    /// use calamine::Sheets;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::open(path).unwrap();
    /// println!("Sheets: {:#?}", workbook.sheet_names());
    /// ```
    pub fn sheet_names(&mut self) -> Result<Vec<String>> {
        self.initialize()?;
        Ok(self.metadata.sheets.clone())
    }

    /// Get all defined names (Ranges names etc)
    pub fn defined_names(&mut self) -> Result<&[(String, String)]> {
        self.initialize()?;
        Ok(&self.metadata.defined_names)
    }
}

/// A trait to share spreadsheets reader functions accross different `FileType`s
trait Reader: Sized {
    /// Creates a new instance based on the actual file
    fn new(f: File) -> Result<Self>;
    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Result<Cow<VbaProject>>;
    /// Initialize
    fn initialize(&mut self) -> Result<Metadata>;
    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>>;
    /// Read worksheet formula in corresponding worksheet path
    fn read_worksheet_formula(&mut self, _: &str) -> Result<Range<String>> {
        Err("Formula reading is not implemented for this extension".into())
    }
}

/// A struct to hold cell position and value
#[derive(Debug, Clone)]
pub struct Cell<T: Default + Clone + PartialEq> {
    /// Position for the cell (row, column)
    pos: (u32, u32),
    /// Value for the cell
    val: T,
}

impl<T: Default + Clone + PartialEq> Cell<T> {
    /// Creates a new `Cell`
    pub fn new(position: (u32, u32), value: T) -> Cell<T> {
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
    pub fn get_value(&self) -> &T {
        &self.val
    }
}

/// A struct which represents a squared selection of cells
#[derive(Debug, Default, Clone)]
pub struct Range<T: Default + Clone + PartialEq> {
    start: (u32, u32),
    end: (u32, u32),
    inner: Vec<T>,
}

impl<T: Default + Clone + PartialEq> Range<T> {
    /// Creates a new `Range`
    ///
    /// When possible, prefer the more efficient `Range::from_sparse`
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range<T> {
        Range {
            start: start,
            end: end,
            inner: vec![T::default(); ((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize],
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

    /// Creates a `Range` from a coo sparse vector of `Cell`s.
    ///
    /// Coordinate list (COO) is the natural way cells are stored
    /// Inner size is defined only by non empty.
    ///
    /// cells: `Vec` of non empty `Cell`s, sorted by row
    ///
    /// # Panics
    ///
    /// panics when a `Cell` row is lower than the first `Cell` row or
    /// bigger than the last `Cell` row.
    fn from_sparse(cells: Vec<Cell<T>>) -> Range<T> {
        if cells.is_empty() {
            Range {
                start: (0, 0),
                end: (0, 0),
                inner: Vec::new(),
            }
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
            let mut v = vec![T::default(); len];
            v.shrink_to_fit();
            for c in cells {
                let idx = ((c.pos.0 - row_start) * width + (c.pos.1 - col_start)) as usize;
                v[idx] = c.val;
            }
            Range {
                start: (row_start, col_start),
                end: (row_end, col_end),
                inner: v,
            }
        }
    }

    /// Set inner value from absolute position
    ///
    /// Will try to resize inner structure if the value is out of bounds.
    /// For relative positions, use Index trait
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
    pub fn set_value(&mut self, absolute_position: (u32, u32), value: T) -> Result<()> {
        if self.start > absolute_position {
            return Err(ErrorKind::CellOutOfRange(absolute_position, self.start).into());
        }

        // check if we need to change range dimension (strangely happens sometimes ...)
        match (self.end.0 < absolute_position.0, self.end.1 < absolute_position.1) {
            (false, false) => (), // regular case, position within bounds
            (true, false) => {
                let len = (absolute_position.0 - self.end.0 + 1) as usize * self.width();
                self.inner.extend_from_slice(&vec![T::default(); len]);
                self.end.0 = absolute_position.0;
            } // missing some rows
            (e, true) => {
                let height = if e {
                    (absolute_position.0 - self.start.0 + 1) as usize
                } else {
                    self.height()
                };
                let width = (absolute_position.1 - self.start.1 + 1) as usize;
                let old_width = self.width();
                let mut data = Vec::with_capacity(width * height);
                let empty = vec![T::default(); width - old_width];
                for sce in self.inner.chunks(old_width) {
                    data.extend_from_slice(sce);
                    data.extend_from_slice(&empty);
                }
                data.extend_from_slice(&vec![T::default(); width * (height - self.height())]);
                if e {
                    self.end = absolute_position
                } else {
                    self.end.1 = absolute_position.1
                }
                self.inner = data;
            } // missing some columns
        }

        let pos = (absolute_position.0 - self.start.0, absolute_position.1 - self.start.1);
        let idx = pos.0 as usize * self.width() + pos.1 as usize;
        self.inner[idx] = value;
        Ok(())
    }

    /// Get cell value from absolute position
    ///
    /// For relative positions, use Index trait
    ///
    /// Panics if indexes are out of range bounds
    pub fn get_value(&self, absolute_position: (u32, u32)) -> &T {
        assert!(absolute_position <= self.end);
        let idx = (absolute_position.0 - self.start.0) as usize * self.width() +
                  (absolute_position.1 - self.start.1) as usize;
        &self.inner[idx]
    }

    /// Get an iterator over inner rows
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let range: Range<DataType> = Range::new((0, 0), (5, 2));
    /// // with rows item row: &[DataType]
    /// assert_eq!(range.rows().map(|r| r.len()).sum::<usize>(), 18);
    /// ```
    pub fn rows(&self) -> Rows<T> {
        if self.inner.is_empty() {
            Rows { inner: None }
        } else {
            let width = self.width();
            Rows { inner: Some(self.inner.chunks(width)) }
        }
    }

    /// Get an iterator over used cells only
    pub fn used_cells(&self) -> UsedCells<T> {
        UsedCells {
            width: self.width(),
            inner: self.inner.iter().enumerate(),
        }
    }
}

impl<T: Default + Clone + PartialEq> Index<usize> for Range<T> {
    type Output = [T];
    fn index(&self, index: usize) -> &[T] {
        let width = self.width();
        &self.inner[index * width..(index + 1) * width]
    }
}

impl<T: Default + Clone + PartialEq> Index<(usize, usize)> for Range<T> {
    type Output = T;
    fn index(&self, index: (usize, usize)) -> &T {
        let width = self.width();
        &self.inner[index.0 * width + index.1]
    }
}

impl<T: Default + Clone + PartialEq> IndexMut<usize> for Range<T> {
    fn index_mut(&mut self, index: usize) -> &mut [T] {
        let width = self.width();
        &mut self.inner[index * width..(index + 1) * width]
    }
}

impl<T: Default + Clone + PartialEq> IndexMut<(usize, usize)> for Range<T> {
    fn index_mut(&mut self, index: (usize, usize)) -> &mut T {
        let width = self.width();
        &mut self.inner[index.0 * width + index.1]
    }
}

/// A struct to iterate over used cells
#[derive(Debug)]
pub struct UsedCells<'a, T: 'a + Default + Clone + PartialEq> {
    width: usize,
    inner: ::std::iter::Enumerate<::std::slice::Iter<'a, T>>,
}

impl<'a, T: 'a + Default + Clone + PartialEq> Iterator for UsedCells<'a, T> {
    type Item = (usize, usize, &'a T);
    fn next(&mut self) -> Option<Self::Item> {
        self.inner
            .by_ref()
            .find(|&(_, v)| v != &T::default())
            .map(|(i, v)| {
                     let row = i / self.width;
                     let col = i % self.width;
                     (row, col, v)
                 })
    }
}

/// An iterator to read `Range` struct row by row
#[derive(Debug)]
pub struct Rows<'a, T: 'a + Default + Clone + PartialEq> {
    inner: Option<::std::slice::Chunks<'a, T>>,
}

impl<'a, T: 'a + Default + Clone + PartialEq> Iterator for Rows<'a, T> {
    type Item = &'a [T];
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.as_mut().and_then(|c| c.next())
    }
}

#[test]
fn test_parse_error() {
    assert_eq!(CellErrorType::from_str("#DIV/0!").unwrap(),
               CellErrorType::Div0);
    assert_eq!(CellErrorType::from_str("#N/A").unwrap(), CellErrorType::NA);
    assert_eq!(CellErrorType::from_str("#NAME?").unwrap(),
               CellErrorType::Name);
    assert_eq!(CellErrorType::from_str("#NULL!").unwrap(),
               CellErrorType::Null);
    assert_eq!(CellErrorType::from_str("#NUM!").unwrap(),
               CellErrorType::Num);
    assert_eq!(CellErrorType::from_str("#REF!").unwrap(),
               CellErrorType::Ref);
    assert_eq!(CellErrorType::from_str("#VALUE!").unwrap(),
               CellErrorType::Value);
}
