//! Rust Excel/OpenDocument reader
//!
//! # Status
//!
//! **calamine** is a pure Rust library to read Excel and OpenDocument Spreadsheet files.
//!
//! Read both cell values and vba project.
//!
//! # Examples
//! ```
//! use calamine::{Sheets, DataType};
//! use std::fs::File;
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook = Sheets::<File>::open(path).expect("Cannot open file");
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
//!
//! // Now get all formula!
//! let sheets = workbook.sheet_names().expect("Cannot get sheet names");
//! for s in sheets {
//!     println!("found {} formula in '{}'",
//!              workbook
//!                 .worksheet_formula(&s)
//!                 .expect("error while getting formula")
//!                 .rows().flat_map(|r| r.iter().filter(|f| !f.is_empty()))
//!                 .count(),
//!              s);
//! }
//! ```
#![deny(missing_docs)]
#![recursion_limit = "128"]

extern crate byteorder;
extern crate encoding_rs;
#[macro_use]
extern crate failure;
extern crate quick_xml;
#[macro_use]
extern crate serde;
extern crate zip;

#[macro_use]
extern crate log;

#[macro_use]
mod utils;
mod datatype;
mod xlsb;
mod xlsx;
mod xls;
mod cfb;
mod ods;

mod de;
pub mod errors;
pub mod vba;

use std::borrow::Cow;
use std::fmt;
use std::fs::File;
use std::io::{Read, Seek};
use std::ops::{Index, IndexMut};
use std::path::Path;
use serde::de::DeserializeOwned;

pub use datatype::DataType;
pub use de::{DeError, RangeDeserializer, RangeDeserializerBuilder, ToCellDeserializer};
pub use errors::Error;

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

impl fmt::Display for CellErrorType {
    fn fmt(&self, f: &mut fmt::Formatter) -> Result<(), fmt::Error> {
        match *self {
            CellErrorType::Div0 => write!(f, "#DIV/0!"),
            CellErrorType::NA => write!(f, "#N/A"),
            CellErrorType::Name => write!(f, "#NAME?"),
            CellErrorType::Null => write!(f, "#NULL!"),
            CellErrorType::Num => write!(f, "#NUM!"),
            CellErrorType::Ref => write!(f, "#REF!"),
            CellErrorType::Value => write!(f, "#VALUE!"),
            CellErrorType::GettingData => write!(f, "#DATA!"),
        }
    }
}

/// File types
enum FileType<RS>
where
    RS: Read + Seek,
{
    /// Compound File Binary Format [MS-CFB] (xls, xla)
    Xls(xls::Xls<RS>),
    /// Regular xml zipped file (xlsx, xlsm, xlam)
    Xlsx(xlsx::Xlsx<RS>),
    /// Binary zipped file (xlsb)
    Xlsb(xlsb::Xlsb<RS>),
    /// OpenDocument Spreadsheet Document
    Ods(ods::Ods<RS>),
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
pub struct Sheets<RS>
where
    RS: Read + Seek,
{
    file: FileType<RS>,
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

impl<RS> Sheets<RS>
where
    RS: Read + Seek,
{
    /// Opens a new workbook from a file.
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    /// assert!(Sheets::<File>::open(path).is_ok());
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Sheets<File>, Error> {
        let f: File = File::open(&path)?;
        let file: FileType<File> = match path.as_ref().extension().and_then(|s| s.to_str()) {
            Some("xls") | Some("xla") => FileType::Xls(xls::Xls::new(f)?),
            Some("xlsx") | Some("xlsm") | Some("xlam") => FileType::Xlsx(xlsx::Xlsx::new(f)?),
            Some("xlsb") => FileType::Xlsb(xlsb::Xlsb::new(f)?),
            Some("ods") => FileType::Ods(ods::Ods::new(f)?),
            Some(e) => return Err(Error::InvalidExtension(e.to_string())),
            None => return Err(Error::InvalidExtension("".to_string())),
        };
        Ok(Sheets {
            file: file,
            metadata: Metadata::default(),
        })
    }

    /// Creates a new workbook from a reader.
    pub fn new(reader: RS, extension: &str) -> Result<Sheets<RS>, Error>
    where
        RS: Read + Seek,
    {
        let filetype = match extension {
            "xls" | "xla" => FileType::Xls(xls::Xls::new(reader)?),
            "xlsx" | "xlsm" | "xlam" => FileType::Xlsx(xlsx::Xlsx::new(reader)?),
            "xlsb" => FileType::Xlsb(xlsb::Xlsb::new(reader)?),
            "ods" => FileType::Ods(ods::Ods::new(reader)?),
            _ => return Err(Error::InvalidExtension("".to_string())),
        };
        Ok(Sheets {
            file: filetype,
            metadata: Metadata::default(),
        })
    }

    /// Get all data from worksheet
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::<File>::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range("Sheet1").expect("Cannot find Sheet1");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_range(&mut self, name: &str) -> Result<Range<DataType>, Error> {
        self.initialize()?;
        inner!(self, read_worksheet_range(name))
    }

    /// Get all formula from worksheet
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::<File>::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_formula("Sheet1").expect("Cannot find Sheet1");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, Error> {
        self.initialize()?;
        inner!(self, read_worksheet_formula(name))
    }

    /// Get all data from `Worksheet` at index `idx` (0 based)
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::<File>::open(path).expect("Cannot open file");
    /// let range = workbook.worksheet_range_by_index(0).expect("Cannot find first sheet");
    /// println!("Used range size: {:?}", range.get_size());
    /// ```
    pub fn worksheet_range_by_index(&mut self, idx: usize) -> Result<Range<DataType>, Error> {
        self.initialize()?;
        let name = self.metadata
            .sheets
            .get(idx)
            .ok_or_else(|| Error::WorksheetIndex { idx: idx })?;
        inner!(self, read_worksheet_range(name))
    }

    fn initialize(&mut self) -> Result<(), Error> {
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
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::<File>::open(path).unwrap();
    /// if workbook.has_vba() {
    ///     let vba = workbook.vba_project().expect("Cannot find vba project");
    ///     println!("References: {:?}", vba.get_references());
    ///     println!("Modules: {:?}", vba.get_module_names());
    /// }
    /// ```
    pub fn vba_project(&mut self) -> Result<Cow<VbaProject>, Error> {
        inner!(self, vba_project())
    }

    /// Get all sheet names of this workbook
    ///
    /// # Examples
    /// ```
    /// use calamine::Sheets;
    /// use std::fs::File;
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook = Sheets::<File>::open(path).unwrap();
    /// println!("Sheets: {:#?}", workbook.sheet_names());
    /// ```
    pub fn sheet_names(&mut self) -> Result<Vec<String>, Error> {
        self.initialize()?;
        Ok(self.metadata.sheets.clone())
    }

    /// Get all defined names (Ranges names etc)
    pub fn defined_names(&mut self) -> Result<&[(String, String)], Error> {
        self.initialize()?;
        Ok(&self.metadata.defined_names)
    }
}

// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API this limits
// the kinds of readers (other) data in formats can be read from.
/// A trait to share spreadsheets reader functions accross different `FileType`s
trait Reader<RS>: Sized
where
    RS: Read + Seek,
{
    /// Creates a new instance.
    fn new(reader: RS) -> Result<Self, Error>;
    /// Does the workbook contain a vba project
    fn has_vba(&mut self) -> bool;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Result<Cow<VbaProject>, Error>;
    /// Initialize
    fn initialize(&mut self) -> Result<Metadata, Error>;
    /// Read worksheet data in corresponding worksheet path
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>, Error>;
    /// Read worksheet formula in corresponding worksheet path
    fn read_worksheet_formula(&mut self, _: &str) -> Result<Range<String>, Error>;
}

/// A trait to constrain cells
pub trait CellType: Default + Clone + PartialEq {}
impl<T: Default + Clone + PartialEq> CellType for T {}

/// A struct to hold cell position and value
#[derive(Debug, Clone)]
pub struct Cell<T: CellType> {
    /// Position for the cell (row, column)
    pos: (u32, u32),
    /// Value for the cell
    val: T,
}

impl<T: CellType> Cell<T> {
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
pub struct Range<T: CellType> {
    start: (u32, u32),
    end: (u32, u32),
    inner: Vec<T>,
}

impl<T: CellType> Range<T> {
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

    /// Get size in (height, width) format
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
    pub fn set_value(&mut self, absolute_position: (u32, u32), value: T) -> Result<(), Error> {
        if self.start > absolute_position {
            return Err(Error::CellOutOfRange {
                try_pos: absolute_position,
                min_pos: self.start,
            });
        }

        // check if we need to change range dimension (strangely happens sometimes ...)
        match (
            self.end.0 < absolute_position.0,
            self.end.1 < absolute_position.1,
        ) {
            (false, false) => (), // regular case, position within bounds
            (true, false) => {
                let len = (absolute_position.0 - self.end.0 + 1) as usize * self.width();
                self.inner.extend_from_slice(&vec![T::default(); len]);
                self.end.0 = absolute_position.0;
            }
            // missing some rows
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

        let pos = (
            absolute_position.0 - self.start.0,
            absolute_position.1 - self.start.1,
        );
        let idx = pos.0 as usize * self.width() + pos.1 as usize;
        self.inner[idx] = value;
        Ok(())
    }

    /// Get cell value from absolute position
    ///
    /// The coordinate format is (row, column). For relative positions, use Index trait
    ///
    /// Panics if indexes are out of range bounds
    pub fn get_value(&self, absolute_position: (u32, u32)) -> &T {
        assert!(absolute_position <= self.end);
        let idx = (absolute_position.0 - self.start.0) as usize * self.width()
            + (absolute_position.1 - self.start.1) as usize;
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
            Rows {
                inner: Some(self.inner.chunks(width)),
            }
        }
    }

    /// Get an iterator over used cells only
    pub fn used_cells(&self) -> UsedCells<T> {
        UsedCells {
            width: self.width(),
            inner: self.inner.iter().enumerate(),
        }
    }

    /// Build a `RangeDeserializer` from this configuration.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{Sheets, RangeDeserializerBuilder};
    /// # use calamine::errors::Error;
    /// # use std::fs::File;
    /// # fn main() { example().unwrap(); }
    /// fn example() -> Result<(), Error> {
    ///     let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook = Sheets::<File>::open(path)?;
    ///     let mut sheet = workbook.worksheet_range("Sheet1")?;
    ///     let mut iter = sheet.deserialize()?;
    ///
    ///     if let Some(result) = iter.next() {
    ///         let (label, value): (String, f64) = result?;
    ///         assert_eq!(label, "celcius");
    ///         assert_eq!(value, 22.2222);
    ///
    ///         Ok(())
    ///     } else {
    ///         return Err(From::from("expected at least one record but got none"));
    ///     }
    /// }
    /// ```
    pub fn deserialize<'a, D>(&'a self) -> Result<RangeDeserializer<'a, T, D>, DeError>
    where
        T: ToCellDeserializer<'a>,
        D: DeserializeOwned,
    {
        RangeDeserializerBuilder::new().from_range(self)
    }
}

impl<T: CellType> Index<usize> for Range<T> {
    type Output = [T];
    fn index(&self, index: usize) -> &[T] {
        let width = self.width();
        &self.inner[index * width..(index + 1) * width]
    }
}

impl<T: CellType> Index<(usize, usize)> for Range<T> {
    type Output = T;
    fn index(&self, index: (usize, usize)) -> &T {
        let width = self.width();
        &self.inner[index.0 * width + index.1]
    }
}

impl<T: CellType> IndexMut<usize> for Range<T> {
    fn index_mut(&mut self, index: usize) -> &mut [T] {
        let width = self.width();
        &mut self.inner[index * width..(index + 1) * width]
    }
}

impl<T: CellType> IndexMut<(usize, usize)> for Range<T> {
    fn index_mut(&mut self, index: (usize, usize)) -> &mut T {
        let width = self.width();
        &mut self.inner[index.0 * width + index.1]
    }
}

/// A struct to iterate over used cells
#[derive(Debug)]
pub struct UsedCells<'a, T: 'a + CellType> {
    width: usize,
    inner: ::std::iter::Enumerate<::std::slice::Iter<'a, T>>,
}

impl<'a, T: 'a + CellType> Iterator for UsedCells<'a, T> {
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
pub struct Rows<'a, T: 'a + CellType> {
    inner: Option<::std::slice::Chunks<'a, T>>,
}

impl<'a, T: 'a + CellType> Iterator for Rows<'a, T> {
    type Item = &'a [T];
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.as_mut().and_then(|c| c.next())
    }
}
