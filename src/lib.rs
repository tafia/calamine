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
//! use calamine::{Reader, open_workbook, Xlsx, DataType};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
//!
//! // Read whole worksheet data and provide some statistics
//! if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
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
//! if let Some(Ok(mut vba)) = workbook.vba_project() {
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
//! for name in workbook.defined_names() {
//!     println!("name: {}, formula: {}", name.0, name.1);
//! }
//!
//! // Now get all formula!
//! let sheets = workbook.sheet_names().to_owned();
//! for s in sheets {
//!     println!("found {} formula in '{}'",
//!              workbook
//!                 .worksheet_formula(&s)
//!                 .expect("sheet not found")
//!                 .expect("error while getting formula")
//!                 .rows().flat_map(|r| r.iter().filter(|f| !f.is_empty()))
//!                 .count(),
//!              s);
//! }
//! ```
#![deny(missing_docs)]

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
mod auto;

mod de;
mod errors;
pub mod vba;

use std::borrow::Cow;
use std::fmt;
use std::io::{BufReader, Read, Seek};
use std::ops::{Index, IndexMut};
use std::fs::File;
use std::path::Path;
use serde::de::DeserializeOwned;

pub use datatype::DataType;
pub use de::{DeError, RangeDeserializer, RangeDeserializerBuilder, ToCellDeserializer};
pub use xls::{Xls, XlsError};
pub use xlsx::{Xlsx, XlsxError};
pub use xlsb::{Xlsb, XlsbError};
pub use ods::{Ods, OdsError};
pub use auto::{open_workbook_auto, Sheets};
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

/// Common file metadata
///
/// Depending on file type, some extra information may be stored
/// in the Reader implementations
#[derive(Debug, Default)]
pub struct Metadata {
    sheets: Vec<String>,
    /// Map of sheet names/sheet path within zip archive
    names: Vec<(String, String)>,
}

// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API this limits
// the kinds of readers (other) data in formats can be read from.
/// A trait to share spreadsheets reader functions accross different `FileType`s
pub trait Reader: Sized {
    /// Inner reader type
    type RS: Read + Seek;
    /// Error specific to file type
    type Error: ::std::fmt::Debug + From<::std::io::Error>;

    /// Creates a new instance.
    fn new(reader: Self::RS) -> Result<Self, Self::Error>;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<VbaProject>, Self::Error>>;
    /// Initialize
    fn metadata(&self) -> &Metadata;
    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, Self::Error>>;
    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, _: &str) -> Option<Result<Range<String>, Self::Error>>;

    /// Get all sheet names of this workbook
    ///
    /// # Examples
    /// ```
    /// use calamine::{Xlsx, open_workbook, Reader};
    ///
    /// # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut workbook: Xlsx<_> = open_workbook(path).unwrap();
    /// println!("Sheets: {:#?}", workbook.sheet_names());
    /// ```
    fn sheet_names(&self) -> &[String] {
        &self.metadata().sheets
    }

    /// Get all defined names (Ranges names etc)
    fn defined_names(&self) -> &[(String, String)] {
        &self.metadata().names
    }
}

/// Convenient function to open a file with a BufReader<File>
pub fn open_workbook<R, P>(path: P) -> Result<R, R::Error>
where
    R: Reader<RS = BufReader<File>>,
    P: AsRef<Path>,
{
    let file = BufReader::new(File::open(path)?);
    R::new(file)
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
    /// Creates a new non-empty `Range`
    ///
    /// When possible, prefer the more efficient `Range::from_sparse`
    ///
    /// # Panics
    ///
    /// Panics if start.0 > end.0 or start.1 > end.1
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range<T> {
        assert!(start <= end, "invalid range bounds");
        Range {
            start: start,
            end: end,
            inner: vec![T::default(); ((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize],
        }
    }

    /// Creates a new empty range
    #[inline]
    pub fn empty() -> Range<T> {
        Range {
            start: (0, 0),
            end: (0, 0),
            inner: Vec::new(),
        }
    }

    /// Get top left cell position (row, column)
    #[inline]
    pub fn start(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.start)
        }
    }

    /// Get bottom right cell position (row, column)
    #[inline]
    pub fn end(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.end)
        }
    }

    /// Get column width
    #[inline]
    pub fn width(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.1 - self.start.1 + 1) as usize
        }
    }

    /// Get column width
    #[inline]
    pub fn height(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.0 - self.start.0 + 1) as usize
        }
    }

    /// Get size in (height, width) format
    #[inline]
    pub fn get_size(&self) -> (usize, usize) {
        (self.height(), self.width())
    }

    /// Is range empty
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
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
    pub fn from_sparse(cells: Vec<Cell<T>>) -> Range<T> {
        if cells.is_empty() {
            Range::empty()
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
    /// # Panics
    ///
    /// If absolute_position > Cell start
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    /// assert_eq!(range.get_value((2, 1)), &DataType::Empty);
    /// range.set_value((2, 1), DataType::Float(1.0));
    /// assert_eq!(range.get_value((2, 1)), &DataType::Float(1.0));
    /// ```
    pub fn set_value(&mut self, absolute_position: (u32, u32), value: T) {
        assert!(self.start <= absolute_position);

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
    }

    /// Get cell value from absolute position
    ///
    /// The coordinate format is (row, column). For relative positions, use Index trait
    ///
    /// Panics if indexes are out of range bounds
    pub fn get_value(&self, absolute_position: (u32, u32)) -> &T {
        assert!(
            absolute_position <= self.end,
            "absolute_position out of range boundary"
        );
        assert!(
            absolute_position >= self.start,
            "absolute_position out of range boundary"
        );
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
    /// # use calamine::{Reader, Error, open_workbook, Xlsx, RangeDeserializerBuilder};
    /// # fn main() { example().unwrap(); }
    /// fn example() -> Result<(), Error> {
    ///     let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let mut sheet = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
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
