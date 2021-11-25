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

#[macro_use]
mod utils;

mod auto;
mod cfb;
mod datatype;
mod ods;
mod xls;
mod xlsb;
mod xlsx;

mod de;
mod errors;
pub mod vba;

use serde::de::DeserializeOwned;
use std::borrow::Cow;
use std::cmp::{max, min};
use std::fmt;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::ops::{Index, IndexMut};
use std::path::Path;

pub use crate::auto::{open_workbook_auto, Sheets};
pub use crate::datatype::DataType;
pub use crate::de::{DeError, RangeDeserializer, RangeDeserializerBuilder, ToCellDeserializer};
pub use crate::errors::Error;
pub use crate::ods::{Ods, OdsError};
pub use crate::xls::{Xls, XlsError, XlsOptions};
pub use crate::xlsb::{Xlsb, XlsbError};
pub use crate::xlsx::{Xlsx, XlsxError};

use crate::vba::VbaProject;

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
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> Result<(), fmt::Error> {
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
/// A trait to share spreadsheets reader functions across different `FileType`s
pub trait Reader: Sized {
    /// Inner reader type
    type RS: Read + Seek;
    /// Error specific to file type
    type Error: std::fmt::Debug + From<std::io::Error>;

    /// Creates a new instance.
    fn new(reader: Self::RS) -> Result<Self, Self::Error>;
    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>>;
    /// Initialize
    fn metadata(&self) -> &Metadata;
    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, Self::Error>>;

    /// Fetch all worksheet data & paths
    fn worksheets(&mut self) -> Vec<(String, Range<DataType>)>;

    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, _: &str) -> Option<Result<Range<String>, Self::Error>>;

    /// Get all sheet names of this workbook, in workbook order
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

    /// Get the nth worksheet. Shortcut for getting the nth
    /// sheet_name, then the corresponding worksheet.
    fn worksheet_range_at(&mut self, n: usize) -> Option<Result<Range<DataType>, Self::Error>> {
        let name = self.sheet_names().get(n)?.to_string();
        self.worksheet_range(&name)
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
    #[inline]
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range<T> {
        assert!(start <= end, "invalid range bounds");
        Range {
            start,
            end,
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

    /// Get column height
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
            let mut col_start = std::u32::MAX;
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
    /// # Remarks
    ///
    /// Will try to resize inner structure if the value is out of bounds.
    /// For relative positions, use Index trait
    ///
    /// Try to avoid this method as much as possible and prefer initializing
    /// the `Range` with `from_sparse` constructor.
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
    /// assert_eq!(range.get_value((2, 1)), Some(&DataType::Empty));
    /// range.set_value((2, 1), DataType::Float(1.0));
    /// assert_eq!(range.get_value((2, 1)), Some(&DataType::Float(1.0)));
    /// ```
    pub fn set_value(&mut self, absolute_position: (u32, u32), value: T) {
        assert!(
            self.start.0 <= absolute_position.0 && self.start.1 <= absolute_position.1,
            "absolute_position out of bounds"
        );

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

    /// Get cell value from **absolute position**.
    ///
    /// If the `absolute_position` is out of range, returns `None`, else returns the cell value.
    /// The coordinate format is (row, column).
    ///
    /// # Warnings
    ///
    /// For relative positions, use Index trait
    ///
    /// # Remarks
    ///
    /// Absolute position is in *sheet* referential while relative position is in *range* referential.
    ///
    /// For instance if we consider range *C2:H38*:
    /// - `(0, 0)` absolute is "A1" and thus this function returns `None`
    /// - `(0, 0)` relative is "C2" and is returned by the `Index` trait (i.e `my_range[(0, 0)]`)
    ///
    /// # Examples
    /// ```
    /// use calamine::{Range, DataType};
    ///
    /// let range: Range<usize> = Range::new((1, 0), (5, 2));
    /// assert_eq!(range.get_value((0, 0)), None);
    /// assert_eq!(range[(0, 0)], 0);
    /// ```
    pub fn get_value(&self, absolute_position: (u32, u32)) -> Option<&T> {
        let p = absolute_position;
        if p.0 >= self.start.0 && p.0 <= self.end.0 && p.1 >= self.start.1 && p.1 <= self.end.1 {
            return self.get((
                (absolute_position.0 - self.start.0) as usize,
                (absolute_position.1 - self.start.1) as usize,
            ));
        }
        None
    }

    /// Get cell value from **relative position**.
    pub fn get(&self, relative_position: (usize, usize)) -> Option<&T> {
        let (row, col) = relative_position;
        self.inner.get(row * self.width() + col)
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
    pub fn rows(&self) -> Rows<'_, T> {
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
    pub fn used_cells(&self) -> UsedCells<'_, T> {
        UsedCells {
            width: self.width(),
            inner: self.inner.iter().enumerate(),
        }
    }

    /// Get an iterator over all cells in this range
    pub fn cells(&self) -> Cells<'_, T> {
        Cells {
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
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let mut sheet = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///     let mut iter = sheet.deserialize()?;
    ///
    ///     if let Some(result) = iter.next() {
    ///         let (label, value): (String, f64) = result?;
    ///         assert_eq!(label, "celsius");
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

    /// Build a new `Range` out of this range
    ///
    /// # Remarks
    ///
    /// Cells within this range will be cloned, cells out of it will be set to Empty
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{Range, DataType};
    /// let mut a = Range::new((1, 1), (3, 3));
    /// a.set_value((1, 1), DataType::Bool(true));
    /// a.set_value((2, 2), DataType::Bool(true));
    ///
    /// let b = a.range((2, 2), (5, 5));
    /// assert_eq!(b.get_value((2, 2)), Some(&DataType::Bool(true)));
    /// assert_eq!(b.get_value((3, 3)), Some(&DataType::Empty));
    ///
    /// let c = a.range((0, 0), (2, 2));
    /// assert_eq!(c.get_value((0, 0)), Some(&DataType::Empty));
    /// assert_eq!(c.get_value((1, 1)), Some(&DataType::Bool(true)));
    /// assert_eq!(c.get_value((2, 2)), Some(&DataType::Bool(true)));
    /// ```
    pub fn range(&self, start: (u32, u32), end: (u32, u32)) -> Range<T> {
        let mut other = Range::new(start, end);
        let (self_start_row, self_start_col) = self.start;
        let (self_end_row, self_end_col) = self.end;
        let (other_start_row, other_start_col) = other.start;
        let (other_end_row, other_end_col) = other.end;

        // copy data from self to other
        let start_row = max(self_start_row, other_start_row);
        let end_row = min(self_end_row, other_end_row);
        let start_col = max(self_start_col, other_start_col);
        let end_col = min(self_end_col, other_end_col);

        if start_row > end_row || start_col > end_col {
            return other;
        }

        let self_width = self.width();
        let other_width = other.width();

        // change referential
        //
        // we want to copy range: start_row..(end_row + 1)
        // In self referential it is (start_row - self_start_row)..(end_row + 1 - self_start_row)
        let self_row_start = (start_row - self_start_row) as usize;
        let self_row_end = (end_row + 1 - self_start_row) as usize;
        let self_col_start = (start_col - self_start_col) as usize;
        let self_col_end = (end_col + 1 - self_start_col) as usize;

        let other_row_start = (start_row - other_start_row) as usize;
        let other_row_end = (end_row + 1 - other_start_row) as usize;
        let other_col_start = (start_col - other_start_col) as usize;
        let other_col_end = (end_col + 1 - other_start_col) as usize;

        {
            let self_rows = self
                .inner
                .chunks(self_width)
                .take(self_row_end)
                .skip(self_row_start);

            let other_rows = other
                .inner
                .chunks_mut(other_width)
                .take(other_row_end)
                .skip(other_row_start);

            for (self_row, other_row) in self_rows.zip(other_rows) {
                let self_cols = &self_row[self_col_start..self_col_end];
                let other_cols = &mut other_row[other_col_start..other_col_end];
                other_cols.clone_from_slice(self_cols);
            }
        }

        other
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
        let (height, width) = self.get_size();
        assert!(index.1 < width && index.0 < height, "index out of bounds");
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
        let (height, width) = self.get_size();
        assert!(index.1 < width && index.0 < height, "index out of bounds");
        &mut self.inner[index.0 * width + index.1]
    }
}

/// A struct to iterate over all cells
#[derive(Debug)]
pub struct Cells<'a, T: CellType> {
    width: usize,
    inner: std::iter::Enumerate<std::slice::Iter<'a, T>>,
}

impl<'a, T: 'a + CellType> Iterator for Cells<'a, T> {
    type Item = (usize, usize, &'a T);
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next().map(|(i, v)| {
            let row = i / self.width;
            let col = i % self.width;
            (row, col, v)
        })
    }
    fn size_hint(&self) -> (usize, Option<usize>) {
        self.inner.size_hint()
    }
}

impl<'a, T: 'a + CellType> DoubleEndedIterator for Cells<'a, T> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner.next_back().map(|(i, v)| {
            let row = i / self.width;
            let col = i % self.width;
            (row, col, v)
        })
    }
}

impl<'a, T: 'a + CellType> ExactSizeIterator for Cells<'a, T> {}

/// A struct to iterate over used cells
#[derive(Debug)]
pub struct UsedCells<'a, T: CellType> {
    width: usize,
    inner: std::iter::Enumerate<std::slice::Iter<'a, T>>,
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
    fn size_hint(&self) -> (usize, Option<usize>) {
        let (_, up) = self.inner.size_hint();
        (0, up)
    }
}

impl<'a, T: 'a + CellType> DoubleEndedIterator for UsedCells<'a, T> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner
            .by_ref()
            .rfind(|&(_, v)| v != &T::default())
            .map(|(i, v)| {
                let row = i / self.width;
                let col = i % self.width;
                (row, col, v)
            })
    }
}

/// An iterator to read `Range` struct row by row
#[derive(Debug)]
pub struct Rows<'a, T: CellType> {
    inner: Option<std::slice::Chunks<'a, T>>,
}

impl<'a, T: 'a + CellType> Iterator for Rows<'a, T> {
    type Item = &'a [T];
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.as_mut().and_then(|c| c.next())
    }
    fn size_hint(&self) -> (usize, Option<usize>) {
        self.inner
            .as_ref()
            .map_or((0, Some(0)), |ch| ch.size_hint())
    }
}

impl<'a, T: 'a + CellType> DoubleEndedIterator for Rows<'a, T> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner.as_mut().and_then(|c| c.next_back())
    }
}

impl<'a, T: 'a + CellType> ExactSizeIterator for Rows<'a, T> {}

/// Struct with the key elements of a table
pub struct Table<T>
where
    T: Default + Clone + PartialEq,
{
    pub(crate) name: String,
    pub(crate) sheet_name: String,
    pub(crate) columns: Vec<String>,
    pub(crate) data: Range<T>,
}
impl<T> Table<T>
where
    T: Default + Clone + PartialEq,
{
    /// Get the name of the table
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Get the name of the sheet that table exists within
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
    /// Get the names of the columns in the order they occur
    pub fn columns(&self) -> &[String] {
        &self.columns
    }
    /// Get a range representing the data from the table (excludes column headers)
    pub fn data(&self) -> &Range<T> {
        &self.data
    }
}
