//! Rust Excel/`OpenDocument` reader
//!
//! # Status
//!
//! **calamine** is a pure Rust library to read Excel and `OpenDocument` Spreadsheet files.
//!
//! Read both cell values and vba project.
//!
//! # Examples
//! ```
//! use calamine::{Reader, open_workbook, Xlsx, Data};
//!
//! // opens a new workbook
//! # let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
//! let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");
//!
//! // Read whole worksheet data and provide some statistics
//! if let Ok(range) = workbook.worksheet_range("Sheet1") {
//!     let total_cells = range.get_size().0 * range.get_size().1;
//!     let non_empty_cells: usize = range.used_cells().count();
//!     println!("Found {} cells in 'Sheet1', including {} non empty cells",
//!              total_cells, non_empty_cells);
//!     // alternatively, we can manually filter rows
//!     assert_eq!(non_empty_cells, range.rows()
//!         .flat_map(|r| r.iter().filter(|&c| c != &Data::Empty)).count());
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
mod formats;
mod ods;
mod xls;
mod xlsb;
mod xlsx;

mod de;
mod errors;
pub mod vba;

use serde::de::{Deserialize, DeserializeOwned, Deserializer};
use std::borrow::Cow;
use std::cmp::{max, min};
use std::fmt;
use std::fs::File;
use std::io::{BufReader, Read, Seek};
use std::ops::{Index, IndexMut};
use std::path::Path;

pub use crate::auto::{open_workbook_auto, open_workbook_auto_from_rs, Sheets};
pub use crate::datatype::{Data, DataRef, DataType, ExcelDateTime, ExcelDateTimeType};
pub use crate::de::{DeError, RangeDeserializer, RangeDeserializerBuilder, ToCellDeserializer};
pub use crate::errors::Error;
pub use crate::formats::{
    builtin_format_by_code, builtin_format_by_id, detect_custom_number_format,
    detect_custom_number_format_with_interner, Alignment, Border, BorderSide, CellFormat,
    CellStyle, Color, Fill, Font, FormatStringInterner, PatternType,
};
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

/// A struct that combines cell data with its formatting information
#[derive(Debug, Clone)]
pub struct DataWithFormatting {
    /// The cell data value
    pub data: Data,
    /// The cell formatting information
    pub formatting: Option<CellStyle>,
}

impl DataWithFormatting {
    /// Creates a new DataWithFormatting with the given data and formatting
    pub fn new(data: Data, formatting: Option<CellStyle>) -> Self {
        Self { data, formatting }
    }

    /// Creates a new DataWithFormatting with data and no formatting
    pub fn from_data(data: Data) -> Self {
        Self {
            data,
            formatting: None,
        }
    }

    /// Gets the data value
    pub fn get_data(&self) -> &Data {
        &self.data
    }

    /// Gets the formatting information
    pub fn get_formatting(&self) -> &Option<CellStyle> {
        &self.formatting
    }

    /// Checks if the underlying data is empty
    pub fn is_empty(&self) -> bool {
        matches!(self.data, Data::Empty)
    }

    /// Gets the data as a string slice if it's string data
    pub fn as_str(&self) -> &str {
        match &self.data {
            Data::String(s) => s,
            _ => "",
        }
    }
}

impl Default for DataWithFormatting {
    fn default() -> Self {
        Self {
            data: Data::Empty,
            formatting: None,
        }
    }
}

impl PartialEq<DataWithFormatting> for DataWithFormatting {
    fn eq(&self, other: &DataWithFormatting) -> bool {
        // For the purpose of range operations, cells with empty data are considered equal
        // regardless of formatting. This preserves the behavior where empty cells are
        // treated consistently even when they have formatting information.
        if matches!(self.data, Data::Empty) && matches!(other.data, Data::Empty) {
            return true;
        }
        self.data == other.data && self.formatting == other.formatting
    }
}

impl CellType for DataWithFormatting {}

impl fmt::Display for DataWithFormatting {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.data)
    }
}

impl PartialEq<Data> for DataWithFormatting {
    fn eq(&self, other: &Data) -> bool {
        self.data == *other
    }
}

impl PartialEq<DataWithFormatting> for Data {
    fn eq(&self, other: &DataWithFormatting) -> bool {
        *self == other.data
    }
}

impl PartialEq<String> for DataWithFormatting {
    fn eq(&self, other: &String) -> bool {
        self.data.to_string() == *other
    }
}

impl PartialEq<DataWithFormatting> for String {
    fn eq(&self, other: &DataWithFormatting) -> bool {
        *self == other.data.to_string()
    }
}

impl PartialEq<f64> for DataWithFormatting {
    fn eq(&self, other: &f64) -> bool {
        match &self.data {
            Data::Float(f) => f == other,
            Data::Int(i) => *i as f64 == *other,
            _ => false,
        }
    }
}

impl PartialEq<DataWithFormatting> for f64 {
    fn eq(&self, other: &DataWithFormatting) -> bool {
        other.eq(self)
    }
}

impl PartialEq<&str> for DataWithFormatting {
    fn eq(&self, other: &&str) -> bool {
        self.data.to_string() == **other
    }
}

impl PartialEq<DataWithFormatting> for &str {
    fn eq(&self, other: &DataWithFormatting) -> bool {
        **self == other.data.to_string()
    }
}

/// Dimensions info
#[derive(Debug, Default, PartialEq, Eq, Hash, Ord, PartialOrd, Copy, Clone)]
pub struct Dimensions {
    /// start: (row, col)
    pub start: (u32, u32),
    /// end: (row, col)
    pub end: (u32, u32),
}

#[allow(clippy::len_without_is_empty)]
impl Dimensions {
    /// create dimensions info with start position and end position
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Self {
        Self { start, end }
    }
    /// check if a position is in it
    pub fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.start.0 && row <= self.end.0 && col >= self.start.1 && col <= self.end.1
    }
    /// len
    pub fn len(&self) -> u64 {
        (self.end.0 - self.start.0 + 1) as u64 * (self.end.1 - self.start.1 + 1) as u64
    }
}

/// Common file metadata
///
/// Depending on file type, some extra information may be stored
/// in the Reader implementations
#[derive(Debug, Default)]
pub struct Metadata {
    sheets: Vec<Sheet>,
    /// Map of sheet names/sheet path within zip archive
    names: Vec<(String, String)>,
}

/// Type of sheet.
///
/// Only Excel formats support this. Default value for ODS is
/// `SheetType::WorkSheet`.
///
/// The property is defined in the following specifications:
///
/// - [ECMA-376 Part 1] 12.3.2, 12.3.7 and 12.3.24.
/// - [MS-XLS `BoundSheet`].
/// - [MS-XLSB `ST_SheetType`].
///
/// [ECMA-376 Part 1]: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
/// [MS-XLS `BoundSheet`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b9ec509a-235d-424e-871d-f8e721106501
/// [MS-XLS `BrtBundleSh`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/1edadf56-b5cd-4109-abe7-76651bbe2722
///
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SheetType {
    /// A worksheet.
    WorkSheet,
    /// A dialog sheet.
    DialogSheet,
    /// A macro sheet.
    MacroSheet,
    /// A chartsheet.
    ChartSheet,
    /// A VBA module.
    Vba,
}

/// Type of visible sheet.
///
/// The property is defined in the following specifications:
///
/// - [ECMA-376 Part 1] 18.18.68 `ST_SheetState` (Sheet Visibility Types).
/// - [MS-XLS `BoundSheet`].
/// - [MS-XLSB `ST_SheetState`].
/// - [OpenDocument v1.2] 19.471 `style:display`.
///
/// [ECMA-376 Part 1]: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
/// [OpenDocument v1.2]: https://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part1.html#property-table_display
/// [MS-XLS `BoundSheet`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b9ec509a-235d-424e-871d-f8e721106501
/// [MS-XLSB `ST_SheetState`]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/74cb1d22-b931-4bf8-997d-17517e2416e9
///
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SheetVisible {
    /// Visible
    Visible,
    /// Hidden
    Hidden,
    /// The sheet is hidden and cannot be displayed using the user interface. It is supported only by Excel formats.
    VeryHidden,
}

/// Metadata of sheet
#[derive(Debug, Clone, PartialEq)]
pub struct Sheet {
    /// Name
    pub name: String,
    /// Type
    /// Only Excel formats support this. Default value for ODS is `SheetType::WorkSheet`.
    pub typ: SheetType,
    /// Visible
    pub visible: SheetVisible,
}

/// Row to use as header
/// By default, the first non-empty row is used as header
#[derive(Debug, Default, Clone, Copy)]
#[non_exhaustive]
pub enum HeaderRow {
    /// First non-empty row
    #[default]
    FirstNonEmptyRow,
    /// Index of the header row
    Row(u32),
}

// FIXME `Reader` must only be seek `Seek` for `Xls::xls`. Because of the present API this limits
// the kinds of readers (other) data in formats can be read from.
/// A trait to share spreadsheets reader functions across different `FileType`s
pub trait Reader<RS>: Sized
where
    RS: Read + Seek,
{
    /// Error specific to file type
    type Error: std::fmt::Debug + From<std::io::Error>;

    /// Creates a new instance.
    fn new(reader: RS) -> Result<Self, Self::Error>;

    /// Set header row (i.e. first row to be read)
    /// If `header_row` is `None`, the first non-empty row will be used as header row
    fn with_header_row(&mut self, header_row: HeaderRow) -> &mut Self;

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>>;

    /// Initialize
    fn metadata(&self) -> &Metadata;

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Result<Range<DataWithFormatting>, Self::Error>;

    /// Fetch all worksheet data & paths
    fn worksheets(&mut self) -> Vec<(String, Range<DataWithFormatting>)>;

    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, _: &str) -> Result<Range<DataWithFormatting>, Self::Error>;



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
    fn sheet_names(&self) -> Vec<String> {
        self.metadata()
            .sheets
            .iter()
            .map(|s| s.name.to_owned())
            .collect()
    }

    /// Fetch all sheets metadata
    fn sheets_metadata(&self) -> &[Sheet] {
        &self.metadata().sheets
    }

    /// Get all defined names (Ranges names etc)
    fn defined_names(&self) -> &[(String, String)] {
        &self.metadata().names
    }

    /// Get the nth worksheet. Shortcut for getting the nth
    /// worksheet name, then the corresponding worksheet.
    fn worksheet_range_at(&mut self, n: usize) -> Option<Result<Range<DataWithFormatting>, Self::Error>> {
        let name = self.sheet_names().get(n)?.to_string();
        Some(self.worksheet_range(&name))
    }

    /// Get all pictures, tuple as (ext: String, data: Vec<u8>)
    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>>;
}

/// A trait to share spreadsheets reader functions across different `FileType`s
pub trait ReaderRef<RS>: Reader<RS>
where
    RS: Read + Seek,
{
    /// Get worksheet range where shared string values are only borrowed.
    ///
    /// This is implemented only for [`calamine::Xlsx`](crate::Xlsx) and [`calamine::Xlsb`](crate::Xlsb), as Xls and Ods formats
    /// do not support lazy iteration.
    fn worksheet_range_ref<'a>(&'a mut self, name: &str)
        -> Result<Range<DataRef<'a>>, Self::Error>;

    /// Get the nth worksheet range where shared string values are only borrowed. Shortcut for getting the nth
    /// worksheet name, then the corresponding worksheet.
    ///
    /// This is implemented only for [`calamine::Xlsx`](crate::Xlsx) and [`calamine::Xlsb`](crate::Xlsb), as Xls and Ods formats
    /// do not support lazy iteration.
    fn worksheet_range_at_ref(
        &mut self,
        n: usize,
    ) -> Option<Result<Range<DataRef<'_>>, Self::Error>> {
        let name = self.sheet_names().get(n)?.to_string();
        Some(self.worksheet_range_ref(&name))
    }
}

/// Convenient function to open a file with a `BufReader<File>`.
pub fn open_workbook<R, P>(path: P) -> Result<R, R::Error>
where
    R: Reader<BufReader<File>>,
    P: AsRef<Path>,
{
    let file = BufReader::new(File::open(path)?);
    R::new(file)
}

/// Convenient function to open a file with a `BufReader<File>`.
pub fn open_workbook_from_rs<R, RS>(rs: RS) -> Result<R, R::Error>
where
    RS: Read + Seek,
    R: Reader<RS>,
{
    R::new(rs)
}

/// A trait to constrain cells
pub trait CellType: Default + Clone + PartialEq {}

impl CellType for Data {}
impl<'a> CellType for DataRef<'a> {}
impl CellType for String {}
impl CellType for usize {} // for tests

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

/// A struct which represents an area of cells and the data within it.
///
/// Ranges are used by `calamine` to represent an area of data in a worksheet. A
/// `Range` is a rectangular area of cells defined by its start and end
/// positions.
///
/// A `Range` is constructed with **absolute positions** in the form of `(row,
/// column)`. The start position for the absolute positioning is the cell `(0,
/// 0)` or `A1`. For the example range "B3:C6", shown below, the start position
/// is `(2, 1)` and the end position is `(5, 2)`. Within the range, the cells
/// are indexed with **relative positions** where `(0, 0)` is the start cell. In
/// the example below the relative positions for the start and end cells are
/// `(0, 0)` and `(3, 1)` respectively.
///
/// ```text
///  ______________________________________________________________________________
/// |         ||                |                |                |                |
/// |         ||       A        |       B        |       C        |       D        |
/// |_________||________________|________________|________________|________________|
/// |    1    ||                |                |                |                |
/// |_________||________________|________________|________________|________________|
/// |    2    ||                |                |                |                |
/// |_________||________________|________________|________________|________________|
/// |    3    ||                | (2, 1), (0, 0) |                |                |
/// |_________||________________|________________|________________|________________|
/// |    4    ||                |                |                |                |
/// |_________||________________|________________|________________|________________|
/// |    5    ||                |                |                |                |
/// |_________||________________|________________|________________|________________|
/// |    6    ||                |                | (5,2), (3, 1)  |                |
/// |_________||________________|________________|________________|________________|
/// |    7    ||                |                |                |                |
/// |_________||________________|________________|________________|________________|
/// |_          ___________________________________________________________________|
///   \ Sheet1 /
///     ------
/// ```
///
/// A `Range` contains a vector of cells of of generic type `T` which implement
/// the [`CellType`] trait. The values are stored in a row-major order.
///
#[derive(Debug, Default, Clone)]
pub struct Range<T> {
    start: (u32, u32),
    end: (u32, u32),
    inner: Vec<T>,
}

impl<T: CellType> Range<T> {
    /// Creates a new `Range` with default values.
    ///
    /// Create a new [`Range`] with the given start and end positions. The
    /// positions are in worksheet absolute coordinates, i.e. `(0, 0)` is cell `A1`.
    ///
    /// The range is populated with default values of type `T`.
    ///
    /// When possible, use the more efficient [`Range::from_sparse()`]
    /// constructor.
    ///
    /// # Parameters
    ///
    /// - `start`: The zero indexed (row, column) tuple.
    /// - `end`: The zero indexed (row, column) tuple.
    ///
    /// # Panics
    ///
    /// Panics if `start` > `end`.
    ///
    ///
    /// # Examples
    ///
    /// An example of creating a new calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_new.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// // Create a 8x1 Range.
    /// let range: Range<Data> = Range::new((2, 2), (9, 2));
    ///
    /// assert_eq!(range.width(), 1);
    /// assert_eq!(range.height(), 8);
    /// assert_eq!(range.cells().count(), 8);
    /// assert_eq!(range.used_cells().count(), 0);
    /// ```
    ///
    ///
    #[inline]
    pub fn new(start: (u32, u32), end: (u32, u32)) -> Range<T> {
        assert!(start <= end, "invalid range bounds");
        Range {
            start,
            end,
            inner: vec![T::default(); ((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize],
        }
    }

    /// Creates a new empty `Range`.
    ///
    /// Creates a new [`Range`] with start and end positions both set to `(0,
    /// 0)` and with an empty inner vector. An empty range can be expanded by
    /// adding data.
    ///
    /// # Examples
    ///
    /// An example of creating a new empty calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_empty.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::empty();
    ///
    /// assert!(range.is_empty());
    /// ```
    ///
    #[inline]
    pub fn empty() -> Range<T> {
        Range {
            start: (0, 0),
            end: (0, 0),
            inner: Vec::new(),
        }
    }

    /// Get top left cell position of a `Range`.
    ///
    /// Get the top left cell position of a range in absolute `(row, column)`
    /// coordinates.
    ///
    /// Returns `None` if the range is empty.
    ///
    /// # Examples
    ///
    /// An example of getting the start position of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_start.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::new((2, 3), (9, 3));
    ///
    /// assert_eq!(range.start(), Some((2, 3)));
    /// ```
    ///
    #[inline]
    pub fn start(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.start)
        }
    }

    /// Get bottom right cell position of a `Range`.
    ///
    /// Get the bottom right cell position of a range in absolute `(row,
    /// column)` coordinates.
    ///
    /// Returns `None` if the range is empty.
    ///
    /// # Examples
    ///
    /// An example of getting the end position of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_end.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::new((2, 3), (9, 3));
    ///
    /// assert_eq!(range.end(), Some((9, 3)));
    /// ```
    ///
    #[inline]
    pub fn end(&self) -> Option<(u32, u32)> {
        if self.is_empty() {
            None
        } else {
            Some(self.end)
        }
    }

    /// Get the column width of a `Range`.
    ///
    /// The width is defined as the number of columns between the start and end
    /// positions.
    ///
    /// # Examples
    ///
    /// An example of getting the column width of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_width.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::new((2, 3), (9, 3));
    ///
    /// assert_eq!(range.width(), 1);
    /// ```
    ///
    #[inline]
    pub fn width(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.1 - self.start.1 + 1) as usize
        }
    }

    /// Get the row height of a `Range`.
    ///
    /// The height is defined as the number of rows between the start and end
    /// positions.
    ///
    /// # Examples
    ///
    /// An example of getting the row height of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_height.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::new((2, 3), (9, 3));
    ///
    /// assert_eq!(range.height(), 8);
    /// ```
    ///
    #[inline]
    pub fn height(&self) -> usize {
        if self.is_empty() {
            0
        } else {
            (self.end.0 - self.start.0 + 1) as usize
        }
    }

    /// Get size of a `Range` in (height, width) format.
    ///
    /// # Examples
    ///
    /// An example of getting the (height, width) size of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_size.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::new((2, 3), (9, 3));
    ///
    /// assert_eq!(range.get_size(), (8, 1));
    /// ```
    ///
    #[inline]
    pub fn get_size(&self) -> (usize, usize) {
        (self.height(), self.width())
    }

    /// Check if a `Range` is empty.
    ///
    /// # Examples
    ///
    /// An example of checking if a calamine `Range` is empty.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_empty.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range: Range<Data> = Range::empty();
    ///
    /// assert!(range.is_empty());
    /// ```
    ///
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Creates a `Range` from a sparse vector of cells.
    ///
    /// The `Range::from_sparse()` constructor can be used to create a Range
    /// from a vector of [`Cell`] data. This is slightly more efficient than
    /// creating a range with [`Range::new()`] and then setting the values.
    ///
    /// # Parameters
    ///
    /// - `cells`: A vector of non-empty [`Cell`] elements, sorted by row. The
    ///   first and last cells define the start and end positions of the range.
    ///
    /// # Panics
    ///
    /// Panics when a `Cell` row is less than the first `Cell` row.
    ///
    /// # Examples
    ///
    /// An example of creating a new calamine `Range` for a sparse vector of
    /// Cells.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_from_sparse.rs
    /// #
    /// use calamine::{Cell, Data, Range};
    ///
    /// let cells = vec![
    ///     Cell::new((2, 2), Data::Int(1)),
    ///     Cell::new((5, 2), Data::Int(1)),
    ///     Cell::new((9, 2), Data::Int(1)),
    /// ];
    ///
    /// let range = Range::from_sparse(cells);
    ///
    /// assert_eq!(range.width(), 1);
    /// assert_eq!(range.height(), 8);
    /// assert_eq!(range.cells().count(), 8);
    /// assert_eq!(range.used_cells().count(), 3);
    /// ```
    ///
    pub fn from_sparse(cells: Vec<Cell<T>>) -> Range<T> {
        let (row_start, row_end) = match &cells[..] {
            [] => return Range::empty(),
            [first] => (first.pos.0, first.pos.0),
            [first, .., last] => (first.pos.0, last.pos.0),
        };
        // search bounds
        let mut col_start = u32::MAX;
        let mut col_end = 0;
        for c in cells.iter().map(|c| c.pos.1) {
            col_start = min(c, col_start);
            col_end = max(c, col_end);
        }
        let cols = (col_end - col_start + 1) as usize;
        let rows = (row_end - row_start + 1) as usize;
        let len = cols.saturating_mul(rows);
        let mut v = vec![T::default(); len];
        v.shrink_to_fit();
        for c in cells {
            let row = (c.pos.0 - row_start) as usize;
            let col = (c.pos.1 - col_start) as usize;
            let idx = row.saturating_mul(cols) + col;
            if let Some(v) = v.get_mut(idx) {
                *v = c.val;
            }
        }
        Range {
            start: (row_start, col_start),
            end: (row_end, col_end),
            inner: v,
        }
    }

    /// Set a value at an absolute position in a `Range`.
    ///
    /// This method sets a value in the range at the given absolute position
    /// (relative to `A1`).
    ///
    /// Try to avoid this method as much as possible and prefer initializing the
    /// `Range` with the [`Range::from_sparse()`] constructor.
    ///
    /// # Parameters
    ///
    /// - `absolute_position`: The absolute position, relative to `A1`, in the
    ///   form of `(row, column)`. It must be greater than or equal to the start
    ///   position of the range. If the position is greater than the end of the range
    ///   the structure will be resized to accommodate the new end position.
    ///
    /// # Panics
    ///
    /// If `absolute_position.0 < self.start.0 || absolute_position.1 < self.start.1`
    ///
    /// # Examples
    ///
    /// An example of setting a value in a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_set_value.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let mut range = Range::new((0, 0), (5, 2));
    ///
    /// // The initial range is empty.
    /// assert_eq!(range.get_value((2, 1)), Some(&Data::Empty));
    ///
    /// // Set a value at a specific position.
    /// range.set_value((2, 1), Data::Float(1.0));
    ///
    /// // The value at the specified position should now be set.
    /// assert_eq!(range.get_value((2, 1)), Some(&Data::Float(1.0)));
    /// ```
    ///
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
                    self.end = absolute_position;
                } else {
                    self.end.1 = absolute_position.1;
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

    /// Get a value at an absolute position in a `Range`.
    ///
    /// If the `absolute_position` is out of range, returns `None`, otherwise
    /// returns the cell value. The coordinate format is `(row, column)`
    /// relative to `A1`.
    ///
    /// For relative positions see the [`Range::get()`] method.
    ///
    /// # Parameters
    ///
    /// - `absolute_position`: The absolute position, relative to `A1`, in the
    ///   form of `(row, column)`.
    ///
    /// # Examples
    ///
    /// An example of getting a value in a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_get_value.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let range = Range::new((1, 1), (5, 5));
    ///
    /// // Get the value for a cell in the range.
    /// assert_eq!(range.get_value((2, 2)), Some(&Data::Empty));
    ///
    /// // Get the value for a cell outside the range.
    /// assert_eq!(range.get_value((0, 0)), None);
    /// ```
    ///
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

    /// Get a value at a relative position in a `Range`.
    ///
    /// If the `relative_position` is out of range, returns `None`, otherwise
    /// returns the cell value. The coordinate format is `(row, column)`
    /// relative to `(0, 0)` in the range.
    ///
    /// For absolute cell positioning see the [`Range::get_value()`] method.
    ///
    /// # Parameters
    ///
    /// - `relative_position`: The position relative to the index `(0, 0)` in
    ///   the range.
    ///
    /// # Examples
    ///
    /// An example of getting a value in a calamine `Range`, using relative
    /// positioning.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_get.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// let mut range = Range::new((1, 1), (5, 5));
    ///
    /// // Set a cell value using the cell absolute position.
    /// range.set_value((2, 3), Data::Int(123));
    ///
    /// // Get the value using the range relative position.
    /// assert_eq!(range.get((1, 2)), Some(&Data::Int(123)));
    /// ```
    ///
    pub fn get(&self, relative_position: (usize, usize)) -> Option<&T> {
        let (row, col) = relative_position;
        let (height, width) = self.get_size();
        if col >= width || row >= height {
            None
        } else {
            self.inner.get(row * width + col)
        }
    }

    /// Get an iterator over the rows of a `Range`.
    ///
    /// # Examples
    ///
    /// An example of using a `Row` iterator with a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_rows.rs
    /// #
    /// use calamine::{Cell, Data, Range};
    ///
    /// let cells = vec![
    ///     Cell::new((1, 1), Data::Int(1)),
    ///     Cell::new((1, 2), Data::Int(2)),
    ///     Cell::new((3, 1), Data::Int(3)),
    /// ];
    ///
    /// // Create a Range from the cells.
    /// let range = Range::from_sparse(cells);
    ///
    /// // Iterate over the rows of the range.
    /// for (row_num, row) in range.rows().enumerate() {
    ///     for (col_num, data) in row.iter().enumerate() {
    ///         // Print the data in each cell of the row.
    ///         println!("({row_num}, {col_num}): {data}");
    ///     }
    /// }
    ///
    /// // Output in relative coordinates:
    /// //
    /// // (0, 0): 1
    /// // (0, 1): 2
    /// // (1, 0):
    /// // (1, 1):
    /// // (2, 0): 3
    /// // (2, 1):
    /// ```
    ///
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

    /// Get an iterator over the used cells in a `Range`.
    ///
    /// This method returns an iterator over the used cells in a range. The
    /// "used" cells are defined as the cells that have a value other than the
    /// default value for `T`. The iterator returns tuples of `(row, column,
    /// value)` for each used cell. The row and column are relative/index values
    /// rather than absolute cell positions.
    ///
    /// # Examples
    ///
    /// An example of iterating over the used cells in a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_used_cells.rs
    /// #
    /// use calamine::{Cell, Data, Range};
    ///
    /// let cells = vec![
    ///     Cell::new((1, 1), Data::Int(1)),
    ///     Cell::new((1, 2), Data::Int(2)),
    ///     Cell::new((3, 1), Data::Int(3)),
    /// ];
    ///
    /// // Create a Range from the cells.
    /// let range = Range::from_sparse(cells);
    ///
    /// // Iterate over the used cells in the range.
    /// for (row, col, data) in range.used_cells() {
    ///     println!("({row}, {col}): {data}");
    /// }
    ///
    /// // Output:
    /// //
    /// // (0, 0): 1
    /// // (0, 1): 2
    /// // (2, 0): 3
    /// ```
    ///
    pub fn used_cells(&self) -> UsedCells<'_, T> {
        UsedCells {
            width: self.width(),
            inner: self.inner.iter().enumerate(),
        }
    }

    /// Get an iterator over all the cells in a `Range`.
    ///
    /// This method returns an iterator over all the cells in a range, including
    /// those that are empty. The iterator returns tuples of `(row, column,
    /// value)` for each cell. The row and column are relative/index values
    /// rather than absolute cell positions.
    ///
    /// # Examples
    ///
    /// An example of iterating over the used cells in a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_cells.rs
    /// #
    /// use calamine::{Cell, Data, Range};
    ///
    /// let cells = vec![
    ///     Cell::new((1, 1), Data::Int(1)),
    ///     Cell::new((1, 2), Data::Int(2)),
    ///     Cell::new((3, 1), Data::Int(3)),
    /// ];
    ///
    /// // Create a Range from the cells.
    /// let range = Range::from_sparse(cells);
    ///
    /// // Iterate over the cells in the range.
    /// for (row, col, data) in range.cells() {
    ///     println!("({row}, {col}): {data}");
    /// }
    ///
    /// // Output:
    /// //
    /// // (0, 0): 1
    /// // (0, 1): 2
    /// // (1, 0):
    /// // (1, 1):
    /// // (2, 0): 3
    /// // (2, 1):
    /// ```
    ///
    pub fn cells(&self) -> Cells<'_, T> {
        Cells {
            width: self.width(),
            inner: self.inner.iter().enumerate(),
        }
    }

    /// Build a `RangeDeserializer` for a `Range`.
    ///
    /// This method returns a [`RangeDeserializer`] that can be used to
    /// deserialize the data in the range.
    ///
    /// # Errors
    ///
    /// - [`DeError`] if the range cannot be deserialized.
    ///
    /// # Examples
    ///
    /// An example of creating a deserializer fora calamine `Range`.
    ///
    /// The sample Excel file `temperature.xlsx` contains a single sheet named
    /// "Sheet1" with the following data:
    ///
    /// ```text
    ///  ____________________________________________
    /// |         ||                |                |
    /// |         ||       A        |       B        |
    /// |_________||________________|________________|
    /// |    1    || label          | value          |
    /// |_________||________________|________________|
    /// |    2    || celsius        | 22.2222        |
    /// |_________||________________|________________|
    /// |    3    || fahrenheit     | 72             |
    /// |_________||________________|________________|
    /// |_          _________________________________|
    ///   \ Sheet1 /
    ///     ------
    /// ```
    ///
    /// ```
    /// # // This code is available in examples/doc_range_deserialize.rs
    /// #
    /// use calamine::{open_workbook, Error, Reader, Xlsx};
    ///
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///
    ///     // Open the workbook.
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///
    ///     // Get the data range from the first sheet.
    ///     let sheet_range = workbook.worksheet_range("Sheet1")?;
    ///
    ///     // Get an iterator over data in the range.
    ///     let mut iter = sheet_range.deserialize()?;
    ///
    ///     // Get the next record in the range. The first row is assumed to be the
    ///     // header.
    ///     if let Some(result) = iter.next() {
    ///         let (label, value): (String, f64) = result?;
    ///
    ///         assert_eq!(label, "celsius");
    ///         assert_eq!(value, 22.2222);
    ///
    ///         Ok(())
    ///     } else {
    ///         Err(From::from("Expected at least one record but got none"))
    ///     }
    /// }
    /// ```
    ///
    pub fn deserialize<'a, D>(&'a self) -> Result<RangeDeserializer<'a, T, D>, DeError>
    where
        T: ToCellDeserializer<'a>,
        D: DeserializeOwned,
    {
        RangeDeserializerBuilder::new().from_range(self)
    }

    /// Build a new `Range` out of the current range.
    ///
    /// This method returns a new `Range` with cloned data. In general it is
    /// used to get a subset of an existing range. However, if the new range is
    /// larger than the existing range the new cells will be filled with default
    /// values.
    ///
    /// # Examples
    ///
    /// An example of getting a sub range of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_range.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// // Create a range with some values.
    /// let mut a = Range::new((1, 1), (3, 3));
    /// a.set_value((1, 1), Data::Bool(true));
    /// a.set_value((2, 2), Data::Bool(true));
    /// a.set_value((3, 3), Data::Bool(true));
    ///
    /// // Get a sub range of the main range.
    /// let b = a.range((1, 1), (2, 2));
    /// assert_eq!(b.get_value((1, 1)), Some(&Data::Bool(true)));
    /// assert_eq!(b.get_value((2, 2)), Some(&Data::Bool(true)));
    ///
    /// // Get a larger range with default values.
    /// let c = a.range((0, 0), (5, 5));
    /// assert_eq!(c.get_value((0, 0)), Some(&Data::Empty));
    /// assert_eq!(c.get_value((3, 3)), Some(&Data::Bool(true)));
    /// assert_eq!(c.get_value((5, 5)), Some(&Data::Empty));
    /// ```
    ///
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

impl<T: CellType + fmt::Display> Range<T> {
    /// Get headers for a `Range`.
    ///
    /// This method returns the first row of the range as an optional vector of
    /// strings. The data type `T` in the range must support the [`ToString`]
    /// trait.
    ///
    /// # Examples
    ///
    /// An example of getting the header row of a calamine `Range`.
    ///
    /// ```
    /// # // This code is available in examples/doc_range_headers.rs
    /// #
    /// use calamine::{Data, Range};
    ///
    /// // Create a range with some values.
    /// let mut range = Range::new((0, 0), (5, 2));
    /// range.set_value((0, 0), Data::String(String::from("a")));
    /// range.set_value((0, 1), Data::Int(1));
    /// range.set_value((0, 2), Data::Bool(true));
    ///
    /// // Get the headers of the range.
    /// let headers = range.headers();
    ///
    /// assert_eq!(
    ///     headers,
    ///     Some(vec![
    ///         String::from("a"),
    ///         String::from("1"),
    ///         String::from("true")
    ///     ])
    /// );
    /// ```
    ///
    pub fn headers(&self) -> Option<Vec<String>> {
        self.rows()
            .next()
            .map(|row| row.iter().map(ToString::to_string).collect())
    }
}

/// Implementation of the `Index` trait for `Range` rows.
///
/// # Examples
///
/// An example of row indexing for a calamine `Range`.
///
/// ```
/// # // This code is available in examples/doc_range_index_row.rs
/// #
/// use calamine::{Data, Range};
///
/// // Create a range with a value.
/// let mut range = Range::new((1, 1), (3, 3));
/// range.set_value((2, 2), Data::Int(123));
///
/// // Get the second row via indexing.
/// assert_eq!(range[1], [Data::Empty, Data::Int(123), Data::Empty]);
/// ```
///
impl<T: CellType> Index<usize> for Range<T> {
    type Output = [T];
    fn index(&self, index: usize) -> &[T] {
        let width = self.width();
        &self.inner[index * width..(index + 1) * width]
    }
}

/// Implementation of the `Index` trait for `Range` cells.
///
/// # Examples
///
/// An example of cell indexing for a calamine `Range`.
///
/// ```
/// # // This code is available in examples/doc_range_index_cell.rs
/// #
/// use calamine::{Data, Range};
///
/// // Create a range with a value.
/// let mut range = Range::new((1, 1), (3, 3));
/// range.set_value((2, 2), Data::Int(123));
///
/// // Get the value via cell indexing.
/// assert_eq!(range[(1, 1)], Data::Int(123));
/// ```
///
impl<T: CellType> Index<(usize, usize)> for Range<T> {
    type Output = T;
    fn index(&self, index: (usize, usize)) -> &T {
        let (height, width) = self.get_size();
        assert!(index.1 < width && index.0 < height, "index out of bounds");
        &self.inner[index.0 * width + index.1]
    }
}

/// Implementation of the `IndexMut` trait for `Range` rows.
impl<T: CellType> IndexMut<usize> for Range<T> {
    fn index_mut(&mut self, index: usize) -> &mut [T] {
        let width = self.width();
        &mut self.inner[index * width..(index + 1) * width]
    }
}

/// Implementation of the `IndexMut` trait for `Range` cells.
///
/// # Examples
///
/// An example of mutable cell indexing for a calamine `Range`.
///
/// ```
/// # // This code is available in examples/doc_range_index_mut_cell.rs
/// #
/// use calamine::{Data, Range};
///
/// // Create a new empty range.
/// let mut range = Range::new((1, 1), (3, 3));
///
/// // Set a value in the range using cell indexing.
/// range[(1, 1)] = Data::Int(123);
///
/// // Test the value was set correctly.
/// assert_eq!(range.get((1, 1)), Some(&Data::Int(123)));
/// ```
///
impl<T: CellType> IndexMut<(usize, usize)> for Range<T> {
    fn index_mut(&mut self, index: (usize, usize)) -> &mut T {
        let (height, width) = self.get_size();
        assert!(index.1 < width && index.0 < height, "index out of bounds");
        &mut self.inner[index.0 * width + index.1]
    }
}

/// A struct to iterate over all cells
#[derive(Clone, Debug)]
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
#[derive(Clone, Debug)]
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
#[derive(Clone, Debug)]
pub struct Rows<'a, T: CellType> {
    inner: Option<std::slice::Chunks<'a, T>>,
}

impl<'a, T: 'a + CellType> Iterator for Rows<'a, T> {
    type Item = &'a [T];
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.as_mut().and_then(std::iter::Iterator::next)
    }
    fn size_hint(&self) -> (usize, Option<usize>) {
        self.inner
            .as_ref()
            .map_or((0, Some(0)), std::iter::Iterator::size_hint)
    }
}

impl<'a, T: 'a + CellType> DoubleEndedIterator for Rows<'a, T> {
    fn next_back(&mut self) -> Option<Self::Item> {
        self.inner
            .as_mut()
            .and_then(std::iter::DoubleEndedIterator::next_back)
    }
}

impl<'a, T: 'a + CellType> ExactSizeIterator for Rows<'a, T> {}

/// Struct with the key elements of a table
pub struct Table<T> {
    pub(crate) name: String,
    pub(crate) sheet_name: String,
    pub(crate) columns: Vec<String>,
    pub(crate) data: Range<T>,
}
impl<T> Table<T> {
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

impl<T: CellType> From<Table<T>> for Range<T> {
    fn from(table: Table<T>) -> Range<T> {
        table.data
    }
}

impl From<Table<DataWithFormatting>> for Range<Data> {
    fn from(table: Table<DataWithFormatting>) -> Range<Data> {
        let inner = table.data.inner.into_iter().map(|dwf| dwf.data).collect();
        Range {
            start: table.data.start,
            end: table.data.end,
            inner,
        }
    }
}

/// A helper function to deserialize cell values as `i64`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_i64`](crate::datatype::DataType::as_i64) method to the cell value, and returns
/// `Ok(Some(value_as_i64))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
pub fn deserialize_as_i64_or_none<'de, D>(deserializer: D) -> Result<Option<i64>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_i64())
}

/// A helper function to deserialize cell values as `i64`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_i64`](crate::datatype::DataType::as_i64) method to the cell value, and returns
/// `Ok(Ok(value_as_i64))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
pub fn deserialize_as_i64_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<i64, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_i64().ok_or_else(|| data.to_string()))
}

/// A helper function to deserialize cell values as `f64`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_f64`](crate::datatype::DataType::as_f64) method to the cell value, and returns
/// `Ok(Some(value_as_f64))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
pub fn deserialize_as_f64_or_none<'de, D>(deserializer: D) -> Result<Option<f64>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_f64())
}

/// A helper function to deserialize cell values as `f64`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_f64`](crate::datatype::DataType::as_f64) method to the cell value, and returns
/// `Ok(Ok(value_as_f64))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
pub fn deserialize_as_f64_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<f64, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_f64().ok_or_else(|| data.to_string()))
}

/// A helper function to deserialize cell values as `chrono::NaiveDate`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_date`] method to the cell value, and returns
/// `Ok(Some(value_as_date))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_date_or_none<'de, D>(
    deserializer: D,
) -> Result<Option<chrono::NaiveDate>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_date())
}

/// A helper function to deserialize cell values as `chrono::NaiveDate`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_date`] method to the cell value, and returns
/// `Ok(Ok(value_as_date))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_date_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<chrono::NaiveDate, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_date().ok_or_else(|| data.to_string()))
}

/// A helper function to deserialize cell values as `chrono::NaiveTime`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_time`] method to the cell value, and returns
/// `Ok(Some(value_as_time))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_time_or_none<'de, D>(
    deserializer: D,
) -> Result<Option<chrono::NaiveTime>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_time())
}

/// A helper function to deserialize cell values as `chrono::NaiveTime`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_time`] method to the cell value, and returns
/// `Ok(Ok(value_as_time))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_time_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<chrono::NaiveTime, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_time().ok_or_else(|| data.to_string()))
}

/// A helper function to deserialize cell values as `chrono::Duration`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_duration`] method to the cell value, and returns
/// `Ok(Some(value_as_duration))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_duration_or_none<'de, D>(
    deserializer: D,
) -> Result<Option<chrono::Duration>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_duration())
}

/// A helper function to deserialize cell values as `chrono::Duration`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_duration`] method to the cell value, and returns
/// `Ok(Ok(value_as_duration))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_duration_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<chrono::Duration, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_duration().ok_or_else(|| data.to_string()))
}

/// A helper function to deserialize cell values as `chrono::NaiveDateTime`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_datetime`] method to the cell value, and returns
/// `Ok(Some(value_as_datetime))` if successful or `Ok(None)` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_datetime_or_none<'de, D>(
    deserializer: D,
) -> Result<Option<chrono::NaiveDateTime>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_datetime())
}

/// A helper function to deserialize cell values as `chrono::NaiveDateTime`,
/// useful when cells may also contain invalid values (i.e. strings).
/// It applies the [`as_datetime`] method to the cell value, and returns
/// `Ok(Ok(value_as_datetime))` if successful or `Ok(Err(value_to_string))` if unsuccessful,
/// therefore never failing. This function is intended to be used with Serde's
/// [`deserialize_with`](https://serde.rs/field-attrs.html) field attribute.
#[cfg(feature = "dates")]
pub fn deserialize_as_datetime_or_string<'de, D>(
    deserializer: D,
) -> Result<Result<chrono::NaiveDateTime, String>, D::Error>
where
    D: Deserializer<'de>,
{
    let data = Data::deserialize(deserializer)?;
    Ok(data.as_datetime().ok_or_else(|| data.to_string()))
}
