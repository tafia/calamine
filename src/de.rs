use serde::de::value::BorrowedStrDeserializer;
use serde::de::{self, DeserializeOwned, DeserializeSeed, SeqAccess, Visitor};
use serde::{self, Deserialize};
use std::marker::PhantomData;
use std::{fmt, slice, str};

use super::{CellErrorType, CellType, DataType, Range, Rows};

/// A cell deserialization specific error enum
#[derive(Debug)]
pub enum DeError {
    /// Cell out of range
    CellOutOfRange {
        /// Position tried
        try_pos: (u32, u32),
        /// Minimum position
        min_pos: (u32, u32),
    },
    /// The cell value is an error
    CellError {
        /// Cell value error
        err: CellErrorType,
        /// Cell position
        pos: (u32, u32),
    },
    /// Unexpected end of row
    UnexpectedEndOfRow {
        /// Cell position
        pos: (u32, u32),
    },
    /// Serde specific error
    Custom(String),
}

impl fmt::Display for DeError {
    fn fmt(&self, f: &mut fmt::Formatter) -> Result<(), fmt::Error> {
        match *self {
            DeError::CellOutOfRange {
                ref try_pos,
                ref min_pos,
            } => write!(
                f,
                "there is no cell at position '{:?}'.Minimum position is '{:?}'",
                try_pos, min_pos
            ),
            DeError::CellError { ref pos, ref err } => {
                write!(f, "Cell error at position '{:?}': {}", pos, err)
            }
            DeError::UnexpectedEndOfRow { ref pos } => {
                write!(f, "Unexpected end of row at position '{:?}'", pos)
            }
            DeError::Custom(ref s) => write!(f, "{}", s),
        }
    }
}

impl ::std::error::Error for DeError {
    fn description(&self) -> &str {
        match *self {
            DeError::CellOutOfRange { .. } => "cell out of range",
            DeError::CellError { .. } => "error in cell value",
            DeError::UnexpectedEndOfRow { .. } => "unexpected end of row",
            DeError::Custom(ref s) => &**s,
        }
    }
    fn cause(&self) -> Option<&::std::error::Error> {
        None
    }
}

impl de::Error for DeError {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        DeError::Custom(msg.to_string())
    }
}

/// Builds a `Range` deserializer with some configuration options.
///
/// This can be used to optionally parse the first row as a header. Once built,
/// a `RangeDeserializer`s cannot be changed.
#[derive(Clone)]
pub struct RangeDeserializerBuilder {
    has_headers: bool,
}

impl Default for RangeDeserializerBuilder {
    fn default() -> Self {
        RangeDeserializerBuilder { has_headers: true }
    }
}

impl RangeDeserializerBuilder {
    /// Constructs a new builder for configuring `Range` deserialization.
    pub fn new() -> Self {
        Default::default()
    }

    /// Build a `RangeDeserializer` from this configuration.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
    /// # fn main() { example().unwrap(); }
    /// fn example() -> Result<(), Error> {
    ///     let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let range = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///     let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;
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
    pub fn from_range<'cell, T, D>(
        &self,
        range: &'cell Range<T>,
    ) -> Result<RangeDeserializer<'cell, T, D>, DeError>
    where
        T: ToCellDeserializer<'cell>,
        D: DeserializeOwned,
    {
        RangeDeserializer::new(self, range)
    }

    /// Decide whether to treat the first row as a special header row.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{DataType, Error, open_workbook, Xlsx, Reader, RangeDeserializerBuilder};
    /// # fn main() { example().unwrap(); }
    /// fn example() -> Result<(), Error> {
    ///     let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let range = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///
    ///     let mut iter = RangeDeserializerBuilder::new()
    ///         .has_headers(false)
    ///         .from_range(&range)?;
    ///
    ///     if let Some(result) = iter.next() {
    ///         let row: Vec<DataType> = result?;
    ///         assert_eq!(row, [DataType::from("label"), DataType::from("value")]);
    ///     } else {
    ///         return Err(From::from("expected at least three records but got none"));
    ///     }
    ///
    ///     if let Some(result) = iter.next() {
    ///         let row: Vec<DataType> = result?;
    ///         assert_eq!(row, [DataType::from("celcius"), DataType::from(22.2222)]);
    ///     } else {
    ///         return Err(From::from("expected at least three records but got one"));
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn has_headers(&mut self, yes: bool) -> &mut RangeDeserializerBuilder {
        self.has_headers = yes;
        self
    }
}

/// A configured `Range` deserializer.
///
/// # Example
///
/// ```
/// # use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
/// # fn main() { example().unwrap(); }
/// fn example() -> Result<(), Error> {
///     let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
///     let mut workbook: Xlsx<_> = open_workbook(path)?;
///     let range = workbook.worksheet_range("Sheet1")
///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
///
///     let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;
///
///     if let Some(result) = iter.next() {
///         let (label, value): (String, f64) = result?;
///         assert_eq!(label, "celcius");
///         assert_eq!(value, 22.2222);
///         Ok(())
///     } else {
///         Err(From::from("expected at least one record but got none"))
///     }
/// }
/// ```
pub struct RangeDeserializer<'cell, T, D>
where
    T: 'cell + ToCellDeserializer<'cell>,
    D: DeserializeOwned,
{
    headers: Option<Vec<String>>,
    rows: Rows<'cell, T>,
    current_pos: (u32, u32),
    end_pos: (u32, u32),
    _priv: PhantomData<D>,
}

impl<'cell, T, D> RangeDeserializer<'cell, T, D>
where
    T: ToCellDeserializer<'cell>,
    D: DeserializeOwned,
{
    fn new(builder: &RangeDeserializerBuilder, range: &'cell Range<T>) -> Result<Self, DeError> {
        let mut rows = range.rows();

        let mut current_pos = range.start();
        let end_pos = range.end();

        let headers = if builder.has_headers {
            if let Some(row) = rows.next() {
                let de = RowDeserializer::new(None, row, current_pos);
                current_pos.0 += 1;
                Some(Deserialize::deserialize(de)?)
            } else {
                None
            }
        } else {
            None
        };

        Ok(RangeDeserializer {
            headers: headers,
            rows: rows,
            current_pos: current_pos,
            end_pos: end_pos,
            _priv: PhantomData,
        })
    }
}

impl<'cell, T, D> Iterator for RangeDeserializer<'cell, T, D>
where
    T: ToCellDeserializer<'cell>,
    D: DeserializeOwned,
{
    type Item = Result<D, DeError>;

    fn next(&mut self) -> Option<Self::Item> {
        let RangeDeserializer {
            ref headers,
            ref mut rows,
            mut current_pos,
            ..
        } = *self;

        if let Some(row) = rows.next() {
            current_pos.0 += 1;

            let headers = headers.as_ref();
            let de = RowDeserializer::new(headers, row, current_pos);
            Some(Deserialize::deserialize(de))
        } else {
            None
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = (self.end_pos.0 - self.current_pos.0) as usize;

        (remaining, Some(remaining))
    }
}

struct RowDeserializer<'header, 'cell, T>
where
    T: 'cell + ToCellDeserializer<'cell>,
{
    headers: Option<slice::Iter<'header, String>>,
    iter: slice::Iter<'cell, T>,
    pos: (u32, u32),
}

impl<'header, 'cell, T> RowDeserializer<'header, 'cell, T>
where
    T: 'cell + ToCellDeserializer<'cell>,
{
    fn new(headers: Option<&'header Vec<String>>, record: &'cell [T], pos: (u32, u32)) -> Self {
        RowDeserializer {
            iter: record.into_iter(),
            headers: headers.map(|headers| headers.into_iter()),
            pos: pos,
        }
    }

    fn has_headers(&self) -> bool {
        self.headers.is_some()
    }

    fn next_header(&mut self) -> Option<&'header str> {
        self.headers
            .as_mut()
            .and_then(|it| it.next().map(|header| &**header))
    }

    fn next_cell(&mut self) -> Result<&'cell T, DeError> {
        if let Some(cell) = self.iter.next() {
            self.pos.1 += 1;
            Ok(cell)
        } else {
            return Err(DeError::UnexpectedEndOfRow { pos: self.pos });
        }
    }
}

impl<'de, 'header, 'cell, T> serde::Deserializer<'de> for RowDeserializer<'header, 'cell, T>
where
    'header: 'de,
    'cell: 'de,
    T: 'cell + ToCellDeserializer<'cell>,
{
    type Error = DeError;

    fn deserialize_any<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        visitor.visit_seq(self)
    }

    fn deserialize_map<V: Visitor<'de>>(self, visitor: V) -> Result<V::Value, Self::Error> {
        if !self.has_headers() {
            visitor.visit_seq(self)
        } else {
            visitor.visit_map(self)
        }
    }

    fn deserialize_struct<V: Visitor<'de>>(
        self,
        _name: &'static str,
        _cells: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Self::Error> {
        if !self.has_headers() {
            visitor.visit_seq(self)
        } else {
            visitor.visit_map(self)
        }
    }

    forward_to_deserialize_any! {
        bool i8 i16 i32 i64 u8 u16 u32 u64 f32 f64 char str string bytes
        byte_buf option unit unit_struct newtype_struct seq tuple
        tuple_struct enum identifier ignored_any
    }
}

impl<'de, 'header, 'cell, T> SeqAccess<'de> for RowDeserializer<'header, 'cell, T>
where
    'header: 'de,
    'cell: 'de,
    T: ToCellDeserializer<'cell>,
{
    type Error = DeError;

    fn next_element_seed<D>(&mut self, seed: D) -> Result<Option<D::Value>, Self::Error>
    where
        D: DeserializeSeed<'de>,
    {
        match self.iter.next() {
            Some(value) => {
                let de = value.to_cell_deserializer(self.pos);
                seed.deserialize(de).map(Some)
            }
            None => Ok(None),
        }
    }

    fn size_hint(&self) -> Option<usize> {
        match self.iter.size_hint() {
            (lower, Some(upper)) if lower == upper => Some(upper),
            _ => None,
        }
    }
}

impl<'de, 'header: 'de, 'cell: 'de, T> de::MapAccess<'de> for RowDeserializer<'header, 'cell, T>
where
    'header: 'de,
    'cell: 'de,
    T: ToCellDeserializer<'cell>,
{
    type Error = DeError;

    fn next_key_seed<K: DeserializeSeed<'de>>(
        &mut self,
        seed: K,
    ) -> Result<Option<K::Value>, Self::Error> {
        assert!(self.has_headers());

        if let Some(header) = self.next_header() {
            let de = BorrowedStrDeserializer::<Self::Error>::new(header);
            seed.deserialize(de).map(Some)
        } else {
            Ok(None)
        }
    }

    fn next_value_seed<K: DeserializeSeed<'de>>(
        &mut self,
        seed: K,
    ) -> Result<K::Value, Self::Error> {
        let cell = self.next_cell()?;
        let de = cell.to_cell_deserializer(self.pos);
        seed.deserialize(de)
    }
}

/// Constructs a deserializer for a `CellType`.
pub trait ToCellDeserializer<'a>: CellType {
    /// The deserializer.
    type Deserializer: for<'de> serde::Deserializer<'de, Error = DeError>;

    /// Construct a `CellType` deserializer at the specified position.
    fn to_cell_deserializer(&'a self, pos: (u32, u32)) -> Self::Deserializer;
}

impl<'a> ToCellDeserializer<'a> for DataType {
    type Deserializer = DataTypeDeserializer<'a>;

    fn to_cell_deserializer(&'a self, pos: (u32, u32)) -> DataTypeDeserializer<'a> {
        DataTypeDeserializer {
            data_type: self,
            pos: pos,
        }
    }
}

/// A deserializer for the `DataType` type.
pub struct DataTypeDeserializer<'a> {
    data_type: &'a DataType,
    pos: (u32, u32),
}

impl<'a, 'de> serde::Deserializer<'de> for DataTypeDeserializer<'a> {
    type Error = DeError;

    fn deserialize_any<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match *self.data_type {
            DataType::Empty => visitor.visit_unit(),
            DataType::Bool(v) => visitor.visit_bool(v),
            DataType::Int(v) => visitor.visit_i64(v),
            DataType::Float(v) => visitor.visit_f64(v),
            DataType::String(ref v) => visitor.visit_str(v),
            DataType::Error(ref err) => {
                return Err(DeError::CellError {
                    err: err.clone(),
                    pos: self.pos,
                })
            }
        }
    }

    fn deserialize_option<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match *self.data_type {
            DataType::Empty => visitor.visit_none(),
            _ => visitor.visit_some(self),
        }
    }

    fn deserialize_newtype_struct<V>(
        self,
        _name: &'static str,
        visitor: V,
    ) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        visitor.visit_newtype_struct(self)
    }

    forward_to_deserialize_any! {
        bool i8 i16 i32 i64 u8 u16 u32 u64 f32 f64 char str string bytes
        byte_buf unit unit_struct seq tuple
        tuple_struct map struct enum identifier ignored_any
    }
}
