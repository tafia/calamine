use serde::de::value::BorrowedStrDeserializer;
use serde::de::{self, DeserializeOwned, DeserializeSeed, SeqAccess, Visitor};
use serde::{self, forward_to_deserialize_any, Deserialize};
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
    /// Required header not found
    HeaderNotFound(String),
    /// Serde specific error
    Custom(String),
}

impl fmt::Display for DeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> Result<(), fmt::Error> {
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
            DeError::HeaderNotFound(ref header) => {
                write!(f, "Cannot find header named '{}'", header)
            }
            DeError::Custom(ref s) => write!(f, "{}", s),
        }
    }
}

impl std::error::Error for DeError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        None
    }
}

impl de::Error for DeError {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        DeError::Custom(msg.to_string())
    }
}

#[derive(Clone)]
pub enum Headers<'h, H> {
    None,
    All,
    Custom(&'h [H]),
}

/// Builds a `Range` deserializer with some configuration options.
///
/// This can be used to optionally parse the first row as a header. Once built,
/// a `RangeDeserializer`s cannot be changed.
#[derive(Clone)]
pub struct RangeDeserializerBuilder<'h, H> {
    headers: Headers<'h, H>,
}

impl Default for RangeDeserializerBuilder<'static, &'static str> {
    fn default() -> Self {
        RangeDeserializerBuilder {
            headers: Headers::All,
        }
    }
}

impl RangeDeserializerBuilder<'static, &'static str> {
    /// Constructs a new builder for configuring `Range` deserialization.
    pub fn new() -> Self {
        Default::default()
    }

    /// Decide whether to treat the first row as a special header row.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{DataType, Error, open_workbook, Xlsx, Reader, RangeDeserializerBuilder};
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
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
    ///         assert_eq!(row, [DataType::from("celsius"), DataType::from(22.2222)]);
    ///     } else {
    ///         return Err(From::from("expected at least three records but got one"));
    ///     }
    ///
    ///     Ok(())
    /// }
    /// ```
    pub fn has_headers(&mut self, yes: bool) -> &mut Self {
        if yes {
            self.headers = Headers::All;
        } else {
            self.headers = Headers::None;
        }
        self
    }
}

impl<'h, H: AsRef<str> + Clone + 'h> RangeDeserializerBuilder<'h, H> {
    /// Build a `RangeDeserializer` from this configuration and keep only selected headers.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let range = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///     let mut iter = RangeDeserializerBuilder::with_headers(&["value", "label"]).from_range(&range)?;
    ///
    ///     if let Some(result) = iter.next() {
    ///         let (value, label): (f64, String) = result?;
    ///         assert_eq!(label, "celsius");
    ///         assert_eq!(value, 22.2222);
    ///
    ///         Ok(())
    ///     } else {
    ///         return Err(From::from("expected at least one record but got none"));
    ///     }
    /// }
    /// ```
    pub fn with_headers(headers: &'h [H]) -> Self {
        RangeDeserializerBuilder {
            headers: Headers::Custom(headers),
        }
    }

    /// Build a `RangeDeserializer` from this configuration.
    ///
    /// # Example
    ///
    /// ```
    /// # use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
    /// fn main() -> Result<(), Error> {
    ///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    ///     let mut workbook: Xlsx<_> = open_workbook(path)?;
    ///     let range = workbook.worksheet_range("Sheet1")
    ///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
    ///     let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;
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
}

/// A configured `Range` deserializer.
///
/// # Example
///
/// ```
/// # use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};
/// fn main() -> Result<(), Error> {
///     let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
///     let mut workbook: Xlsx<_> = open_workbook(path)?;
///     let range = workbook.worksheet_range("Sheet1")
///         .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;
///
///     let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;
///
///     if let Some(result) = iter.next() {
///         let (label, value): (String, f64) = result?;
///         assert_eq!(label, "celsius");
///         assert_eq!(value, 22.2222);
///         Ok(())
///     } else {
///         Err(From::from("expected at least one record but got none"))
///     }
/// }
/// ```
pub struct RangeDeserializer<'cell, T, D>
where
    T: ToCellDeserializer<'cell>,
    D: DeserializeOwned,
{
    column_indexes: Vec<usize>,
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
    fn new<'h, H: AsRef<str> + Clone + 'h>(
        builder: &RangeDeserializerBuilder<'h, H>,
        range: &'cell Range<T>,
    ) -> Result<Self, DeError> {
        let mut rows = range.rows();

        let mut current_pos = range.start().unwrap_or((0, 0));
        let end_pos = range.end().unwrap_or((0, 0));

        let (column_indexes, headers) = match builder.headers {
            Headers::None => ((0..range.width()).collect(), None),
            Headers::All => {
                if let Some(row) = rows.next() {
                    let all_indexes = (0..row.len()).collect::<Vec<_>>();
                    let all_headers = {
                        let de = RowDeserializer::new(&all_indexes, None, row, current_pos);
                        current_pos.0 += 1;
                        Deserialize::deserialize(de)?
                    };
                    (all_indexes, Some(all_headers))
                } else {
                    (Vec::new(), None)
                }
            }
            Headers::Custom(headers) => {
                if let Some(row) = rows.next() {
                    let all_indexes = (0..row.len()).collect::<Vec<_>>();
                    let de = RowDeserializer::new(&all_indexes, None, row, current_pos);
                    current_pos.0 += 1;
                    let all_headers: Vec<String> = Deserialize::deserialize(de)?;
                    let custom_indexes = headers
                        .iter()
                        .map(|h| h.as_ref().trim())
                        .map(|h| {
                            all_headers
                                .iter()
                                .position(|header| header.trim() == h)
                                .ok_or_else(|| DeError::HeaderNotFound(h.to_owned()))
                        })
                        .collect::<Result<Vec<_>, DeError>>()?;
                    (custom_indexes, Some(all_headers))
                } else {
                    (Vec::new(), None)
                }
            }
        };

        Ok(RangeDeserializer {
            column_indexes,
            headers,
            rows,
            current_pos,
            end_pos,
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
            ref column_indexes,
            ref headers,
            ref mut rows,
            mut current_pos,
            ..
        } = *self;

        if let Some(row) = rows.next() {
            current_pos.0 += 1;
            let headers = headers.as_ref().map(|h| &**h);
            let de = RowDeserializer::new(column_indexes, headers, row, current_pos);
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
    T: ToCellDeserializer<'cell>,
{
    cells: &'cell [T],
    headers: Option<&'header [String]>,
    iter: slice::Iter<'header, usize>, // iterator over column indexes
    peek: Option<usize>,
    pos: (u32, u32),
}

impl<'header, 'cell, T> RowDeserializer<'header, 'cell, T>
where
    T: 'cell + ToCellDeserializer<'cell>,
{
    fn new(
        column_indexes: &'header [usize],
        headers: Option<&'header [String]>,
        cells: &'cell [T],
        pos: (u32, u32),
    ) -> Self {
        RowDeserializer {
            iter: column_indexes.iter(),
            headers,
            cells,
            pos,
            peek: None,
        }
    }

    fn has_headers(&self) -> bool {
        self.headers.is_some()
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
        match self.iter.next().map(|i| &self.cells[*i]) {
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
        let headers = self
            .headers
            .expect("Cannot map-deserialize range without headers");

        while let Some(i) = self.iter.next() {
            if !self.cells[*i].is_empty() {
                self.peek = Some(*i);
                let de = BorrowedStrDeserializer::<Self::Error>::new(&headers[*i]);
                return seed.deserialize(de).map(Some);
            }
        }
        Ok(None)
    }

    fn next_value_seed<K: DeserializeSeed<'de>>(
        &mut self,
        seed: K,
    ) -> Result<K::Value, Self::Error> {
        let cell = self
            .peek
            .take()
            .map(|i| &self.cells[i])
            .ok_or(DeError::UnexpectedEndOfRow { pos: self.pos })?;
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

    /// Assess if the cell is empty.
    fn is_empty(&self) -> bool;
}

impl<'a> ToCellDeserializer<'a> for DataType {
    type Deserializer = DataTypeDeserializer<'a>;

    fn to_cell_deserializer(&'a self, pos: (u32, u32)) -> DataTypeDeserializer<'a> {
        DataTypeDeserializer {
            data_type: self,
            pos,
        }
    }

    #[inline]
    fn is_empty(&self) -> bool {
        if let DataType::Empty = self {
            true
        } else {
            false
        }
    }
}

macro_rules! deserialize_num {
    ($typ:ty, $method:ident, $visit:ident) => {
        fn $method<V>(self, visitor: V) -> Result<V::Value, Self::Error>
        where
            V: Visitor<'de>,
        {
            match self.data_type {
                DataType::Float(v) => visitor.$visit(*v as $typ),
                DataType::Int(v) => visitor.$visit(*v as $typ),
                DataType::String(ref s) => {
                    let v = s.parse().map_err(|_| {
                        DeError::Custom(format!("Expecting {}, got '{}'", stringify!($typ), s))
                    })?;
                    visitor.$visit(v)
                }
                DataType::Error(ref err) => Err(DeError::CellError {
                    err: err.clone(),
                    pos: self.pos,
                }),
                ref d => Err(DeError::Custom(format!(
                    "Expecting {}, got {:?}",
                    stringify!($typ),
                    d
                ))),
            }
        }
    };
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
        match self.data_type {
            DataType::String(v) => visitor.visit_str(v),
            DataType::Float(v) => visitor.visit_f64(*v),
            DataType::Bool(v) => visitor.visit_bool(*v),
            DataType::Int(v) => visitor.visit_i64(*v),
            DataType::Empty => visitor.visit_unit(),
            DataType::DateTime(v) => visitor.visit_f64(*v),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
        }
    }

    fn deserialize_str<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
            DataType::String(v) => visitor.visit_str(v),
            DataType::Empty => visitor.visit_str(""),
            DataType::Float(v) => visitor.visit_str(&v.to_string()),
            DataType::Int(v) => visitor.visit_str(&v.to_string()),
            DataType::Bool(v) => visitor.visit_str(&v.to_string()),
            DataType::DateTime(v) => visitor.visit_str(&v.to_string()),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
        }
    }

    fn deserialize_bytes<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
            DataType::String(v) => visitor.visit_bytes(v.as_bytes()),
            DataType::Empty => visitor.visit_bytes(&[]),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
            ref d => Err(DeError::Custom(format!("Expecting bytes, got {:?}", d))),
        }
    }

    fn deserialize_byte_buf<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        self.deserialize_bytes(visitor)
    }

    fn deserialize_string<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        self.deserialize_str(visitor)
    }

    fn deserialize_bool<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
            DataType::Bool(v) => visitor.visit_bool(*v),
            DataType::String(ref v) => match &**v {
                "TRUE" | "true" | "True" => visitor.visit_bool(true),
                "FALSE" | "false" | "False" => visitor.visit_bool(false),
                d => Err(DeError::Custom(format!("Expecting bool, got '{}'", d))),
            },
            DataType::Empty => visitor.visit_bool(false),
            DataType::Float(v) => visitor.visit_bool(*v != 0.),
            DataType::Int(v) => visitor.visit_bool(*v != 0),
            DataType::DateTime(v) => visitor.visit_bool(*v != 0.),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
        }
    }

    fn deserialize_char<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
            DataType::String(ref s) if s.len() == 1 => {
                visitor.visit_char(s.chars().next().expect("s not empty"))
            }
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
            ref d => Err(DeError::Custom(format!("Expecting unit, got {:?}", d))),
        }
    }

    fn deserialize_unit<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
            DataType::Empty => visitor.visit_unit(),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
            ref d => Err(DeError::Custom(format!("Expecting unit, got {:?}", d))),
        }
    }

    fn deserialize_option<V>(self, visitor: V) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        match self.data_type {
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

    fn deserialize_enum<V>(
        self,
        _name: &'static str,
        _variants: &'static [&'static str],
        visitor: V,
    ) -> Result<V::Value, Self::Error>
    where
        V: Visitor<'de>,
    {
        use serde::de::IntoDeserializer;

        match self.data_type {
            DataType::String(s) => visitor.visit_enum(s.as_str().into_deserializer()),
            DataType::Error(ref err) => Err(DeError::CellError {
                err: err.clone(),
                pos: self.pos,
            }),
            ref d => Err(DeError::Custom(format!("Expecting enum, got {:?}", d))),
        }
    }

    deserialize_num!(i64, deserialize_i64, visit_i64);
    deserialize_num!(i32, deserialize_i32, visit_i32);
    deserialize_num!(i16, deserialize_i16, visit_i16);
    deserialize_num!(i8, deserialize_i8, visit_i8);
    deserialize_num!(u64, deserialize_u64, visit_u64);
    deserialize_num!(u32, deserialize_u32, visit_u32);
    deserialize_num!(u16, deserialize_u16, visit_u16);
    deserialize_num!(u8, deserialize_u8, visit_u8);
    deserialize_num!(f64, deserialize_f64, visit_f64);
    deserialize_num!(f32, deserialize_f32, visit_f32);

    forward_to_deserialize_any! {
        unit_struct seq tuple tuple_struct map struct identifier ignored_any
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn test_deserialize_enum() {
        use crate::ToCellDeserializer;
        use serde::Deserialize;

        #[derive(Debug, serde_derive::Deserialize, PartialEq)]
        enum Content {
            Foo,
        }

        assert_eq!(
            Content::deserialize(
                super::DataType::String("Foo".to_string()).to_cell_deserializer((0, 0))
            )
            .unwrap(),
            Content::Foo
        );
    }
}
