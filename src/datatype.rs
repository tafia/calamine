use std::fmt;

use serde::{self, Deserialize};
use serde::de::Visitor;

use super::CellErrorType;

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

impl fmt::Display for DataType {
    fn fmt(&self, f: &mut fmt::Formatter) -> ::std::result::Result<(), fmt::Error> {
        match *self {
            DataType::Int(ref e) => write!(f, "{}", e),
            DataType::Float(ref e) => write!(f, "{}", e),
            DataType::String(ref e) => write!(f, "{}", e),
            DataType::Bool(ref e) => write!(f, "{}", e),
            DataType::Error(ref e) => write!(f, "{}", e),
            DataType::Empty => Ok(()),
        }
    }
}

impl<'de> Deserialize<'de> for DataType {
    #[inline]
    fn deserialize<D>(deserializer: D) -> Result<DataType, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct DataTypeVisitor;

        impl<'de> Visitor<'de> for DataTypeVisitor {
            type Value = DataType;

            fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
                formatter.write_str("any valid JSON value")
            }

            #[inline]
            fn visit_bool<E>(self, value: bool) -> Result<DataType, E> {
                Ok(DataType::Bool(value))
            }

            #[inline]
            fn visit_i64<E>(self, value: i64) -> Result<DataType, E> {
                Ok(DataType::Int(value))
            }

            #[inline]
            fn visit_u64<E>(self, value: u64) -> Result<DataType, E> {
                Ok(DataType::Int(value as i64))
            }

            #[inline]
            fn visit_f64<E>(self, value: f64) -> Result<DataType, E> {
                Ok(DataType::Float(value))
            }

            #[inline]
            fn visit_str<E>(self, value: &str) -> Result<DataType, E>
            where
                E: serde::de::Error,
            {
                self.visit_string(String::from(value))
            }

            #[inline]
            fn visit_string<E>(self, value: String) -> Result<DataType, E> {
                Ok(DataType::String(value))
            }

            #[inline]
            fn visit_none<E>(self) -> Result<DataType, E> {
                Ok(DataType::Empty)
            }

            #[inline]
            fn visit_some<D>(self, deserializer: D) -> Result<DataType, D::Error>
            where
                D: serde::Deserializer<'de>,
            {
                Deserialize::deserialize(deserializer)
            }

            #[inline]
            fn visit_unit<E>(self) -> Result<DataType, E> {
                Ok(DataType::Empty)
            }
        }

        deserializer.deserialize_any(DataTypeVisitor)
    }
}

macro_rules! define_from {
    ($variant:path, $ty:ty) => {
        impl From<$ty> for DataType {
            fn from(v: $ty) -> Self {
                $variant(v)
            }
        }
    }
}

define_from!(DataType::Int, i64);
define_from!(DataType::Float, f64);
define_from!(DataType::String, String);
define_from!(DataType::Bool, bool);
define_from!(DataType::Error, CellErrorType);

impl<'a> From<&'a str> for DataType {
    fn from(v: &'a str) -> Self {
        DataType::String(String::from(v))
    }
}

impl From<()> for DataType {
    fn from(_: ()) -> Self {
        DataType::Empty
    }
}

impl<T> From<Option<T>> for DataType
where
    DataType: From<T>,
{
    fn from(v: Option<T>) -> Self {
        match v {
            Some(v) => From::from(v),
            None => DataType::Empty,
        }
    }
}
