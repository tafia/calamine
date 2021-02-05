use std::fmt;

use serde::de::Visitor;
use serde::{self, Deserialize};

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
    /// Date or Time
    DateTime(f64),
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

impl DataType {
    /// Assess if datatype is empty
    pub fn is_empty(&self) -> bool {
        *self == DataType::Empty
    }
    /// Assess if datatype is a int
    pub fn is_int(&self) -> bool {
        matches!(*self, DataType::Int(_))
    }
    /// Assess if datatype is a float
    pub fn is_float(&self) -> bool {
        matches!(*self, DataType::Float(_))
    }
    /// Assess if datatype is a bool
    pub fn is_bool(&self) -> bool {
        matches!(*self, DataType::Bool(_))
    }
    /// Assess if datatype is a string
    pub fn is_string(&self) -> bool {
        matches!(*self, DataType::String(_))
    }

    /// Try getting int value
    pub fn get_int(&self) -> Option<i64> {
        if let DataType::Int(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    /// Try getting float value
    pub fn get_float(&self) -> Option<f64> {
        if let DataType::Float(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    /// Try getting bool value
    pub fn get_bool(&self) -> Option<bool> {
        if let DataType::Bool(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    /// Try getting string value
    pub fn get_string(&self) -> Option<&str> {
        if let DataType::String(v) = self {
            Some(&**v)
        } else {
            None
        }
    }

    /// Try converting data type into a date
    #[cfg(feature = "dates")]
    pub fn as_date(&self) -> Option<chrono::NaiveDate> {
        self.as_datetime().map(|dt| dt.date())
    }

    /// Try converting data type into a time
    #[cfg(feature = "dates")]
    pub fn as_time(&self) -> Option<chrono::NaiveTime> {
        self.as_datetime().map(|dt| dt.time())
    }

    /// Try converting data type into a datetime
    #[cfg(feature = "dates")]
    pub fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        match self {
            DataType::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            DataType::Float(f) | DataType::DateTime(f) => {
                let unix_days = f - 25569.;
                let unix_secs = unix_days * 86400.;
                let secs = unix_secs.trunc() as i64;
                let nsecs = (unix_secs.fract().abs() * 1e9) as u32;
                chrono::NaiveDateTime::from_timestamp_opt(secs, nsecs)
            }
            _ => None,
        }
    }
}

impl PartialEq<str> for DataType {
    fn eq(&self, other: &str) -> bool {
        matches!(*self, DataType::String(ref s) if s == other)
    }
}

impl PartialEq<f64> for DataType {
    fn eq(&self, other: &f64) -> bool {
        matches!(*self, DataType::Float(ref s) if *s == *other)
    }
}

impl PartialEq<bool> for DataType {
    fn eq(&self, other: &bool) -> bool {
        matches!(*self, DataType::Bool(ref s) if *s == *other)
    }
}

impl PartialEq<i64> for DataType {
    fn eq(&self, other: &i64) -> bool {
        matches!(*self, DataType::Int(ref s) if *s == *other)
    }
}

impl fmt::Display for DataType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> std::result::Result<(), fmt::Error> {
        match *self {
            DataType::Int(ref e) => write!(f, "{}", e),
            DataType::Float(ref e) => write!(f, "{}", e),
            DataType::String(ref e) => write!(f, "{}", e),
            DataType::Bool(ref e) => write!(f, "{}", e),
            DataType::DateTime(ref e) => write!(f, "{}", e),
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

            fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
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
    };
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

#[cfg(all(test, feature = "dates"))]
mod tests {
    use super::*;

    #[test]
    fn test_dates() {
        use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

        let unix_epoch = DataType::Float(25569.);
        assert_eq!(
            unix_epoch.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd(1970, 1, 1),
                NaiveTime::from_hms(0, 0, 0)
            ))
        );

        let unix_epoch_15h30m = DataType::Float(25569.645833333333333);
        let chrono_dt = NaiveDateTime::new(
            NaiveDate::from_ymd(1970, 1, 1),
            NaiveTime::from_hms(15, 30, 0),
        );
        let micro = Duration::microseconds(1);
        assert!(unix_epoch_15h30m.as_datetime().unwrap() - chrono_dt < micro);
    }

    #[test]
    fn test_int_dates() {
        use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

        let unix_epoch = DataType::Int(25569);
        assert_eq!(
            unix_epoch.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd(1970, 1, 1),
                NaiveTime::from_hms(0, 0, 0)
            ))
        );

        let time = DataType::Int(44060);
        assert_eq!(
            time.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd(2020, 8, 17),
                NaiveTime::from_hms(0, 0, 0),
            ))
        );
    }
}
