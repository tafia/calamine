use std::fmt;

#[cfg(feature = "dates")]
use once_cell::sync::OnceCell;
use serde::de::Visitor;
use serde::{self, Deserialize};

use super::CellErrorType;

#[cfg(feature = "dates")]
static EXCEL_EPOCH: OnceCell<chrono::NaiveDateTime> = OnceCell::new();

/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq)]
pub enum DataType {
    /// Signed integer
    Int(i64),
    /// Float
    Float(f64),
    /// String
    String(String),
    /// Boolean
    Bool(bool),
    /// Date or Time
    DateTime(f64),
    /// Date, Time or DateTime in ISO 8601
    DateTimeIso(String),
    /// Duration in ISO 8601
    DurationIso(String),
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
        if let DataType::Int(_) = *self {
            true
        } else {
            false
        }
    }
    /// Assess if datatype is a float
    pub fn is_float(&self) -> bool {
        if let DataType::Float(_) = *self {
            true
        } else {
            false
        }
    }
    /// Assess if datatype is a bool
    pub fn is_bool(&self) -> bool {
        if let DataType::Bool(_) = *self {
            true
        } else {
            false
        }
    }
    /// Assess if datatype is a string
    pub fn is_string(&self) -> bool {
        if let DataType::String(_) = *self {
            true
        } else {
            false
        }
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
        use std::str::FromStr;
        match self {
            DataType::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.date())
                .or_else(|| chrono::NaiveDate::from_str(s).ok()),
            _ => self.as_datetime().map(|dt| dt.date()),
        }
    }

    /// Try converting data type into a time
    #[cfg(feature = "dates")]
    pub fn as_time(&self) -> Option<chrono::NaiveTime> {
        use std::str::FromStr;
        match self {
            DataType::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.time())
                .or_else(|| chrono::NaiveTime::from_str(s).ok()),
            DataType::DurationIso(s) => chrono::NaiveTime::parse_from_str(s, "PT%HH%MM%S%.fS").ok(),
            _ => self.as_datetime().map(|dt| dt.time()),
        }
    }

    /// Try converting data type into a datetime
    #[cfg(feature = "dates")]
    pub fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        use std::str::FromStr;
        const MS_MULTIPLIER: f64 = 24f64 * 60f64 * 60f64 * 1e+3f64;

        match self {
            DataType::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            DataType::Float(f) | DataType::DateTime(f) => {
                let excel_epoch = EXCEL_EPOCH
                    .get_or_init(|| chrono::NaiveDate::from_ymd(1899, 12, 30).and_hms(0, 0, 0));
                let ms = f * MS_MULTIPLIER;
                let excel_duration = chrono::Duration::milliseconds(ms.round() as i64);
                excel_epoch.checked_add_signed(excel_duration)
            }
            DataType::DateTimeIso(s) => chrono::NaiveDateTime::from_str(s).ok(),
            _ => None,
        }
    }
}

impl PartialEq<str> for DataType {
    fn eq(&self, other: &str) -> bool {
        match *self {
            DataType::String(ref s) if s == other => true,
            _ => false,
        }
    }
}

impl PartialEq<f64> for DataType {
    fn eq(&self, other: &f64) -> bool {
        match *self {
            DataType::Float(ref s) if *s == *other => true,
            _ => false,
        }
    }
}

impl PartialEq<bool> for DataType {
    fn eq(&self, other: &bool) -> bool {
        match *self {
            DataType::Bool(ref s) if *s == *other => true,
            _ => false,
        }
    }
}

impl PartialEq<i64> for DataType {
    fn eq(&self, other: &i64) -> bool {
        match *self {
            DataType::Int(ref s) if *s == *other => true,
            _ => false,
        }
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
            DataType::DateTimeIso(ref e) => write!(f, "{}", e),
            DataType::DurationIso(ref e) => write!(f, "{}", e),
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

        // test for https://github.com/tafia/calamine/issues/251
        let unix_epoch_precision = DataType::Float(44484.7916666667);
        assert_eq!(
            unix_epoch_precision.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd(2021, 10, 15),
                NaiveTime::from_hms(19, 0, 0)
            ))
        );

        // test rounding
        assert_eq!(
            DataType::Float(0.18737500000000001).as_time(),
            Some(NaiveTime::from_hms_milli_opt(4, 29, 49, 200).unwrap())
        );
        assert_eq!(
            DataType::Float(0.25951736111111101).as_time(),
            Some(NaiveTime::from_hms_milli_opt(6, 13, 42, 300).unwrap())
        );

        // test overflow
        assert_eq!(DataType::Float(1e20).as_time(), None);

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
        use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

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
