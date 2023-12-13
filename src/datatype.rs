use std::fmt;

#[cfg(feature = "dates")]
use once_cell::sync::OnceCell;
use serde::de::Visitor;
use serde::{self, Deserialize};

use super::CellErrorType;

#[cfg(feature = "dates")]
static EXCEL_EPOCH: OnceCell<chrono::NaiveDateTime> = OnceCell::new();

#[cfg(feature = "dates")]
const MS_MULTIPLIER: f64 = 24f64 * 60f64 * 60f64 * 1e+3f64;

/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq, Default)]
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
    /// Duration
    Duration(f64),
    /// Date, Time or DateTime in ISO 8601
    DateTimeIso(String),
    /// Duration in ISO 8601
    DurationIso(String),
    /// Error
    Error(CellErrorType),
    /// Empty cell
    #[default]
    Empty,
}

/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
impl DataTypeTrait for DataType {
    fn is_empty(&self) -> bool {
        *self == DataType::Empty
    }
    fn is_int(&self) -> bool {
        matches!(*self, DataType::Int(_))
    }
    fn is_float(&self) -> bool {
        matches!(*self, DataType::Float(_))
    }
    fn is_bool(&self) -> bool {
        matches!(*self, DataType::Bool(_))
    }
    fn is_string(&self) -> bool {
        matches!(*self, DataType::String(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration(&self) -> bool {
        matches!(*self, DataType::Duration(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, DataType::DurationIso(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime(&self) -> bool {
        matches!(*self, DataType::DateTime(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime_iso(&self) -> bool {
        matches!(*self, DataType::DateTimeIso(_))
    }

    fn get_int(&self) -> Option<i64> {
        if let DataType::Int(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_float(&self) -> Option<f64> {
        if let DataType::Float(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_bool(&self) -> Option<bool> {
        if let DataType::Bool(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_string(&self) -> Option<&str> {
        if let DataType::String(v) = self {
            Some(&**v)
        } else {
            None
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            DataType::Float(v) => Some(v.to_string()),
            DataType::Int(v) => Some(v.to_string()),
            DataType::String(v) => Some(v.clone()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            DataType::Int(v) => Some(*v),
            DataType::Float(v) => Some(*v as i64),
            DataType::String(v) => v.parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            DataType::Int(v) => Some(*v as f64),
            DataType::Float(v) => Some(*v),
            DataType::String(v) => v.parse::<f64>().ok(),
            _ => None,
        }
    }
    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        match self {
            DataType::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.date())
                .or_else(|| chrono::NaiveDate::from_str(s).ok()),
            _ => self.as_datetime().map(|dt| dt.date()),
        }
    }

    #[cfg(feature = "dates")]
    fn as_time(&self) -> Option<chrono::NaiveTime> {
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

    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        match self {
            DataType::Duration(days) => {
                let ms = days * MS_MULTIPLIER;
                Some(chrono::Duration::milliseconds(ms.round() as i64))
            }
            // need replace in the future to smth like chrono::Duration::from_str()
            // https://github.com/chronotope/chrono/issues/579
            DataType::DurationIso(_) => self.as_time().map(|t| {
                chrono::Duration::nanoseconds(
                    t.num_seconds_from_midnight() as i64 * 1_000_000_000 + t.nanosecond() as i64,
                )
            }),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        use std::str::FromStr;

        match self {
            DataType::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            DataType::Float(f) | DataType::DateTime(f) => {
                let excel_epoch = EXCEL_EPOCH.get_or_init(|| {
                    chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                        .unwrap()
                        .and_hms_opt(0, 0, 0)
                        .unwrap()
                });
                let f = if *f >= 60.0 { *f } else { *f + 1.0 };
                let ms = f * MS_MULTIPLIER;
                let excel_duration = chrono::Duration::milliseconds(ms.round() as i64);
                excel_epoch.checked_add_signed(excel_duration)
            }
            DataType::DateTimeIso(s) => chrono::NaiveDateTime::from_str(s).ok(),
            _ => None,
        }
    }
}

impl PartialEq<&str> for DataType {
    fn eq(&self, other: &&str) -> bool {
        match *self {
            DataType::String(ref s) if s == other => true,
            _ => false,
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
            DataType::Duration(ref e) => write!(f, "{}", e),
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

/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq, Default)]
pub enum DataTypeRef<'a> {
    /// Signed integer
    Int(i64),
    /// Float
    Float(f64),
    /// String
    String(String),
    /// Shared String
    SharedString(&'a str),
    /// Boolean
    Bool(bool),
    /// Date or Time
    DateTime(f64),
    /// Duration
    Duration(f64),
    /// Date, Time or DateTime in ISO 8601
    DateTimeIso(String),
    /// Duration in ISO 8601
    DurationIso(String),
    /// Error
    Error(CellErrorType),
    /// Empty cell
    #[default]
    Empty,
}

impl DataTypeTrait for DataTypeRef<'_> {
    fn is_empty(&self) -> bool {
        *self == DataTypeRef::Empty
    }

    fn is_int(&self) -> bool {
        matches!(*self, DataTypeRef::Int(_))
    }

    fn is_float(&self) -> bool {
        matches!(*self, DataTypeRef::Float(_))
    }

    fn is_bool(&self) -> bool {
        matches!(*self, DataTypeRef::Bool(_))
    }

    fn is_string(&self) -> bool {
        matches!(*self, DataTypeRef::String(_) | DataTypeRef::SharedString(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration(&self) -> bool {
        matches!(*self, DataTypeRef::Duration(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, DataTypeRef::DurationIso(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime(&self) -> bool {
        matches!(*self, DataTypeRef::DateTime(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime_iso(&self) -> bool {
        matches!(*self, DataTypeRef::DateTimeIso(_))
    }

    fn get_int(&self) -> Option<i64> {
        if let DataTypeRef::Int(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_float(&self) -> Option<f64> {
        if let DataTypeRef::Float(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_bool(&self) -> Option<bool> {
        if let DataTypeRef::Bool(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_string(&self) -> Option<&str> {
        match self {
            DataTypeRef::String(v) => Some(&**v),
            DataTypeRef::SharedString(v) => Some(v),
            _ => None,
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            DataTypeRef::Float(v) => Some(v.to_string()),
            DataTypeRef::Int(v) => Some(v.to_string()),
            DataTypeRef::String(v) => Some(v.clone()),
            DataTypeRef::SharedString(v) => Some(v.to_string()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            DataTypeRef::Int(v) => Some(*v),
            DataTypeRef::Float(v) => Some(*v as i64),
            DataTypeRef::String(v) => v.parse::<i64>().ok(),
            DataTypeRef::SharedString(v) => v.parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            DataTypeRef::Int(v) => Some(*v as f64),
            DataTypeRef::Float(v) => Some(*v),
            DataTypeRef::String(v) => v.parse::<f64>().ok(),
            DataTypeRef::SharedString(v) => v.parse::<f64>().ok(),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        match self {
            DataTypeRef::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.date())
                .or_else(|| chrono::NaiveDate::from_str(s).ok()),
            _ => self.as_datetime().map(|dt| dt.date()),
        }
    }

    #[cfg(feature = "dates")]
    fn as_time(&self) -> Option<chrono::NaiveTime> {
        use std::str::FromStr;
        match self {
            DataTypeRef::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.time())
                .or_else(|| chrono::NaiveTime::from_str(s).ok()),
            DataTypeRef::DurationIso(s) => {
                chrono::NaiveTime::parse_from_str(s, "PT%HH%MM%S%.fS").ok()
            }
            _ => self.as_datetime().map(|dt| dt.time()),
        }
    }

    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        match self {
            DataTypeRef::Duration(days) => {
                let ms = days * MS_MULTIPLIER;
                Some(chrono::Duration::milliseconds(ms.round() as i64))
            }
            // need replace in the future to smth like chrono::Duration::from_str()
            // https://github.com/chronotope/chrono/issues/579
            DataTypeRef::DurationIso(_) => self.as_time().map(|t| {
                chrono::Duration::nanoseconds(
                    t.num_seconds_from_midnight() as i64 * 1_000_000_000 + t.nanosecond() as i64,
                )
            }),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        use std::str::FromStr;

        match self {
            DataTypeRef::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            DataTypeRef::Float(f) | DataTypeRef::DateTime(f) => {
                let excel_epoch = EXCEL_EPOCH.get_or_init(|| {
                    chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                        .unwrap()
                        .and_hms_opt(0, 0, 0)
                        .unwrap()
                });
                let f = if *f >= 60.0 { *f } else { *f + 1.0 };
                let ms = f * MS_MULTIPLIER;
                let excel_duration = chrono::Duration::milliseconds(ms.round() as i64);
                excel_epoch.checked_add_signed(excel_duration)
            }
            DataTypeRef::DateTimeIso(s) => chrono::NaiveDateTime::from_str(s).ok(),
            _ => None,
        }
    }
}

/// A trait to represent all different data types that can appear as
/// a value in a worksheet cell
pub trait DataTypeTrait {
    /// Assess if datatype is empty
    fn is_empty(&self) -> bool;

    /// Assess if datatype is a int
    fn is_int(&self) -> bool;

    /// Assess if datatype is a float
    fn is_float(&self) -> bool;

    /// Assess if datatype is a bool
    fn is_bool(&self) -> bool;

    /// Assess if datatype is a string
    fn is_string(&self) -> bool;

    /// Assess if datatype is a duration
    #[cfg(feature = "dates")]
    fn is_duration(&self) -> bool;

    /// Assess if datatype is an ISO8601 duration
    #[cfg(feature = "dates")]
    fn is_duration_iso(&self) -> bool;

    /// Assess if datatype is a datetime
    #[cfg(feature = "dates")]
    fn is_datetime(&self) -> bool;

    /// Assess if datatype is an ISO8601 datetime
    #[cfg(feature = "dates")]
    fn is_datetime_iso(&self) -> bool;

    /// Try getting int value
    fn get_int(&self) -> Option<i64>;

    /// Try getting float value
    fn get_float(&self) -> Option<f64>;

    /// Try getting bool value
    fn get_bool(&self) -> Option<bool>;

    /// Try getting string value
    fn get_string(&self) -> Option<&str>;

    /// Try converting data type into a string
    fn as_string(&self) -> Option<String>;

    /// Try converting data type into an int
    fn as_i64(&self) -> Option<i64>;

    /// Try converting data type into a float
    fn as_f64(&self) -> Option<f64>;

    /// Try converting data type into a date
    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate>;

    /// Try converting data type into a time
    #[cfg(feature = "dates")]
    fn as_time(&self) -> Option<chrono::NaiveTime>;

    /// Try converting data type into a duration
    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration>;

    /// Try converting data type into a datetime
    #[cfg(feature = "dates")]
    fn as_datetime(&self) -> Option<chrono::NaiveDateTime>;
}

impl<'a> From<DataTypeRef<'a>> for DataType {
    fn from(value: DataTypeRef<'a>) -> Self {
        match value {
            DataTypeRef::Int(v) => DataType::Int(v),
            DataTypeRef::Float(v) => DataType::Float(v),
            DataTypeRef::String(v) => DataType::String(v),
            DataTypeRef::SharedString(v) => DataType::String(v.into()),
            DataTypeRef::Bool(v) => DataType::Bool(v),
            DataTypeRef::DateTime(v) => DataType::DateTime(v),
            DataTypeRef::Duration(v) => DataType::Duration(v),
            DataTypeRef::DateTimeIso(v) => DataType::DateTimeIso(v),
            DataTypeRef::DurationIso(v) => DataType::DurationIso(v),
            DataTypeRef::Error(v) => DataType::Error(v),
            DataTypeRef::Empty => DataType::Empty,
        }
    }
}

#[cfg(all(test, feature = "dates"))]
mod date_tests {
    use super::*;

    #[test]
    fn test_dates() {
        use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

        let unix_epoch = DataType::Float(25569.);
        assert_eq!(
            unix_epoch.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(1970, 1, 1).unwrap(),
                NaiveTime::from_hms_opt(0, 0, 0).unwrap(),
            ))
        );

        // test for https://github.com/tafia/calamine/issues/251
        let unix_epoch_precision = DataType::Float(44484.7916666667);
        assert_eq!(
            unix_epoch_precision.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(2021, 10, 15).unwrap(),
                NaiveTime::from_hms_opt(19, 0, 0).unwrap(),
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
            NaiveDate::from_ymd_opt(1970, 1, 1).unwrap(),
            NaiveTime::from_hms_opt(15, 30, 0).unwrap(),
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
                NaiveDate::from_ymd_opt(1970, 1, 1).unwrap(),
                NaiveTime::from_hms_opt(0, 0, 0).unwrap(),
            ))
        );

        let time = DataType::Int(44060);
        assert_eq!(
            time.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(2020, 8, 17).unwrap(),
                NaiveTime::from_hms_opt(0, 0, 0).unwrap(),
            ))
        );
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_partial_eq() {
        assert_eq!(DataType::String("value".to_string()), "value");
        assert_eq!(DataType::String("value".to_string()), "value"[..]);
        assert_eq!(DataType::Float(100.0), 100.0f64);
        assert_eq!(DataType::Bool(true), true);
        assert_eq!(DataType::Int(100), 100i64);
    }
}
