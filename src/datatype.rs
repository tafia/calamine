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
pub enum Data {
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
impl DataType for Data {
    fn is_empty(&self) -> bool {
        *self == Data::Empty
    }
    fn is_int(&self) -> bool {
        matches!(*self, Data::Int(_))
    }
    fn is_float(&self) -> bool {
        matches!(*self, Data::Float(_))
    }
    fn is_bool(&self) -> bool {
        matches!(*self, Data::Bool(_))
    }
    fn is_string(&self) -> bool {
        matches!(*self, Data::String(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration(&self) -> bool {
        matches!(*self, Data::Duration(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, Data::DurationIso(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime(&self) -> bool {
        matches!(*self, Data::DateTime(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime_iso(&self) -> bool {
        matches!(*self, Data::DateTimeIso(_))
    }

    fn get_int(&self) -> Option<i64> {
        if let Data::Int(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_float(&self) -> Option<f64> {
        if let Data::Float(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_bool(&self) -> Option<bool> {
        if let Data::Bool(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    fn get_string(&self) -> Option<&str> {
        if let Data::String(v) = self {
            Some(&**v)
        } else {
            None
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            Data::Float(v) => Some(v.to_string()),
            Data::Int(v) => Some(v.to_string()),
            Data::String(v) => Some(v.clone()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            Data::Int(v) => Some(*v),
            Data::Float(v) => Some(*v as i64),
            Data::String(v) => v.parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            Data::Int(v) => Some(*v as f64),
            Data::Float(v) => Some(*v),
            Data::String(v) => v.parse::<f64>().ok(),
            _ => None,
        }
    }
    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        match self {
            Data::DateTimeIso(s) => self
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
            Data::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.time())
                .or_else(|| chrono::NaiveTime::from_str(s).ok()),
            Data::DurationIso(s) => chrono::NaiveTime::parse_from_str(s, "PT%HH%MM%S%.fS").ok(),
            _ => self.as_datetime().map(|dt| dt.time()),
        }
    }

    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        match self {
            Data::Duration(days) => {
                let ms = days * MS_MULTIPLIER;
                Some(chrono::Duration::milliseconds(ms.round() as i64))
            }
            // need replace in the future to smth like chrono::Duration::from_str()
            // https://github.com/chronotope/chrono/issues/579
            Data::DurationIso(_) => self.as_time().map(|t| {
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
            Data::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            Data::Float(f) | Data::DateTime(f) => {
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
            Data::DateTimeIso(s) => chrono::NaiveDateTime::from_str(s).ok(),
            _ => None,
        }
    }
}

impl PartialEq<&str> for Data {
    fn eq(&self, other: &&str) -> bool {
        match *self {
            Data::String(ref s) if s == other => true,
            _ => false,
        }
    }
}

impl PartialEq<str> for Data {
    fn eq(&self, other: &str) -> bool {
        matches!(*self, Data::String(ref s) if s == other)
    }
}

impl PartialEq<f64> for Data {
    fn eq(&self, other: &f64) -> bool {
        matches!(*self, Data::Float(ref s) if *s == *other)
    }
}

impl PartialEq<bool> for Data {
    fn eq(&self, other: &bool) -> bool {
        matches!(*self, Data::Bool(ref s) if *s == *other)
    }
}

impl PartialEq<i64> for Data {
    fn eq(&self, other: &i64) -> bool {
        matches!(*self, Data::Int(ref s) if *s == *other)
    }
}

impl fmt::Display for Data {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> std::result::Result<(), fmt::Error> {
        match *self {
            Data::Int(ref e) => write!(f, "{}", e),
            Data::Float(ref e) => write!(f, "{}", e),
            Data::String(ref e) => write!(f, "{}", e),
            Data::Bool(ref e) => write!(f, "{}", e),
            Data::DateTime(ref e) => write!(f, "{}", e),
            Data::Duration(ref e) => write!(f, "{}", e),
            Data::DateTimeIso(ref e) => write!(f, "{}", e),
            Data::DurationIso(ref e) => write!(f, "{}", e),
            Data::Error(ref e) => write!(f, "{}", e),
            Data::Empty => Ok(()),
        }
    }
}

impl<'de> Deserialize<'de> for Data {
    #[inline]
    fn deserialize<D>(deserializer: D) -> Result<Data, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        struct DataVisitor;

        impl<'de> Visitor<'de> for DataVisitor {
            type Value = Data;

            fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                formatter.write_str("any valid JSON value")
            }

            #[inline]
            fn visit_bool<E>(self, value: bool) -> Result<Data, E> {
                Ok(Data::Bool(value))
            }

            #[inline]
            fn visit_i64<E>(self, value: i64) -> Result<Data, E> {
                Ok(Data::Int(value))
            }

            #[inline]
            fn visit_u64<E>(self, value: u64) -> Result<Data, E> {
                Ok(Data::Int(value as i64))
            }

            #[inline]
            fn visit_f64<E>(self, value: f64) -> Result<Data, E> {
                Ok(Data::Float(value))
            }

            #[inline]
            fn visit_str<E>(self, value: &str) -> Result<Data, E>
            where
                E: serde::de::Error,
            {
                self.visit_string(String::from(value))
            }

            #[inline]
            fn visit_string<E>(self, value: String) -> Result<Data, E> {
                Ok(Data::String(value))
            }

            #[inline]
            fn visit_none<E>(self) -> Result<Data, E> {
                Ok(Data::Empty)
            }

            #[inline]
            fn visit_some<D>(self, deserializer: D) -> Result<Data, D::Error>
            where
                D: serde::Deserializer<'de>,
            {
                Deserialize::deserialize(deserializer)
            }

            #[inline]
            fn visit_unit<E>(self) -> Result<Data, E> {
                Ok(Data::Empty)
            }
        }

        deserializer.deserialize_any(DataVisitor)
    }
}

macro_rules! define_from {
    ($variant:path, $ty:ty) => {
        impl From<$ty> for Data {
            fn from(v: $ty) -> Self {
                $variant(v)
            }
        }
    };
}

define_from!(Data::Int, i64);
define_from!(Data::Float, f64);
define_from!(Data::String, String);
define_from!(Data::Bool, bool);
define_from!(Data::Error, CellErrorType);

impl<'a> From<&'a str> for Data {
    fn from(v: &'a str) -> Self {
        Data::String(String::from(v))
    }
}

impl From<()> for Data {
    fn from(_: ()) -> Self {
        Data::Empty
    }
}

impl<T> From<Option<T>> for Data
where
    Data: From<T>,
{
    fn from(v: Option<T>) -> Self {
        match v {
            Some(v) => From::from(v),
            None => Data::Empty,
        }
    }
}

/// An enum to represent all different data types that can appear as
/// a value in a worksheet cell
#[derive(Debug, Clone, PartialEq, Default)]
pub enum DataRef<'a> {
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

impl DataType for DataRef<'_> {
    fn is_empty(&self) -> bool {
        *self == DataRef::Empty
    }

    fn is_int(&self) -> bool {
        matches!(*self, DataRef::Int(_))
    }

    fn is_float(&self) -> bool {
        matches!(*self, DataRef::Float(_))
    }

    fn is_bool(&self) -> bool {
        matches!(*self, DataRef::Bool(_))
    }

    fn is_string(&self) -> bool {
        matches!(*self, DataRef::String(_) | DataRef::SharedString(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration(&self) -> bool {
        matches!(*self, DataRef::Duration(_))
    }

    #[cfg(feature = "dates")]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, DataRef::DurationIso(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime(&self) -> bool {
        matches!(*self, DataRef::DateTime(_))
    }

    #[cfg(feature = "dates")]
    fn is_datetime_iso(&self) -> bool {
        matches!(*self, DataRef::DateTimeIso(_))
    }

    fn get_int(&self) -> Option<i64> {
        if let DataRef::Int(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_float(&self) -> Option<f64> {
        if let DataRef::Float(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_bool(&self) -> Option<bool> {
        if let DataRef::Bool(v) = self {
            Some(*v)
        } else {
            None
        }
    }

    fn get_string(&self) -> Option<&str> {
        match self {
            DataRef::String(v) => Some(&**v),
            DataRef::SharedString(v) => Some(v),
            _ => None,
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            DataRef::Float(v) => Some(v.to_string()),
            DataRef::Int(v) => Some(v.to_string()),
            DataRef::String(v) => Some(v.clone()),
            DataRef::SharedString(v) => Some(v.to_string()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            DataRef::Int(v) => Some(*v),
            DataRef::Float(v) => Some(*v as i64),
            DataRef::String(v) => v.parse::<i64>().ok(),
            DataRef::SharedString(v) => v.parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            DataRef::Int(v) => Some(*v as f64),
            DataRef::Float(v) => Some(*v),
            DataRef::String(v) => v.parse::<f64>().ok(),
            DataRef::SharedString(v) => v.parse::<f64>().ok(),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        match self {
            DataRef::DateTimeIso(s) => self
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
            DataRef::DateTimeIso(s) => self
                .as_datetime()
                .map(|dt| dt.time())
                .or_else(|| chrono::NaiveTime::from_str(s).ok()),
            DataRef::DurationIso(s) => chrono::NaiveTime::parse_from_str(s, "PT%HH%MM%S%.fS").ok(),
            _ => self.as_datetime().map(|dt| dt.time()),
        }
    }

    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        match self {
            DataRef::Duration(days) => {
                let ms = days * MS_MULTIPLIER;
                Some(chrono::Duration::milliseconds(ms.round() as i64))
            }
            // need replace in the future to smth like chrono::Duration::from_str()
            // https://github.com/chronotope/chrono/issues/579
            DataRef::DurationIso(_) => self.as_time().map(|t| {
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
            DataRef::Int(x) => {
                let days = x - 25569;
                let secs = days * 86400;
                chrono::NaiveDateTime::from_timestamp_opt(secs, 0)
            }
            DataRef::Float(f) | DataRef::DateTime(f) => {
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
            DataRef::DateTimeIso(s) => chrono::NaiveDateTime::from_str(s).ok(),
            _ => None,
        }
    }
}

/// A trait to represent all different data types that can appear as
/// a value in a worksheet cell
pub trait DataType {
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

impl<'a> From<DataRef<'a>> for Data {
    fn from(value: DataRef<'a>) -> Self {
        match value {
            DataRef::Int(v) => Data::Int(v),
            DataRef::Float(v) => Data::Float(v),
            DataRef::String(v) => Data::String(v),
            DataRef::SharedString(v) => Data::String(v.into()),
            DataRef::Bool(v) => Data::Bool(v),
            DataRef::DateTime(v) => Data::DateTime(v),
            DataRef::Duration(v) => Data::Duration(v),
            DataRef::DateTimeIso(v) => Data::DateTimeIso(v),
            DataRef::DurationIso(v) => Data::DurationIso(v),
            DataRef::Error(v) => Data::Error(v),
            DataRef::Empty => Data::Empty,
        }
    }
}

#[cfg(all(test, feature = "dates"))]
mod date_tests {
    use super::*;

    #[test]
    fn test_dates() {
        use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

        let unix_epoch = Data::Float(25569.);
        assert_eq!(
            unix_epoch.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(1970, 1, 1).unwrap(),
                NaiveTime::from_hms_opt(0, 0, 0).unwrap(),
            ))
        );

        // test for https://github.com/tafia/calamine/issues/251
        let unix_epoch_precision = Data::Float(44484.7916666667);
        assert_eq!(
            unix_epoch_precision.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(2021, 10, 15).unwrap(),
                NaiveTime::from_hms_opt(19, 0, 0).unwrap(),
            ))
        );

        // test rounding
        assert_eq!(
            Data::Float(0.18737500000000001).as_time(),
            Some(NaiveTime::from_hms_milli_opt(4, 29, 49, 200).unwrap())
        );
        assert_eq!(
            Data::Float(0.25951736111111101).as_time(),
            Some(NaiveTime::from_hms_milli_opt(6, 13, 42, 300).unwrap())
        );

        // test overflow
        assert_eq!(Data::Float(1e20).as_time(), None);

        let unix_epoch_15h30m = Data::Float(25569.645833333333333);
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

        let unix_epoch = Data::Int(25569);
        assert_eq!(
            unix_epoch.as_datetime(),
            Some(NaiveDateTime::new(
                NaiveDate::from_ymd_opt(1970, 1, 1).unwrap(),
                NaiveTime::from_hms_opt(0, 0, 0).unwrap(),
            ))
        );

        let time = Data::Int(44060);
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
        assert_eq!(Data::String("value".to_string()), "value");
        assert_eq!(Data::String("value".to_string()), "value"[..]);
        assert_eq!(Data::Float(100.0), 100.0f64);
        assert_eq!(Data::Bool(true), true);
        assert_eq!(Data::Int(100), 100i64);
    }
}
