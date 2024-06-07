use std::fmt;
#[cfg(feature = "dates")]
use std::sync::OnceLock;

use serde::de::Visitor;
use serde::{self, Deserialize};

use super::CellErrorType;
use crate::style::RichText;

#[cfg(feature = "dates")]
static EXCEL_EPOCH: OnceLock<chrono::NaiveDateTime> = OnceLock::new();

#[cfg(feature = "dates")]
/// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
const EXCEL_1900_1904_DIFF: f64 = 1462.;

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
    /// Rich (formatted) text
    RichText(RichText),
    /// Boolean
    Bool(bool),
    /// Date or Time
    DateTime(ExcelDateTime),
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
        matches!(*self, Data::String(_) | Data::RichText(_))
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

    fn is_error(&self) -> bool {
        matches!(*self, Data::Error(_))
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
        match self {
            Data::String(v) => Some(v),
            Data::RichText(v) => Some(v.text()),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_datetime(&self) -> Option<ExcelDateTime> {
        match self {
            Data::DateTime(v) => Some(*v),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_datetime_iso(&self) -> Option<&str> {
        match self {
            Data::DateTimeIso(v) => Some(&**v),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_duration_iso(&self) -> Option<&str> {
        match self {
            Data::DurationIso(v) => Some(&**v),
            _ => None,
        }
    }

    fn get_error(&self) -> Option<&CellErrorType> {
        match self {
            Data::Error(e) => Some(e),
            _ => None,
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            Data::Float(v) => Some(v.to_string()),
            Data::Int(v) => Some(v.to_string()),
            Data::String(v) => Some(v.clone()),
            Data::RichText(v) => Some(v.text().clone()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            Data::Int(v) => Some(*v),
            Data::Float(v) => Some(*v as i64),
            Data::Bool(v) => Some(*v as i64),
            Data::String(v) => v.parse::<i64>().ok(),
            Data::RichText(v) => v.text().parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            Data::Int(v) => Some(*v as f64),
            Data::Float(v) => Some(*v),
            Data::Bool(v) => Some((*v as i32).into()),
            Data::String(v) => v.parse::<f64>().ok(),
            Data::RichText(v) => v.text().parse::<f64>().ok(),
            _ => None,
        }
    }
}

impl PartialEq<&str> for Data {
    fn eq(&self, other: &&str) -> bool {
        match *self {
            Data::String(ref s) if s == other => true,
            Data::RichText(ref s) if s.text() == other => true,
            _ => false,
        }
    }
}

impl PartialEq<str> for Data {
    fn eq(&self, other: &str) -> bool {
        match self {
            Data::String(s) if s == other => true,
            Data::RichText(s) if s.text() == other => true,
            _ => false,
        }
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
            Data::RichText(ref e) => write!(f, "{}", e.text()),
            Data::Bool(ref e) => write!(f, "{}", e),
            Data::DateTime(ref e) => write!(f, "{}", e),
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
define_from!(Data::RichText, RichText);
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
    /// Rich text
    RichText(RichText),
    /// Shared String
    SharedString(&'a RichText),
    /// Boolean
    Bool(bool),
    /// Date or Time
    DateTime(ExcelDateTime),
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
        matches!(
            *self,
            DataRef::String(_) | DataRef::RichText(_) | DataRef::SharedString(_)
        )
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

    fn is_error(&self) -> bool {
        matches!(*self, DataRef::Error(_))
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
            DataRef::String(v) => Some(v),
            DataRef::RichText(v) => Some(v.text()),
            DataRef::SharedString(v) => Some(v.text()),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_datetime(&self) -> Option<ExcelDateTime> {
        match self {
            DataRef::DateTime(v) => Some(*v),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_datetime_iso(&self) -> Option<&str> {
        match self {
            DataRef::DateTimeIso(v) => Some(&**v),
            _ => None,
        }
    }

    #[cfg(feature = "dates")]
    fn get_duration_iso(&self) -> Option<&str> {
        match self {
            DataRef::DurationIso(v) => Some(&**v),
            _ => None,
        }
    }

    fn get_error(&self) -> Option<&CellErrorType> {
        match self {
            DataRef::Error(e) => Some(e),
            _ => None,
        }
    }

    fn as_string(&self) -> Option<String> {
        match self {
            DataRef::Float(v) => Some(v.to_string()),
            DataRef::Int(v) => Some(v.to_string()),
            DataRef::String(v) => Some(v.clone()),
            DataRef::RichText(v) => Some(v.text().clone()),
            DataRef::SharedString(v) => Some(v.text().clone()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            DataRef::Int(v) => Some(*v),
            DataRef::Float(v) => Some(*v as i64),
            DataRef::Bool(v) => Some(*v as i64),
            DataRef::String(v) => v.parse::<i64>().ok(),
            DataRef::RichText(v) => v.text().parse::<i64>().ok(),
            DataRef::SharedString(v) => v.text().parse::<i64>().ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            DataRef::Int(v) => Some(*v as f64),
            DataRef::Float(v) => Some(*v),
            DataRef::Bool(v) => Some((*v as i32).into()),
            DataRef::String(v) => v.parse::<f64>().ok(),
            DataRef::RichText(v) => v.text().parse::<f64>().ok(),
            DataRef::SharedString(v) => v.text().parse::<f64>().ok(),
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

    /// Assess if datatype is a CellErrorType
    fn is_error(&self) -> bool;

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

    /// Try getting datetime value
    #[cfg(feature = "dates")]
    fn get_datetime(&self) -> Option<ExcelDateTime>;

    /// Try getting datetime ISO8601 value
    #[cfg(feature = "dates")]
    fn get_datetime_iso(&self) -> Option<&str>;

    /// Try getting duration ISO8601 value
    #[cfg(feature = "dates")]
    fn get_duration_iso(&self) -> Option<&str>;

    /// Try getting Error value
    fn get_error(&self) -> Option<&CellErrorType>;

    /// Try converting data type into a string
    fn as_string(&self) -> Option<String>;

    /// Try converting data type into an int
    fn as_i64(&self) -> Option<i64>;

    /// Try converting data type into a float
    fn as_f64(&self) -> Option<f64>;

    /// Try converting data type into a date
    #[cfg(feature = "dates")]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        if self.is_datetime_iso() {
            self.as_datetime().map(|dt| dt.date()).or_else(|| {
                self.get_datetime_iso()
                    .map(|s| chrono::NaiveDate::from_str(&s).ok())
                    .flatten()
            })
        } else {
            self.as_datetime().map(|dt| dt.date())
        }
    }

    /// Try converting data type into a time
    #[cfg(feature = "dates")]
    fn as_time(&self) -> Option<chrono::NaiveTime> {
        use std::str::FromStr;
        if self.is_datetime_iso() {
            self.as_datetime().map(|dt| dt.time()).or_else(|| {
                self.get_datetime_iso()
                    .map(|s| chrono::NaiveTime::from_str(&s).ok())
                    .flatten()
            })
        } else if self.is_duration_iso() {
            self.get_duration_iso()
                .map(|s| chrono::NaiveTime::parse_from_str(&s, "PT%HH%MM%S%.fS").ok())
                .flatten()
        } else {
            self.as_datetime().map(|dt| dt.time())
        }
    }

    /// Try converting data type into a duration
    #[cfg(feature = "dates")]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        if self.is_datetime() {
            self.get_datetime().map(|dt| dt.as_duration()).flatten()
        } else if self.is_duration_iso() {
            // need replace in the future to smth like chrono::Duration::from_str()
            // https://github.com/chronotope/chrono/issues/579
            self.as_time().map(|t| {
                chrono::Duration::nanoseconds(
                    t.num_seconds_from_midnight() as i64 * 1_000_000_000 + t.nanosecond() as i64,
                )
            })
        } else {
            None
        }
    }

    /// Try converting data type into a datetime
    #[cfg(feature = "dates")]
    fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        use std::str::FromStr;

        if self.is_int() || self.is_float() {
            self.as_f64()
                .map(|f| ExcelDateTime::from_value_only(f).as_datetime())
        } else if self.is_datetime() {
            self.get_datetime().map(|d| d.as_datetime())
        } else if self.is_datetime_iso() {
            self.get_datetime_iso()
                .map(|s| chrono::NaiveDateTime::from_str(&s).ok())
        } else {
            None
        }
        .flatten()
    }
}

impl<'a> From<DataRef<'a>> for Data {
    fn from(value: DataRef<'a>) -> Self {
        match value {
            DataRef::Int(v) => Data::Int(v),
            DataRef::Float(v) => Data::Float(v),
            DataRef::String(v) => Data::String(v),
            DataRef::RichText(v) => Data::RichText(v),
            DataRef::SharedString(v) => {
                if v.is_plain() {
                    Data::String(v.text().clone())
                } else {
                    Data::RichText(v.clone())
                }
            }
            DataRef::Bool(v) => Data::Bool(v),
            DataRef::DateTime(v) => Data::DateTime(v),
            DataRef::DateTimeIso(v) => Data::DateTimeIso(v),
            DataRef::DurationIso(v) => Data::DurationIso(v),
            DataRef::Error(v) => Data::Error(v),
            DataRef::Empty => Data::Empty,
        }
    }
}

/// Excel datetime type. Possible: date, time, datetime, duration.
/// At this time we can only determine datetime (date and time are datetime too) and duration.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ExcelDateTimeType {
    /// DateTime
    DateTime,
    /// TimeDelta (Duration)
    TimeDelta,
}

/// Structure for Excel date and time representation.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ExcelDateTime {
    value: f64,
    datetime_type: ExcelDateTimeType,
    is_1904: bool,
}

impl ExcelDateTime {
    /// Creates a new `ExcelDateTime`
    pub fn new(value: f64, datetime_type: ExcelDateTimeType, is_1904: bool) -> Self {
        ExcelDateTime {
            value,
            datetime_type,
            is_1904,
        }
    }

    /// Is used only for converting excel value to chrono
    #[cfg(feature = "dates")]
    fn from_value_only(value: f64) -> Self {
        ExcelDateTime {
            value,
            ..Default::default()
        }
    }

    /// True if excel datetime has duration format ([hh]:mm:ss, for example)
    #[cfg(feature = "dates")]
    pub fn is_duration(&self) -> bool {
        matches!(self.datetime_type, ExcelDateTimeType::TimeDelta)
    }

    /// True if excel datetime has datetime format (not duration)
    #[cfg(feature = "dates")]
    pub fn is_datetime(&self) -> bool {
        matches!(self.datetime_type, ExcelDateTimeType::DateTime)
    }

    /// Converting data type into a float
    pub fn as_f64(&self) -> f64 {
        self.value
    }

    /// Try converting data type into a duration
    #[cfg(feature = "dates")]
    pub fn as_duration(&self) -> Option<chrono::Duration> {
        let ms = self.value * MS_MULTIPLIER;
        Some(chrono::Duration::milliseconds(ms.round() as i64))
    }

    /// Try converting data type into a datetime
    #[cfg(feature = "dates")]
    pub fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        let excel_epoch = EXCEL_EPOCH.get_or_init(|| {
            chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .unwrap()
                .and_time(chrono::NaiveTime::MIN)
        });
        let f = if self.is_1904 {
            self.value + EXCEL_1900_1904_DIFF
        } else {
            self.value
        };
        let f = if f >= 60.0 { f } else { f + 1.0 };
        let ms = f * MS_MULTIPLIER;
        let excel_duration = chrono::Duration::milliseconds(ms.round() as i64);
        excel_epoch.checked_add_signed(excel_duration)
    }
}

impl Default for ExcelDateTime {
    fn default() -> Self {
        ExcelDateTime {
            value: 0.,
            datetime_type: ExcelDateTimeType::DateTime,
            is_1904: false,
        }
    }
}

impl fmt::Display for ExcelDateTime {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> std::result::Result<(), fmt::Error> {
        write!(f, "{}", self.value)
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

    #[test]
    fn test_as_i64_with_bools() {
        assert_eq!(Data::Bool(true).as_i64(), Some(1));
        assert_eq!(Data::Bool(false).as_i64(), Some(0));
        assert_eq!(DataRef::Bool(true).as_i64(), Some(1));
        assert_eq!(DataRef::Bool(false).as_i64(), Some(0));
    }

    #[test]
    fn test_as_f64_with_bools() {
        assert_eq!(Data::Bool(true).as_f64(), Some(1.0));
        assert_eq!(Data::Bool(false).as_f64(), Some(0.0));
        assert_eq!(DataRef::Bool(true).as_f64(), Some(1.0));
        assert_eq!(DataRef::Bool(false).as_f64(), Some(0.0));
    }
}
