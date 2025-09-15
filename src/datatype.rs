// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use std::fmt;
#[cfg(feature = "chrono")]
use std::sync::OnceLock;

use serde::de::Visitor;
use serde::Deserialize;

use super::CellErrorType;

// Constants used in Excel date calculations.
const DAY_SECONDS: f64 = 24.0 * 60.0 * 60.;
const HOUR_SECONDS: u64 = 60 * 60;
const MINUTE_SECONDS: u64 = 60;
const YEAR_DAYS: u64 = 365;
const YEAR_DAYS_4: u64 = YEAR_DAYS * 4 + 1;
const YEAR_DAYS_100: u64 = YEAR_DAYS * 100 + 25;
const YEAR_DAYS_400: u64 = YEAR_DAYS * 400 + 97;

#[cfg(feature = "chrono")]
static EXCEL_EPOCH: OnceLock<chrono::NaiveDateTime> = OnceLock::new();

#[cfg(feature = "chrono")]
// https://learn.microsoft.com/en-us/office/troubleshoot/excel/1900-and-1904-date-system
const EXCEL_1900_1904_DIFF: f64 = 1462.;

#[cfg(feature = "chrono")]
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
    DateTime(ExcelDateTime),
    /// Date, Time or Date/Time in ISO 8601
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

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, Data::DurationIso(_))
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_datetime(&self) -> bool {
        matches!(*self, Data::DateTime(_))
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
        if let Data::String(v) = self {
            Some(&**v)
        } else {
            None
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime(&self) -> Option<ExcelDateTime> {
        match self {
            Data::DateTime(v) => Some(*v),
            _ => None,
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime_iso(&self) -> Option<&str> {
        match self {
            Data::DateTimeIso(v) => Some(&**v),
            _ => None,
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            Data::Int(v) => Some(*v),
            Data::Float(v) => Some(*v as i64),
            Data::Bool(v) => Some(*v as i64),
            Data::String(v) => atoi_simd::parse::<i64>(v.as_bytes()).ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            Data::Int(v) => Some(*v as f64),
            Data::Float(v) => Some(*v),
            Data::Bool(v) => Some((*v as i32).into()),
            Data::String(v) => fast_float2::parse(v).ok(),
            _ => None,
        }
    }
}

impl PartialEq<&str> for Data {
    fn eq(&self, other: &&str) -> bool {
        matches!(*self, Data::String(ref s) if s == other)
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
            Data::Int(ref e) => write!(f, "{e}"),
            Data::Float(ref e) => write!(f, "{e}"),
            Data::String(ref e) => write!(f, "{e}"),
            Data::Bool(ref e) => write!(f, "{e}"),
            Data::DateTime(ref e) => write!(f, "{e}"),
            Data::DateTimeIso(ref e) => write!(f, "{e}"),
            Data::DurationIso(ref e) => write!(f, "{e}"),
            Data::Error(ref e) => write!(f, "{e}"),
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
    DateTime(ExcelDateTime),
    /// Date, Time or Date/Time in ISO 8601
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

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_duration_iso(&self) -> bool {
        matches!(*self, DataRef::DurationIso(_))
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_datetime(&self) -> bool {
        matches!(*self, DataRef::DateTime(_))
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
            DataRef::String(v) => Some(&**v),
            DataRef::SharedString(v) => Some(v),
            _ => None,
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime(&self) -> Option<ExcelDateTime> {
        match self {
            DataRef::DateTime(v) => Some(*v),
            _ => None,
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime_iso(&self) -> Option<&str> {
        match self {
            DataRef::DateTimeIso(v) => Some(&**v),
            _ => None,
        }
    }

    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
            DataRef::SharedString(v) => Some(v.to_string()),
            _ => None,
        }
    }

    fn as_i64(&self) -> Option<i64> {
        match self {
            DataRef::Int(v) => Some(*v),
            DataRef::Float(v) => Some(*v as i64),
            DataRef::Bool(v) => Some(*v as i64),
            DataRef::String(v) => atoi_simd::parse::<i64>(v.as_bytes()).ok(),
            DataRef::SharedString(v) => atoi_simd::parse::<i64>(v.as_bytes()).ok(),
            _ => None,
        }
    }

    fn as_f64(&self) -> Option<f64> {
        match self {
            DataRef::Int(v) => Some(*v as f64),
            DataRef::Float(v) => Some(*v),
            DataRef::Bool(v) => Some((*v as i32).into()),
            DataRef::String(v) => fast_float2::parse(v).ok(),
            DataRef::SharedString(v) => fast_float2::parse(v).ok(),
            _ => None,
        }
    }
}

impl PartialEq<&str> for DataRef<'_> {
    fn eq(&self, other: &&str) -> bool {
        matches!(*self, DataRef::String(ref s) if s == other)
    }
}

impl PartialEq<str> for DataRef<'_> {
    fn eq(&self, other: &str) -> bool {
        matches!(*self, DataRef::String(ref s) if s == other)
    }
}

impl PartialEq<f64> for DataRef<'_> {
    fn eq(&self, other: &f64) -> bool {
        matches!(*self, DataRef::Float(ref s) if *s == *other)
    }
}

impl PartialEq<bool> for DataRef<'_> {
    fn eq(&self, other: &bool) -> bool {
        matches!(*self, DataRef::Bool(ref s) if *s == *other)
    }
}

impl PartialEq<i64> for DataRef<'_> {
    fn eq(&self, other: &i64) -> bool {
        matches!(*self, DataRef::Int(ref s) if *s == *other)
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

    /// Assess if datatype is a `CellErrorType`
    fn is_error(&self) -> bool;

    /// Assess if datatype is an ISO8601 duration
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_duration_iso(&self) -> bool;

    /// Assess if datatype is a datetime
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn is_datetime(&self) -> bool;

    /// Assess if datatype is an ISO8601 datetime
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime(&self) -> Option<ExcelDateTime>;

    /// Try getting datetime ISO8601 value
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn get_datetime_iso(&self) -> Option<&str>;

    /// Try getting duration ISO8601 value
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn as_date(&self) -> Option<chrono::NaiveDate> {
        use std::str::FromStr;
        if self.is_datetime_iso() {
            self.as_datetime().map(|dt| dt.date()).or_else(|| {
                self.get_datetime_iso()
                    .and_then(|s| chrono::NaiveDate::from_str(s).ok())
            })
        } else {
            self.as_datetime().map(|dt| dt.date())
        }
    }

    /// Try converting data type into a time
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn as_time(&self) -> Option<chrono::NaiveTime> {
        use std::str::FromStr;
        if self.is_datetime_iso() {
            self.as_datetime().map(|dt| dt.time()).or_else(|| {
                self.get_datetime_iso()
                    .and_then(|s| chrono::NaiveTime::from_str(s).ok())
            })
        } else if self.is_duration_iso() {
            self.get_duration_iso()
                .and_then(|s| chrono::NaiveTime::parse_from_str(s, "PT%HH%MM%S%.fS").ok())
        } else {
            self.as_datetime().map(|dt| dt.time())
        }
    }

    /// Try converting data type into a duration
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    fn as_duration(&self) -> Option<chrono::Duration> {
        use chrono::Timelike;

        if self.is_datetime() {
            self.get_datetime().and_then(|dt| dt.as_duration())
        } else if self.is_duration_iso() {
            // need replace in the future to something like chrono::Duration::from_str()
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

    // Try converting data type into a datetime.
    #[cfg(feature = "chrono")]
    fn as_datetime(&self) -> Option<chrono::NaiveDateTime> {
        use std::str::FromStr;

        if self.is_int() || self.is_float() {
            self.as_f64()
                .map(|f| ExcelDateTime::from_value_only(f).as_datetime())
        } else if self.is_datetime() {
            self.get_datetime().map(|d| d.as_datetime())
        } else if self.is_datetime_iso() {
            self.get_datetime_iso()
                .map(|s| chrono::NaiveDateTime::from_str(s).ok())
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
            DataRef::SharedString(v) => Data::String(v.into()),
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
    /// `DateTime`
    DateTime,
    /// `TimeDelta` (Duration)
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

    // Is used only for converting excel value to chrono.
    #[cfg(feature = "chrono")]
    fn from_value_only(value: f64) -> Self {
        ExcelDateTime {
            value,
            ..Default::default()
        }
    }

    /// True if excel datetime has duration format (`[hh]:mm:ss`, for example)
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    pub fn is_duration(&self) -> bool {
        matches!(self.datetime_type, ExcelDateTimeType::TimeDelta)
    }

    /// True if excel datetime has datetime format (not duration)
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    pub fn is_datetime(&self) -> bool {
        matches!(self.datetime_type, ExcelDateTimeType::DateTime)
    }

    /// Converting data type into a float
    pub fn as_f64(&self) -> f64 {
        self.value
    }

    /// Convert an Excel serial datetime to standard date components.
    ///
    /// Datetimes in Excel are serial dates with days counted from an epoch
    /// (usually 1900-01-01) and where the time is a percentage/decimal of the
    /// milliseconds in the day. Both the date and time are stored in the same
    /// f64 value. For example, 2025/10/13 12:00:00 is stored as 45943.5.
    ///
    /// This function returns a tuple of (year, month, day, hour, minutes,
    /// seconds, milliseconds). It works for serial dates in both the 1900 and
    /// 1904 epochs.
    ///
    /// This function always returns a date, even if the serial value is outside
    /// of Excel's range of `0.0 <= datetime < 10000.0`. It also returns, as
    /// Excel does, the invalid date 1900/02/29 due to the [Excel 1900 leap year
    /// bug](https://en.wikipedia.org/wiki/Leap_year_problem#Occurrences).
    ///
    /// Excel only supports millisecond precision and it also doesn't use or
    /// encode timezone information in any way.
    ///
    /// # Examples
    ///
    /// An example of converting an Excel date/time to standard components.
    ///
    /// ```
    /// use calamine::{ExcelDateTime, ExcelDateTimeType};
    ///
    /// // Create an Excel datetime from the serial value 45943.541 which is
    /// // equivalent to the date "2025/10/13 12:59:02.400".
    /// let excel_datetime = ExcelDateTime::new(
    ///     45943.541,
    ///     ExcelDateTimeType::DateTime,
    ///     false, // Using 1900 epoch (not 1904).
    /// );
    ///
    /// // Convert to standard date/time components.
    /// let (year, month, day, hour, min, sec, milli) = excel_datetime.to_ymd_hms_milli();
    ///
    /// assert_eq!(year, 2025);
    /// assert_eq!(month, 10);
    /// assert_eq!(day, 13);
    /// assert_eq!(hour, 12);
    /// assert_eq!(min, 59);
    /// assert_eq!(sec, 2);
    /// assert_eq!(milli, 400);
    /// ```
    ///
    pub fn to_ymd_hms_milli(&self) -> (u16, u8, u8, u8, u8, u8, u16) {
        Self::excel_to_standard_datetime(self.value, self.is_1904)
    }

    /// Try converting data type into a duration.
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
    pub fn as_duration(&self) -> Option<chrono::Duration> {
        let ms = self.value * MS_MULTIPLIER;
        Some(chrono::Duration::milliseconds(ms.round() as i64))
    }

    /// Try converting data type into a datetime.
    #[cfg(feature = "chrono")]
    #[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
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

    // Convert an Excel serial datetime to its date components.
    //
    // Datetimes in Excel are serial dates with days counted from an epoch and
    // where the time is a percentage/decimal of the milliseconds in the day.
    // Both the date and time are stored in the same f64 value.
    //
    // The calculation back to standard date and time components is deceptively
    // tricky since simple division doesn't work due to the 4/100/400 year leap
    // day changes. The basic approach is to divide the range into 400 year
    // blocks, 100 year blocks, 4 year blocks and 1 year blocks to calculate the
    // year (relative to the epoch). The remaining days and seconds are used to
    // calculate the year day and time. To make the leap year calculations
    // easier we move the effective epoch back to 1600-01-01 which is the
    // closest 400 year epoch before 1900/1904.
    //
    // In addition we need to handle both a 1900 and 1904 epoch and we need to
    // account for the Excel date bug where it treats 1900 as a leap year.
    //
    // Works in the range 1899-12-31/1904-01-01 to 9999-12-31.
    //
    // Leap seconds and the timezone aren't taken into account since Excel
    // doesn't handle them.
    //
    fn excel_to_standard_datetime(
        excel_datetime: f64,
        is_1904: bool,
    ) -> (u16, u8, u8, u8, u8, u8, u16) {
        let mut months = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

        // Convert the seconds to a whole number of days.
        let mut days = excel_datetime.floor() as u64;

        // Move the epoch to 1600-01-01 to make the leap calculations easier.
        if is_1904 {
            // 1904 epoch dates.
            days += 111_033;
        } else if days > YEAR_DAYS {
            // 1900 epoch years other than 1900.
            days += 109_571;
        } else {
            // Adjust for the Excel 1900 leap year bug.
            days += 109_572;
        }

        // Get the number of 400 year blocks.
        let year_days_400 = days / YEAR_DAYS_400;
        let mut days = days % YEAR_DAYS_400;

        // Get the number of 100 year blocks. There are 2 kinds: those starting
        // from a %400 year with an extra leap day (36,525 days) and those
        // starting from other 100 year intervals with 1 day less (36,524 days).
        let year_days_100;
        if days < YEAR_DAYS_100 {
            year_days_100 = days / YEAR_DAYS_100;
            days %= YEAR_DAYS_100;
        } else {
            year_days_100 = 1 + (days - YEAR_DAYS_100) / (YEAR_DAYS_100 - 1);
            days = (days - YEAR_DAYS_100) % (YEAR_DAYS_100 - 1);
        }

        // Get the number of 4 year blocks. There are 2 kinds: a 4 year block
        // with a leap day (1461 days) and a 4 year block starting from non-leap
        // %100 years without a leap day (1460 days). We also need to account
        // for whether a 1461 day block was preceded by a 1460 day block at the
        // start of the 100 year block.
        let year_days_4;
        let mut non_leap_year_block = false;
        if year_days_100 == 0 {
            // Any 4 year block in a 36,525 day 100 year block. Has extra leap.
            year_days_4 = days / YEAR_DAYS_4;
            days %= YEAR_DAYS_4;
        } else if days < YEAR_DAYS_4 {
            // A 4 year block at the start of a 36,524 day 100 year block.
            year_days_4 = days / (YEAR_DAYS_4 - 1);
            days %= YEAR_DAYS_4 - 1;
            non_leap_year_block = true;
        } else {
            // A non-initial 4 year block in a 36,524 day 100 year block.
            year_days_4 = 1 + (days - (YEAR_DAYS_4 - 1)) / YEAR_DAYS_4;
            days = (days - (YEAR_DAYS_4 - 1)) % YEAR_DAYS_4;
        }

        // Get the number of 1 year blocks. We need to account for leap years
        // and non-leap years and whether the non-leap occurs after a leap year.
        let year_days_1;
        if non_leap_year_block {
            // A non-leap block not preceded by a leap block.
            year_days_1 = days / YEAR_DAYS;
            days %= YEAR_DAYS;
        } else if days < YEAR_DAYS + 1 {
            // A leap year block.
            year_days_1 = days / (YEAR_DAYS + 1);
            days %= YEAR_DAYS + 1;
        } else {
            // A non-leap block preceded by a leap block.
            year_days_1 = 1 + (days - (YEAR_DAYS + 1)) / YEAR_DAYS;
            days = (days - (YEAR_DAYS + 1)) % YEAR_DAYS;
        }

        // Calculate the year as the number of blocks*days since the epoch.
        let year = 1600 + year_days_400 * 400 + year_days_100 * 100 + year_days_4 * 4 + year_days_1;

        // Convert from 0 indexed to 1 indexed days.
        days += 1;

        // Adjust February day count for leap years.
        if Self::is_leap_year(year) {
            months[1] = 29;
        }

        // Handle edge cases due to Excel erroneously treating 1900 as a leap year.
        if !is_1904 && year == 1900 {
            months[1] = 29;

            // Adjust last day of 1900.
            if excel_datetime.trunc() == 366.0 {
                days += 1;
            }
        }

        // Calculate the relevant month based on the sequential number of days.
        let mut month = 1;
        for month_days in months {
            if days > month_days {
                days -= month_days;
                month += 1;
            } else {
                break;
            }
        }

        // The final remainder is the day of the month.
        let day = days;

        // Get the time part of the Excel datetime.
        let time = excel_datetime.fract();
        let milli = ((time * DAY_SECONDS).fract() * 1000.0).round() as u64;
        let day_as_seconds = (time * DAY_SECONDS) as u64;

        // Calculate the hours, minutes and seconds in the day.
        let hour = day_as_seconds / HOUR_SECONDS;
        let min = (day_as_seconds - hour * HOUR_SECONDS) / MINUTE_SECONDS;
        let sec = (day_as_seconds - hour * HOUR_SECONDS - min * MINUTE_SECONDS) % MINUTE_SECONDS;

        // Return the date and time components.
        (
            year as u16,
            month as u8,
            day as u8,
            hour as u8,
            min as u8,
            sec as u8,
            milli as u16,
        )
    }

    // Check if a year is a leap year.
    fn is_leap_year(year: u64) -> bool {
        year % 4 == 0 && (year % 100 != 0 || year % 400 == 0)
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

#[cfg(all(test, feature = "chrono"))]
mod date_tests {
    use super::*;

    #[test]
    fn test_dates() {
        use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

        #[allow(clippy::excessive_precision)]
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
        #[allow(clippy::excessive_precision)]
        let date = Data::Float(0.18737500000000001);
        assert_eq!(
            date.as_time(),
            Some(NaiveTime::from_hms_milli_opt(4, 29, 49, 200).unwrap())
        );

        #[allow(clippy::excessive_precision)]
        let date = Data::Float(0.25951736111111101);
        assert_eq!(
            date.as_time(),
            Some(NaiveTime::from_hms_milli_opt(6, 13, 42, 300).unwrap())
        );

        // test overflow
        assert_eq!(Data::Float(1e20).as_time(), None);

        #[allow(clippy::excessive_precision)]
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

    #[test]
    fn test_datetimes_1900_epoch() {
        #[allow(clippy::excessive_precision)]
        let test_data = vec![
            (0.0, (1899, 12, 31, 0, 0, 0, 0)),
            (30188.010650613425, (1982, 8, 25, 0, 15, 20, 213)),
            (60376.011670023145, (2065, 4, 19, 0, 16, 48, 290)),
            (90565.038488958337, (2147, 12, 15, 0, 55, 25, 446)),
            (120753.04359827546, (2230, 8, 10, 1, 2, 46, 891)),
            (150942.04462496529, (2313, 4, 6, 1, 4, 15, 597)),
            (181130.04838991899, (2395, 11, 30, 1, 9, 40, 889)),
            (211318.04968240741, (2478, 7, 25, 1, 11, 32, 560)),
            (241507.06272186342, (2561, 3, 21, 1, 30, 19, 169)),
            (271695.07529606484, (2643, 11, 15, 1, 48, 25, 580)),
            (301884.08578609955, (2726, 7, 12, 2, 3, 31, 919)),
            (332072.09111094906, (2809, 3, 6, 2, 11, 11, 986)),
            (362261.10042934027, (2891, 10, 31, 2, 24, 37, 95)),
            (392449.10772245371, (2974, 6, 26, 2, 35, 7, 220)),
            (422637.11472348380, (3057, 2, 19, 2, 45, 12, 109)),
            (452826.12962951389, (3139, 10, 17, 3, 6, 39, 990)),
            (483014.13065105322, (3222, 6, 11, 3, 8, 8, 251)),
            (513203.13834000000, (3305, 2, 5, 3, 19, 12, 576)),
            (543391.14563164348, (3387, 10, 1, 3, 29, 42, 574)),
            (573579.15105107636, (3470, 5, 27, 3, 37, 30, 813)),
            (603768.17683137732, (3553, 1, 21, 4, 14, 38, 231)),
            (633956.17810832174, (3635, 9, 16, 4, 16, 28, 559)),
            (664145.17914608796, (3718, 5, 13, 4, 17, 58, 222)),
            (694333.18173372687, (3801, 1, 6, 4, 21, 41, 794)),
            (724522.20596981479, (3883, 9, 2, 4, 56, 35, 792)),
            (754710.22586672450, (3966, 4, 28, 5, 25, 14, 885)),
            (784898.22645513888, (4048, 12, 21, 5, 26, 5, 724)),
            (815087.24078782403, (4131, 8, 18, 5, 46, 44, 68)),
            (845275.24167987274, (4214, 4, 13, 5, 48, 1, 141)),
            (875464.24574438657, (4296, 12, 7, 5, 53, 52, 315)),
            (905652.26028449077, (4379, 8, 3, 6, 14, 48, 580)),
            (935840.28212659725, (4462, 3, 28, 6, 46, 15, 738)),
            (966029.31343063654, (4544, 11, 22, 7, 31, 20, 407)),
            (996217.33233511576, (4627, 7, 19, 7, 58, 33, 754)),
            (1026406.3386936343, (4710, 3, 15, 8, 7, 43, 130)),
            (1056594.3536005903, (4792, 11, 7, 8, 29, 11, 91)),
            (1086783.3807329629, (4875, 7, 4, 9, 8, 15, 328)),
            (1116971.3963169097, (4958, 2, 27, 9, 30, 41, 781)),
            (1147159.3986627546, (5040, 10, 23, 9, 34, 4, 462)),
            (1177348.4009715857, (5123, 6, 20, 9, 37, 23, 945)),
            (1207536.4013501736, (5206, 2, 12, 9, 37, 56, 655)),
            (1237725.4063915510, (5288, 10, 8, 9, 45, 12, 230)),
            (1267913.4126710880, (5371, 6, 4, 9, 54, 14, 782)),
            (1298101.4127558796, (5454, 1, 28, 9, 54, 22, 108)),
            (1328290.4177795255, (5536, 9, 24, 10, 1, 36, 151)),
            (1358478.5068125231, (5619, 5, 20, 12, 9, 48, 602)),
            (1388667.5237100578, (5702, 1, 14, 12, 34, 8, 549)),
            (1418855.5389640625, (5784, 9, 8, 12, 56, 6, 495)),
            (1449044.5409515856, (5867, 5, 6, 12, 58, 58, 217)),
            (1479232.5416002662, (5949, 12, 30, 12, 59, 54, 263)),
            (1509420.5657561459, (6032, 8, 24, 13, 34, 41, 331)),
            (1539609.5822754744, (6115, 4, 21, 13, 58, 28, 601)),
            (1569797.5849178126, (6197, 12, 14, 14, 2, 16, 899)),
            (1599986.6085352316, (6280, 8, 10, 14, 36, 17, 444)),
            (1630174.6096927200, (6363, 4, 6, 14, 37, 57, 451)),
            (1660363.6234115392, (6445, 11, 30, 14, 57, 42, 757)),
            (1690551.6325035533, (6528, 7, 26, 15, 10, 48, 307)),
            (1720739.6351839120, (6611, 3, 22, 15, 14, 39, 890)),
            (1750928.6387498612, (6693, 11, 15, 15, 19, 47, 988)),
            (1781116.6697262037, (6776, 7, 11, 16, 4, 24, 344)),
            (1811305.6822216667, (6859, 3, 7, 16, 22, 23, 952)),
            (1841493.6874536921, (6941, 10, 31, 16, 29, 55, 999)),
            (1871681.7071789235, (7024, 6, 26, 16, 58, 20, 259)),
            (1901870.7111390624, (7107, 2, 21, 17, 4, 2, 415)),
            (1932058.7211762732, (7189, 10, 16, 17, 18, 29, 630)),
            (1962247.7412190163, (7272, 6, 11, 17, 47, 21, 323)),
            (1992435.7454845603, (7355, 2, 5, 17, 53, 29, 866)),
            (2022624.7456143056, (7437, 10, 2, 17, 53, 41, 76)),
            (2052812.7465977315, (7520, 5, 28, 17, 55, 6, 44)),
            (2083000.7602910995, (7603, 1, 21, 18, 14, 49, 151)),
            (2113189.7623349307, (7685, 9, 16, 18, 17, 45, 738)),
            (2143377.7708298611, (7768, 5, 12, 18, 29, 59, 700)),
            (2173566.7731624190, (7851, 1, 7, 18, 33, 21, 233)),
            (2203754.8016744559, (7933, 9, 2, 19, 14, 24, 673)),
            (2233942.8036205554, (8016, 4, 27, 19, 17, 12, 816)),
            (2264131.8080603937, (8098, 12, 22, 19, 23, 36, 418)),
            (2294319.8239109721, (8181, 8, 17, 19, 46, 25, 908)),
            (2324508.8387420601, (8264, 4, 13, 20, 7, 47, 314)),
            (2354696.8552963310, (8346, 12, 8, 20, 31, 37, 603)),
            (2384885.8610853008, (8429, 8, 3, 20, 39, 57, 770)),
            (2415073.8682530904, (8512, 3, 29, 20, 50, 17, 67)),
            (2445261.8770581828, (8594, 11, 22, 21, 2, 57, 827)),
            (2475450.8910360998, (8677, 7, 19, 21, 23, 5, 519)),
            (2505638.8991848612, (8760, 3, 14, 21, 34, 49, 572)),
            (2535827.9021521294, (8842, 11, 8, 21, 39, 5, 944)),
            (2566015.9022965971, (8925, 7, 4, 21, 39, 18, 426)),
            (2596203.9070343636, (9008, 2, 28, 21, 46, 7, 769)),
            (2626392.9152275696, (9090, 10, 24, 21, 57, 55, 662)),
            (2656580.9299968979, (9173, 6, 19, 22, 19, 11, 732)),
            (2686769.9332335186, (9256, 2, 13, 22, 23, 51, 376)),
            (2716957.9360968866, (9338, 10, 9, 22, 27, 58, 771)),
            (2747146.9468795368, (9421, 6, 5, 22, 43, 30, 392)),
            (2777334.9502990046, (9504, 1, 30, 22, 48, 25, 834)),
            (2807522.9540709145, (9586, 9, 24, 22, 53, 51, 727)),
            (2837711.9673210187, (9669, 5, 20, 23, 12, 56, 536)),
            (2867899.9693762613, (9752, 1, 14, 23, 15, 54, 109)),
            (2898088.9702850925, (9834, 9, 10, 23, 17, 12, 632)),
            (2958465.9999884260, (9999, 12, 31, 23, 59, 59, 0)),
        ];

        for test in test_data {
            let (excel_serial_datetime, expected) = test;
            let datetime =
                ExcelDateTime::new(excel_serial_datetime, ExcelDateTimeType::DateTime, false);
            let got = datetime.to_ymd_hms_milli();

            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_dates_only_1900_epoch() {
        let test_data = vec![
            (0.0, (1899, 12, 31)),
            (1.0, (1900, 1, 1)),
            (58.0, (1900, 2, 27)),
            (59.0, (1900, 2, 28)),
            (60.0, (1900, 2, 29)),
            (61.0, (1900, 3, 1)),
            (62.0, (1900, 3, 2)),
            (71.0, (1900, 3, 11)),
            (99.0, (1900, 4, 8)),
            (256.0, (1900, 9, 12)),
            (364.0, (1900, 12, 29)),
            (365.0, (1900, 12, 30)),
            (366.0, (1900, 12, 31)),
            (367.0, (1901, 1, 1)),
            (489.0, (1901, 5, 3)),
            (652.0, (1901, 10, 13)),
            (777.0, (1902, 2, 15)),
            (888.0, (1902, 6, 6)),
            (999.0, (1902, 9, 25)),
            (1001.0, (1902, 9, 27)),
            (1212.0, (1903, 4, 26)),
            (1313.0, (1903, 8, 5)),
            (1461.0, (1903, 12, 31)),
            (1462.0, (1904, 1, 1)),
            (1520.0, (1904, 2, 28)),
            (1521.0, (1904, 2, 29)),
            (1522.0, (1904, 3, 1)),
            (2615.0, (1907, 2, 27)),
            (2616.0, (1907, 2, 28)),
            (2617.0, (1907, 3, 1)),
            (2618.0, (1907, 3, 2)),
            (2619.0, (1907, 3, 3)),
            (2620.0, (1907, 3, 4)),
            (2621.0, (1907, 3, 5)),
            (2622.0, (1907, 3, 6)),
            (36161.0, (1999, 1, 1)),
            (36191.0, (1999, 1, 31)),
            (36192.0, (1999, 2, 1)),
            (36219.0, (1999, 2, 28)),
            (36220.0, (1999, 3, 1)),
            (36250.0, (1999, 3, 31)),
            (36251.0, (1999, 4, 1)),
            (36280.0, (1999, 4, 30)),
            (36281.0, (1999, 5, 1)),
            (36311.0, (1999, 5, 31)),
            (36312.0, (1999, 6, 1)),
            (36341.0, (1999, 6, 30)),
            (36342.0, (1999, 7, 1)),
            (36372.0, (1999, 7, 31)),
            (36373.0, (1999, 8, 1)),
            (36403.0, (1999, 8, 31)),
            (36404.0, (1999, 9, 1)),
            (36433.0, (1999, 9, 30)),
            (36434.0, (1999, 10, 1)),
            (36464.0, (1999, 10, 31)),
            (36465.0, (1999, 11, 1)),
            (36494.0, (1999, 11, 30)),
            (36495.0, (1999, 12, 1)),
            (36525.0, (1999, 12, 31)),
            (36526.0, (2000, 1, 1)),
            (36556.0, (2000, 1, 31)),
            (36557.0, (2000, 2, 1)),
            (36585.0, (2000, 2, 29)),
            (36586.0, (2000, 3, 1)),
            (36616.0, (2000, 3, 31)),
            (36617.0, (2000, 4, 1)),
            (36646.0, (2000, 4, 30)),
            (36647.0, (2000, 5, 1)),
            (36677.0, (2000, 5, 31)),
            (36678.0, (2000, 6, 1)),
            (36707.0, (2000, 6, 30)),
            (36708.0, (2000, 7, 1)),
            (36738.0, (2000, 7, 31)),
            (36739.0, (2000, 8, 1)),
            (36769.0, (2000, 8, 31)),
            (36770.0, (2000, 9, 1)),
            (36799.0, (2000, 9, 30)),
            (36800.0, (2000, 10, 1)),
            (36830.0, (2000, 10, 31)),
            (36831.0, (2000, 11, 1)),
            (36860.0, (2000, 11, 30)),
            (36861.0, (2000, 12, 1)),
            (36891.0, (2000, 12, 31)),
            (36892.0, (2001, 1, 1)),
            (36922.0, (2001, 1, 31)),
            (36923.0, (2001, 2, 1)),
            (36950.0, (2001, 2, 28)),
            (36951.0, (2001, 3, 1)),
            (36981.0, (2001, 3, 31)),
            (36982.0, (2001, 4, 1)),
            (37011.0, (2001, 4, 30)),
            (37012.0, (2001, 5, 1)),
            (37042.0, (2001, 5, 31)),
            (37043.0, (2001, 6, 1)),
            (37072.0, (2001, 6, 30)),
            (37073.0, (2001, 7, 1)),
            (37103.0, (2001, 7, 31)),
            (37104.0, (2001, 8, 1)),
            (37134.0, (2001, 8, 31)),
            (37135.0, (2001, 9, 1)),
            (37164.0, (2001, 9, 30)),
            (37165.0, (2001, 10, 1)),
            (37195.0, (2001, 10, 31)),
            (37196.0, (2001, 11, 1)),
            (37225.0, (2001, 11, 30)),
            (37226.0, (2001, 12, 1)),
            (37256.0, (2001, 12, 31)),
            (182623.0, (2400, 1, 1)),
            (182653.0, (2400, 1, 31)),
            (182654.0, (2400, 2, 1)),
            (182682.0, (2400, 2, 29)),
            (182683.0, (2400, 3, 1)),
            (182713.0, (2400, 3, 31)),
            (182714.0, (2400, 4, 1)),
            (182743.0, (2400, 4, 30)),
            (182744.0, (2400, 5, 1)),
            (182774.0, (2400, 5, 31)),
            (182775.0, (2400, 6, 1)),
            (182804.0, (2400, 6, 30)),
            (182805.0, (2400, 7, 1)),
            (182835.0, (2400, 7, 31)),
            (182836.0, (2400, 8, 1)),
            (182866.0, (2400, 8, 31)),
            (182867.0, (2400, 9, 1)),
            (182896.0, (2400, 9, 30)),
            (182897.0, (2400, 10, 1)),
            (182927.0, (2400, 10, 31)),
            (182928.0, (2400, 11, 1)),
            (182957.0, (2400, 11, 30)),
            (182958.0, (2400, 12, 1)),
            (182988.0, (2400, 12, 31)),
            (767011.0, (4000, 1, 1)),
            (767041.0, (4000, 1, 31)),
            (767042.0, (4000, 2, 1)),
            (767070.0, (4000, 2, 29)),
            (767071.0, (4000, 3, 1)),
            (767101.0, (4000, 3, 31)),
            (767102.0, (4000, 4, 1)),
            (767131.0, (4000, 4, 30)),
            (767132.0, (4000, 5, 1)),
            (767162.0, (4000, 5, 31)),
            (767163.0, (4000, 6, 1)),
            (767192.0, (4000, 6, 30)),
            (767193.0, (4000, 7, 1)),
            (767223.0, (4000, 7, 31)),
            (767224.0, (4000, 8, 1)),
            (767254.0, (4000, 8, 31)),
            (767255.0, (4000, 9, 1)),
            (767284.0, (4000, 9, 30)),
            (767285.0, (4000, 10, 1)),
            (767315.0, (4000, 10, 31)),
            (767316.0, (4000, 11, 1)),
            (767345.0, (4000, 11, 30)),
            (767346.0, (4000, 12, 1)),
            (767376.0, (4000, 12, 31)),
            (884254.0, (4321, 1, 1)),
            (884284.0, (4321, 1, 31)),
            (884285.0, (4321, 2, 1)),
            (884312.0, (4321, 2, 28)),
            (884313.0, (4321, 3, 1)),
            (884343.0, (4321, 3, 31)),
            (884344.0, (4321, 4, 1)),
            (884373.0, (4321, 4, 30)),
            (884374.0, (4321, 5, 1)),
            (884404.0, (4321, 5, 31)),
            (884405.0, (4321, 6, 1)),
            (884434.0, (4321, 6, 30)),
            (884435.0, (4321, 7, 1)),
            (884465.0, (4321, 7, 31)),
            (884466.0, (4321, 8, 1)),
            (884496.0, (4321, 8, 31)),
            (884497.0, (4321, 9, 1)),
            (884526.0, (4321, 9, 30)),
            (884527.0, (4321, 10, 1)),
            (884557.0, (4321, 10, 31)),
            (884558.0, (4321, 11, 1)),
            (884587.0, (4321, 11, 30)),
            (884588.0, (4321, 12, 1)),
            (884618.0, (4321, 12, 31)),
            (2958101.0, (9999, 1, 1)),
            (2958131.0, (9999, 1, 31)),
            (2958132.0, (9999, 2, 1)),
            (2958159.0, (9999, 2, 28)),
            (2958160.0, (9999, 3, 1)),
            (2958190.0, (9999, 3, 31)),
            (2958191.0, (9999, 4, 1)),
            (2958220.0, (9999, 4, 30)),
            (2958221.0, (9999, 5, 1)),
            (2958251.0, (9999, 5, 31)),
            (2958252.0, (9999, 6, 1)),
            (2958281.0, (9999, 6, 30)),
            (2958282.0, (9999, 7, 1)),
            (2958312.0, (9999, 7, 31)),
            (2958313.0, (9999, 8, 1)),
            (2958343.0, (9999, 8, 31)),
            (2958344.0, (9999, 9, 1)),
            (2958373.0, (9999, 9, 30)),
            (2958374.0, (9999, 10, 1)),
            (2958404.0, (9999, 10, 31)),
            (2958405.0, (9999, 11, 1)),
            (2958434.0, (9999, 11, 30)),
            (2958435.0, (9999, 12, 1)),
            (2958465.0, (9999, 12, 31)),
        ];

        for test in test_data {
            let (excel_serial_datetime, expected) = test;
            let datetime =
                ExcelDateTime::new(excel_serial_datetime, ExcelDateTimeType::DateTime, false);
            let got = datetime.to_ymd_hms_milli();
            let got = (got.0, got.1, got.2); // Date parts only.
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_dates_only_1904_epoch() {
        let test_data = vec![(0.0, (1904, 1, 1))];

        for test in test_data {
            let (excel_serial_datetime, expected) = test;
            let datetime =
                ExcelDateTime::new(excel_serial_datetime, ExcelDateTimeType::DateTime, true);
            let got = datetime.to_ymd_hms_milli();
            let got = (got.0, got.1, got.2); // Date parts only.
            assert_eq!(expected, got);
        }
    }

    #[test]
    fn test_times_only_both_epochs() {
        #[allow(clippy::excessive_precision)]
        let test_data = vec![
            (0.0, (0, 0, 0, 0)),
            (1.0650613425925924e-2, (0, 15, 20, 213)),
            (1.1670023148148148e-2, (0, 16, 48, 290)),
            (3.8488958333333337e-2, (0, 55, 25, 446)),
            (4.3598275462962965e-2, (1, 2, 46, 891)),
            (4.4624965277777782e-2, (1, 4, 15, 597)),
            (4.8389918981481483e-2, (1, 9, 40, 889)),
            (4.9682407407407404e-2, (1, 11, 32, 560)),
            (6.2721863425925936e-2, (1, 30, 19, 169)),
            (7.5296064814814809e-2, (1, 48, 25, 580)),
            (8.5786099537037031e-2, (2, 3, 31, 919)),
            (9.1110949074074077e-2, (2, 11, 11, 986)),
            (0.10042934027777778, (2, 24, 37, 95)),
            (0.10772245370370370, (2, 35, 7, 220)),
            (0.11472348379629631, (2, 45, 12, 109)),
            (0.12962951388888888, (3, 6, 39, 990)),
            (0.13065105324074075, (3, 8, 8, 251)),
            (0.13833999999999999, (3, 19, 12, 576)),
            (0.14563164351851851, (3, 29, 42, 574)),
            (0.15105107638888890, (3, 37, 30, 813)),
            (0.17683137731481480, (4, 14, 38, 231)),
            (0.17810832175925925, (4, 16, 28, 559)),
            (0.17914608796296297, (4, 17, 58, 222)),
            (0.18173372685185185, (4, 21, 41, 794)),
            (0.20596981481481480, (4, 56, 35, 792)),
            (0.22586672453703704, (5, 25, 14, 885)),
            (0.22645513888888891, (5, 26, 5, 724)),
            (0.24078782407407406, (5, 46, 44, 68)),
            (0.24167987268518520, (5, 48, 1, 141)),
            (0.24574438657407408, (5, 53, 52, 315)),
            (0.26028449074074073, (6, 14, 48, 580)),
            (0.28212659722222222, (6, 46, 15, 738)),
            (0.31343063657407405, (7, 31, 20, 407)),
            (0.33233511574074076, (7, 58, 33, 754)),
            (0.33869363425925925, (8, 7, 43, 130)),
            (0.35360059027777774, (8, 29, 11, 91)),
            (0.38073296296296300, (9, 8, 15, 328)),
            (0.39631690972222228, (9, 30, 41, 781)),
            (0.39866275462962958, (9, 34, 4, 462)),
            (0.40097158564814817, (9, 37, 23, 945)),
            (0.40135017361111114, (9, 37, 56, 655)),
            (0.40639155092592594, (9, 45, 12, 230)),
            (0.41267108796296298, (9, 54, 14, 782)),
            (0.41275587962962962, (9, 54, 22, 108)),
            (0.41777952546296299, (10, 1, 36, 151)),
            (0.50681252314814818, (12, 9, 48, 602)),
            (0.52371005787037039, (12, 34, 8, 549)),
            (0.53896406249999995, (12, 56, 6, 495)),
            (0.54095158564814816, (12, 58, 58, 217)),
            (0.54160026620370372, (12, 59, 54, 263)),
            (0.56575614583333333, (13, 34, 41, 331)),
            (0.58227547453703699, (13, 58, 28, 601)),
            (0.58491781249999997, (14, 2, 16, 899)),
            (0.60853523148148148, (14, 36, 17, 444)),
            (0.60969271990740748, (14, 37, 57, 451)),
            (0.62341153935185190, (14, 57, 42, 757)),
            (0.63250355324074070, (15, 10, 48, 307)),
            (0.63518391203703706, (15, 14, 39, 890)),
            (0.63874986111111109, (15, 19, 47, 988)),
            (0.66972620370370362, (16, 4, 24, 344)),
            (0.68222166666666662, (16, 22, 23, 952)),
            (0.68745369212962970, (16, 29, 55, 999)),
            (0.70717892361111112, (16, 58, 20, 259)),
            (0.71113906250000003, (17, 4, 2, 415)),
            (0.72117627314814825, (17, 18, 29, 630)),
            (0.74121901620370367, (17, 47, 21, 323)),
            (0.74548456018518516, (17, 53, 29, 866)),
            (0.74561430555555563, (17, 53, 41, 76)),
            (0.74659773148148145, (17, 55, 6, 44)),
            (0.76029109953703700, (18, 14, 49, 151)),
            (0.76233493055555546, (18, 17, 45, 738)),
            (0.77082986111111118, (18, 29, 59, 700)),
            (0.77316241898148153, (18, 33, 21, 233)),
            (0.80167445601851861, (19, 14, 24, 673)),
            (0.80362055555555545, (19, 17, 12, 816)),
            (0.80806039351851855, (19, 23, 36, 418)),
            (0.82391097222222232, (19, 46, 25, 908)),
            (0.83874206018518516, (20, 7, 47, 314)),
            (0.85529633101851854, (20, 31, 37, 603)),
            (0.86108530092592594, (20, 39, 57, 770)),
            (0.86825309027777775, (20, 50, 17, 67)),
            (0.87705818287037041, (21, 2, 57, 827)),
            (0.89103609953703700, (21, 23, 5, 519)),
            (0.89918486111111118, (21, 34, 49, 572)),
            (0.90215212962962965, (21, 39, 5, 944)),
            (0.90229659722222222, (21, 39, 18, 426)),
            (0.90703436342592603, (21, 46, 7, 769)),
            (0.91522756944444439, (21, 57, 55, 662)),
            (0.92999689814814823, (22, 19, 11, 732)),
            (0.93323351851851843, (22, 23, 51, 376)),
            (0.93609688657407408, (22, 27, 58, 771)),
            (0.94687953703703709, (22, 43, 30, 392)),
            (0.95029900462962968, (22, 48, 25, 834)),
            (0.95407091435185187, (22, 53, 51, 727)),
            (0.96732101851851848, (23, 12, 56, 536)),
            (0.96937626157407408, (23, 15, 54, 109)),
            (0.97028509259259266, (23, 17, 12, 632)),
            (0.99999998842592586, (23, 59, 59, 999)),
        ];

        for test in test_data {
            let (excel_serial_datetime, expected) = test;

            // 1900 epoch.
            let datetime =
                ExcelDateTime::new(excel_serial_datetime, ExcelDateTimeType::DateTime, false);
            let got = datetime.to_ymd_hms_milli();
            let got = (got.3, got.4, got.5, got.6); // Time parts only.
            assert_eq!(expected, got);

            // 1904 epoch.
            let datetime =
                ExcelDateTime::new(excel_serial_datetime, ExcelDateTimeType::DateTime, true);
            let got = datetime.to_ymd_hms_milli();
            let got = (got.3, got.4, got.5, got.6); // Time parts only.
            assert_eq!(expected, got);
        }
    }
}
