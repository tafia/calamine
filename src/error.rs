//! ExcelError management module
//! 
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, ExcelError>

use std::fmt;
use std::io;
use std::borrow::Cow;
use zip::result::ZipError;
use quick_xml::error::Error as XmlError;
use std::num::{ParseIntError, ParseFloatError};
use std::str::Utf8Error;
use std::string::FromUtf8Error;

/// An error produced by an operation on CSV data.
#[derive(Debug)]
pub enum ExcelError {
    /// An error originating from reading or writing to the underlying buffer.
    Io(io::Error),
    /// An error occured while reading zip
    Zip(ZipError),
    /// An error occured when parsing xml
    Xml((XmlError, usize)),
    /// Error while parsing int
    ParseInt(ParseIntError),
    /// Error while parsing float
    ParseFloat(ParseFloatError),
    /// Utf8
    Utf8(Utf8Error),
    /// FromUtf18
    Utf16(Cow<'static, str>),
    /// Unexpected error
    Unexpected(String),
}

/// Result type
pub type ExcelResult<T> = ::std::result::Result<T, ExcelError>;

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match *self {
            ExcelError::Io(ref err) => write!(f, "{}", err),
            ExcelError::Zip(ref err) => write!(f, "{}", err),
            ExcelError::Xml((ref err, i)) => write!(f, "{} at position {}", err, i),
            ExcelError::ParseInt(ref err) => write!(f, "{}", err),
            ExcelError::ParseFloat(ref err) => write!(f, "{}", err),
            ExcelError::Utf8(ref err) => write!(f, "{}", err),
            ExcelError::Utf16(ref err) => write!(f, "{}", err),
            ExcelError::Unexpected(ref err) => write!(f, "{}", err),
        }
    }
}

impl ::std::error::Error for ExcelError {
    fn description(&self) -> &str {
        match *self {
            ExcelError::Io(..) => "CSV IO error",
            ExcelError::Zip(..) => "Zip error",
            ExcelError::Xml(..) => "Xml error",
            ExcelError::ParseInt(..) => "Parse int error",
            ExcelError::ParseFloat(..) => "Parse float error",
            ExcelError::Utf8(..) => "Decode utf8 string errorr",
            ExcelError::Utf16(..) => "Decode utf16 string errorr",
            ExcelError::Unexpected(..) => "Unexpected error",
        }
    }

    fn cause(&self) -> Option<&::std::error::Error> {
        match *self {
            ExcelError::Io(ref err) => Some(err),
            _ => None,
        }
    }
}

impl From<io::Error> for ExcelError {
    fn from(err: io::Error) -> ExcelError { ExcelError::Io(err) }
}

impl From<ZipError> for ExcelError {
    fn from(err: ZipError) -> ExcelError { ExcelError::Zip(err) }
}

impl From<(XmlError, usize)> for ExcelError {
    fn from(err: (XmlError, usize)) -> ExcelError { ExcelError::Xml(err) }
}

impl From<XmlError> for ExcelError {
    fn from(err: XmlError) -> ExcelError { ExcelError::Xml((err, 0)) }
}

impl From<ParseIntError> for ExcelError {
    fn from(err: ParseIntError) -> ExcelError { ExcelError::ParseInt(err) }
}

impl From<ParseFloatError> for ExcelError {
    fn from(err: ParseFloatError) -> ExcelError { ExcelError::ParseFloat(err) }
}

impl From<FromUtf8Error> for ExcelError {
    fn from(err: FromUtf8Error) -> ExcelError { ExcelError::Utf8(err.utf8_error()) }
}

impl From<Utf8Error> for ExcelError {
    fn from(err: Utf8Error) -> ExcelError { ExcelError::Utf8(err) }
}
