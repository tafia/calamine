//! `Error` management module
//!
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, Error>

#![allow(missing_docs)]

use std::fmt;

use quick_xml::errors::Error as XmlError;
use serde::de;

use super::CellErrorType;

error_chain! {
    foreign_links {
        Io(::std::io::Error);
        Zip(::zip::result::ZipError);
        Xml(XmlError);
        ParseInt(::std::num::ParseIntError);
        ParseFloat(::std::num::ParseFloatError);
        Utf8(::std::str::Utf8Error);
        FromUtf8(::std::string::FromUtf8Error);
    }

    errors {
        InvalidExtension(ext: String) {
            description("invalid extension")
            display("invalid extension: '{}'", ext)
        }
        CellOutOfRange(try_pos: (u32, u32), min_pos: (u32, u32)) {
            description("no cell found at this position")
            display("there is no cell at position '{:?}'. Minimum position is '{:?}'",
                    try_pos, min_pos)
        }
        WorksheetName(name: String) {
            description("invalid worksheet name")
            display("invalid worksheet name: '{}'", name)
        }
        WorksheetIndex(idx: usize) {
            description("invalid worksheet index")
            display("invalid worksheet index: {}", idx)
        }
        UnexpectedEndOfRow(pos: (u32, u32)) {
            description("unexpected end of row")
            display("unexpected end of row at position '{:?}'", pos)
        }
        CellError(err: CellErrorType, pos: (u32, u32)) {
            description("cell error")
            display("cell error at position '{:?}'", pos)
        }
    }
}

impl From<(XmlError, usize)> for Error {
    fn from(err: (XmlError, usize)) -> Error {
        err.0.into()
    }
}

impl de::Error for Error {
    fn custom<T: fmt::Display>(msg: T) -> Self {
        Error::from(msg.to_string())
    }
}
