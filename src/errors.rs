//! `ExcelError` management module
//!
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, ExcelError>

#![allow(missing_docs)]

use quick_xml::error::Error as XmlError;

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
}

impl From<(XmlError, usize)> for Error {
    fn from(err: (XmlError, usize)) -> Error {
        err.0.into()
    }
}
