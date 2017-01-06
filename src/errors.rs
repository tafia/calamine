//! `ExcelError` management module
//!
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, ExcelError>

#![allow(missing_docs)]

use quick_xml::error::Error as XmlError;

error_chain! {
//    types {
//        Error, ErrorKind, ResultExt, Result;
//    }

//     links {
//         rustup_dist::Error, rustup_dist::ErrorKind, Dist;
//         rustup_utils::Error, rustup_utils::ErrorKind, Utils;
//     }

    foreign_links {
        Io(::std::io::Error);
        Zip(::zip::result::ZipError);
        Xml(XmlError);
        ParseInt(::std::num::ParseIntError);
        ParseFloat(::std::num::ParseFloatError);
        Utf8(::std::str::Utf8Error);
        FromUtf8(::std::string::FromUtf8Error);
    }

//     errors {
//         InvalidToolchainName(t: String) {
//             description("invalid toolchain name")
//             display("invalid toolchain name: '{}'", t)
//         }
//     }
}

impl From<(XmlError, usize)> for Error {
    fn from(err: (XmlError, usize)) -> Error {
        err.0.into()
    }
}
