//! ExcelError management module
//! 
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, ExcelError>

use quick_xml::error::Error as XmlError;

error_chain! {
    types {
        Error, ErrorKind, ChainErr, Result;
    }

//     links {
//         rustup_dist::Error, rustup_dist::ErrorKind, Dist;
//         rustup_utils::Error, rustup_utils::ErrorKind, Utils;
//     }

    foreign_links {
        ::std::io::Error, Io;
        ::zip::result::ZipError, Zip;
        XmlError, Xml;
        ::std::num::ParseIntError, ParseInt;
        ::std::num::ParseFloatError, ParseFloat;
        ::std::str::Utf8Error, Utf8;
        ::std::string::FromUtf8Error, FromUtf8;
    }

//     errors {
//         InvalidToolchainName(t: String) {
//             description("invalid toolchain name")
//             display("invalid toolchain name: '{}'", t)
//         }
//     }
}

impl From<(XmlError, usize)> for Error {
    fn from(err: (XmlError, usize)) -> Error { err.0.into() }
}
