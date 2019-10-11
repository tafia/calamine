//! A module to provide a convenient wrapper around all error types

/// A struct to handle any error and a message
#[derive(Debug)]
pub enum Error {
    /// IO error
    Io(::std::io::Error),

    /// Ods specific error
    Ods(::ods::OdsError),
    /// xls specific error
    Xls(::xls::XlsError),
    /// xlsb specific error
    Xlsb(::xlsb::XlsbError),
    /// xlsx specific error
    Xlsx(::xlsx::XlsxError),
    /// vba specific error
    Vba(::vba::VbaError),
    /// cfb specific error
    De(::de::DeError),

    /// General error message
    Msg(&'static str),
}

from_err!(::std::io::Error, Error, Io);
from_err!(::ods::OdsError, Error, Ods);
from_err!(::xls::XlsError, Error, Xls);
from_err!(::xlsb::XlsbError, Error, Xlsb);
from_err!(::xlsx::XlsxError, Error, Xlsx);
from_err!(::vba::VbaError, Error, Vba);
from_err!(::de::DeError, Error, De);
from_err!(&'static str, Error, Msg);

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {}", e),
            Error::Ods(e) => write!(f, "Ods error: {}", e),
            Error::Xls(e) => write!(f, "Xls error: {}", e),
            Error::Xlsx(e) => write!(f, "Xlsx error: {}", e),
            Error::Xlsb(e) => write!(f, "Xlsb error: {}", e),
            Error::Vba(e) => write!(f, "Vba error: {}", e),
            Error::De(e) => write!(f, "Deserializer error: {}", e),
            Error::Msg(msg) => write!(f, "{}", msg),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Ods(e) => Some(e),
            Error::Xls(e) => Some(e),
            Error::Xlsb(e) => Some(e),
            Error::Xlsx(e) => Some(e),
            Error::Vba(e) => Some(e),
            Error::De(e) => Some(e),
            _ => None,
        }
    }
}
