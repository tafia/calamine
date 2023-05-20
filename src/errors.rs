//! A module to provide a convenient wrapper around all error types

/// A struct to handle any error and a message
#[derive(Debug)]
pub enum Error {
    /// IO error
    Io(std::io::Error),

    /// Ods specific error
    Ods(crate::ods::OdsError),
    /// xls specific error
    Xls(crate::xls::XlsError),
    /// xlsb specific error
    Xlsb(crate::xlsb::XlsbError),
    /// xlsx specific error
    Xlsx(crate::xlsx::XlsxError),
    /// vba specific error
    Vba(crate::vba::VbaError),
    /// cfb specific error
    De(crate::de::DeError),

    /// General error message
    Msg(&'static str),
}

from_err!(std::io::Error, Error, Io);
from_err!(crate::ods::OdsError, Error, Ods);
from_err!(crate::xls::XlsError, Error, Xls);
from_err!(crate::xlsb::XlsbError, Error, Xlsb);
from_err!(crate::xlsx::XlsxError, Error, Xlsx);
from_err!(crate::vba::VbaError, Error, Vba);
from_err!(crate::de::DeError, Error, De);
from_err!(&'static str, Error, Msg);

impl std::fmt::Display for Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
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
            Error::Msg(_) => None,
        }
    }
}
