//! A module to provide a convenient wrapper around all error types

/// A struct to handle any error and a message
#[derive(Fail, Debug)]
pub enum Error {
    /// IO error
    #[fail(display = "{}", _0)]
    Io(#[cause] ::std::io::Error),

    /// Ods specific error
    #[fail(display = "{}", _0)]
    Ods(#[cause] ::ods::OdsError),
    /// xls specific error
    #[fail(display = "{}", _0)]
    Xls(#[cause] ::xls::XlsError),
    /// xlsb specific error
    #[fail(display = "{}", _0)]
    Xlsb(#[cause] ::xlsb::XlsbError),
    /// xlsx specific error
    #[fail(display = "{}", _0)]
    Xlsx(#[cause] ::xlsx::XlsxError),
    /// vba specific error
    #[fail(display = "{}", _0)]
    Vba(#[cause] ::vba::VbaError),
    /// Auto error
    #[fail(display = "{}", _0)]
    Auto(#[cause] ::auto::AutoError),
    /// cfb specific error
    #[fail(display = "{}", _0)]
    De(#[cause] ::de::DeError),

    /// General error message
    #[fail(display = "{}", _0)]
    Msg(&'static str),
}

from_err!(::std::io::Error, Error, Io);
from_err!(::ods::OdsError, Error, Ods);
from_err!(::xls::XlsError, Error, Xls);
from_err!(::xlsb::XlsbError, Error, Xlsb);
from_err!(::xlsx::XlsxError, Error, Xlsx);
from_err!(::vba::VbaError, Error, Vba);
from_err!(::auto::AutoError, Error, Auto);
from_err!(::de::DeError, Error, De);
from_err!(&'static str, Error, Msg);
