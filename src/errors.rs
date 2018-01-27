//! `Error` management module
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, Error>

/// A struct to handle calamine specific errors
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

    /// Invalid extension: unrecognized input file
    #[fail(display = "invalid extension: '{}'", _0)]
    InvalidExtension(String),
    /// Cell out of range
    #[fail(display = "there is no cell at position '{:?}'.\
                      Minimum position is '{:?}'",
           try_pos, min_pos)]
    CellOutOfRange {
        /// position tried
        try_pos: (u32, u32),
        /// minimum position
        min_pos: (u32, u32),
    },
    /// Worksheet does not exist
    #[fail(display = "invalid worksheet name: '{}'", _0)]
    WorksheetName(String),
    /// Worksheet index does not exist
    #[fail(display = "invalid worksheet index: {}", idx)]
    WorksheetIndex {
        /// worksheet index
        idx: usize,
    },
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
from_err!(::de::DeError, Error, De);
from_err!(&'static str, Error, Msg);
