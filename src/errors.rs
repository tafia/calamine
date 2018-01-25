//! `Error` management module
//! Provides all excel error conversion and description
//! Also provides `Result` as a alias of `Result<_, Error>

#![allow(missing_docs)]

/// A struct to handle calamine specific errors
#[derive(Fail, Debug)]
pub enum Error {
    #[fail(display = "{}", _0)] Io(#[cause] ::std::io::Error),

    #[fail(display = "{}", _0)] Ods(#[cause] ::ods::OdsError),
    #[fail(display = "{}", _0)] Xls(#[cause] ::xls::XlsError),
    #[fail(display = "{}", _0)] Xlsb(#[cause] ::xlsb::XlsbError),
    #[fail(display = "{}", _0)] Xlsx(#[cause] ::xlsx::XlsxError),
    #[fail(display = "{}", _0)] Vba(#[cause] ::vba::VbaError),
    #[fail(display = "{}", _0)] De(#[cause] ::de::DeError),

    #[fail(display = "invalid extension: '{}'", _0)] InvalidExtension(String),
    #[fail(display = "there is no cell at position '{:?}'.\
                      Minimum position is '{:?}'",
           try_pos, min_pos)]
    CellOutOfRange {
        try_pos: (u32, u32),
        min_pos: (u32, u32),
    },
    #[fail(display = "invalid worksheet name: '{}'", _0)] WorksheetName(String),
    #[fail(display = "invalid worksheet index: {}", idx)]
    WorksheetIndex {
        idx: usize,
    },
    #[fail(display = "{}", _0)] StaticMsg(&'static str),
}

from_err!(::std::io::Error, Error, Io);
from_err!(::ods::OdsError, Error, Ods);
from_err!(::xls::XlsError, Error, Xls);
from_err!(::xlsb::XlsbError, Error, Xlsb);
from_err!(::xlsx::XlsxError, Error, Xlsx);
from_err!(::vba::VbaError, Error, Vba);
from_err!(::de::DeError, Error, De);
from_err!(&'static str, Error, StaticMsg);
