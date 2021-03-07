//! A module to convert file extension to reader

use crate::errors::Error;
use crate::vba::VbaProject;
use crate::{open_workbook, DataType, Metadata, Ods, Range, Reader, Xls, Xlsb, Xlsx};
use std::borrow::Cow;
use std::fs::File;
use std::io::BufReader;
use std::path::Path;

/// A wrapper over all sheets when the file type is not known at static time
pub enum Sheets {
    /// Xls reader
    Xls(Xls<BufReader<File>>),
    /// Xlsx reader
    Xlsx(Xlsx<BufReader<File>>),
    /// Xlsb reader
    Xlsb(Xlsb<BufReader<File>>),
    /// Ods reader
    Ods(Ods<BufReader<File>>),
}

/// Opens a workbook and define the file type at runtime.
///
/// Whenever possible use the statically known `open_workbook` function instead
pub fn open_workbook_auto<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
    Ok(match path.as_ref().extension().and_then(|e| e.to_str()) {
        Some("xls") | Some("xla") => Sheets::Xls(open_workbook(&path).map_err(Error::Xls)?),
        Some("xlsx") | Some("xlsm") | Some("xlam") => {
            Sheets::Xlsx(open_workbook(&path).map_err(Error::Xlsx)?)
        }
        Some("xlsb") => Sheets::Xlsb(open_workbook(&path).map_err(Error::Xlsb)?),
        Some("ods") => Sheets::Ods(open_workbook(&path).map_err(Error::Ods)?),
        _ => {
            fn open_workbook_xlsx<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
                Ok(Sheets::Xlsx(open_workbook(&path).map_err(Error::Xlsx)?))
            }
            fn open_workbook_xls<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
                Ok(Sheets::Xls(open_workbook(&path).map_err(Error::Xls)?))
            }
            fn open_workbook_xlsb<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
                Ok(Sheets::Xlsb(open_workbook(&path).map_err(Error::Xlsb)?))
            }
            fn open_workbook_ods<P: AsRef<Path>>(path: P) -> Result<Sheets, Error> {
                Ok(Sheets::Ods(open_workbook(&path).map_err(Error::Ods)?))
            }

            return if let Ok(ret) = open_workbook_xlsx(&path) {
                Ok(ret)
            } else if let Ok(ret) = open_workbook_xls(&path) {
                Ok(ret)
            } else if let Ok(ret) = open_workbook_xlsb(&path) {
                Ok(ret)
            } else if let Ok(ret) = open_workbook_ods(&path) {
                Ok(ret)
            } else {
                Err(Error::Msg("Cannot detect file format"))
            };
        }
    })
}

impl Reader for Sheets {
    type RS = BufReader<File>;
    type Error = Error;

    /// Creates a new instance.
    fn new(_reader: Self::RS) -> Result<Self, Self::Error> {
        Err(Error::Msg("Sheets must be created from a Path"))
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>> {
        match *self {
            Sheets::Xls(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xls)),
            Sheets::Xlsx(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xlsx)),
            Sheets::Xlsb(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xlsb)),
            Sheets::Ods(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Ods)),
        }
    }

    /// Initialize
    fn metadata(&self) -> &Metadata {
        match *self {
            Sheets::Xls(ref e) => e.metadata(),
            Sheets::Xlsx(ref e) => e.metadata(),
            Sheets::Xlsb(ref e) => e.metadata(),
            Sheets::Ods(ref e) => e.metadata(),
        }
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, Self::Error>> {
        match *self {
            Sheets::Xls(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xls)),
            Sheets::Xlsx(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xlsx)),
            Sheets::Xlsb(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Xlsb)),
            Sheets::Ods(ref mut e) => e.worksheet_range(name).map(|r| r.map_err(Error::Ods)),
        }
    }

    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, Self::Error>> {
        match *self {
            Sheets::Xls(ref mut e) => e.worksheet_formula(name).map(|r| r.map_err(Error::Xls)),
            Sheets::Xlsx(ref mut e) => e.worksheet_formula(name).map(|r| r.map_err(Error::Xlsx)),
            Sheets::Xlsb(ref mut e) => e.worksheet_formula(name).map(|r| r.map_err(Error::Xlsb)),
            Sheets::Ods(ref mut e) => e.worksheet_formula(name).map(|r| r.map_err(Error::Ods)),
        }
    }

    fn worksheets(&mut self) -> Vec<(String, Range<DataType>)> {
        match *self {
            Sheets::Xls(ref mut e) => e.worksheets(),
            Sheets::Xlsx(ref mut e) => e.worksheets(),
            Sheets::Xlsb(ref mut e) => e.worksheets(),
            Sheets::Ods(ref mut e) => e.worksheets(),
        }
    }
}
