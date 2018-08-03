//! A module to convert file extension to reader

use errors::Error;
use std::borrow::Cow;
use std::fs::File;
use std::io::BufReader;
use std::path::Path;
use vba::VbaProject;
use {open_workbook, DataType, Metadata, Ods, Range, Reader, Xls, Xlsb, Xlsx};

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
        _ => return Err(Error::Msg("Unknown extension")),
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
    fn vba_project(&mut self) -> Option<Result<Cow<VbaProject>, Self::Error>> {
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
}
