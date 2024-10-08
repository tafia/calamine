//! A module to convert file extension to reader

use crate::errors::Error;
use crate::vba::VbaProject;
use crate::{
    open_workbook, open_workbook_from_rs, Data, DataRef, HeaderRow, Metadata, Ods, Range, Reader,
    ReaderRef, Xls, Xlsb, Xlsx,
};
use std::borrow::Cow;
use std::fs::File;
use std::io::BufReader;
use std::path::Path;

/// A wrapper over all sheets when the file type is not known at static time
pub enum Sheets<RS> {
    /// Xls reader
    Xls(Xls<RS>),
    /// Xlsx reader
    Xlsx(Xlsx<RS>),
    /// Xlsb reader
    Xlsb(Xlsb<RS>),
    /// Ods reader
    Ods(Ods<RS>),
}

/// Opens a workbook and define the file type at runtime.
///
/// Whenever possible use the statically known `open_workbook` function instead
pub fn open_workbook_auto<P>(path: P) -> Result<Sheets<BufReader<File>>, Error>
where
    P: AsRef<Path>,
{
    let path = path.as_ref();
    Ok(match path.extension().and_then(|e| e.to_str()) {
        Some("xls") | Some("xla") => Sheets::Xls(open_workbook(path).map_err(Error::Xls)?),
        Some("xlsx") | Some("xlsm") | Some("xlam") => {
            Sheets::Xlsx(open_workbook(path).map_err(Error::Xlsx)?)
        }
        Some("xlsb") => Sheets::Xlsb(open_workbook(path).map_err(Error::Xlsb)?),
        Some("ods") => Sheets::Ods(open_workbook(path).map_err(Error::Ods)?),
        _ => {
            if let Ok(ret) = open_workbook::<Xls<_>, _>(path) {
                return Ok(Sheets::Xls(ret));
            } else if let Ok(ret) = open_workbook::<Xlsx<_>, _>(path) {
                return Ok(Sheets::Xlsx(ret));
            } else if let Ok(ret) = open_workbook::<Xlsb<_>, _>(path) {
                return Ok(Sheets::Xlsb(ret));
            } else if let Ok(ret) = open_workbook::<Ods<_>, _>(path) {
                return Ok(Sheets::Ods(ret));
            } else {
                return Err(Error::Msg("Cannot detect file format"));
            };
        }
    })
}

/// Opens a workbook from the given bytes.
///
/// Whenever possible use the statically known `open_workbook_from_rs` function instead
pub fn open_workbook_auto_from_rs<RS>(data: RS) -> Result<Sheets<RS>, Error>
where
    RS: std::io::Read + std::io::Seek + Clone,
{
    if let Ok(ret) = open_workbook_from_rs::<Xls<RS>, RS>(data.clone()) {
        Ok(Sheets::Xls(ret))
    } else if let Ok(ret) = open_workbook_from_rs::<Xlsx<RS>, RS>(data.clone()) {
        Ok(Sheets::Xlsx(ret))
    } else if let Ok(ret) = open_workbook_from_rs::<Xlsb<RS>, RS>(data.clone()) {
        Ok(Sheets::Xlsb(ret))
    } else if let Ok(ret) = open_workbook_from_rs::<Ods<RS>, RS>(data) {
        Ok(Sheets::Ods(ret))
    } else {
        Err(Error::Msg("Cannot detect file format"))
    }
}

impl<RS> Reader<RS> for Sheets<RS>
where
    RS: std::io::Read + std::io::Seek,
{
    type Error = Error;

    /// Creates a new instance.
    fn new(_reader: RS) -> Result<Self, Self::Error> {
        Err(Error::Msg("Sheets must be created from a Path"))
    }

    fn with_header_row(&mut self, header_row: HeaderRow) -> &mut Self {
        match self {
            Sheets::Xls(ref mut e) => {
                e.with_header_row(header_row);
            }
            Sheets::Xlsx(ref mut e) => {
                e.with_header_row(header_row);
            }
            Sheets::Xlsb(ref mut e) => {
                e.with_header_row(header_row);
            }
            Sheets::Ods(ref mut e) => {
                e.with_header_row(header_row);
            }
        }
        self
    }

    /// Gets `VbaProject`
    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, Self::Error>> {
        match self {
            Sheets::Xls(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xls)),
            Sheets::Xlsx(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xlsx)),
            Sheets::Xlsb(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Xlsb)),
            Sheets::Ods(ref mut e) => e.vba_project().map(|vba| vba.map_err(Error::Ods)),
        }
    }

    /// Initialize
    fn metadata(&self) -> &Metadata {
        match self {
            Sheets::Xls(ref e) => e.metadata(),
            Sheets::Xlsx(ref e) => e.metadata(),
            Sheets::Xlsb(ref e) => e.metadata(),
            Sheets::Ods(ref e) => e.metadata(),
        }
    }

    /// Read worksheet data in corresponding worksheet path
    fn worksheet_range(&mut self, name: &str) -> Result<Range<Data>, Self::Error> {
        match self {
            Sheets::Xls(ref mut e) => e.worksheet_range(name).map_err(Error::Xls),
            Sheets::Xlsx(ref mut e) => e.worksheet_range(name).map_err(Error::Xlsx),
            Sheets::Xlsb(ref mut e) => e.worksheet_range(name).map_err(Error::Xlsb),
            Sheets::Ods(ref mut e) => e.worksheet_range(name).map_err(Error::Ods),
        }
    }

    /// Read worksheet formula in corresponding worksheet path
    fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, Self::Error> {
        match self {
            Sheets::Xls(ref mut e) => e.worksheet_formula(name).map_err(Error::Xls),
            Sheets::Xlsx(ref mut e) => e.worksheet_formula(name).map_err(Error::Xlsx),
            Sheets::Xlsb(ref mut e) => e.worksheet_formula(name).map_err(Error::Xlsb),
            Sheets::Ods(ref mut e) => e.worksheet_formula(name).map_err(Error::Ods),
        }
    }

    fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
        match self {
            Sheets::Xls(ref mut e) => e.worksheets(),
            Sheets::Xlsx(ref mut e) => e.worksheets(),
            Sheets::Xlsb(ref mut e) => e.worksheets(),
            Sheets::Ods(ref mut e) => e.worksheets(),
        }
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        match self {
            Sheets::Xls(ref e) => e.pictures(),
            Sheets::Xlsx(ref e) => e.pictures(),
            Sheets::Xlsb(ref e) => e.pictures(),
            Sheets::Ods(ref e) => e.pictures(),
        }
    }
}

impl<RS> ReaderRef<RS> for Sheets<RS>
where
    RS: std::io::Read + std::io::Seek,
{
    fn worksheet_range_ref<'a>(
        &'a mut self,
        name: &str,
    ) -> Result<Range<DataRef<'a>>, Self::Error> {
        match self {
            Sheets::Xlsx(ref mut e) => e.worksheet_range_ref(name).map_err(Error::Xlsx),
            Sheets::Xlsb(ref mut e) => e.worksheet_range_ref(name).map_err(Error::Xlsb),
            Sheets::Xls(_) => unimplemented!(),
            Sheets::Ods(_) => unimplemented!(),
        }
    }
}
