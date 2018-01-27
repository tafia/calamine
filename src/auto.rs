//! A module to convert file extension to reader

use std::borrow::Cow;
use std::fs::File;
use std::path::Path;
use std::io::{BufReader, Error as IoError, ErrorKind, Read, Seek, SeekFrom};
use vba::VbaProject;
use {DataType, Metadata, Range, Reader};

/// A wrapper around path which resolves extension
pub enum Extension {
    /// wrapper for extension .xls
    Xls,
    /// wrapper for extension .xlsx and .xlsm
    Xlsx,
    /// wrapper for extension .xlsb
    Xlsb,
    /// wrapper for extension .ods
    Ods,
}

impl Extension {
    /// Converts a path into an Extension
    pub fn from_path(p: &Path) -> Result<Extension, IoError> {
        Ok(match p.extension().and_then(|e| e.to_str()) {
            Some("xls") => Extension::Xls,
            Some("xlsx") | Some("xlsm") => Extension::Xlsx,
            Some("xlsb") => Extension::Xlsb,
            Some("ods") => Extension::Ods,
            _ => return Err(IoError::new(ErrorKind::InvalidInput, "unknown extension")),
        })
    }
}

/// A wrapper over a reader which hold original file extension
pub struct ExtensionReader<R> {
    extension: Extension,
    inner: R,
}

impl<R> ExtensionReader<R> {
    /// Creates a new `ExtensionReader`
    pub fn new(extension: Extension, reader: R) -> Self {
        ExtensionReader {
            extension: extension,
            inner: reader,
        }
    }
}

impl ExtensionReader<BufReader<File>> {
    /// Open a path and return a ExtensionReader<BufReader<File>>
    pub fn open<P: AsRef<Path>>(p: P) -> Result<Self, IoError> {
        let extension = Extension::from_path(p.as_ref())?;
        Ok(ExtensionReader::new(
            extension,
            BufReader::new(File::open(p)?),
        ))
    }
}

impl<R: Read> Read for ExtensionReader<R> {
    fn read(&mut self, buf: &mut [u8]) -> Result<usize, IoError> {
        self.inner.read(buf)
    }
}

impl<R: Seek> Seek for ExtensionReader<R> {
    fn seek(&mut self, pos: SeekFrom) -> Result<u64, IoError> {
        self.inner.seek(pos)
    }
}

/// A reader wrapper based on file extension
pub enum AutoReader<RS>
where
    RS: Read + Seek,
{
    /// wrapper for extension .xls
    Xls(::xls::Xls<RS>),
    /// wrapper for extension .xlsx or .xlsm
    Xlsx(::xlsx::Xlsx<RS>),
    /// wrapper for extension .xlsb
    Xlsb(::xlsb::Xlsb<RS>),
    /// wrapper for extension .ods
    Ods(::ods::Ods<RS>),
}

/// An error wrapper based on file extension
#[derive(Debug, Fail)]
pub enum AutoError {
    /// wrapper for extension .xls
    #[fail(display = "{}", _0)]
    Xls(#[cause] ::xls::XlsError),
    /// wrapper for extension .xlsx or .xlsm
    #[fail(display = "{}", _0)]
    Xlsx(#[cause] ::xlsx::XlsxError),
    /// wrapper for extension .xlsb
    #[fail(display = "{}", _0)]
    Xlsb(#[cause] ::xlsb::XlsbError),
    /// wrapper for extension .ods
    #[fail(display = "{}", _0)]
    Ods(#[cause] ::ods::OdsError),
    /// special io error
    #[fail(display = "{}", _0)]
    Io(#[cause] ::std::io::Error),
}

from_err!(IoError, AutoError, Io);

macro_rules! auto {
    ($self:expr, $fn:tt $(, $arg:expr)*) => {
        match *$self {
            AutoReader::Xls(ref mut e) => e.$fn($($arg,)*).map_err(AutoError::Xls),
            AutoReader::Xlsx(ref mut e) => e.$fn($($arg,)*).map_err(AutoError::Xlsx),
            AutoReader::Xlsb(ref mut e) => e.$fn($($arg,)*).map_err(AutoError::Xlsb),
            AutoReader::Ods(ref mut e) => e.$fn($($arg,)*).map_err(AutoError::Ods),
        }
    }
}

impl<RS: Read + Seek> Reader for AutoReader<RS> {
    type Error = AutoError;
    type RS = ExtensionReader<RS>;

    fn new(reader: Self::RS) -> Result<Self, Self::Error> {
        Ok(match reader.extension {
            Extension::Xls => {
                AutoReader::Xls(::xls::Xls::new(reader.inner).map_err(AutoError::Xls)?)
            }
            Extension::Xlsx => {
                AutoReader::Xlsx(::xlsx::Xlsx::new(reader.inner).map_err(AutoError::Xlsx)?)
            }
            Extension::Xlsb => {
                AutoReader::Xlsb(::xlsb::Xlsb::new(reader.inner).map_err(AutoError::Xlsb)?)
            }
            Extension::Ods => {
                AutoReader::Ods(::ods::Ods::new(reader.inner).map_err(AutoError::Ods)?)
            }
        })
    }

    fn has_vba(&mut self) -> bool {
        match *self {
            AutoReader::Xls(ref mut e) => e.has_vba(),
            AutoReader::Xlsx(ref mut e) => e.has_vba(),
            AutoReader::Xlsb(ref mut e) => e.has_vba(),
            AutoReader::Ods(ref mut e) => e.has_vba(),
        }
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>, Self::Error> {
        auto!(self, vba_project)
    }

    fn initialize(&mut self) -> Result<Metadata, Self::Error> {
        auto!(self, initialize)
    }

    fn read_worksheet_range(&mut self, name: &str) -> Result<Option<Range<DataType>>, Self::Error> {
        auto!(self, read_worksheet_range, name)
    }

    fn read_worksheet_formula(&mut self, name: &str) -> Result<Option<Range<String>>, Self::Error> {
        auto!(self, read_worksheet_formula, name)
    }
}
