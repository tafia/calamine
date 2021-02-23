//! Parse vbaProject.bin file
//!
//! Retranscription from:
//! https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py

use std::collections::HashMap;
use std::io::Read;
use std::path::PathBuf;

use byteorder::{LittleEndian, ReadBytesExt};
use log::{debug, log_enabled, warn, Level};

use crate::cfb::{Cfb, XlsEncoding};
use crate::utils::read_u16;

/// A VBA specific error enum
#[derive(Debug)]
pub enum VbaError {
    /// Error comes from a cfb parsing
    Cfb(crate::cfb::CfbError),
    /// Io error
    Io(std::io::Error),

    /// Cannot find module
    ModuleNotFound(String),
    /// Generic unknown u16 value
    Unknown {
        /// error type
        typ: &'static str,
        /// value found
        val: u16,
    },
    /// Invalid libid format
    LibId,
    /// Invalid record id
    InvalidRecordId {
        /// expected record id
        expected: u16,
        /// record if found
        found: u16,
    },
}

from_err!(crate::cfb::CfbError, VbaError, Cfb);
from_err!(std::io::Error, VbaError, Io);

impl std::fmt::Display for VbaError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            VbaError::Io(e) => write!(f, "I/O error: {}", e),
            VbaError::Cfb(e) => write!(f, "Cfb error: {}", e),

            VbaError::ModuleNotFound(e) => write!(f, "Cannot find module '{}'", e),
            VbaError::Unknown { typ, val } => write!(f, "Unknown {} '{:X}'", typ, val),
            VbaError::LibId => write!(f, "Unexpected libid format"),
            VbaError::InvalidRecordId { expected, found } => write!(
                f,
                "Invalid record id: expecting {:X} found {:X}",
                expected, found
            ),
        }
    }
}

impl std::error::Error for VbaError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            VbaError::Io(e) => Some(e),
            VbaError::Cfb(e) => Some(e),
            _ => None,
        }
    }
}

/// A struct for managing VBA reading
#[allow(dead_code)]
#[derive(Clone)]
pub struct VbaProject {
    references: Vec<Reference>,
    modules: HashMap<String, Vec<u8>>,
    encoding: XlsEncoding,
}

impl VbaProject {
    /// Create a new `VbaProject` out of the vbaProject.bin `ZipFile` or xls file
    ///
    /// Starts reading project metadata (header, directories, sectors and minisectors).
    pub fn new<R: Read>(r: &mut R, len: usize) -> Result<VbaProject, VbaError> {
        let mut cfb = Cfb::new(r, len)?;
        VbaProject::from_cfb(r, &mut cfb)
    }

    /// Creates a new `VbaProject` out of a Compound File Binary and the corresponding reader
    pub fn from_cfb<R: Read>(r: &mut R, cfb: &mut Cfb) -> Result<VbaProject, VbaError> {
        // dir stream
        let stream = cfb.get_stream("dir", r)?;
        let stream = crate::cfb::decompress_stream(&*stream)?;
        let stream = &mut &*stream;

        // read dir information record (not used)
        let encoding = read_dir_information(stream)?;

        // array of REFERENCE records
        let refs = Reference::from_stream(stream, &encoding)?;

        // modules
        let mods: Vec<Module> = read_modules(stream, &encoding)?;

        // read all modules
        let modules: HashMap<String, Vec<u8>> = mods
            .into_iter()
            .map(|m| {
                cfb.get_stream(&m.stream_name, r).and_then(|s| {
                    crate::cfb::decompress_stream(&s[m.text_offset..]).map(move |s| (m.name, s))
                })
            })
            .collect::<Result<HashMap<_, _>, _>>()?;

        Ok(VbaProject {
            references: refs,
            modules,
            encoding,
        })
    }

    /// Gets the list of `Reference`s
    pub fn get_references(&self) -> &[Reference] {
        &self.references
    }

    /// Gets the list of `Module` names
    pub fn get_module_names(&self) -> Vec<&str> {
        self.modules.keys().map(|k| &**k).collect()
    }

    /// Reads module content and tries to convert to utf8
    ///
    /// While it works most of the time, the modules are MBCS encoding and the conversion
    /// may fail. If this is the case you should revert to `read_module_raw` as there is
    /// no built in decoding provided in this crate
    ///
    /// # Examples
    /// ```
    /// use calamine::{Reader, open_workbook, Xlsx};
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut xl: Xlsx<_> = open_workbook(path).expect("Cannot find excel file");
    /// if let Some(Ok(mut vba)) = xl.vba_project() {
    ///     let vba = vba.to_mut();
    ///     let modules = vba.get_module_names().into_iter()
    ///                      .map(|s| s.to_string()).collect::<Vec<_>>();
    ///     for m in modules {
    ///         println!("Module {}:", m);
    ///         println!("{}", vba.get_module(&m)
    ///                           .expect(&format!("cannot read {:?} module", m)));
    ///     }
    /// }
    /// ```
    pub fn get_module(&self, name: &str) -> Result<String, VbaError> {
        debug!("read module {}", name);
        let data = self.get_module_raw(name)?;
        Ok(self.encoding.decode_all(data, None))
    }

    /// Reads module content (MBCS encoded) and output it as-is (binary output)
    pub fn get_module_raw(&self, name: &str) -> Result<&[u8], VbaError> {
        match self.modules.get(name) {
            Some(m) => Ok(&**m),
            None => Err(VbaError::ModuleNotFound(name.into())),
        }
    }
}

/// A vba reference
#[derive(Debug, Clone, Hash, Eq, PartialEq)]
pub struct Reference {
    /// name
    pub name: String,
    /// description
    pub description: String,
    /// location of the reference
    pub path: PathBuf,
}

impl Reference {
    /// Check if the reference location is accessible
    pub fn is_missing(&self) -> bool {
        !self.path.exists()
    }

    /// Gets the list of references from the dir_stream relevant part
    fn from_stream(stream: &mut &[u8], encoding: &XlsEncoding) -> Result<Vec<Reference>, VbaError> {
        debug!("read all references metadata");

        let mut references = Vec::new();
        let mut reference = Reference {
            name: "".to_string(),
            description: "".to_string(),
            path: "".into(),
        };

        loop {
            let check = stream.read_u16::<LittleEndian>();
            match check? {
                0x000F => {
                    // termination of references array
                    if !reference.name.is_empty() {
                        references.push(reference);
                    }
                    break;
                }
                0x0016 => {
                    // REFERENCENAME
                    if !reference.name.is_empty() {
                        references.push(reference);
                    }
                    let name = read_variable_record(stream, 1)?;
                    let name = encoding.decode_all(name, None);
                    reference = Reference {
                        name: name.clone(),
                        description: name,
                        path: "".into(),
                    };
                    check_variable_record(0x003E, stream)?; // unicode
                }
                0x0033 => {
                    // REFERENCEORIGINAL (followed by REFERENCECONTROL)
                    reference.set_libid(stream, encoding)?;
                }
                0x002F => {
                    // REFERENCECONTROL
                    *stream = &stream[4..]; // SizeTwiddled: len of total ref control
                    reference.set_libid(stream, encoding)?;

                    *stream = &stream[6..];
                    match stream.read_u16::<LittleEndian>()? {
                        0x0016 => {
                            // optional name record extended
                            read_variable_record(stream, 1)?; // name extended
                            check_variable_record(0x003E, stream)?; // name extended unicode
                            check_record(0x0030, stream)?;
                        }
                        0x0030 => (),
                        e => {
                            return Err(VbaError::Unknown {
                                typ: "token in reference control",
                                val: e,
                            });
                        }
                    }
                    *stream = &stream[4..];
                    reference.set_libid(stream, encoding)?;
                    *stream = &stream[26..];
                }
                0x000D => {
                    // REFERENCEREGISTERED
                    *stream = &stream[4..];
                    reference.set_libid(stream, encoding)?;
                    *stream = &stream[6..];
                }
                0x000E => {
                    // REFERENCEPROJECT
                    *stream = &stream[4..];
                    let absolute = read_variable_record(stream, 1)?; // project libid absolute
                    {
                        let absolute = encoding.decode_all(absolute, None);
                        reference.path = if absolute.starts_with("*\\C") {
                            absolute[3..].into()
                        } else {
                            absolute.into()
                        };
                    }
                    read_variable_record(stream, 1)?; // project libid relative
                    *stream = &stream[6..];
                }
                c => {
                    return Err(VbaError::Unknown {
                        typ: "check id",
                        val: c,
                    });
                }
            }
        }

        debug!("references: {:#?}", references);
        Ok(references)
    }

    fn set_libid(&mut self, stream: &mut &[u8], encoding: &XlsEncoding) -> Result<(), VbaError> {
        let libid = read_variable_record(stream, 1)?; //libid twiddled
        if libid.is_empty() || libid.ends_with(b"##") {
            return Ok(());
        }
        let libid = encoding.decode_all(libid, None);
        let mut parts = libid.rsplit('#');
        match (parts.next(), parts.next()) {
            (Some(desc), Some(path)) => {
                self.description = desc.into();
                // use original path if already set
                if !path.is_empty() && self.path.as_os_str().is_empty() {
                    self.path = path.into();
                }
                Ok(())
            }
            _ => Err(VbaError::LibId),
        }
    }
}

/// A vba module
#[derive(Debug, Clone, Default)]
struct Module {
    /// module name as it appears in vba project
    name: String,
    stream_name: String,
    text_offset: usize,
}

fn read_dir_information(stream: &mut &[u8]) -> Result<XlsEncoding, VbaError> {
    debug!("read dir header");

    // PROJECTSYSKIND, PROJECTLCID and PROJECTLCIDINVOKE Records
    *stream = &stream[30..];

    // PROJECT Codepage
    let encoding = XlsEncoding::from_codepage(read_u16(&stream[6..8]))?;
    *stream = &stream[8..];

    // PROJECTNAME Record
    check_variable_record(0x0004, stream)?;

    // PROJECTDOCSTRING Record
    check_variable_record(0x0005, stream)?;
    check_variable_record(0x0040, stream)?; // unicode

    // PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
    check_variable_record(0x0006, stream)?;
    check_variable_record(0x003D, stream)?;

    // PROJECTHELPCONTEXT PROJECTLIBFLAGS and PROJECTVERSION Records
    *stream = &stream[32..];

    // PROJECTCONSTANTS Record
    check_variable_record(0x000C, stream)?;
    check_variable_record(0x003C, stream)?; // unicode

    Ok(encoding)
}

fn read_modules(stream: &mut &[u8], encoding: &XlsEncoding) -> Result<Vec<Module>, VbaError> {
    debug!("read all modules metadata");
    *stream = &stream[4..];

    let module_len = stream.read_u16::<LittleEndian>()? as usize;

    *stream = &stream[8..]; // PROJECTCOOKIE record
    let mut modules = Vec::with_capacity(module_len);

    for _ in 0..module_len {
        // name
        let name = check_variable_record(0x0019, stream)?;
        let name = encoding.decode_all(name, None);

        check_variable_record(0x0047, stream)?; // unicode

        let stream_name = check_variable_record(0x001A, stream)?; // stream name
        let stream_name = encoding.decode_all(stream_name, None);

        check_variable_record(0x0032, stream)?; // stream name unicode
        check_variable_record(0x001C, stream)?; // doc string
        check_variable_record(0x0048, stream)?; // doc string unicode

        // offset
        check_record(0x0031, stream)?;
        *stream = &stream[4..];
        let offset = stream.read_u32::<LittleEndian>()? as usize;

        // help context
        check_record(0x001E, stream)?;
        *stream = &stream[8..];

        // cookie
        check_record(0x002C, stream)?;
        *stream = &stream[6..];

        match stream.read_u16::<LittleEndian>()? {
            0x0021 /* procedural module */ |
            0x0022 /* document, class or designer module */ => (),
            e => return Err(VbaError::Unknown { typ: "module typ", val: e }),
        }

        loop {
            *stream = &stream[4..]; // reserved
            match stream.read_u16::<LittleEndian>() {
                Ok(0x0025) /* readonly */ | Ok(0x0028) /* private */ => (),
                Ok(0x002B) => break,
                Ok(e) => return Err(VbaError::Unknown { typ: "record id", val: e }),
                Err(e) => return Err(VbaError::Io(e)),
            }
        }
        *stream = &stream[4..]; // reserved

        modules.push(Module {
            name,
            stream_name,
            text_offset: offset,
        });
    }

    Ok(modules)
}

/// Reads a variable length record
///
/// `mult` is a multiplier of the length (e.g 2 when parsing XLWideString)
fn read_variable_record<'a>(r: &mut &'a [u8], mult: usize) -> Result<&'a [u8], VbaError> {
    let len = r.read_u32::<LittleEndian>()? as usize * mult;
    let (read, next) = r.split_at(len);
    *r = next;
    Ok(read)
}

/// Check that next record matches `id` and returns a variable length record
fn check_variable_record<'a>(id: u16, r: &mut &'a [u8]) -> Result<&'a [u8], VbaError> {
    check_record(id, r)?;
    let record = read_variable_record(r, 1)?;
    if log_enabled!(Level::Warn) && record.len() > 100_000 {
        warn!(
            "record id {} as a suspicious huge length of {} (hex: {:x})",
            id,
            record.len(),
            record.len() as u32
        );
    }
    Ok(record)
}

/// Check that next record matches `id`
fn check_record(id: u16, r: &mut &[u8]) -> Result<(), VbaError> {
    debug!("check record {:x}", id);
    let record_id = r.read_u16::<LittleEndian>()?;
    if record_id != id {
        Err(VbaError::InvalidRecordId {
            expected: id,
            found: record_id,
        })
    } else {
        Ok(())
    }
}
