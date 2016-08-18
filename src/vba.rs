//! Parse vbaProject.bin file
//!
//! Retranscription from: 
//! https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py

use zip::read::ZipFile;
use std::io::{Read, BufRead};
use std::collections::HashMap;
use std::path::PathBuf;
use error::{ExcelResult, ExcelError};
use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;
use byteorder::{LittleEndian, ReadBytesExt};
use log::LogLevel;

const OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;
const CLASS_EXTENSION: &'static str = "cls";
const MODULE_EXTENSION: &'static str = "bas";
const FORM_EXTENSION: &'static str = "frm";

const POWER_2: [usize; 16] = [1   , 1<<1, 1<<2,  1<<3,  1<<4,  1<<5,  1<<6,  1<<7, 
                              1<<8, 1<<9, 1<<10, 1<<11, 1<<12, 1<<13, 1<<14, 1<<15];

#[allow(dead_code)]
pub struct VbaProject {
    header: Header,
    directories: Vec<Directory>,
    sectors: Sector,
    mini_sectors: Option<Sector>,
}

impl VbaProject {

    /// Create a new `VbaProject` out of the vbaProject.bin `ZipFile`.
    ///
    /// Starts reading project metadata (header, directories, sectors and minisectors).
    /// Warning: Buffers the entire ZipFile in memory, it may be bad for huge projects
    pub fn new(mut f: ZipFile) -> ExcelResult<VbaProject> {
        debug!("new vba project");

        // load header
        let header = try!(Header::from_reader(&mut f));

        // check signature
        if header.ab_sig != OLE_SIGNATURE {
            return Err(ExcelError::Unexpected("invalid OLE signature (not an office document?)".to_string()));
        }

        let sector_size = 2u64.pow(header.sector_shift as u32) as usize;
        if (f.size() as usize - 512) % sector_size != 0 {
            return Err(ExcelError::Unexpected("last sector has invalid size".to_string()));
        }

        // Read whole file in memory (the file is delimited by sectors)
        let mut data = Vec::with_capacity(f.size() as usize - 512);
        try!(f.read_to_end(&mut data));
        let sector = Sector::new(data, sector_size);

        // load fat and dif sectors
        debug!("load dif");
        let mut fat_sectors = header.sect_fat.to_vec();
        let mut sector_id = header.sect_dif_start;
        while sector_id != FREESECT && sector_id != ENDOFCHAIN {
            fat_sectors.extend_from_slice(&try!(to_u32_vec(sector.get(sector_id))));
            sector_id = fat_sectors.pop().unwrap(); //TODO: check if in infinite loop
        }

        // load the FATs
        debug!("load fat");
        let mut fat = Vec::with_capacity(fat_sectors.len() * sector_size);
        for sector_id in fat_sectors.into_iter().filter(|id| *id != FREESECT) {
            fat.extend_from_slice(&try!(to_u32_vec(sector.get(sector_id))));
        }
        
        // set sector fats
        let sectors = sector.with_fats(fat);

        // get the list of directory sectors
        debug!("load dirs");
        let buffer = sectors.read_chain(header.sect_dir_start);
        let mut directories = Vec::with_capacity(buffer.len() / 128);
        for c in buffer.chunks(128) {
            directories.push(try!(Directory::from_slice(c)));
        }

        // load the mini streams
        let mini_sectors = if directories[0].sect_start == ENDOFCHAIN {
            None
        } else {
            debug!("load minis");
            let mut ministream = sectors.read_chain(directories[0].sect_start);
//             assert_eq!(ministream.len(), directories[0].ul_size as usize);
            ministream.truncate(directories[0].ul_size as usize); // should not be needed

            debug!("load minifat");
            let minifat = try!(to_u32_vec(&sectors.read_chain(header.sect_mini_fat_start)));

            let mini_sector_size = 2usize.pow(header.mini_sector_shift as u32);
            assert!(directories[0].ul_size as usize % mini_sector_size == 0);
            Some(Sector::new(ministream, mini_sector_size).with_fats(minifat))
        };

        Ok(VbaProject {
            header: header,
            directories: directories,
            sectors: sectors,
            mini_sectors: mini_sectors,
        })

    }

    /// Gets a stream by name out of directories
    fn get_stream(&self, name: &str) -> Option<Vec<u8>> {
        debug!("get stream {}", name);
        self.directories.iter()
            .find(|d| d.get_name().map(|n| &*n == name).unwrap_or(false))
            .map(|d| {
                let mut data = if d.ul_size < self.header.mini_sector_cutoff {
                    self.mini_sectors.as_ref()
                        .map_or_else(|| Vec::new(), |s| s.read_chain(d.sect_start))
                } else {
                    self.sectors.read_chain(d.sect_start)
                };
                data.truncate(d.ul_size as usize);
                data
            })
    }

    /// Gets `Module` extensions, in case one wants to output the results
    pub fn get_code_modules(&self) -> ExcelResult<HashMap<String, &'static str>> {
        let mut stream = &*match self.get_stream("PROJECT") {
            Some(s) => s,
            None => return Err(ExcelError::Unexpected("cannot find 'PROJECT' stream".to_string())),
        };
        
        let mut code_modules = HashMap::new();
        loop {
            let mut line = String::new();
            if try!(stream.read_line(&mut line)) == 0 { break; }
            let line = line.trim();
            if line.is_empty() || line.starts_with("[") { continue; }
            match line.find('=') {
                None => continue, // invalid or unknown PROJECT property line
                Some(pos) => {
                    let value = match &line[..pos] {
                        "Document" | "Class" => CLASS_EXTENSION,
                        "Module" => MODULE_EXTENSION,
                        "BaseClass" => FORM_EXTENSION,
                        _ => continue,
                    };
                    code_modules.insert(line[pos + 1..].to_string(), value);
                }
            }
        }
        Ok(code_modules)
    }

    /// Reads project `Reference`s and `Module`s
    pub fn read_vba(&self) -> ExcelResult<(Vec<Reference>, Vec<Module>)> {
        debug!("read vba");
        
        // dir stream
        let mut stream = &*match self.get_stream("dir") {
            Some(s) => try!(decompress_stream(&s)),
            None => return Err(ExcelError::Unexpected("cannot find 'dir' stream".to_string())),
        };

        // read header (not used)
        try!(self.read_dir_header(&mut stream));

        // array of REFERENCE records
        let references = try!(self.read_references(&mut stream));

        // modules
        let modules = try!(self.read_modules(&mut stream));
        Ok((references, modules))

    }

    fn read_dir_header(&self, stream: &mut &[u8]) -> ExcelResult<()> {
        debug!("read dir header");
        let mut buf = [0; 2048]; // should be enough as per [MS-OVBA]

        // PROJECTSYSKIND, PROJECTLCID and PROJECTLCIDINVOKE Records
        try!(stream.read_exact(&mut buf[0..38]));
        
        // PROJECTNAME Record
        try!(check_variable_record(0x0004, stream, &mut buf));

        // PROJECTDOCSTRING Record
        try!(check_variable_record(0x0005, stream, &mut buf));
        try!(check_variable_record(0x0040, stream, &mut buf)); // unicode

        // PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
        try!(check_variable_record(0x0006, stream, &mut buf));
        try!(check_variable_record(0x003D, stream, &mut buf));

        // PROJECTHELPCONTEXT PROJECTLIBFLAGS and PROJECTVERSION Records
        try!(stream.read_exact(&mut buf[..32]));

        // PROJECTCONSTANTS Record
        try!(check_variable_record(0x000C, stream, &mut buf));
        try!(check_variable_record(0x003C, stream, &mut buf)); // unicode

        Ok(())
    }

    fn read_references(&self, stream: &mut &[u8]) -> ExcelResult<Vec<Reference>> {
        debug!("read all references metadata");

        let mut references = Vec::new();
        let mut buf = [0; 512];

        let mut reference = Reference { 
            name: "".to_string(), 
            description: "".to_string(), 
            path: "/".into() 
        };

        fn set_module_from_libid(reference: &mut Reference, len: usize, buf: &mut [u8]) 
            -> ExcelResult<()> 
        {
            let libid = try!(::std::str::from_utf8(&buf[..len]));
            let mut parts = libid.split('#').rev();
            parts.next().map(|p| reference.description = p.to_string());
            parts.next().map(|p| reference.path = p.into());
            Ok(())
        }

        loop {

            let check = stream.read_u16::<LittleEndian>();
            match try!(check) {
                0x000F => { // termination of references array
                    if !reference.name.is_empty() { references.push(reference); }
                    break;
                },

                0x0016 => { // REFERENCENAME
                    if !reference.name.is_empty() { references.push(reference); }

                    let len = try!(read_variable_record(stream, &mut buf));
                    let name = try!(::std::string::String::from_utf8(buf[..len].into()));
                    reference = Reference {
                        name: name.clone(),
                        description: name.clone(),
                        path: "/".into(),
                    };

                    try!(check_variable_record(0x003E, stream, &mut buf)); // unicode
                },

                0x0033 => { // REFERENCEORIGINAL (followed by REFERENCECONTROL)
                    try!(read_variable_record(stream, &mut buf));
                },

                0x002F => { // REFERENCECONTROL
                    try!(stream.read_exact(&mut buf[..4])); // len of total ref control

                    let len = try!(read_variable_record(stream, &mut buf)); //libid twiddled
                    try!(set_module_from_libid(&mut reference, len, &mut buf));

                    try!(stream.read_exact(&mut buf[..6]));

                    match try!(stream.read_u16::<LittleEndian>()) {
                        0x0016 => { // optional name record extended
                            try!(read_variable_record(stream, &mut buf)); // name extended
                            try!(check_variable_record(0x003E, stream, &mut buf)); // name extended unicode
                            try!(check_record(0x0030, stream));
                        },
                        0x0030 => (),
                        e => return Err(ExcelError::Unexpected(format!(
                                    "unexpected token in reference control {:x}", e))),
                    } 
                    try!(stream.read_exact(&mut buf[..4]));
                    try!(read_variable_record(stream, &mut buf)); // libid extended
                    try!(stream.read_exact(&mut buf[..26]));
                },

                0x000D => { // REFERENCEREGISTERED
                    try!(stream.read_exact(&mut buf[..4]));

                    let len = try!(read_variable_record(stream, &mut buf)); // libid registered
                    try!(set_module_from_libid(&mut reference, len, &mut buf));

                    try!(stream.read_exact(&mut buf[..6]));
                },

                0x000E => { // REFERENCEPROJECT
                    try!(stream.read_exact(&mut buf[..4]));
                    let len = try!(read_variable_record(stream, &mut buf)); // project libid absolute
                    {
                        let absolute = try!(::std::str::from_utf8(&buf[..len]));
                        reference.path = if absolute.starts_with("*\\C") { 
                            absolute[3..].into()
                        } else {
                            absolute.into()
                        };
                    }
                    try!(read_variable_record(stream, &mut buf)); // project libid relative
                    try!(stream.read_exact(&mut buf[..6]));
                },
                c => return Err(ExcelError::Unexpected(format!("invalid of unknown check Id {}", c))),
            }
        }

        Ok(references)
    }

    fn read_modules(&self, stream: &mut &[u8]) -> ExcelResult<Vec<Module>> {
        debug!("read all modules metadata");
        let mut buf = [0; 4096];
        try!(stream.read_exact(&mut buf[..4]));
        
        let module_len = try!(stream.read_u16::<LittleEndian>()) as usize;

        try!(stream.read_exact(&mut buf[..8])); // PROJECTCOOKIE record
        let mut modules = Vec::with_capacity(module_len);

        for _ in 0..module_len {

            // name
            let len = try!(check_variable_record(0x0019, stream, &mut buf));
            let name = try!(::std::string::String::from_utf8(buf[..len].to_vec()));

            try!(check_variable_record(0x0047, stream, &mut buf));      // unicode

            let len = try!(check_variable_record(0x001A, stream, &mut buf)); // stream name
            let stream_name = try!(::std::string::String::from_utf8(buf[..len].to_vec())); 

            try!(check_variable_record(0x0032, stream, &mut buf));      // stream name unicode
            try!(check_variable_record(0x001C, stream, &mut buf));      // doc string
            try!(check_variable_record(0x0048, stream, &mut buf));      // doc string unicode

            // offset
            try!(check_record(0x0031, stream));
            try!(stream.read_exact(&mut buf[..4]));
            let offset = try!(stream.read_u32::<LittleEndian>()) as usize;

            // help context
            try!(check_record(0x001E, stream));
            try!(stream.read_exact(&mut buf[..8]));

            // cookie
            try!(check_record(0x002C, stream));
            try!(stream.read_exact(&mut buf[..6]));

            match try!(stream.read_u16::<LittleEndian>()) {
                0x0021 => (), // procedural module
                0x0022 => (), // document, class or designer module
                e => return Err(ExcelError::Unexpected(format!(
                            "unknown module type {}", e))),
            }

            loop {
                try!(stream.read_exact(&mut buf[..4])); // reserved
                match stream.read_u16::<LittleEndian>() {
                    Ok(0x0025) /* readonly */ | Ok(0x0028) /* private */ => (),
                    Ok(0x002B) => break,
                    Ok(e) => return Err(ExcelError::Unexpected(format!(
                                "unknown record id {}", e))),
                    Err(e) => return Err(ExcelError::Io(e)),
                }
            }
            try!(stream.read_exact(&mut buf[..4])); // reserved

            modules.push(Module {
                name: name,
                stream_name: stream_name,
                text_offset: offset,
            });
        }

        Ok(modules)
    }

    /// Reads module content and tries to convert to utf8
    ///
    /// While it works most of the time, the modules are MBSC encoding and the conversion
    /// may fail. If this is the case you should revert to `read_module_raw` as there is 
    /// no built in decoding provided in this crate
    pub fn read_module(&self, module: &Module) -> ExcelResult<String> {
        debug!("read module {}", module.name);
        match self.get_stream(&module.stream_name) {
            None => Err(ExcelError::Unexpected(format!("cannot find {} stream", module.stream_name))),
            Some(s) => {
                let data = try!(decompress_stream(&s[module.text_offset..]));
                let data = try!(::std::string::String::from_utf8(data));
                Ok(data)
            }
        }
    }

    /// Reads module content (MBSC encoded) and output it as-is
    pub fn read_module_raw(&self, module: &Module) -> ExcelResult<Vec<u8>> {
        debug!("read module {}", module.name);
        match self.get_stream(&module.stream_name) {
            None => Err(ExcelError::Unexpected(format!("cannot find {} stream", module.stream_name))),
            Some(s) => {
                let data = try!(decompress_stream(&s[module.text_offset..]));
                Ok(data)
            }
        }
    }

}

struct Header {
    ab_sig: [u8; 8],
    _clid: [u8; 16],
    _minor_version: u16,
    _dll_version: u16,
    _byte_order: u16,
    sector_shift: u16,
    mini_sector_shift: u16,
    _reserved: u16,
    _reserved1: u32,
    _reserved2: u32,
    _sect_fat_len: u32,
    sect_dir_start: u32,
    _signature: u32,
    mini_sector_cutoff: u32,
    sect_mini_fat_start: u32,
    _sect_mini_fat_len: u32,
    sect_dif_start: u32,
    _sect_dif_len: u32,
    sect_fat: [u32; 109]
}

impl Header {
    fn from_reader<R: Read>(f: &mut R) -> ExcelResult<Header> {

        let mut ab_sig = [0; 8];
        try!(f.read_exact(&mut ab_sig));
        let mut clid = [0; 16];
        try!(f.read_exact(&mut clid));
        
        let minor_version = try!(f.read_u16::<LittleEndian>());
        let dll_version = try!(f.read_u16::<LittleEndian>());
        let byte_order = try!(f.read_u16::<LittleEndian>());
        let sector_shift = try!(f.read_u16::<LittleEndian>());
        let mini_sector_shift = try!(f.read_u16::<LittleEndian>());
        let reserved = try!(f.read_u16::<LittleEndian>());
        let reserved1 = try!(f.read_u32::<LittleEndian>());
        let reserved2 = try!(f.read_u32::<LittleEndian>());
        let sect_fat_len = try!(f.read_u32::<LittleEndian>());
        let sect_dir_start = try!(f.read_u32::<LittleEndian>());
        let signature = try!(f.read_u32::<LittleEndian>());
        let mini_sector_cutoff = try!(f.read_u32::<LittleEndian>());
        let sect_mini_fat_start = try!(f.read_u32::<LittleEndian>());
        let sect_mini_fat_len = try!(f.read_u32::<LittleEndian>());
        let sect_dif_start = try!(f.read_u32::<LittleEndian>());
        let sect_dif_len = try!(f.read_u32::<LittleEndian>());

        let mut sect_fat = [0u32; 109];
        for i in 0..109 {
            sect_fat[i] = try!(f.read_u32::<LittleEndian>());
        }

        Ok(Header {
            ab_sig: ab_sig, 
            _clid: clid,
            _minor_version: minor_version,
            _dll_version: dll_version,
            _byte_order: byte_order,
            sector_shift: sector_shift,
            mini_sector_shift: mini_sector_shift,
            _reserved: reserved,
            _reserved1: reserved1,
            _reserved2: reserved2,
            _sect_fat_len: sect_fat_len,
            sect_dir_start: sect_dir_start,
            _signature: signature,
            mini_sector_cutoff: mini_sector_cutoff,
            sect_mini_fat_start: sect_mini_fat_start,
            _sect_mini_fat_len: sect_mini_fat_len,
            sect_dif_start: sect_dif_start,
            _sect_dif_len: sect_dif_len,
            sect_fat: sect_fat,
        })
    }
}

/// Decode a buffer into u32 vector
fn to_u32_vec(mut buffer: &[u8]) -> ExcelResult<Vec<u32>> {
    assert!(buffer.len() % 4 == 0);
    let mut res = Vec::with_capacity(buffer.len() / 4);
    for _ in 0..buffer.len() / 4 {
        res.push(try!(buffer.read_u32::<LittleEndian>()));
    }
    Ok(res)
}

/// To better understand what's happening, look at 
/// http://www.wordarticles.com/Articles/Formats/StreamCompression.php
fn decompress_stream(mut r: &[u8]) -> ExcelResult<Vec<u8>> {
    debug!("decompress slice (len {})", r.len());
    let mut res = Vec::new();

    if try!(r.read_u8()) != 0x01 {
        return Err(ExcelError::Unexpected("invalid signature byte".to_string()));
    }

    while !r.is_empty() {

        // each 'chunk' is 4096 wide, let's reserve that space
        let start = res.len();
        res.reserve(4096);

        let chunk_header = try!(r.read_u16::<LittleEndian>());
        let chunk_size = chunk_header & 0x0FFF;
        let chunk_signature = (chunk_header & 0x7000) >> 12;
        let chunk_flag = (chunk_header & 0x8000) >> 15;

        assert_eq!(chunk_signature, 0b011);

        if chunk_flag == 0 { // uncompressed

            res.extend_from_slice(&[0u8; 4096]);
            try!(r.read_exact(&mut res[start..]));

        } else {

            let mut chunk_len = 0;
            'chunk: loop {

                let bit_flags = try!(r.read_u8());
                chunk_len += 1;

                for bit_index in 0..8 {

                    if chunk_len > chunk_size { break 'chunk; }

                    if (bit_flags & (1 << bit_index)) == 0 {
                        // literal token
                        res.push(try!(r.read_u8()));
                        chunk_len += 1;
                    } else {
                        // copy token
                        let token = try!(r.read_u16::<LittleEndian>());
                        chunk_len += 2;

                        let decomp_len = res.len() - start;
                        let bit_count = (4..16).find(|i| POWER_2[*i] >= decomp_len).unwrap();
                        let len_mask = 0xFFFF >> bit_count;
                        let len = (token & len_mask) as usize + 3;
                        let offset = ((token & !len_mask) >> (16 - bit_count)) as usize + 1;

                        for i in (res.len() - offset)..(res.len() - offset + len) {
                            let v = res[i];
                            res.push(v);
                        }

                    }
                }
            }
        }
    }
    Ok(res)

}

struct Sector {
    data: Vec<u8>,
    size: usize,
    fats: Vec<u32>,
}

impl Sector {

    fn new(data: Vec<u8>, size: usize) -> Sector {
        assert!(data.len() % size == 0);
        Sector {
            data: data,
            size: size as usize,
            fats: Vec::new(),
        }
    }

    fn with_fats(mut self, fats: Vec<u32>) -> Sector {
        self.fats = fats;
        self
    }

    fn get(&self, id: u32) -> &[u8] {
        &self.data[id as usize * self.size .. (id as usize + 1) * self.size]
    }

    fn read_chain(&self, mut sector_id: u32) -> Vec<u8> {
        debug!("chain reading sector {}", sector_id);
        let mut buffer = Vec::new();
        while sector_id != ENDOFCHAIN {
            buffer.extend_from_slice(self.get(sector_id));
            sector_id = self.fats[sector_id as usize];
        }
        buffer
    }

}

pub struct Directory {
    ab: [u8; 64],
    _cb: u16,
    _mse: i8,
    _flags: i8,
    _id_left_sib: u32,
    _id_right_sib: u32,
    _id_child: u32,
    _cls_id: [u8; 16],
    _dw_user_flags: u32,
    _time: [u64; 2],
    sect_start: u32,
    ul_size: u32,
    _dpt_prop_type: u16,
}

impl Directory {

    fn from_slice(mut rdr: &[u8]) -> ExcelResult<Directory> {
        let mut ab = [0; 64];
        try!(rdr.read_exact(&mut ab));

        let cb = try!(rdr.read_u16::<LittleEndian>());
        let mse = try!(rdr.read_i8());
        let flags = try!(rdr.read_i8());
        let id_left_sib = try!(rdr.read_u32::<LittleEndian>());
        let id_right_sib = try!(rdr.read_u32::<LittleEndian>());
        let id_child = try!(rdr.read_u32::<LittleEndian>());
        let mut cls_id = [0; 16];
        try!(rdr.read_exact(&mut cls_id));
        let dw_user_flags = try!(rdr.read_u32::<LittleEndian>());
        let time = [try!(rdr.read_u64::<LittleEndian>()),
                    try!(rdr.read_u64::<LittleEndian>())];
        let sect_start = try!(rdr.read_u32::<LittleEndian>());
        let ul_size = try!(rdr.read_u32::<LittleEndian>());
        let dpt_prop_type = try!(rdr.read_u16::<LittleEndian>());

        Ok(Directory {
            ab: ab,
            _cb: cb,
            _mse: mse,
            _flags: flags,
            _id_left_sib: id_left_sib,
            _id_right_sib: id_right_sib,
            _id_child: id_child,
            _cls_id: cls_id,
            _dw_user_flags: dw_user_flags,
            _time: time,
            sect_start: sect_start,
            ul_size: ul_size,
            _dpt_prop_type: dpt_prop_type,
        })

    }

    fn get_name(&self) -> ExcelResult<String> {
        let mut name = try!(UTF_16LE.decode(&self.ab, DecoderTrap::Ignore)
                            .map_err(ExcelError::Utf16));
        if let Some(len) = name.as_bytes().iter().position(|b| *b == 0) {
            name.truncate(len);
        }
        Ok(name)
    }
}

fn read_variable_record(r: &mut &[u8], buf: &mut [u8]) -> ExcelResult<usize> {
    let len = try!(r.read_u32::<LittleEndian>()) as usize;
    try!(r.read_exact(&mut buf[..len]));
    Ok(len)
}

fn check_variable_record(id: u16, r: &mut &[u8], buf: &mut [u8]) -> ExcelResult<usize> {
    try!(check_record(id, r));
    let len = try!(read_variable_record(r, buf));
    if log_enabled!(LogLevel::Warn) {
        if len > 100_000 {
            warn!("record id {} as a suspicious huge length of {} (hex: {:x})", 
                  id, len, len as u32);
        }
    }
    Ok(len)
}

fn check_record(id: u16, r: &mut &[u8]) -> ExcelResult<()> {
    debug!("check record {:x}", id);
    let record_id = try!(r.read_u16::<LittleEndian>());
    if record_id != id {
        return Err(ExcelError::Unexpected(
                format!("invalid record id, found {:x}, expecting {:x}", record_id, id)));
    }
    Ok(())
}

/// A vba reference
#[derive(Debug, Clone)]
pub struct Reference {
    pub name: String,
    pub description: String,
    pub path: PathBuf,
}

/// A vba module
#[derive(Debug, Clone, Default)]
pub struct Module {
    pub name: String,
    stream_name: String,
    text_offset: usize,
}

