//! Parse vbaProject.bin file
//!
//! Retranscription from: 
//! https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py

use std::io::Read;
use std::path::PathBuf;
use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;
use byteorder::{LittleEndian, ReadBytesExt};
use log::LogLevel;
use errors::*;

const OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;

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
    pub fn new<R: Read>(mut f: R, len: usize) -> Result<VbaProject> {
        debug!("new vba project");

        // load header
        let header = try!(Header::from_reader(&mut f));

        // check signature
        if header.ab_sig != OLE_SIGNATURE {
            return Err("invalid OLE signature (not an office document?)".into());
        }

        let sector_size = 2u64.pow(header.sector_shift as u32) as usize;
        if (len - 512) % sector_size != 0 {
            return Err("last sector has invalid size".into());
        }

        // Read whole file in memory (the file is delimited by sectors)
        let mut data = Vec::with_capacity(len - 512);
        try!(f.read_to_end(&mut data));
        let sector = Sector::new(data, sector_size);

        // load fat and dif sectors
        debug!("load dif");
        let mut fat_sectors = header.sect_fat.to_vec();
        let mut sector_id = header.sect_dif_start;
        while sector_id != FREESECT && sector_id != ENDOFCHAIN {
            fat_sectors.extend(to_u32(sector.get(sector_id)));
            sector_id = fat_sectors.pop().unwrap(); //TODO: check if in infinite loop
        }

        // load the FATs
        debug!("load fat");
        let fat = fat_sectors.into_iter()
            .filter(|id| *id != FREESECT)
            .flat_map(|id| to_u32(sector.get(id)))
            .collect::<Vec<_>>();
        debug!("fats: {:?}", fat);
        
        // set sector fats
        let sectors = sector.with_fats(fat);

        // get the list of directory sectors
        debug!("load dirs");
        let directories = try!(sectors.read_chain(header.sect_dir_start)
            .flat_map(|sector| sector.chunks(128).map(|c| Directory::from_slice(c)))
            .collect::<Result<Vec<Directory>>>());

        // load the mini streams
        let mini_sectors = if directories[0].sect_start == ENDOFCHAIN {
            None
        } else {
            debug!("load minis");
            let mut ministream = sectors.read_chain(directories[0].sect_start)
                .collect::<Vec<_>>().concat();
            ministream.truncate(directories[0].ul_size as usize);

            debug!("load minifat");
            let minifat = sectors.read_chain(header.sect_mini_fat_start)
                .flat_map(|s| to_u32(s)).collect::<Vec<_>>();

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
    fn get_stream<'a>(&'a self, name: &str) -> Option<impl Iterator<Item=u8> + 'a> {
        debug!("get stream {}", name);
        match self.directories.iter()
            .find(|d| &*d.name == name) {
            None => None,
            Some(d) => {
                let sectors = if d.ul_size < self.header.mini_sector_cutoff {
                    self.mini_sectors.as_ref()
                } else {
                    Some(&self.sectors)
                };
                sectors.map(|ss| ss.read_chain(d.sect_start)
                            .flat_map(|s| s.iter()).take(d.ul_size as usize).cloned())
            }
        }
    }

    /// Reads project `Reference`s and `Module`s
    pub fn read_vba(&self) -> Result<(Vec<Reference>, Vec<Module>)> {
        debug!("read vba");
        
        // dir stream
        let mut stream = &*match self.get_stream("dir") {
            Some(s) => try!(decompress_stream(s)),
            None => return Err("cannot find 'dir' stream".into()),
        };

        // read header (not used)
        try!(self.read_dir_header(&mut stream));

        // array of REFERENCE records
        let references = try!(self.read_references(&mut stream));

        // modules
        let modules = try!(self.read_modules(&mut stream));
        Ok((references, modules))

    }

    fn read_dir_header(&self, stream: &mut &[u8]) -> Result<()> {
        debug!("read dir header");

        // PROJECTSYSKIND, PROJECTLCID and PROJECTLCIDINVOKE Records
        *stream = &stream[38..];
        
        // PROJECTNAME Record
        try!(check_variable_record(0x0004, stream));

        // PROJECTDOCSTRING Record
        try!(check_variable_record(0x0005, stream));
        try!(check_variable_record(0x0040, stream)); // unicode

        // PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
        try!(check_variable_record(0x0006, stream));
        try!(check_variable_record(0x003D, stream));

        // PROJECTHELPCONTEXT PROJECTLIBFLAGS and PROJECTVERSION Records
        *stream = &stream[32..];

        // PROJECTCONSTANTS Record
        try!(check_variable_record(0x000C, stream));
        try!(check_variable_record(0x003C, stream)); // unicode

        Ok(())
    }

    fn read_references(&self, stream: &mut &[u8]) -> Result<Vec<Reference>> {
        debug!("read all references metadata");

        let mut references = Vec::new();

        let mut reference = Reference { 
            name: "".to_string(), 
            description: "".to_string(), 
            path: "/".into() 
        };

        fn set_module_from_libid(reference: &mut Reference, libid: &[u8]) 
            -> Result<()> 
        {
            let libid = try!(::std::str::from_utf8(libid));
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

                    let name = try!(read_variable_record(stream));
                    let name = try!(::std::string::String::from_utf8(name.to_vec()));
                    reference = Reference {
                        name: name.clone(),
                        description: name.clone(),
                        path: "/".into(),
                    };

                    try!(check_variable_record(0x003E, stream)); // unicode
                },

                0x0033 => { // REFERENCEORIGINAL (followed by REFERENCECONTROL)
                    try!(read_variable_record(stream));
                },

                0x002F => { // REFERENCECONTROL
                    *stream = &stream[4..]; // len of total ref control

                    let libid = try!(read_variable_record(stream)); //libid twiddled
                    try!(set_module_from_libid(&mut reference, libid));

                    *stream = &stream[6..];

                    match try!(stream.read_u16::<LittleEndian>()) {
                        0x0016 => { // optional name record extended
                            try!(read_variable_record(stream)); // name extended
                            try!(check_variable_record(0x003E, stream)); // name extended unicode
                            try!(check_record(0x0030, stream));
                        },
                        0x0030 => (),
                        e => return Err(format!( "unexpected token in reference control {:x}", e).into()),
                    } 
                    *stream = &stream[4..];
                    try!(read_variable_record(stream)); // libid extended
                    *stream = &stream[26..];
                },

                0x000D => { // REFERENCEREGISTERED
                    *stream = &stream[4..];

                    let libid = try!(read_variable_record(stream)); // libid registered
                    try!(set_module_from_libid(&mut reference, libid));

                    *stream = &stream[6..];
                },

                0x000E => { // REFERENCEPROJECT
                    *stream = &stream[4..];
                    let absolute = try!(read_variable_record(stream)); // project libid absolute
                    {
                        let absolute = try!(::std::str::from_utf8(absolute));
                        reference.path = if absolute.starts_with("*\\C") { 
                            absolute[3..].into()
                        } else {
                            absolute.into()
                        };
                    }
                    try!(read_variable_record(stream)); // project libid relative
                    *stream = &stream[6..];
                },
                c => return Err(format!("invalid of unknown check Id {}", c).into()),
            }
        }

        Ok(references)
    }

    fn read_modules(&self, stream: &mut &[u8]) -> Result<Vec<Module>> {
        debug!("read all modules metadata");
        *stream = &stream[4..];
        
        let module_len = try!(stream.read_u16::<LittleEndian>()) as usize;

        *stream = &stream[8..]; // PROJECTCOOKIE record
        let mut modules = Vec::with_capacity(module_len);

        for _ in 0..module_len {

            // name
            let name = try!(check_variable_record(0x0019, stream));
            let name = try!(::std::string::String::from_utf8(name.to_vec()));

            try!(check_variable_record(0x0047, stream));      // unicode

            let stream_name = try!(check_variable_record(0x001A, stream)); // stream name
            let stream_name = try!(::std::string::String::from_utf8(stream_name.to_vec())); 

            try!(check_variable_record(0x0032, stream));      // stream name unicode
            try!(check_variable_record(0x001C, stream));      // doc string
            try!(check_variable_record(0x0048, stream));      // doc string unicode

            // offset
            try!(check_record(0x0031, stream));
            *stream = &stream[4..];
            let offset = try!(stream.read_u32::<LittleEndian>()) as usize;

            // help context
            try!(check_record(0x001E, stream));
            *stream = &stream[8..];

            // cookie
            try!(check_record(0x002C, stream));
            *stream = &stream[6..];

            match try!(stream.read_u16::<LittleEndian>()) {
                0x0021 => (), // procedural module
                0x0022 => (), // document, class or designer module
                e => return Err(format!("unknown module type {}", e).into()),
            }

            loop {
                *stream = &stream[4..]; // reserved
                match stream.read_u16::<LittleEndian>() {
                    Ok(0x0025) /* readonly */ | Ok(0x0028) /* private */ => (),
                    Ok(0x002B) => break,
                    Ok(e) => return Err(format!("unknown record id {}", e).into()),
                    Err(e) => return Err(e.into()),
                }
            }
            *stream = &stream[4..]; // reserved

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
    pub fn read_module(&self, module: &Module) -> Result<String> {
        debug!("read module {}", module.name);
        let data = try!(self.read_module_raw(module));
        let data = try!(::std::string::String::from_utf8(data));
        Ok(data)
    }

    /// Reads module content (MBSC encoded) and output it as-is
    pub fn read_module_raw(&self, module: &Module) -> Result<Vec<u8>> {
        debug!("read module raw {}", module.name);
        match self.get_stream(&module.stream_name) {
            None => Err(format!("cannot find {} stream", module.stream_name).into()),
            Some(s) => {
                let data = try!(decompress_stream(s.skip(module.text_offset)));
                Ok(data)
            }
        }
    }

}

struct Header {
    ab_sig: [u8; 8],
    sector_shift: u16,
    mini_sector_shift: u16,
    sect_dir_start: u32,
    mini_sector_cutoff: u32,
    sect_mini_fat_start: u32,
    sect_dif_start: u32,
    sect_fat: [u32; 109]
}

impl Header {
    fn from_reader<R: Read>(f: &mut R) -> Result<Header> {

        let mut ab_sig = [0; 8];
        try!(f.read_exact(&mut ab_sig));
        let mut clid = [0; 16];
        try!(f.read_exact(&mut clid));
        
        let _minor_version = try!(f.read_u16::<LittleEndian>());
        let _dll_version = try!(f.read_u16::<LittleEndian>());
        let _byte_order = try!(f.read_u16::<LittleEndian>());
        let sector_shift = try!(f.read_u16::<LittleEndian>());
        let mini_sector_shift = try!(f.read_u16::<LittleEndian>());
        let _reserved = try!(f.read_u16::<LittleEndian>());
        let _reserved1 = try!(f.read_u32::<LittleEndian>());
        let _reserved2 = try!(f.read_u32::<LittleEndian>());
        let _sect_fat_len = try!(f.read_u32::<LittleEndian>());
        let sect_dir_start = try!(f.read_u32::<LittleEndian>());
        let _signature = try!(f.read_u32::<LittleEndian>());
        let mini_sector_cutoff = try!(f.read_u32::<LittleEndian>());
        let sect_mini_fat_start = try!(f.read_u32::<LittleEndian>());
        let _sect_mini_fat_len = try!(f.read_u32::<LittleEndian>());
        let sect_dif_start = try!(f.read_u32::<LittleEndian>());
        let _sect_dif_len = try!(f.read_u32::<LittleEndian>());

        let mut sect_fat = [0u8; 109 * 4];
        try!(f.read_exact(&mut sect_fat));
        let sect_fat = unsafe { ::std::ptr::read(sect_fat.as_ptr() as *const [u32; 109]) };

        Ok(Header {
            ab_sig: ab_sig, 
            sector_shift: sector_shift,
            mini_sector_shift: mini_sector_shift,
            sect_dir_start: sect_dir_start,
            mini_sector_cutoff: mini_sector_cutoff,
            sect_mini_fat_start: sect_mini_fat_start,
            sect_dif_start: sect_dif_start,
            sect_fat: sect_fat,
        })
    }
}

/// To better understand what's happening, look at 
/// http://www.wordarticles.com/Articles/Formats/StreamCompression.php
fn decompress_stream<I: Iterator<Item=u8>>(mut r: I) -> Result<Vec<u8>> {
    debug!("decompress stream");
    let mut res = Vec::new();

    match r.next() {
        Some(0x01) => (),
        _ => return Err("invalid signature byte".into()),
    }

    fn read_u16<J: Iterator<Item=u8>>(i: &mut J) -> Option<u16> {
        match (i.next(), i.next()) {
            (Some(i1), Some(i2)) => (&[i1, i2] as &[u8]).read_u16::<LittleEndian>().ok(),
            _ => None,
        }
    }

    fn ok_or<T>(o: Option<T>, err: &'static str) -> Result<T> {
        match o {
            Some(o) => Ok(o),
            None => Err(err.into()),
        }
    }

    while let Some(chunk_header) = read_u16(&mut r) {

        // each 'chunk' is 4096 wide, let's reserve that space
        let start = res.len();
        res.reserve(4096);

        let chunk_size = chunk_header & 0x0FFF;
        let chunk_signature = (chunk_header & 0x7000) >> 12;
        let chunk_flag = (chunk_header & 0x8000) >> 15;

        assert_eq!(chunk_signature, 0b011);

        if chunk_flag == 0 { // uncompressed
            res.extend(r.by_ref().take(4096));
        } else {

            let mut chunk_len = 0;
            'chunk: loop {

                let bit_flags: u8 = try!(ok_or(r.next(), "no bit_flag in compressed stream"));
                chunk_len += 1;

                for bit_index in 0..8 {

                    if chunk_len > chunk_size { break 'chunk; }

                    if (bit_flags & (1 << bit_index)) == 0 {
                        // literal token
                        res.push(try!(ok_or(r.next(), "no literal token in compressed stream")));
                        chunk_len += 1;
                    } else {
                        // copy token
                        let token = try!(ok_or(read_u16(&mut r), "no literal token in compressed stream"));
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

    fn read_chain(&self, sector_id: u32) -> SectorChain {
        debug!("chain reading sector {}", sector_id);
        SectorChain {
            sector: self,
            id: sector_id,
        }
    }

}

struct SectorChain<'a> {
    sector: &'a Sector,
    id: u32,
}

impl<'a> Iterator for SectorChain<'a> {
    type Item = &'a [u8];
    fn next(&mut self) -> Option<&'a [u8]> {
        debug!("read chain next, id: {}", self.id);
        if self.id == ENDOFCHAIN {
            None
        } else {
            let sector = self.sector.get(self.id);
            self.id = self.sector.fats[self.id as usize];
            Some(sector)
        }
    }
}

pub struct Directory {
    sect_start: u32,
    ul_size: u32,
    name: String,
}

impl Directory {

    fn from_slice(rdr: &[u8]) -> Result<Directory> {
//         ab: [u8; 64],
//         _cb: u16,                    |
//         _mse: i8,                    |
//         _flags: i8,                  |
//         _id_left_sib: u32,           |
//         _id_right_sib: u32,          | = [u8; 52] with padding alignment
//         _id_child: u32,              |
//         _cls_id: [u8; 16],           |
//         _dw_user_flags: u32,         |
//         _time: [u64; 2],             |
//         sect_start: u32,
//         ul_size: u32,
//         _dpt_prop_type: u16,
        assert_eq!(rdr.len(), 128);
        let (ab, _, sect_start, ul_size) = unsafe { 
            ::std::ptr::read(rdr.as_ptr() as *const ([u8; 64], [u8; 52], u32, u32))
        };

        let mut name = try!(UTF_16LE.decode(&ab, DecoderTrap::Ignore).map_err(|e| e.to_string()));
        if let Some(len) = name.as_bytes().iter().position(|b| *b == 0) {
            name.truncate(len);
        }

        Ok(Directory {
            sect_start: sect_start,
            ul_size: ul_size,
            name: name,
        })
    }

}

fn read_variable_record<'a>(r: &mut &'a[u8]) -> Result<&'a[u8]> {
    let len = try!(r.read_u32::<LittleEndian>()) as usize;
    let (read, next) = r.split_at(len);
    *r = next;
    Ok(read)
}

fn check_variable_record<'a>(id: u16, r: &mut &'a[u8]) -> Result<&'a[u8]> {
    try!(check_record(id, r));
    let record = try!(read_variable_record(r));
    if log_enabled!(LogLevel::Warn) {
        if record.len() > 100_000 {
            warn!("record id {} as a suspicious huge length of {} (hex: {:x})", 
                  id, record.len(), record.len() as u32);
        }
    }
    Ok(record)
}

fn check_record(id: u16, r: &mut &[u8]) -> Result<()> {
    debug!("check record {:x}", id);
    let record_id = try!(r.read_u16::<LittleEndian>());
    if record_id != id {
        Err(format!("invalid record id, found {:x}, expecting {:x}", record_id, id).into())
    } else {
        Ok(())
    }
}

fn to_u32<'a>(s: &'a [u8]) -> impl Iterator<Item=u32> + 'a {
    s.chunks(4).map(|c| unsafe { ::std::ptr::read(c as *const [u8] as *const u32) })
}

/// A vba reference
#[derive(Debug, Clone, Hash, Eq, PartialEq)]
pub struct Reference {
    pub name: String,
    pub description: String,
    pub path: PathBuf,
}

impl Reference {
    pub fn is_missing(&self) -> bool {
        !self.path.as_os_str().is_empty() && !self.path.exists()
    }
}

/// A vba module
#[derive(Debug, Clone, Default)]
pub struct Module {
    pub name: String,
    stream_name: String,
    text_offset: usize,
}
