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
use utils;
use cfb::Cfb;

const OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;

const POWER_2: [usize; 16] = [1   , 1<<1, 1<<2,  1<<3,  1<<4,  1<<5,  1<<6,  1<<7, 
                              1<<8, 1<<9, 1<<10, 1<<11, 1<<12, 1<<13, 1<<14, 1<<15];

/// A struct for managing VBA reading
#[allow(dead_code)]
pub struct VbaProject {
    cfb: Cfb,
}

impl VbaProject {

    /// Create a new `VbaProject` out of the vbaProject.bin `ZipFile` or xls file
    ///
    /// Starts reading project metadata (header, directories, sectors and minisectors).
    pub fn new<R: Read>(f: &mut R, len: usize) -> Result<VbaProject> {
        Cfb::new(f, len).map(|cfb| VbaProject { cfb: cfb, })
    }

    /// Reads project `Reference`s and `Module`s
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut vba = Excel::open(path)
    ///     .and_then(|mut xl| xl.vba_project())
    ///     .expect("Cannot read vba project");
    /// let (references, modules) = vba.read_vba().unwrap();
    /// println!("References: {:?}", references);
    /// println!("Modules: {:?}", modules);
    /// ```
    fn read_vba<R: Read>(&mut self, r: &mut R) -> Result<(Vec<Reference>, Vec<Module>)> {
        debug!("read vba");
        
        // dir stream
        let mut stream = try!(self.cfb.get_stream("dir", r));

        // read header (not used)
        try!(self.read_dir_header(&mut &*stream));

        // array of REFERENCE records
        let references = try!(self.read_references(&mut &*stream));

        // modules
        let modules = try!(self.read_modules(&mut &*stream));
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

                    let name = try!(read_variable_record(stream, 1));
                    let name = try!(::std::string::String::from_utf8(name.to_vec()));
                    reference = Reference {
                        name: name.clone(),
                        description: name.clone(),
                        path: "/".into(),
                    };

                    try!(check_variable_record(0x003E, stream)); // unicode
                },

                0x0033 => { // REFERENCEORIGINAL (followed by REFERENCECONTROL)
                    try!(read_variable_record(stream, 1));
                },

                0x002F => { // REFERENCECONTROL
                    *stream = &stream[4..]; // len of total ref control

                    let libid = try!(read_variable_record(stream, 1)); //libid twiddled
                    try!(set_module_from_libid(&mut reference, libid));

                    *stream = &stream[6..];

                    match try!(stream.read_u16::<LittleEndian>()) {
                        0x0016 => { // optional name record extended
                            try!(read_variable_record(stream, 1)); // name extended
                            try!(check_variable_record(0x003E, stream)); // name extended unicode
                            try!(check_record(0x0030, stream));
                        },
                        0x0030 => (),
                        e => return Err(format!( "unexpected token in reference control {:x}", e).into()),
                    } 
                    *stream = &stream[4..];
                    try!(read_variable_record(stream, 1)); // libid extended
                    *stream = &stream[26..];
                },

                0x000D => { // REFERENCEREGISTERED
                    *stream = &stream[4..];

                    let libid = try!(read_variable_record(stream, 1)); // libid registered
                    try!(set_module_from_libid(&mut reference, libid));

                    *stream = &stream[6..];
                },

                0x000E => { // REFERENCEPROJECT
                    *stream = &stream[4..];
                    let absolute = try!(read_variable_record(stream, 1)); // project libid absolute
                    {
                        let absolute = try!(::std::str::from_utf8(absolute));
                        reference.path = if absolute.starts_with("*\\C") { 
                            absolute[3..].into()
                        } else {
                            absolute.into()
                        };
                    }
                    try!(read_variable_record(stream, 1)); // project libid relative
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
                0x0021 /* procedural module */ |
                0x0022 /* document, class or designer module */ => (),
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
    ///
    /// # Examples
    /// ```
    /// use office::Excel;
    ///
    /// # let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    /// let mut vba = Excel::open(path)
    ///     .and_then(|mut xl| xl.vba_project())
    ///     .expect("Cannot read vba project");
    /// let (_, modules) = vba.read_vba().unwrap();
    /// for m in modules {
    ///     println!("Module {}:", m.name);
    ///     println!("{}", vba.read_module(&m)
    ///                       .expect(&format!("cannot read {:?} module", m)));
    /// }
    /// ```
    pub fn read_module<R: Read>(&mut self, module: &Module, r: &mut R) -> Result<String> {
        debug!("read module {}", module.name);
        let data = try!(self.read_module_raw(module, r));
        let data = try!(::std::string::String::from_utf8(data));
        Ok(data)
    }

    /// Reads module content (MBSC encoded) and output it as-is (binary output)
    pub fn read_module_raw<R: Read>(&mut self, module: &Module, r: &mut R) -> Result<Vec<u8>> {
        debug!("read module raw {}", module.name);
        self.cfb.get_stream(&module.stream_name, r)
//         match self.cfb.get_stream(&module.stream_name, r) {
//             None => Err(format!("cannot find {} stream", module.stream_name).into()),
//             Some(s) => {
//                 let data = try!(decompress_stream(s.skip(module.text_offset)));
//                 Ok(data)
//             }
//         }
    }

}

/// A hidden struct which defines vba project
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
        let sect_fat = unsafe { *(&sect_fat as *const [u8; 109 * 4] as *const [u32; 109]) };

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

/// To better understand what's happening, look
/// [here](http://www.wordarticles.com/Articles/Formats/StreamCompression.php)
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
            let mut buf = [0u8; 4096];
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
                        let mut len = (token & len_mask) as usize + 3;
                        let offset = ((token & !len_mask) >> (16 - bit_count)) as usize + 1;

                        while len > offset {
                            buf[..offset].copy_from_slice(&res[res.len() - offset..]);
                            res.extend_from_slice(&buf[..offset]);
                            len -= offset;
                        }
                        buf[..len].copy_from_slice(&res[res.len() - offset..res.len() - offset + len]);
                        res.extend_from_slice(&buf[..len]);
                    }
                }
            }
        }
    }
    Ok(res)

}

/// A struct corresponding to the elementary block of memory
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

/// An iterator over `Sector`s which represents the final memory allocated
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

struct StreamIter<'a> {
    len: usize,
    cur_len: usize,
    chain: SectorChain<'a>,
    current: ::std::iter::Cloned<::std::slice::Iter<'a, u8>>,
}

impl<'a> Iterator for StreamIter<'a> {

    type Item=u8;
    fn next(&mut self) -> Option<u8> {
        if self.cur_len == self.len {
            return None;
        }
        self.cur_len += 1;
        self.current.next().or_else(|| {
            self.chain.next().and_then(|c| {
                self.current = c.iter().cloned();
                self.current.next()
            })
        })
    }
}

/// A struct representing sector organizations, behaves similarly to a tree
struct Directory {
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
        !self.path.as_os_str().is_empty() && !self.path.exists()
    }
}

/// A vba module
#[derive(Debug, Clone, Default)]
pub struct Module {
    /// module name as it appears in vba project
    pub name: String,
    stream_name: String,
    text_offset: usize,
}

/// Reads a variable length record
/// 
/// `mult` is a multiplier of the length (e.g 2 when parsing XLWideString)
fn read_variable_record<'a>(r: &mut &'a[u8], mult: usize) -> Result<&'a[u8]> {
    let len = try!(r.read_u32::<LittleEndian>()) as usize * mult;
    let (read, next) = r.split_at(len);
    *r = next;
    Ok(read)
}

/// Check that next record matches `id` and returns a variable length record
fn check_variable_record<'a>(id: u16, r: &mut &'a[u8]) -> Result<&'a[u8]> {
    try!(check_record(id, r));
    let record = try!(read_variable_record(r, 1));
    if log_enabled!(LogLevel::Warn) && record.len() > 100_000 {
        warn!("record id {} as a suspicious huge length of {} (hex: {:x})", 
              id, record.len(), record.len() as u32);
    }
    Ok(record)
}

/// Check that next record matches `id`
fn check_record(id: u16, r: &mut &[u8]) -> Result<()> {
    debug!("check record {:x}", id);
    let record_id = try!(r.read_u16::<LittleEndian>());
    if record_id != id {
        Err(format!("invalid record id, found {:x}, expecting {:x}", record_id, id).into())
    } else {
        Ok(())
    }
}
