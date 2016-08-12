//! Parse vbaProject.bin file
//!
//! Retranscription from: 
//! https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py

use zip::read::ZipFile;
use std::io::{Read, Cursor};
use error::{ExcelResult, ExcelError};
use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;
use byteorder::{LittleEndian, ReadBytesExt};

const OLE_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;

#[allow(dead_code)]
pub struct VbaProject {
    directories: Vec<Directory>,
    sectors: Sector,
    mini_sectors: Option<Sector>,
}

impl VbaProject {

    pub fn new(mut f: ZipFile) -> ExcelResult<VbaProject> {

        // load header
        debug!("loading header");
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
        assert!(sector_size % 4 == 0);
        while sector_id != FREESECT && sector_id != ENDOFCHAIN {
            let mut rdr = Cursor::new(sector.get(sector_id));
            for _ in 0..sector_size / 4 {
                fat_sectors.push(try!(rdr.read_u32::<LittleEndian>()));
            }
            sector_id = fat_sectors.pop().unwrap(); //TODO: check if in infinite loop
        }

        // load the FATs
        debug!("load fat");
        let mut fat = Vec::with_capacity(fat_sectors.len() * sector_size);
        for sector_id in fat_sectors.into_iter().filter(|id| *id != FREESECT) {
            let mut rdr = Cursor::new(sector.get(sector_id));
            for _ in 0..sector_size / 4 {
                fat.push(try!(rdr.read_u32::<LittleEndian>()));
            }
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
            let buffer = sectors.read_chain(header.sect_mini_fat_start);
            let len = buffer.len() / 4;
            let mut rdr = Cursor::new(buffer);
            let mut minifat = Vec::with_capacity(len);
            for _ in 0..len {
                minifat.push(try!(rdr.read_u32::<LittleEndian>()));
            }

            let mini_sector_size = 2usize.pow(header.mini_sector_shift as u32);
            assert!(directories[0].ul_size as usize % mini_sector_size == 0);
            Some(Sector::new(ministream, mini_sector_size).with_fats(minifat))
        };

        Ok(VbaProject {
            directories: directories,
            sectors: sectors,
            mini_sectors: mini_sectors,
        })

    }

    pub fn get_stream(&self, name: &str) -> Option<usize> {
        for (i, d) in self.directories.iter().enumerate() {
            if let Ok(n) = d.get_name() {
                if &*n == name {
                    return Some(i);
                }
            } else {
                return None;
            }
        }
        None
    }
}

#[allow(dead_code)]
struct Header {
    ab_sig: [u8; 8],
    clid: [u8; 16],
    minor_version: u16,
    dll_version: u16,
    byte_order: u16,
    sector_shift: u16,
    mini_sector_shift: u16,
    reserved: u16,
    reserved1: u32,
    reserved2: u32,
    sect_fat_len: u32,
    sect_dir_start: u32,
    signature: u32,
    mini_sector_cutoff: u32,
    sect_mini_fat_start: u32,
    sect_mini_fat_len: u32,
    sect_dif_start: u32,
    sect_dif_len: u32,
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
            clid: clid,
            minor_version: minor_version,
            dll_version: dll_version,
            byte_order: byte_order,
            sector_shift: sector_shift,
            mini_sector_shift: mini_sector_shift,
            reserved: reserved,
            reserved1: reserved1,
            reserved2: reserved2,
            sect_fat_len: sect_fat_len,
            sect_dir_start: sect_dir_start,
            signature: signature,
            mini_sector_cutoff: mini_sector_cutoff,
            sect_mini_fat_start: sect_mini_fat_start,
            sect_mini_fat_len: sect_mini_fat_len,
            sect_dif_start: sect_dif_start,
            sect_dif_len: sect_dif_len,
            sect_fat: sect_fat,
        })
    }
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
        let mut buffer = Vec::new();
        while sector_id != ENDOFCHAIN {
            buffer.extend_from_slice(self.get(sector_id));
            sector_id = self.fats[sector_id as usize];
        }
        buffer
    }
}

#[allow(dead_code)]
pub struct Directory {
    ab: [u8; 64],
    cb: u16,
    mse: i8,
    flags: i8,
    id_left_sib: u32,
    id_right_sib: u32,
    id_child: u32,
    cls_id: [u8; 16],
    dw_user_flags: u32,
    time: [u64; 2],
    sect_start: u32,
    ul_size: u32,
    dpt_prop_type: u16,
}

impl Directory {

    fn from_slice(slice: &[u8]) -> ExcelResult<Directory> {
        let mut rdr = Cursor::new(slice);
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
            cb: cb,
            mse: mse,
            flags: flags,
            id_left_sib: id_left_sib,
            id_right_sib: id_right_sib,
            id_child: id_child,
            cls_id: cls_id,
            dw_user_flags: dw_user_flags,
            time: time,
            sect_start: sect_start,
            ul_size: ul_size,
            dpt_prop_type: dpt_prop_type,
        })

    }

    fn get_name(&self) -> ExcelResult<String> {
        let mut name = try!(UTF_16LE.decode(&self.ab, DecoderTrap::Ignore)
                            .map_err(ExcelError::FromUtf16));
        if let Some(len) = name.as_bytes().iter().position(|b| *b == 0) {
            name.truncate(len);
        }
        println!("{:?}", name);
        Ok(name)
    }
}

