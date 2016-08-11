//! Parse vbaProject.bin file
//!
//! Retranscription from: 
//! https://github.com/unixfreak0037/officeparser/blob/master/officeparser.py

use zip::read::ZipFile;
use std::io::{Read};
use error::{ExcelResult, ExcelError};
use std::mem;

const OLE_SIGNATURE: &'static str = r#"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"#;
// const DIFSECT: u32 = 0xFFFFFFFC;
// const FATSECT: u32 = 0xFFFFFFFD;
const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;
// const MODULE_EXTENSION: &'static str = "bas";
// const CLASS_EXTENSION: &'static str = "cls";
// const FORM_EXTENSION: &'static str = "frm";
// const BINFILE_NAME: &'static str = "/vbaProject.bin";

fn get_sector(data: &[u8], id: u32, sector_count: usize, sector_size: usize) -> ExcelResult<&[u8]> {
    if id as usize >= sector_count {
        return Err(ExcelError::Unexpected(format!("reference to invalid sector {}", id)));
    }
    Ok(&data[id as usize * sector_size .. (id as usize + 1) * sector_size])
}

fn read_signature<R: Read>(f: &mut R) -> ExcelResult<bool> {
    let mut sig = [0; 8];
    try!(f.read_exact(&mut sig));
    Ok(sig.as_ref() == OLE_SIGNATURE.as_bytes())
}

pub struct VbaProject {
    directories: Vec<Directory>,
    sectors: Sector,
    mini_sectors: Option<Sector>,
}

impl VbaProject {

    pub fn new(mut f: ZipFile) -> ExcelResult<VbaProject> {

        try!(read_signature(&mut f));

        // load header
        let header = try!(Header::from_reader(&mut f));

        let sector_size = 2u64.pow(header.sector_shift as u32) as usize;
        
        if (f.size() as usize - 512) % sector_size != 0 {
            return Err(ExcelError::Unexpected("last sector has invalid size".to_string()));
        }

        // Read whole file in memory (the file is delimited by sectors)
        let sector_count = (f.size() as usize - 512) / sector_size;
        let mut data = Vec::with_capacity(f.size() as usize - 512);
        try!(f.read_to_end(&mut data));

        // load fat and dif sectors
        let mut fat_sectors = header.sect_fat.to_vec();
        let mut sector_id = header.sect_dif_start;
        while sector_id != FREESECT && sector_id != ENDOFCHAIN {
            let sector: &[u32] = unsafe {
                mem::transmute(try!(get_sector(&data, sector_id, sector_count, sector_size)))
            };
            fat_sectors.extend_from_slice(&sector[..sector.len() - 1]);
            sector_id = sector[sector.len() - 1];
            //TODO: check if that sector has been read already (infinite loop)
        }

        // load the FATs
        let mut fat = Vec::with_capacity(fat_sectors.len() * sector_size);
        for sector_id in fat_sectors.into_iter().filter(|id| *id != FREESECT) {
            fat.extend_from_slice(try!(get_sector(&data, sector_id, sector_count, sector_size)));
        }
        let fat: Vec<u32> = unsafe { mem::transmute(fat) };
        
        // create sector reader
        let sectors = Sector::new(data, sector_count, sector_size, fat);

        // get the list of directory sectors
        let directories: Vec<Directory> = unsafe { 
            mem::transmute(try!(sectors.read_chain(header.sect_dir_start))) 
        };

        // load the mini streams
        let mini_sectors = if directories[0].sect_start == ENDOFCHAIN {
            None
        } else {
            let mut ministream = try!(sectors.read_chain(directories[0].sect_start)); 
            ministream.truncate(directories[0].ul_size as usize); // should not be needed

            let minifat: Vec<u32> = unsafe {
                mem::transmute(try!(sectors.read_chain(header.sect_mini_fat_start)))
            };
            let mini_sector_size = 2usize.pow(header.mini_sector_shift as u32);
            let mini_sector_count = ministream.len() / mini_sector_size;
            Some(Sector::new(ministream, mini_sector_count, mini_sector_size, minifat))
        };

        Ok(VbaProject {
            directories: directories,
            sectors: sectors,
            mini_sectors: mini_sectors,
        })

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
        let mut data = [0; 512];
        try!(f.read_exact(&mut data));
        Ok(unsafe { mem::transmute(data) })
    }
}

struct Sector {
    data: Vec<u8>,
    size: usize,
    count: usize,
    fats: Vec<u32>,
}

impl Sector {

    fn new(data: Vec<u8>, count: usize, size: usize, fats: Vec<u32>) -> Sector {
        Sector {
            data: data,
            size: size as usize,
            count: count as usize,
            fats: fats,
        }
    }

    fn read_chain(&self, mut sector_id: u32) -> ExcelResult<Vec<u8>> {
        let mut buffer = Vec::new();
        while sector_id != ENDOFCHAIN {
            let sector = try!(get_sector(&self.data, sector_id, self.count, self.size));
            buffer.extend_from_slice(sector);
            sector_id = self.fats[sector_id as usize];
        }
        Ok(buffer)
    }

}

#[allow(dead_code)]
struct Directory {
    ab: [u8; 64],
    cb: u16,
    mse: i8,
    flags: i8,
    id_left_sib: u32,
    id_right_sib: u32,
    id_child: u32,
    cls_id: [u8; 16],
    dw_user_flags: u32,
    time: (u64, u64),
    sect_start: u32,
    ul_size: u32,
    dpt_prop_type: u16,
    padding: [u8; 2],
}

impl Directory {
//     fn get_name(&self) -> {
//         self.name = ''.join([x for x in self._ab[0:self._cb] if ord(x) != 0])
//     }
}

