use std::fs::File;
use std::collections::HashMap;
use std::io::{BufReader, Read};

use errors::*;

use {ExcelReader, Range};
use vba::VbaProject;

/// A struct representing an old xls format file (CFB)
pub struct Xls {
    file: BufReader<File>,
}

impl ExcelReader for Xls {
    fn new(f: File) -> Result<Self> {
        Ok(Xls { file: BufReader::new(f), })
    }
    fn has_vba(&mut self) -> bool {
        true
    }
//     fn vba_project(&mut self) -> Result<VbaProject<BufReader<File>>> {
//         let len = try!(self.file.get_ref().metadata()).len() as usize;
//         VbaProject::new(&mut self.file, len)
//     }
    fn read_sheets_names(&mut self, _: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>> {
        unimplemented!()
    }
    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        unimplemented!()
    }
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        unimplemented!()
    }
    fn read_worksheet_range(&mut self, _: &str, _: &[String]) -> Result<Range> {
        unimplemented!()
    }
}

struct Stream {

}

struct SubStream {
    typ: u16,

}

struct Record {
    typ: u16,
    len: usize,
    data: Vec<u8>,
}


// /// A hidden struct which defines vba project
// struct Header {
//     ab_sig: [u8; 8],
//     sector_shift: u16,
//     mini_sector_shift: u16,
//     sect_dir_start: u32,
//     mini_sector_cutoff: u32,
//     sect_mini_fat_start: u32,
//     sect_dif_start: u32,
//     sect_fat: [u32; 109]
// }
// 
// impl Header {
//     fn from_reader<R: Read>(f: &mut R) -> Result<Header> {
// 
//         let mut ab_sig = [0; 8];
//         try!(f.read_exact(&mut ab_sig));
//         let mut clid = [0; 16];
//         try!(f.read_exact(&mut clid));
//         
//         let _minor_version = try!(f.read_u16::<LittleEndian>());
//         let _dll_version = try!(f.read_u16::<LittleEndian>());
//         let _byte_order = try!(f.read_u16::<LittleEndian>());
//         let sector_shift = try!(f.read_u16::<LittleEndian>());
//         let mini_sector_shift = try!(f.read_u16::<LittleEndian>());
//         let _reserved = try!(f.read_u16::<LittleEndian>());
//         let _reserved1 = try!(f.read_u32::<LittleEndian>());
//         let _reserved2 = try!(f.read_u32::<LittleEndian>());
//         let _sect_fat_len = try!(f.read_u32::<LittleEndian>());
//         let sect_dir_start = try!(f.read_u32::<LittleEndian>());
//         let _signature = try!(f.read_u32::<LittleEndian>());
//         let mini_sector_cutoff = try!(f.read_u32::<LittleEndian>());
//         let sect_mini_fat_start = try!(f.read_u32::<LittleEndian>());
//         let _sect_mini_fat_len = try!(f.read_u32::<LittleEndian>());
//         let sect_dif_start = try!(f.read_u32::<LittleEndian>());
//         let _sect_dif_len = try!(f.read_u32::<LittleEndian>());
// 
//         let mut sect_fat = [0u8; 109 * 4];
//         try!(f.read_exact(&mut sect_fat));
//         let sect_fat = unsafe { *(&sect_fat as *const [u8; 109 * 4] as *const [u32; 109]) };
// 
//         Ok(Header {
//             ab_sig: ab_sig, 
//             sector_shift: sector_shift,
//             mini_sector_shift: mini_sector_shift,
//             sect_dir_start: sect_dir_start,
//             mini_sector_cutoff: mini_sector_cutoff,
//             sect_mini_fat_start: sect_mini_fat_start,
//             sect_dif_start: sect_dif_start,
//             sect_fat: sect_fat,
//         })
//     }
// }
// 
// /// To better understand what's happening, look
// /// [here](http://www.wordarticles.com/Articles/Formats/StreamCompression.php)
// fn decompress_stream<I: Iterator<Item=u8>>(mut r: I) -> Result<Vec<u8>> {
//     debug!("decompress stream");
//     let mut res = Vec::new();
// 
//     match r.next() {
//         Some(0x01) => (),
//         _ => return Err("invalid signature byte".into()),
//     }
// 
//     fn read_u16<J: Iterator<Item=u8>>(i: &mut J) -> Option<u16> {
//         match (i.next(), i.next()) {
//             (Some(i1), Some(i2)) => (&[i1, i2] as &[u8]).read_u16::<LittleEndian>().ok(),
//             _ => None,
//         }
//     }
// 
//     fn ok_or<T>(o: Option<T>, err: &'static str) -> Result<T> {
//         match o {
//             Some(o) => Ok(o),
//             None => Err(err.into()),
//         }
//     }
// 
//     while let Some(chunk_header) = read_u16(&mut r) {
// 
//         // each 'chunk' is 4096 wide, let's reserve that space
//         let start = res.len();
//         res.reserve(4096);
// 
//         let chunk_size = chunk_header & 0x0FFF;
//         let chunk_signature = (chunk_header & 0x7000) >> 12;
//         let chunk_flag = (chunk_header & 0x8000) >> 15;
// 
//         assert_eq!(chunk_signature, 0b011);
// 
//         if chunk_flag == 0 { // uncompressed
//             res.extend(r.by_ref().take(4096));
//         } else {
// 
//             let mut chunk_len = 0;
//             let mut buf = [0u8; 4096];
//             'chunk: loop {
// 
//                 let bit_flags: u8 = try!(ok_or(r.next(), "no bit_flag in compressed stream"));
//                 chunk_len += 1;
// 
//                 for bit_index in 0..8 {
// 
//                     if chunk_len > chunk_size { break 'chunk; }
// 
//                     if (bit_flags & (1 << bit_index)) == 0 {
//                         // literal token
//                         res.push(try!(ok_or(r.next(), "no literal token in compressed stream")));
//                         chunk_len += 1;
//                     } else {
//                         // copy token
//                         let token = try!(ok_or(read_u16(&mut r), "no literal token in compressed stream"));
//                         chunk_len += 2;
// 
//                         let decomp_len = res.len() - start;
//                         let bit_count = (4..16).find(|i| POWER_2[*i] >= decomp_len).unwrap();
//                         let len_mask = 0xFFFF >> bit_count;
//                         let mut len = (token & len_mask) as usize + 3;
//                         let offset = ((token & !len_mask) >> (16 - bit_count)) as usize + 1;
// 
//                         while len > offset {
//                             buf[..offset].copy_from_slice(&res[res.len() - offset..]);
//                             res.extend_from_slice(&buf[..offset]);
//                             len -= offset;
//                         }
//                         buf[..len].copy_from_slice(&res[res.len() - offset..res.len() - offset + len]);
//                         res.extend_from_slice(&buf[..len]);
//                     }
//                 }
//             }
//         }
//     }
//     Ok(res)
// 
// }
// 
// /// A struct corresponding to the elementary block of memory
// struct Sector {
//     data: Vec<u8>,
//     size: usize,
//     fats: Vec<u32>,
// }
// 
// impl Sector {
// 
//     fn new(data: Vec<u8>, size: usize) -> Sector {
//         assert!(data.len() % size == 0);
//         Sector {
//             data: data,
//             size: size as usize,
//             fats: Vec::new(),
//         }
//     }
// 
//     fn with_fats(mut self, fats: Vec<u32>) -> Sector {
//         self.fats = fats;
//         self
//     }
// 
//     fn get(&self, id: u32) -> &[u8] {
//         &self.data[id as usize * self.size .. (id as usize + 1) * self.size]
//     }
// 
//     fn read_chain(&self, sector_id: u32) -> SectorChain {
//         debug!("chain reading sector {}", sector_id);
//         SectorChain {
//             sector: self,
//             id: sector_id,
//         }
//     }
// 
// }
// 
// /// An iterator over `Sector`s which represents the final memory allocated
// struct SectorChain<'a> {
//     sector: &'a Sector,
//     id: u32,
// }
// 
// impl<'a> Iterator for SectorChain<'a> {
//     type Item = &'a [u8];
//     fn next(&mut self) -> Option<&'a [u8]> {
//         debug!("read chain next, id: {}", self.id);
//         if self.id == ENDOFCHAIN {
//             None
//         } else {
//             let sector = self.sector.get(self.id);
//             self.id = self.sector.fats[self.id as usize];
//             Some(sector)
//         }
//     }
// }
// 
// struct StreamIter<'a> {
//     len: usize,
//     cur_len: usize,
//     chain: SectorChain<'a>,
//     current: ::std::iter::Cloned<::std::slice::Iter<'a, u8>>,
// }
// 
// impl<'a> Iterator for StreamIter<'a> {
//     type Item=u8;
//     fn next(&mut self) -> Option<u8> {
//         if self.cur_len == self.len {
//             return None;
//         }
//         self.cur_len += 1;
//         self.current.next().or_else(|| {
//             self.chain.next().and_then(|c| {
//                 self.current = c.iter().cloned();
//                 self.current.next()
//             })
//         })
//     }
// }
// 
// /// A struct representing sector organizations, behaves similarly to a tree
// struct Directory {
//     sect_start: u32,
//     ul_size: u32,
//     name: String,
// }
// 
// impl Directory {
// 
//     fn from_slice(rdr: &[u8]) -> Result<Directory> {
// //         ab: [u8; 64],
// //         _cb: u16,                    |
// //         _mse: i8,                    |
// //         _flags: i8,                  |
// //         _id_left_sib: u32,           |
// //         _id_right_sib: u32,          | = [u8; 52] with padding alignment
// //         _id_child: u32,              |
// //         _cls_id: [u8; 16],           |
// //         _dw_user_flags: u32,         |
// //         _time: [u64; 2],             |
// //         sect_start: u32,
// //         ul_size: u32,
// //         _dpt_prop_type: u16,
//         assert_eq!(rdr.len(), 128);
//         let (ab, _, sect_start, ul_size) = unsafe { 
//             ::std::ptr::read(rdr.as_ptr() as *const ([u8; 64], [u8; 52], u32, u32))
//         };
// 
//         let mut name = try!(UTF_16LE.decode(&ab, DecoderTrap::Ignore).map_err(|e| e.to_string()));
//         if let Some(len) = name.as_bytes().iter().position(|b| *b == 0) {
//             name.truncate(len);
//         }
// 
//         Ok(Directory {
//             sect_start: sect_start,
//             ul_size: ul_size,
//             name: name,
//         })
//     }
// 
// }

/// reads slice in LittleEndian
fn read_slice<T>(s: &[u8]) -> T {
    unsafe { ::std::ptr::read(&s[..::std::mem::size_of::<T>()] as *const [u8] as *const T) }
}

fn read_u16(s: &[u8]) -> u16 {
    read_slice(s)
}

fn read_u32(s: &[u8]) -> u32 {
    read_slice(s)
}

fn read_usize(s: &[u8]) -> usize {
    read_u32(s) as usize
}
