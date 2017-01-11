//! Coumpound File Binary format MS-CFB

use std::borrow::Cow;
use std::cmp::min;
use std::io::Read;

use encoding::{Encoding};
use encoding::all::UTF_16LE;

use encoding::{DecoderTrap, EncodingRef, StringWriter};
use encoding::label::encoding_from_windows_code_page;

use errors::*;
use utils::*;

const ENDOFCHAIN: u32 = 0xFFFFFFFE;
const FREESECT: u32 = 0xFFFFFFFF;

/// A struct for managing Compound File Binary format
#[derive(Debug, Clone)]
pub struct Cfb {
    directories: Vec<Directory>,
    sectors: Sectors,
    fats: Vec<u32>,
    mini_sectors: Sectors,
    mini_fats: Vec<u32>,
}

impl Cfb {
    /// Create a new `Cfb`
    ///
    /// Starts reading project metadata (header, directories, sectors and minisectors).
    pub fn new<R: Read>(mut f: &mut R, len: usize) -> Result<Cfb> {

        // load header
        let (h, mut difat) = Header::from_reader(&mut f)?;
        let mut sectors = Sectors::new(h.sector_size, Vec::with_capacity(len));

        // load fat and dif sectors
        debug!("load difat");
        let mut sector_id = h.difat_start;
        while sector_id != FREESECT && sector_id != ENDOFCHAIN {
            difat.extend_from_slice(to_u32(sectors.get(sector_id, f)?));
            sector_id = difat.pop().unwrap(); //TODO: check if in infinite loop
        }

        // load the FATs
        debug!("load fat");
        let mut fats = Vec::with_capacity(h.fat_len);
        for id in difat.into_iter().filter(|id| *id != FREESECT) {
            fats.extend_from_slice(to_u32(sectors.get(id, f)?));
        }

        // get the list of directory sectors
        debug!("load directories");
        let dirs = sectors.get_chain(h.dir_start, &fats, f, h.dir_len * h.sector_size)?;
        let dirs = dirs.chunks(128).map(|c| Directory::from_slice(c, h.sector_size))
            .collect::<Result<Vec<_>>>()?;

        if dirs.is_empty() || (h.version != 3 && dirs[0].start == ENDOFCHAIN) {
            return Err("Unexpected empty root directory".into());
        }
        debug!("{:?}", dirs);

        // load the mini streams
        debug!("load minis");
        let ministream = sectors.get_chain(dirs[0].start, &fats, f, dirs[0].len)?;
        let minifat = sectors.get_chain(h.mini_fat_start, &fats, f, h.mini_fat_len * h.sector_size)?;
        let minifat = to_u32(&minifat).to_vec();
        Ok(Cfb {
            directories: dirs,
            sectors: sectors,
            fats: fats,
            mini_sectors: Sectors::new(64, ministream),
            mini_fats: minifat,
        })
    }

    /// Checks if directory exists
    pub fn has_directory(&self, name: &str) -> bool {
        self.directories.iter().any(|d| &*d.name == name)
    }

    /// Gets a stream by name out of directories
    pub fn get_stream<R: Read>(&mut self, name: &str, r: &mut R) -> Result<Vec<u8>> {
        match self.directories.iter().find(|d| &*d.name == name) {
            None => Err(format!("Cannot find {} stream", name).into()),
            Some(d) => {
                if d.len < 4096 {
                    // TODO: Study the possibility to return a `VecArray` (stack allocated)
                    self.mini_sectors.get_chain(d.start, &self.mini_fats, r, d.len)
                } else {
                    self.sectors.get_chain(d.start, &self.fats, r, d.len)
                }
            }
        }
    }
}

/// A hidden struct which defines cfb files structure
#[derive(Debug)]
struct Header {
    version: u16,
    sector_size: usize,
    dir_len: usize,
    dir_start: u32,
    fat_len: usize,
    mini_fat_len: usize,
    mini_fat_start: u32,
    difat_start: u32,
}

impl Header {
    fn from_reader<R: Read>(f: &mut R) -> Result<(Header, Vec<u32>)> {
        let mut buf = [0u8; 512];
        f.read_exact(&mut buf)?;

        // check ole signature
        if read_slice::<u64>(buf.as_ref()) != 0xE11AB1A1E011CFD0 {
            return Err("invalid OLE signature (not an office document?)".into());
        }

        let version = read_u16(&buf[26..28]);

        let sector_size = match read_u16(&buf[30..32]) {
            0x0009 => 512,
            0x000C => {
                // sector size is 4096 bytes, but header is 512 bytes,
                // so the remaining sector bytes have to be read
                let mut buf_end = [0u8; 3584];
                f.read_exact(&mut buf_end)?;
                4096
            },
            s => return Err(format!("Invalid sector shift, expecting 0x09 \
                                     or 0x0C, got {:x}", s).into()),
        };

        if read_u16(&buf[32..34]) != 0x0006 {
            return Err("Invalid minisector shift".into());
        }

        let dir_len = read_usize(&buf[40..44]);
        let fat_len = read_usize(&buf[44..48]);
        let dir_start = read_u32(&buf[48..52]);
        let mini_fat_start = read_u32(&buf[60..64]);
        let mini_fat_len = read_usize(&buf[64..68]);
        let difat_start = read_u32(&buf[68..72]);
        let difat_len = read_usize(&buf[62..76]);

        let mut difat = Vec::with_capacity(difat_len);
        difat.extend_from_slice(to_u32(&buf[76..512]));

        Ok((Header {
            version: version,
            sector_size: sector_size,
            dir_len: dir_len,
            fat_len: fat_len,
            dir_start: dir_start,
            mini_fat_len: mini_fat_len,
            mini_fat_start: mini_fat_start,
            difat_start: difat_start,
        }, difat))
    }
}

/// A struct corresponding to the elementary block of memory
///
/// `data` will persist in memory to ensure the file is read once
#[derive(Debug, Clone)]
struct Sectors {
    data: Vec<u8>,
    size: usize,
}

impl Sectors {

    fn new(size: usize, data: Vec<u8>) -> Sectors {
        Sectors {
            data: data,
            size: size,
        }
    }

    fn get<R: Read>(&mut self, id: u32, r: &mut R) -> Result<&[u8]> {
        let start = id as usize * self.size;
        let end = start + self.size;
        if end > self.data.len() {
            let len = self.data.len();
            unsafe { self.data.set_len(end); }
            r.read_exact(&mut self.data[len..end])?;
        }
        Ok(&self.data[start..end])
    }

    fn get_chain<R: Read>(&mut self, mut sector_id: u32, fats: &[u32], 
                          r: &mut R, len: usize) -> Result<Vec<u8>> {
        let mut chain = if len > 0 { Vec::with_capacity(len) } else { Vec::new() };
        while sector_id != ENDOFCHAIN {
            chain.extend_from_slice(self.get(sector_id, r)?);
            sector_id = fats[sector_id as usize];
        }
        if len > 0 { chain.truncate(len); }
        Ok(chain)
    }

}

/// A struct representing sector organizations, behaves similarly to a tree
#[derive(Debug, Clone)]
struct Directory {
    name: String,
    start: u32,
    len: usize,
}

impl Directory {
    fn from_slice(buf: &[u8], sector_size: usize) -> Result<Directory> {
        let mut name = UTF_16LE.decode(&buf[..64], DecoderTrap::Ignore)
                            .map_err(|e| e.to_string())?;
        if let Some(l) = name.as_bytes().iter().position(|b| *b == 0) {
            name.truncate(l);
        }
        let start = read_u32(&buf[116..120]);
        let len = if sector_size == 512 {
            read_slice::<u32>(&buf[120..124]) as usize
        } else {
            read_slice::<u64>(&buf[120..128]) as usize
        };

        Ok(Directory {
            start: start,
            len: len,
            name: name,
        })
    }
}

/// Decompresses stream
pub fn decompress_stream(s: &[u8]) -> Result<Vec<u8>> {

    const POWER_2: [usize; 16] = [1   , 1<<1, 1<<2,  1<<3,  1<<4,  1<<5,  1<<6,  1<<7, 
                                  1<<8, 1<<9, 1<<10, 1<<11, 1<<12, 1<<13, 1<<14, 1<<15];

    debug!("decompress stream");
    let mut res = Vec::new();

    if s[0] != 0x01 {
        return Err("invalid signature byte".into());
    }

    let mut i = 1;
    while i < s.len() {

        let chunk_header = read_u16(&s[i..]);
        i += 2;

        // each 'chunk' is 4096 wide, let's reserve that space
        let start = res.len();
        res.reserve(4096);

        let chunk_size = chunk_header & 0x0FFF;
        let chunk_signature = (chunk_header & 0x7000) >> 12;
        let chunk_flag = (chunk_header & 0x8000) >> 15;

        assert_eq!(chunk_signature, 0b011, "i={}, len={}", i, s.len());

        if chunk_flag == 0 { // uncompressed
            res.extend_from_slice(&s[i..i + 4096]);
            i += 4096;
        } else {

            let mut chunk_len = 0;
            let mut buf = [0u8; 4096];
            'chunk: loop {

                let bit_flags = s[i];
                i += 1;
                chunk_len += 1;

                for bit_index in 0..8 {

                    if chunk_len > chunk_size { break 'chunk; }

                    if (bit_flags & (1 << bit_index)) == 0 {
                        // literal token
                        res.push(s[i]);
                        i += 1;
                        chunk_len += 1;
                    } else {
                        // copy token
                        let token = read_u16(&s[i..]);
                        i += 2;
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

#[derive(Clone)]
pub struct XlsEncoding {
    encoding: EncodingRef,
    pub high_byte: Option<bool>, // None if single byte encoding
}

impl XlsEncoding {
    pub fn from_codepage(codepage: u16) -> Result<XlsEncoding> {
        let e = match encoding_from_windows_code_page(codepage as usize) {
            Some(e) => e,
            None => return Err(format!("Cannot find {} codepage", codepage).into()),
        };
        let high_byte = match codepage {
            20127 | 65000 | 65001 | 20866 | 21866 | 10000 | 10007 | 
            874 | 1250...1258 | 28591...28605 => None, // SingleByte encodings
            _ => Some(false),
        };

        Ok(XlsEncoding { encoding: e, high_byte: high_byte })
    }

    pub fn decode_to(&self, stream: &[u8], len: usize, s: &mut StringWriter) 
        -> Result<(usize, usize)> 
    {
        let (l, ub, bytes) = match self.high_byte {
            None => {
                let l = min(stream.len(), len);
                (l, l, Cow::Borrowed(&stream[..l]))
            },
            Some(false) => {
                let l = min(stream.len(), len);

                // add 0x00 high bytes to unicodes
                let mut bytes = vec![0; l * 2];
                for (i, sce) in stream.iter().take(l).enumerate() {
                    bytes[2 * i] = *sce;
                }
                (l, l, Cow::Owned(bytes))
            },
            Some(true) => {
                let l = min(stream.len() / 2, len);
                (l, 2*l, Cow::Borrowed(&stream[..2*l]))
            }
        };
        
        self.encoding.decode_to(&bytes, DecoderTrap::Ignore, s)
            .map_err(|e| e.to_string().into())
            .map(|_| (l, ub))
    }

    pub fn decode_all(&self, stream: &[u8]) -> Result<String> {
        let mut s = String::with_capacity(stream.len());
        let _ = self.decode_to(stream, stream.len(), &mut s)?;
        Ok(s)
    }
}
