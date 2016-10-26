use std::string::String;
use std::fs::File;
use std::io::{BufReader, Read};
use std::collections::HashMap;
use std::ptr;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use byteorder::{LittleEndian, ReadBytesExt};
use quick_xml::{XmlReader, Event, AsStr};
use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;

use {DataType, ExcelReader, Range, CellErrorType};
use vba::VbaProject;
use utils;
use errors::*;
 
pub struct Xlsb {
    zip: ZipArchive<File>,
}

impl Xlsb {
    fn iter<'a>(&'a mut self, path: &str) -> Result<RecordIter<'a>> {
        match self.zip.by_name(path) {
            Ok(f) => Ok(RecordIter { r: BufReader::new(f) }),
            Err(ZipError::FileNotFound) => Err(format!("file {} does not exist", path).into()),
            Err(e) => Err(e.into()),
        }
    }
}

impl ExcelReader for Xlsb {

    fn new(f: File) -> Result<Self> {
        Ok(Xlsb { zip: try!(ZipArchive::new(f)) })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<VbaProject> {
        let mut f = try!(self.zip.by_name("xl/vbaProject.bin"));
        let len = f.size() as usize;
        VbaProject::new(&mut f, len)
    }

    /// MS-XLSB
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let mut relationships = HashMap::new();
        match self.zip.by_name("xl/_rels/workbook.bin.rels") {
            Ok(f) => {
                let xml = XmlReader::from_reader(BufReader::new(f))
                    .with_check(false)
                    .trim_text(false);

                for res_event in xml {
                    match res_event {
                        Ok(Event::Start(ref e)) if e.name() == b"Relationship" => {
                            let mut id = Vec::new();
                            let mut target = String::new();
                            for a in e.attributes() {
                                match try!(a) {
                                    (b"Id", v) => id.extend_from_slice(v),
                                    (b"Target", v) => target = try!(v.as_str()).to_string(),
                                    _ => (),
                                }
                            }
                            relationships.insert(id, target);
                        }
                        Err(e) => return Err(e.into()),
                        _ => (),
                    }
                }
            },
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(relationships)
    }

    /// MS-XLSB 2.1.7.45
    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        let mut iter = try!(self.iter("xl/sharedStrings.bin"));
        let mut buf = vec![0; 1024];

        let _ = try!(iter.fill_next(0x009F, &mut buf)); // BrtBeginSst
        let len = try!((&mut &*buf).read_u32::<LittleEndian>()) as usize;
        let mut strings = Vec::with_capacity(len);

        // BrtSSTItems
        for _ in 0..len {
            let _ = try!(iter.next_skip_blocks(0x0013, &[
                                               (0x0023, Some(0x0024)) // future
                                               ], &mut buf)); // BrtSSTItem
            let flags = buf[0];
            if flags & 0b11000000 != 0 {
                // suppose A and B == 0
                return Err("only regular shared strings are supported".into());
            }
            strings.push(try!(wide_str(&buf[1..])));
        }
        Ok(strings)
    }

    /// MS-XLSB 2.1.7.61
    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>>
    {
        let mut iter = try!(self.iter("xl/workbook.bin"));
        let mut buf = vec![0; 1024];

        let _ = try!(iter.fill_next(0x0083, &mut buf)); // BrtBeginBook

        // BrtBeginBundleShs
        let _ = try!(iter.next_skip_blocks(0x008F, &[
                                          (0x0080, None),         // BrtFileVersion
                                          (0x0099, None),         // BrtWbProp
                                          (0x02A4, Some(0x0224)), // File Sharing
                                          (0x0025, Some(0x0026)), // AC blocks
                                          (0x02A5, Some(0x0216)), // Book protection(iso)
                                          (0x0087, Some(0x0088)), // BOOKVIEWS
                                          ], &mut buf)); 
        let mut sheets = HashMap::new();
        loop {
            match try!(iter.read_type()) {
                0x0090 => return Ok(sheets), // BrtEndBundleShs
                0x009C => (), // BrtBundleSh
                typ => return Err(format!("Expecting end of sheet, got {:x}", typ).into()),
            }
            let len = try!(iter.fill_buffer(&mut buf));
            let rel_len = utils::start_u32(&buf[8..len]);
            if rel_len != 0xFFFFFFFF {
                let rel_len = rel_len as usize * 2;
                let relid = &buf[12..12 + rel_len];
                // converts utf16le to utf8 for HashMap search
                let relid = try!(UTF_16LE.decode(relid, DecoderTrap::Ignore).map_err(|e| e.to_string()));
                let path = format!("xl/{}", relationships[relid.as_bytes()]);
                let name = try!(wide_str(&buf[12 + rel_len..len]));
                sheets.insert(name, path);
            }
        }
    }

    /// MS-XLSB 2.1.7.62
    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        let mut iter = try!(self.iter(path));
        let mut buf = vec![0; 1024];

        let _ = try!(iter.fill_next(0x0081, &mut buf)); // BrtBeginSheet

        // BrtWsDim
        let _ = try!(iter.next_skip_blocks(0x0094, &[
                                          (0x0093, None),         // BrtWsProp
                                          ], &mut buf)); 
        let (position, size) = unchecked_rfx(&buf[..16]);

        if size.0 == 0 || size.1 == 0 {
            return Ok(Range::default());
        }

        // BrtBeginSheetData
        let _ = try!(iter.next_skip_blocks(0x0091, &[
                                          (0x0085, Some(0x0086)), // Views
                                          (0x0025, Some(0x0026)), // AC blocks
                                          (0x01E5, None),         // BrtWsFmtInfo
                                          (0x0186, Some(0x0187)), // Col Infos
                                          ], &mut buf)); 

        let mut data = vec![DataType::Empty; (size.0 * size.1) as usize];
        
        // loop through all non empty rows
        loop {
            // BrtRowHdr
            let _ = try!(iter.next_skip_blocks(0x0000, &[
                                              (0x0025, Some(0x0026)), // AC blocks
                                              (0x0023, Some(0x0024)), // future
                                              ], &mut buf)); 
            let row = utils::start_u32(&buf);

            // get column indexes
            let cols = {
                let span_len = utils::start_u32(&buf[13..]) as usize;
                let mut span_iter = utils::to_u32(&buf[17..]).take(span_len * 2);
                let mut cols = Vec::with_capacity(size.1);
                for _ in 0..span_len {
                    cols.extend(span_iter.next().unwrap()..span_iter.next().unwrap());
                }
                cols
            };

            // read all values for the row
            for idx in cols.iter()
                .map(|col| (row - position.0) as usize * size.1 + (*col - position.1) as usize) {
                loop {
                    // read record
                    let typ = try!(iter.read_type());
                    let _ = try!(iter.fill_buffer(&mut buf));

                    match typ {
                        0x0092 => return Err("Expecting cell, got BrtEndSheetData".into()),
                        0x0001 => (), // BrtCellBlank: nothing to do as it is the default value
                        0x0002 => { // BrtCellRk MS-XLSB 2.5.122
                            let rk = utils::start_u32(&buf[8..12]);
                            let d100 = (rk & 0b10000000) == 0;
                            // TODO: use unchecked if too slow ...
                            data[idx] = if (rk & 0b01000000) == 0 {
                                let v = (rk & 0b00111111) as i64;
                                DataType::Int( if d100 { v } else { v/100 })
                            } else {
                                let v = ((rk & 0b00111111) as i64) << 34; 
                                let v = unsafe { ptr::read(&v as *const i64 as *const f64) };
                                DataType::Float( if d100 { v } else { v/100f64 })
                            };
                        },
                        0x0003 => { // BrtCellError
                            data[idx] = DataType::Error(match buf[8] {
                                0x00 => CellErrorType::Null,
                                0x07 => CellErrorType::Div0,
                                0x0F => CellErrorType::Value,
                                0x17 => CellErrorType::Ref,
                                0x1D => CellErrorType::Name,
                                0x24 => CellErrorType::Num,
                                0x2A => CellErrorType::NA,
                                0x2B => CellErrorType::GettingData,
                                c => return Err(format!("Unrecognised cell error code 0x{:x}", c).into()),
                            });
                        },
                        0x0004 => { // BrtCellBool
                            data[idx] = DataType::Bool(buf[8] != 0);
                        },
                        0x0005 => { // BrtCellReal
                            let v = unsafe { ptr::read(&buf[8..16] as *const [u8] as *const f64) };
                            data[idx] = DataType::Float(v);
                        },
                        0x0006 => { // BrtCellSt
                            data[idx] = DataType::String(try!(wide_str(&buf[8..])));
                        },
                        0x0007 => { // BrtCellIsst
                            let isst = utils::start_u32(&buf[8..12]) as usize;
                            data[idx] = DataType::String(strings[isst].clone());
                        },
                        _ => continue, // anything else, ignore and try next
                    }
                    break;
                }
            }

            if (row, *cols.last().unwrap()) == position {
                return Ok(Range::new(position, size, data))
            }
        }
    }

}

struct RecordIter<'a> {
    r: BufReader<ZipFile<'a>>,
}

impl<'a> RecordIter<'a> {

    fn read_type(&mut self) -> Result<u16> {
        let b = try!(self.r.read_u8());
        let typ = if (b & 0x80) == 0x80 {
            (b & 0x7F) as u16 + (((try!(self.r.read_u8()) & 0x7F) as u16)<<7)
        } else {
            b as u16
        };
        Ok(typ)
    }

    fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize> {
        let mut b = try!(self.r.read_u8());
        let mut len = b as usize;
        for i in 0..3 {
            if (b & 0x80) == 0 { break; }
            b = try!(self.r.read_u8());
            len += (b as usize) << (7 * i);
        } 
        if buf.len() < len { *buf = vec![0; len]; }

        let _ = try!(self.r.read_exact(&mut buf[..len]));
        Ok(len)
    }

    fn fill_next(&mut self, record_type: u16, buf: &mut Vec<u8>) -> Result<usize> {
        let typ = try!(self.read_type());
        if record_type != typ {
            Err(format!("Unexpected record: expecting 0x{:x}, found 0x{:x}", 
                        record_type, typ).into())
        } else {
            self.fill_buffer(buf)
        }
    }

    /// Reads next type, and discard block between `start` and `end`
    fn next_skip_blocks(&mut self, record_type: u16, bounds: &[(u16, Option<u16>)], 
                       buf: &mut Vec<u8>) -> Result<usize> 
    {
        let mut end = None;
        loop {
            let typ = try!(self.read_type());
            match end {
                Some(e) if e == typ => end = None,
                Some(_) => (),
                None if typ == record_type => {
                    return self.fill_buffer(buf);
                },
                None => match bounds.iter().position(|b| b.0 == typ) {
                    Some(i) => {
                        end = bounds[i].1;
                    },
                    None => return Err(format!("Unexpected record after block: \
                        expecting 0x{:x} found 0x{:x}", record_type, typ).into())
                }
            }
            let _ = try!(self.fill_buffer(buf));
        }
    }

}

fn wide_str(buf: &[u8]) -> Result<String> {
    let len = utils::start_u32(buf);
    let s = &buf[4..4 + len as usize * 2];
    UTF_16LE.decode(s, DecoderTrap::Ignore).map_err(|e| e.to_string().into())
}

fn unchecked_rfx(buf: &[u8]) -> ((u32, u32), (usize, usize)) {
    let mut iter = utils::to_u32(&buf[..16]);
    let rw_first = iter.next().unwrap();
    let rw_last = iter.next().unwrap();
    let col_first = iter.next().unwrap();
    let col_last = iter.next().unwrap();

    ((rw_first, col_first), 
     ((rw_last - rw_first + 1) as usize, (col_last - col_first + 1) as usize))
}

// fn nullable_wide_str(buf: &[u8]) -> Result<Option<String>> {
//     let len = utils::start_u32(buf);
//     if len == 0xFFFFFFFF {
//         Ok(None)
//     } else {
//         let s = &buf[4..4 + len as usize * 2];
//         let s = try!(UTF_16LE.decode(s, DecoderTrap::Ignore).map_err(|e| e.to_string()));
//         Ok(Some(s))
//     }
// }
