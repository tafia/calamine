use std::string::String;
use std::fs::File;
use std::io::{BufReader, Read};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::{XmlReader, Event, AsStr};
use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;

use {DataType, ExcelReader, Cell, Range, CellErrorType};
use vba::VbaProject;
use utils::{read_u32, read_usize, read_slice};
use errors::*;

pub struct Xlsb {
    zip: ZipArchive<File>,
}

impl Xlsb {
    fn iter<'a>(&'a mut self, path: &str) -> Result<RecordIter<'a>> {
        match self.zip.by_name(path) {
            Ok(f) => {
                Ok(RecordIter {
                    r: BufReader::new(f),
                    b: [0],
                })
            }
            Err(ZipError::FileNotFound) => Err(format!("file {} does not exist", path).into()),
            Err(e) => Err(e.into()),
        }
    }
}

impl ExcelReader for Xlsb {
    fn new(f: File) -> Result<Self> {
        Ok(Xlsb { zip: ZipArchive::new(f)? })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        let mut f = self.zip.by_name("xl/vbaProject.bin")?;
        let len = f.size() as usize;
        VbaProject::new(&mut f, len).map(|v| Cow::Owned(v))
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
                                match a? {
                                    (b"Id", v) => id.extend_from_slice(v),
                                    (b"Target", v) => target = v.as_str()?.to_string(),
                                    _ => (),
                                }
                            }
                            relationships.insert(id, target);
                        }
                        Err(e) => return Err(e.into()),
                        _ => (),
                    }
                }
            }
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(relationships)
    }

    /// MS-XLSB 2.1.7.45
    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        let mut iter = match self.iter("xl/sharedStrings.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(Vec::new()), // it is fine if path does not exists
        };
        let mut buf = vec![0; 1024];

        let _ = iter.next_skip_blocks(0x009F, &[], &mut buf)?; // BrtBeginSst
        let len = read_usize(&buf[4..8]);
        let mut strings = Vec::with_capacity(len);

        // BrtSSTItems
        for _ in 0..len {
            let _ = iter.next_skip_blocks(0x0013,
                                  &[ 
                                               (0x0023, Some(0x0024)) // future 
                                               ],
                                  &mut buf)?; // BrtSSTItem
            strings.push(wide_str(&buf[1..])?);
        }
        Ok(strings)
    }

    /// MS-XLSB 2.1.7.61
    fn read_sheets_names(&mut self,
                         relationships: &HashMap<Vec<u8>, String>)
                         -> Result<Vec<(String, String)>> {
        let mut iter = self.iter("xl/workbook.bin")?;
        let mut buf = vec![0; 1024];

        // BrtBeginBundleShs
        let _ = iter.next_skip_blocks(0x008F,
                              &[
                                          (0x0083, None),         // BrtBeginBook
                                          (0x0080, None),         // BrtFileVersion
                                          (0x0099, None),         // BrtWbProp
                                          (0x02A4, Some(0x0224)), // File Sharing
                                          (0x0025, Some(0x0026)), // AC blocks
                                          (0x02A5, Some(0x0216)), // Book protection(iso)
                                          (0x0087, Some(0x0088)), // BOOKVIEWS
                                          ],
                              &mut buf)?;
        let mut sheets = Vec::new();
        loop {
            match iter.read_type()? {
                0x0090 => return Ok(sheets), // BrtEndBundleShs
                0x009C => (), // BrtBundleSh
                typ => return Err(format!("Expecting end of sheet, got {:x}", typ).into()),
            }
            let len = iter.fill_buffer(&mut buf)?;
            let rel_len = read_u32(&buf[8..len]);
            if rel_len != 0xFFFFFFFF {
                let rel_len = rel_len as usize * 2;
                let relid = &buf[12..12 + rel_len];
                // converts utf16le to utf8 for HashMap search
                let relid = UTF_16LE.decode(relid, DecoderTrap::Ignore).map_err(|e| e.to_string())?;
                let path = format!("xl/{}", relationships[relid.as_bytes()]);
                let name = wide_str(&buf[12 + rel_len..len])?;
                sheets.push((name, path));
            }
        }
    }

    /// MS-XLSB 2.1.7.62
    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        let mut iter = self.iter(path)?;
        let mut buf = vec![0; 1024];

        // BrtWsDim
        let _ = iter.next_skip_blocks(0x0094,
                              &[
                                           (0x0081, None), // BrtBeginSheet
                                           (0x0093, None), // BrtWsProp
                                           ],
                              &mut buf)?;
        let (start, end) = parse_dimensions(&buf[..16]);
        let mut cells = Vec::with_capacity((((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as
                                            usize));

        // BrtBeginSheetData
        let _ = iter.next_skip_blocks(0x0091,
                              &[
                                          (0x0085, Some(0x0086)), // Views
                                          (0x0025, Some(0x0026)), // AC blocks
                                          (0x01E5, None),         // BrtWsFmtInfo
                                          (0x0186, Some(0x0187)), // Col Infos
                                          ],
                              &mut buf)?;

        // Initialization: first BrtRowHdr
        let mut typ: u16;
        let mut row = 0u32;

        // loop until end of sheet
        loop {
            typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;

            let value = match typ {
                0x0001 => continue, // DataType::Empty, // BrtCellBlank
                0x0002 => {
                    // BrtCellRk MS-XLSB 2.5.122
                    let d100 = (buf[8] & 1) != 0;
                    let is_int = (buf[8] & 2) != 0;
                    buf[8] &= 0xFC;
                    if is_int {
                        let v = (read_slice::<i32>(&buf[8..12]) >> 2) as i64;
                        DataType::Int(if d100 { v / 100 } else { v })
                    } else {
                        let mut v = [0u8; 8];
                        v[4..].copy_from_slice(&buf[8..12]);
                        let v = read_slice(&v);
                        DataType::Float(if d100 { v / 100.0 } else { v })
                    }
                }
                0x0003 => {
                    // BrtCellError
                    DataType::Error(match buf[8] {
                        0x00 => CellErrorType::Null,
                        0x07 => CellErrorType::Div0,
                        0x0F => CellErrorType::Value,
                        0x17 => CellErrorType::Ref,
                        0x1D => CellErrorType::Name,
                        0x24 => CellErrorType::Num,
                        0x2A => CellErrorType::NA,
                        0x2B => CellErrorType::GettingData,
                        c => return Err(format!("Unrecognised cell error code 0x{:x}", c).into()),
                    })
                }
                0x0004 => DataType::Bool(buf[8] != 0), // BrtCellBool 
                0x0005 => DataType::Float(read_slice(&buf[8..16])), // BrtCellReal 
                0x0006 => DataType::String(wide_str(&buf[8..])?), // BrtCellSt 
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&buf[8..12]);
                    DataType::String(strings[isst].clone())
                }
                0x0000 => {
                    // BrtRowHdr
                    row = read_u32(&buf);
                    if row > 0x00100000 {
                        return Ok(Range::from_sparse(cells)); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(Range::from_sparse(cells)),  // BrtEndSheetData
                _ => continue,  // anything else, ignore and try next, without changing idx
            };

            let col = read_u32(&buf);
            cells.push(Cell::new((row, col), value));
        }
    }
}

struct RecordIter<'a> {
    b: [u8; 1],
    r: BufReader<ZipFile<'a>>,
}

impl<'a> RecordIter<'a> {
    fn read_u8(&mut self) -> Result<u8> {
        self.r.read_exact(&mut self.b)?;
        Ok(self.b[0])
    }

    fn read_type(&mut self) -> Result<u16> {
        let b = self.read_u8()?;
        let typ = if (b & 0x80) == 0x80 {
            (b & 0x7F) as u16 + (((self.read_u8()? & 0x7F) as u16) << 7)
        } else {
            b as u16
        };
        Ok(typ)
    }

    fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize> {
        let mut b = self.read_u8()?;
        let mut len = (b & 0x7F) as usize;
        for i in 1..4 {
            if (b & 0x80) == 0 {
                break;
            }
            b = self.read_u8()?;
            len += ((b & 0x7F) as usize) << (7 * i);
        }
        if buf.len() < len {
            *buf = vec![0; len];
        }

        let _ = self.r.read_exact(&mut buf[..len])?;
        Ok(len)
    }

    /// Reads next type, and discard blocks between `start` and `end`
    fn next_skip_blocks(&mut self,
                        record_type: u16,
                        bounds: &[(u16, Option<u16>)],
                        buf: &mut Vec<u8>)
                        -> Result<usize> {
        let mut end = None;
        loop {
            let typ = self.read_type()?;
            match end {
                Some(e) if e == typ => end = None,
                Some(_) => (),
                None if typ == record_type => return self.fill_buffer(buf),
                None => {
                    match bounds.iter().position(|b| b.0 == typ) {
                        Some(i) => end = bounds[i].1,
                        None => {
                            return Err(format!("Unexpected record after block: expecting 0x{:x} \
                                                found 0x{:x}",
                                               record_type,
                                               typ)
                                .into())
                        }
                    }
                }
            }
            let _ = self.fill_buffer(buf)?;
        }
    }
}

fn wide_str(buf: &[u8]) -> Result<String> {
    let len = read_u32(buf) as usize;
    if buf.len() < 4 + len * 2 {
        return Err(format!("Wide string length ({}) exceeds buffer length ({})",
                           4 + len * 2,
                           buf.len())
            .into());
    }
    let s = &buf[4..4 + len * 2];
    UTF_16LE.decode(s, DecoderTrap::Ignore).map_err(|e| e.to_string().into())
}

fn parse_dimensions(buf: &[u8]) -> ((u32, u32), (u32, u32)) {
    ((read_u32(&buf[0..4]), read_u32(&buf[8..12])), (read_u32(&buf[4..8]), read_u32(&buf[12..16])))
}
