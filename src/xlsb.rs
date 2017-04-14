use std::string::String;
use std::fs::File;
use std::io::{BufReader, Read};
use std::collections::HashMap;
use std::borrow::Cow;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::reader::Reader as XmlReader;
use quick_xml::events::Event;
use quick_xml::events::attributes::Attribute;
use encoding_rs::UTF_16LE;

use {Metadata, DataType, Reader, Cell, Range, CellErrorType};
use vba::VbaProject;
use utils::{read_u16, read_u32, read_usize, read_slice};
use errors::*;

pub struct Xlsb {
    zip: ZipArchive<File>,
    sheets: Vec<(String, String)>,
    strings: Vec<String>,
}

impl Xlsb {
    /// MS-XLSB
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let mut relationships = HashMap::new();
        match self.zip.by_name("xl/_rels/workbook.bin.rels") {
            Ok(f) => {
                let mut xml = XmlReader::from_reader(BufReader::new(f));
                xml.check_end_names(false)
                    .trim_text(false)
                    .check_comments(false)
                    .expand_empty_elements(true);
                let mut buf = Vec::new();

                loop {
                    match xml.read_event(&mut buf) {
                        Ok(Event::Start(ref e)) if e.name() == b"Relationship" => {
                            let mut id = None;
                            let mut target = None;
                            for a in e.attributes() {
                                match a? {
                                    Attribute {
                                        key: b"Id",
                                        value: v,
                                    } => {
                                        id = Some(v.to_vec());
                                    }
                                    Attribute {
                                        key: b"Target",
                                        value: v,
                                    } => {
                                        target = Some(xml.decode(v).into_owned());
                                    }
                                    _ => (),
                                }
                            }
                            if let (Some(id), Some(target)) = (id, target) {
                                relationships.insert(id, target);
                            }
                        }
                        Ok(Event::Eof) => break,
                        Err(e) => return Err(e.into()),
                        _ => (),
                    }
                    buf.clear();
                }
            }
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(relationships)
    }

    /// MS-XLSB 2.1.7.45
    fn read_shared_strings(&mut self) -> Result<()> {
        let mut iter = match RecordIter::from_zip(&mut self.zip, "xl/sharedStrings.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(()), // it is fine if path does not exists
        };
        let mut buf = vec![0; 1024];

        let _ = iter.next_skip_blocks(0x009F, &[], &mut buf)?; // BrtBeginSst
        let len = read_usize(&buf[4..8]);

        // BrtSSTItems
        for _ in 0..len {
            let _ = iter.next_skip_blocks(0x0013,
                                          &[
                                           (0x0023, Some(0x0024)) // future
                                           ],
                                          &mut buf)?; // BrtSSTItem
            self.strings
                .push(wide_str(&buf[1..], &mut 0)?.into_owned());
        }
        Ok(())
    }

    /// MS-XLSB 2.1.7.61
    fn read_workbook(&mut self,
                     relationships: &HashMap<Vec<u8>, String>)
                     -> Result<Vec<(String, String)>> {
        let mut iter = RecordIter::from_zip(&mut self.zip, "xl/workbook.bin")?;
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
        loop {
            match iter.read_type()? {
                0x0090 => break, // BrtEndBundleShs
                0x009C => {
                    // BrtBundleSh
                    let len = iter.fill_buffer(&mut buf)?;
                    let rel_len = read_u32(&buf[8..len]);
                    if rel_len != 0xFFFFFFFF {
                        let rel_len = rel_len as usize * 2;
                        let relid = &buf[12..12 + rel_len];
                        // converts utf16le to utf8 for HashMap search
                        let relid = UTF_16LE.decode(relid).0;
                        let path = format!("xl/{}", relationships[relid.as_bytes()]);
                        let name = wide_str(&buf[12 + rel_len..len], &mut 0)?;
                        self.sheets.push((name.into_owned(), path));
                    }
                }
                typ => return Err(format!("Expecting end of sheet, got {:x}", typ).into()),
            }
        }
        // BrtName
        let mut defined_names = Vec::new();
        let mut extern_sheets = Vec::new();
        loop {
            match iter.read_type()? {
                0x016A => { // BrtExternSheet
                    let len = iter.fill_buffer(&mut buf)?;
                    let cxti = read_u32(&buf[..4]) as usize;
                    extern_sheets.reserve(cxti);
                    let mut start = 4;
                    for _ in 0..cxti {
                        let first = read_u32(&buf[start + 4..len]) as usize;
                        extern_sheets.push(&*self.sheets[first].0);
                        start += 12;
                    }
                }
                0x0027 => {
                    let len = iter.fill_buffer(&mut buf)?;
                    let mut str_len = 0;
                    let name = wide_str(&buf[9..len], &mut str_len)?.into_owned();
                    let rgce_len = read_u32(&buf[9 + str_len..]) as usize;
                    let rgce = &buf[13 + str_len..13 + str_len + rgce_len];
                    let formula = parse_area3d(rgce, &extern_sheets)?; // formula
                    defined_names.push((name, formula));
                }
                0x018D | // BrtUserBookView
                    0x0084 => return Ok(defined_names), // BrtEndBook
                _ => (),
            }
        }
    }
}

impl Reader for Xlsb {
    fn new(f: File) -> Result<Self> {
        Ok(Xlsb {
               zip: ZipArchive::new(f)?,
               sheets: Vec::new(),
               strings: Vec::new(),
           })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        let mut f = self.zip.by_name("xl/vbaProject.bin")?;
        let len = f.size() as usize;
        VbaProject::new(&mut f, len).map(|v| Cow::Owned(v))
    }

    fn initialize(&mut self) -> Result<Metadata> {
        self.read_shared_strings()?;
        let relationships = self.read_relationships()?;
        let defined_names = self.read_workbook(&relationships)?;
        Ok(Metadata {
               sheets: self.sheets
                   .iter()
                   .map(|&(ref s, _)| s.clone())
                   .collect(),
               defined_names: defined_names,
           })
    }

    /// MS-XLSB 2.1.7.62
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range> {

        let path = {
            let &(_, ref path) = self.sheets
                .iter()
                .find(|&&(ref n, _)| n == name)
                .ok_or_else(|| ErrorKind::WorksheetName(name.to_string()))?;
            path.clone()
        };

        let mut iter = RecordIter::from_zip(&mut self.zip, &path)?;
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
                    let error = match buf[8] {
                        0x00 => CellErrorType::Null,
                        0x07 => CellErrorType::Div0,
                        0x0F => CellErrorType::Value,
                        0x17 => CellErrorType::Ref,
                        0x1D => CellErrorType::Name,
                        0x24 => CellErrorType::Num,
                        0x2A => CellErrorType::NA,
                        0x2B => CellErrorType::GettingData,
                        c => return Err(format!("Unrecognised cell error code 0x{:x}", c).into()),
                    };
                    // BrtCellError
                    DataType::Error(error)
                }
                0x0004 => DataType::Bool(buf[8] != 0),                         // BrtCellBool
                0x0005 => DataType::Float(read_slice(&buf[8..16])),            // BrtCellReal
                0x0006 => DataType::String(wide_str(&buf[8..], &mut 0)?.into_owned()), // BrtCellSt
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&buf[8..12]);
                    DataType::String(self.strings[isst].clone())
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
    fn from_zip(zip: &'a mut ZipArchive<File>, path: &str) -> Result<RecordIter<'a>> {
        match zip.by_name(path) {
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

fn wide_str<'a, 'b>(buf: &'a [u8], str_len: &'b mut usize) -> Result<Cow<'a, str>> {
    let len = read_u32(buf) as usize;
    if buf.len() < 4 + len * 2 {
        return Err(format!("Wide string length ({}) exceeds buffer length ({})",
                           4 + len * 2,
                           buf.len())
                           .into());
    }
    *str_len = 4 + len * 2;
    let s = &buf[4..*str_len];
    Ok(UTF_16LE.decode(s).0)
}

fn parse_dimensions(buf: &[u8]) -> ((u32, u32), (u32, u32)) {
    ((read_u32(&buf[0..4]), read_u32(&buf[8..12])), (read_u32(&buf[4..8]), read_u32(&buf[12..16])))
}

fn push_column(mut col: u32, buf: &mut String) {
    if col < 26 {
        buf.push((b'A' + col as u8) as char);
    } else {
        let mut rev = String::new();
        while col >= 26 {
            let c = col % 26;
            rev.push((b'A' + c as u8) as char);
            col -= c;
            col /= 26;
        }
        buf.extend(rev.chars().rev());
    }
}

/// Formula parsing
///
/// Does not implement ALL possibilities, only Area are parsed
///
/// [MS-XLSB 2.2.2]
/// [MS-XLSB 2.5.97.88]
fn parse_area3d(rgce: &[u8], sheets: &[&str]) -> Result<String> {
    println!("parsing {}",
             rgce.iter()
                 .map(|x| format!("{:x} ", x))
                 .collect::<String>());
    let mut stack = Vec::with_capacity(rgce.len());

    let ptg = rgce[0];
    match ptg {
        0x3b | 0x5b | 0x7b => {
            // PtgArea3d
            let ixti = read_u16(&rgce[1..3]);
            let mut f = String::new();
            f.push_str(sheets[ixti as usize]);
            f.push('!');
            // TODO: check with relative columns
            f.push('$');
            push_column(read_u16(&rgce[11..13]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u32(&rgce[3..7]) + 1));
            f.push(':');
            f.push('$');
            push_column(read_u16(&rgce[13..15]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u32(&rgce[7..11]) + 1));
            stack.push(f);
        }
        0x3d | 0x5d | 0x7d => {
            // PtgArea3dErr
            let ixti = read_u16(&rgce[1..3]);
            println!("ixti err: {}", ixti);
            let mut f = String::new();
            f.push_str(sheets[ixti as usize]);
            f.push('!');
            f.push_str("#REF!");
            stack.push(f);
        }
        _ => {
            stack.push(format!("Unsupported ptg: {:x}", ptg));
        }
    }

    if stack.len() != 1 {
        bail!("Invalid formula stack");
    }

    Ok(stack.pop().unwrap())
}
