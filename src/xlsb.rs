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
use utils::{read_u16, read_u32, read_usize, read_slice, push_column};
use errors::*;

pub struct Xlsb {
    zip: ZipArchive<File>,
    extern_sheets: Vec<String>,
    sheets: Vec<(String, String)>,
    strings: Vec<String>,
    defined_names: Vec<(String, String)>,
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
    fn read_workbook(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<()> {
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
        loop {
            match iter.read_type()? {
                0x016A => {
                    // BrtExternSheet
                    let len = iter.fill_buffer(&mut buf)?;
                    let cxti = read_u32(&buf[..4]) as usize;
                    if len < 4 + cxti as usize * 12 {
                        bail!("BrtExternSheet buffer too small");
                    }
                    self.extern_sheets.reserve(cxti);
                    let mut start = 4;
                    for _ in 0..cxti {
                        let sh = match read_slice::<i32>(&buf[start + 4..len]) {
                            -2 => "#ThisWorkbook",
                            -1 => "#InvalidWorkSheet",
                            p if p >= 0 && (p as usize) < self.sheets.len() => {
                                &self.sheets[p as usize].0
                            }
                            _ => "#Unknown",
                        };
                        self.extern_sheets.push(sh.to_string());
                        start += 12;
                    }
                }
                0x0027 => {
                    let len = iter.fill_buffer(&mut buf)?;
                    let mut str_len = 0;
                    let name = wide_str(&buf[9..len], &mut str_len)?.into_owned();
                    let rgce_len = read_u32(&buf[9 + str_len..]) as usize;
                    let rgce = &buf[13 + str_len..13 + str_len + rgce_len];
                    let formula = parse_formula(rgce, &self.extern_sheets, &[])?; // formula
                    defined_names.push((name, formula));
                }
                0x018D | 0x0084 => {
                    // BrtUserBookView
                    self.defined_names = defined_names;
                    return Ok(()); // BrtEndBook
                }
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
               extern_sheets: Vec::new(),
               defined_names: Vec::new(),
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
        self.read_workbook(&relationships)?;
        Ok(Metadata {
               sheets: self.sheets.iter().map(|s| s.0.clone()).collect(),
               defined_names: self.defined_names.clone(),
           })
    }

    /// MS-XLSB 2.1.7.62
    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>> {

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

    /// MS-XLSB 2.1.7.62
    fn read_worksheet_formula(&mut self, name: &str) -> Result<Range<String>> {

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
                0x0008 => {
                    // BrtFmlaString
                    let cch = read_u32(&buf[8..]) as usize;
                    let formula = &buf[14 + cch * 2..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.defined_names)?
                }
                0x0009 => {
                    // BrtFmlaNum
                    let formula = &buf[18..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.defined_names)?
                }
                0x000A | 0x000B => {
                    // BrtFmlaBool | BrtFmlaError
                    let formula = &buf[11..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.defined_names)?
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

/// Formula parsing
///
/// [MS-XLSB 2.2.2]
/// [MS-XLSB 2.5.97]
fn parse_formula(mut rgce: &[u8], sheets: &[String], names: &[(String, String)]) -> Result<String> {
    if rgce.is_empty() {
        return Ok(String::new());
    }

    let mut stack = Vec::new();
    let mut formula = String::with_capacity(rgce.len());
    while !rgce.is_empty() {
        let ptg = rgce[0];
        rgce = &rgce[1..];
        match ptg {
            0x3a | 0x5a | 0x7a => {
                // PtgRef3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                // TODO: check with relative columns
                formula.push('$');
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u32(&rgce[2..6]) + 1));
                rgce = &rgce[8..];
            }
            0x3b | 0x5b | 0x7b => {
                // PtgArea3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                // TODO: check with relative columns
                formula.push('$');
                push_column(read_u16(&rgce[10..12]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u32(&rgce[2..6]) + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[12..14]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u32(&rgce[6..10]) + 1));
                rgce = &rgce[14..];
            }
            0x3c | 0x5c | 0x7c => {
                // PtfRefErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[8..];
            }
            0x3d | 0x5d | 0x7d => {
                // PtgAreaErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[14..];
            }
            0x03...0x11 => {
                // binary operation
                let e2 = stack
                    .pop()
                    .ok_or::<Error>("Invalid stack length".into())?;
                let e2 = formula.split_off(e2);
                // imaginary 'e1' will actually already be the start of the binary op
                let op = match ptg {
                    0x03 => "+",
                    0x04 => "-",
                    0x05 => "*",
                    0x06 => "/",
                    0x07 => "^",
                    0x08 => "&",
                    0x09 => "<",
                    0x0A => "<=",
                    0x0B => "=",
                    0x0C => ">",
                    0x0D => ">=",
                    0x0E => "<>",
                    0x0F => " ",
                    0x10 => ",",
                    0x11 => ":",
                    _ => unreachable!(),
                };
                formula.push_str(op);
                formula.push_str(&e2);
            }
            0x12 => {
                let e = stack
                    .last()
                    .ok_or::<Error>("Invalid stack length".into())?;
                formula.insert(*e, '+');
            }
            0x13 => {
                let e = stack
                    .last()
                    .ok_or::<Error>("Invalid stack length".into())?;
                formula.insert(*e, '-');
            }
            0x14 => {
                formula.push('%');
            }
            0x15 => {
                let e = stack
                    .last()
                    .ok_or::<Error>("Invalid stack length".into())?;
                formula.insert(*e, '(');
                formula.push(')');
            }
            0x16 => {
                stack.push(formula.len());
            }
            0x17 => {
                stack.push(formula.len());
                formula.push('\"');
                let cch = read_u16(&rgce[0..2]) as usize;
                formula.push_str(&*UTF_16LE.decode(&rgce[2..2 + 2 * cch]).0);
                formula.push('\"');
                rgce = &rgce[2 + 2 * cch..];
            }
            0x18 | 0x19 => {
                // ignore most of these ptgs ...
                let etpg = rgce[0];
                rgce = &rgce[1..];
                match etpg {
                    0x19 => rgce = &rgce[12..],
                    0x1D => rgce = &rgce[4..],
                    0x01 => rgce = &rgce[2..],
                    0x02 => rgce = &rgce[2..],
                    0x04 => rgce = &rgce[10..],
                    0x08 => rgce = &rgce[2..],
                    0x10 => {
                        rgce = &rgce[2..];
                        let e = *stack
                                     .last()
                                     .ok_or::<Error>("Invalid stack length".into())?;
                        let e = formula.split_off(e);
                        formula.push_str("SUM(");
                        formula.push_str(&e);
                        formula.push(')');
                    }
                    0x20 => rgce = &rgce[2..],
                    0x21 => rgce = &rgce[2..],
                    0x40 => rgce = &rgce[2..],
                    0x41 => rgce = &rgce[2..],
                    0x80 => rgce = &rgce[2..],
                    e => bail!("Unsupported etpg: 0x{:x}", e),
                }
            }
            0x1C => {
                stack.push(formula.len());
                let err = rgce[0];
                rgce = &rgce[1..];
                match err {
                    0x00 => formula.push_str("#NULL!"),
                    0x07 => formula.push_str("#DIV/0!"),
                    0x0F => formula.push_str("#VALUE!"),
                    0x17 => formula.push_str("#REF!"),
                    0x1D => formula.push_str("#NAME?"),
                    0x24 => formula.push_str("#NUM!"),
                    0x2A => formula.push_str("#N/A"),
                    0x2B => formula.push_str("#GETTING_DATA"),
                    e => bail!("Unrecognosed BErr 0x{:x}", e),
                }
            }
            0x1D => {
                stack.push(formula.len());
                formula.push_str(if rgce[0] == 0 { "FALSE" } else { "TRUE" });
                rgce = &rgce[1..];
            }
            0x1E => {
                stack.push(formula.len());
                formula.push_str(&format!("{}", read_u16(rgce)));
                rgce = &rgce[2..];
            }
            0x1F => {
                stack.push(formula.len());
                formula.push_str(&format!("{}", read_slice::<f64>(rgce)));
                rgce = &rgce[8..];
            }
            0x20 | 0x40 | 0x60 => {
                // PtgArray: ignore
                stack.push(formula.len());
                rgce = &rgce[14..];
            }
            0x21 | 0x22 | 0x41 | 0x42 | 0x61 | 0x62 => {
                let (iftab, argc) = match ptg {
                    0x22 | 0x42 | 0x62 => {
                        let iftab = read_u16(&rgce[1..]) as usize;
                        let argc = rgce[0] as usize;
                        rgce = &rgce[3..];
                        (iftab, argc)
                    }
                    _ => {
                        let iftab = read_u16(rgce) as usize;
                        if iftab > ::utils::FTAB_LEN {
                            bail!("Invalid iftab");
                        }
                        rgce = &rgce[2..];
                        let argc = ::utils::FTAB_ARGC[iftab] as usize;
                        (iftab, argc)
                    }
                };
                if stack.len() < argc {
                    bail!("Invalid formula, stack is too small");
                }
                if argc > 0 {
                    let args_start = stack.len() - argc;
                    let mut args = stack.split_off(args_start);
                    let start = args[0];
                    for s in &mut args {
                        *s -= start;
                    }
                    let fargs = formula.split_off(start);
                    stack.push(formula.len());
                    args.push(fargs.len());
                    formula.push_str(::utils::FTAB[iftab]);
                    formula.push('(');
                    for w in args.windows(2) {
                        formula.push_str(&fargs[w[0]..w[1]]);
                        formula.push(',');
                    }
                    formula.pop();
                    formula.push(')');
                } else {
                    stack.push(formula.len());
                    formula.push_str(::utils::FTAB[iftab]);
                    formula.push_str("()");
                }
            }
            0x23 | 0x43 | 0x63 => {
                let iname = read_u32(rgce) as usize - 1; // one-based
                stack.push(formula.len());
                formula.push_str(&names[iname].0);
                rgce = &rgce[4..];
            }
            0x24 | 0x44 | 0x64 => {
                let row = read_u32(rgce) + 1;
                let col = [rgce[4], rgce[5] & 0x3F];
                let col = read_u16(&col);
                stack.push(formula.len());
                if rgce[5] & 0x80 != 0x80 {
                    formula.push('$');
                }
                push_column(col as u32, &mut formula);
                if rgce[5] & 0x40 != 0x40 {
                    formula.push('$');
                }
                formula.push_str(&format!("{}", row));
                rgce = &rgce[6..];
            }
            0x25 | 0x45 | 0x65 => {
                stack.push(formula.len());
                formula.push('$');
                push_column(read_u16(&rgce[8..10]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u32(&rgce[0..4]) + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[10..12]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u32(&rgce[4..8]) + 1));
                rgce = &rgce[12..];
            }
            0x2A | 0x4A | 0x6A => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[6..];
            }
            0x2B | 0x4B | 0x6B => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[12..];
            }
            _ => bail!("Unsupported ptg: 0x{:x}", ptg),
        }
    }

    if stack.len() != 1 {
        Err(format!("Invalid formula, final stack size: {}", stack.len()).into())
    } else {
        Ok(formula)
    }
}
