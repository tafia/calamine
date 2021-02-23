use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{BufReader, Read, Seek};
use std::string::String;

use log::debug;

use encoding_rs::UTF_16LE;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::utils::{push_column, read_f64, read_i32, read_u16, read_u32, read_usize};
use crate::vba::VbaProject;
use crate::{Cell, CellErrorType, DataType, Metadata, Range, Reader};

/// A Xlsb specific error
#[derive(Debug)]
pub enum XlsbError {
    /// Io error
    Io(std::io::Error),
    /// Zip error
    Zip(zip::result::ZipError),
    /// Xml error
    Xml(quick_xml::Error),
    /// Vba error
    Vba(crate::vba::VbaError),

    /// Mismatch value
    Mismatch {
        /// expected
        expected: &'static str,
        /// found
        found: u16,
    },
    /// File not found
    FileNotFound(String),
    /// Invalid formula, stack length too short
    StackLen,

    /// Unsupported type
    UnsupportedType(u16),
    /// Unsupported etpg
    Etpg(u8),
    /// Unsupported iftab
    IfTab(usize),
    /// Unsupported BErr
    BErr(u8),
    /// Unsupported Ptg
    Ptg(u8),
    /// Unsupported cell error code
    CellError(u8),
    /// Wide str length too long
    WideStr {
        /// wide str length
        ws_len: usize,
        /// buffer length
        buf_len: usize,
    },
}

from_err!(std::io::Error, XlsbError, Io);
from_err!(zip::result::ZipError, XlsbError, Zip);
from_err!(quick_xml::Error, XlsbError, Xml);

impl std::fmt::Display for XlsbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsbError::Io(e) => write!(f, "I/O error: {}", e),
            XlsbError::Zip(e) => write!(f, "Zip error: {}", e),
            XlsbError::Xml(e) => write!(f, "Xml error: {}", e),
            XlsbError::Vba(e) => write!(f, "Vba error: {}", e),
            XlsbError::Mismatch { expected, found } => {
                write!(f, "Expecting {}, got {:X}", expected, found)
            }
            XlsbError::FileNotFound(file) => write!(f, "File not found: '{}'", file),
            XlsbError::StackLen => write!(f, "Invalid stack length"),
            XlsbError::UnsupportedType(t) => write!(f, "Unsupported type {:X}", t),
            XlsbError::Etpg(t) => write!(f, "Unsupported etpg {:X}", t),
            XlsbError::IfTab(t) => write!(f, "Unsupported iftab {:X}", t),
            XlsbError::BErr(t) => write!(f, "Unsupported BErr {:X}", t),
            XlsbError::Ptg(t) => write!(f, "Unsupported Ptf {:X}", t),
            XlsbError::CellError(t) => write!(f, "Unsupported Cell Error code {:X}", t),
            XlsbError::WideStr { ws_len, buf_len } => write!(
                f,
                "Wide str length exceeds buffer length ({} > {})",
                ws_len, buf_len
            ),
        }
    }
}

impl std::error::Error for XlsbError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsbError::Io(e) => Some(e),
            XlsbError::Zip(e) => Some(e),
            XlsbError::Xml(e) => Some(e),
            XlsbError::Vba(e) => Some(e),
            _ => None,
        }
    }
}

/// A Xlsb reader
pub struct Xlsb<RS>
where
    RS: Read + Seek,
{
    zip: ZipArchive<RS>,
    extern_sheets: Vec<String>,
    sheets: Vec<(String, String)>,
    strings: Vec<String>,
    metadata: Metadata,
}

impl<RS: Read + Seek> Xlsb<RS> {
    /// MS-XLSB
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>, XlsbError> {
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
                                        target = Some(xml.decode(&v).into_owned());
                                    }
                                    _ => (),
                                }
                            }
                            if let (Some(id), Some(target)) = (id, target) {
                                relationships.insert(id, target);
                            }
                        }
                        Ok(Event::Eof) => break,
                        Err(e) => return Err(XlsbError::Xml(e)),
                        _ => (),
                    }
                    buf.clear();
                }
            }
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(XlsbError::Zip(e)),
        }
        Ok(relationships)
    }

    /// MS-XLSB 2.1.7.45
    fn read_shared_strings(&mut self) -> Result<(), XlsbError> {
        let mut iter = match RecordIter::from_zip(&mut self.zip, "xl/sharedStrings.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(()), // it is fine if path does not exists
        };
        let mut buf = vec![0; 1024];

        let _ = iter.next_skip_blocks(0x009F, &[], &mut buf)?; // BrtBeginSst
        let len = read_usize(&buf[4..8]);

        // BrtSSTItems
        for _ in 0..len {
            let _ = iter.next_skip_blocks(
                0x0013,
                &[
                    (0x0023, Some(0x0024)), // future
                ],
                &mut buf,
            )?; // BrtSSTItem
            self.strings.push(wide_str(&buf[1..], &mut 0)?.into_owned());
        }
        Ok(())
    }

    /// MS-XLSB 2.1.7.61
    fn read_workbook(&mut self, relationships: &HashMap<Vec<u8>, String>) -> Result<(), XlsbError> {
        let mut iter = RecordIter::from_zip(&mut self.zip, "xl/workbook.bin")?;
        let mut buf = vec![0; 1024];

        // BrtBeginBundleShs
        let _ = iter.next_skip_blocks(
            0x008F,
            &[
                (0x0083, None),         // BrtBeginBook
                (0x0080, None),         // BrtFileVersion
                (0x0099, None),         // BrtWbProp
                (0x02A4, Some(0x0224)), // File Sharing
                (0x0025, Some(0x0026)), // AC blocks
                (0x02A5, Some(0x0216)), // Book protection(iso)
                (0x0087, Some(0x0088)), // BOOKVIEWS
            ],
            &mut buf,
        )?;
        loop {
            match iter.read_type()? {
                0x0090 => break, // BrtEndBundleShs
                0x009C => {
                    // BrtBundleSh
                    let len = iter.fill_buffer(&mut buf)?;
                    let rel_len = read_u32(&buf[8..len]);
                    if rel_len != 0xFFFF_FFFF {
                        let rel_len = rel_len as usize * 2;
                        let relid = &buf[12..12 + rel_len];
                        // converts utf16le to utf8 for HashMap search
                        let relid = UTF_16LE.decode(relid).0;
                        let path = format!("xl/{}", relationships[relid.as_bytes()]);
                        let name = wide_str(&buf[12 + rel_len..len], &mut 0)?;
                        self.metadata.sheets.push(name.to_string());
                        self.sheets.push((name.into_owned(), path));
                    }
                }
                typ => {
                    return Err(XlsbError::Mismatch {
                        expected: "end of sheet",
                        found: typ,
                    });
                }
            }
        }

        // BrtName
        let mut defined_names = Vec::new();
        loop {
            let typ = iter.read_type()?;
            match typ {
                0x016A => {
                    // BrtExternSheet
                    let _len = iter.fill_buffer(&mut buf)?;
                    let cxti = read_u32(&buf[..4]) as usize;
                    if cxti < 1_000_000 {
                        self.extern_sheets.reserve(cxti);
                    }
                    let sheets = &self.sheets;
                    let extern_sheets = buf[4..]
                        .chunks(12)
                        .map(|xti| {
                            match read_i32(&xti[4..8]) {
                                -2 => "#ThisWorkbook",
                                -1 => "#InvalidWorkSheet",
                                p if p >= 0 && (p as usize) < sheets.len() => &sheets[p as usize].0,
                                _ => "#Unknown",
                            }
                            .to_string()
                        })
                        .take(cxti)
                        .collect();
                    self.extern_sheets = extern_sheets;
                }
                0x0027 => {
                    // BrtName
                    let len = iter.fill_buffer(&mut buf)?;
                    let mut str_len = 0;
                    let name = wide_str(&buf[9..len], &mut str_len)?.into_owned();
                    let rgce_len = read_u32(&buf[9 + str_len..]) as usize;
                    let rgce = &buf[13 + str_len..13 + str_len + rgce_len];
                    let formula = parse_formula(rgce, &self.extern_sheets, &defined_names)?;
                    defined_names.push((name, formula));
                }
                0x009D | 0x0225 | 0x018D | 0x0180 | 0x009A | 0x0252 | 0x0229 | 0x009B | 0x0084 => {
                    // record supposed to happen AFTER BrtNames
                    self.metadata.names = defined_names;
                    return Ok(());
                }
                _ => debug!("Unsupported type {:X}", typ),
            }
        }
    }

    fn worksheet_range_from_path(&mut self, path: String) -> Result<Range<DataType>, XlsbError> {
        let mut iter = RecordIter::from_zip(&mut self.zip, &path)?;
        let mut buf = vec![0; 1024];

        // BrtWsDim
        let _ = iter.next_skip_blocks(
            0x0094,
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
        let (start, end) = parse_dimensions(&buf[..16]);
        let len = (end.0 - start.0 + 1) * (end.1 - start.1 + 1);
        let mut cells = if len < 1_000_000 {
            Vec::with_capacity(len as usize)
        } else {
            Vec::new()
        };

        // BrtBeginSheetData
        let _ = iter.next_skip_blocks(
            0x0091,
            &[
                (0x0085, Some(0x0086)), // Views
                (0x0025, Some(0x0026)), // AC blocks
                (0x01E5, None),         // BrtWsFmtInfo
                (0x0186, Some(0x0187)), // Col Infos
            ],
            &mut buf,
        )?;

        // Initialization: first BrtRowHdr
        let mut typ: u16;
        let mut row = 0u32;

        // loop until end of sheet
        loop {
            typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;

            let value = match typ {
                // 0x0001 => continue, // DataType::Empty, // BrtCellBlank
                0x0002 => {
                    // BrtCellRk MS-XLSB 2.5.122
                    let d100 = (buf[8] & 1) != 0;
                    let is_int = (buf[8] & 2) != 0;
                    buf[8] &= 0xFC;
                    if is_int {
                        let v = (read_i32(&buf[8..12]) >> 2) as i64;
                        if d100 {
                            DataType::Float((v as f64) / 100.0)
                        } else {
                            DataType::Int(v)
                        }
                    } else {
                        let mut v = [0u8; 8];
                        v[4..].copy_from_slice(&buf[8..12]);
                        let v = read_f64(&v);
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
                        c => return Err(XlsbError::CellError(c)),
                    };
                    // BrtCellError
                    DataType::Error(error)
                }
                0x0004 | 0x000A => DataType::Bool(buf[8] != 0), // BrtCellBool or BrtFmlaBool
                0x0005 | 0x0009 => DataType::Float(read_f64(&buf[8..16])), // BrtCellReal or BrtFmlaFloat
                0x0006 | 0x0008 => DataType::String(wide_str(&buf[8..], &mut 0)?.into_owned()), // BrtCellSt or BrtFmlaString
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&buf[8..12]);
                    DataType::String(self.strings[isst].clone())
                }
                0x0000 => {
                    // BrtRowHdr
                    row = read_u32(&buf);
                    if row > 0x0010_0000 {
                        return Ok(Range::from_sparse(cells)); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(Range::from_sparse(cells)), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };

            let col = read_u32(&buf);
            cells.push(Cell::new((row, col), value));
        }
    }

    fn worksheet_formula_from_path(&mut self, path: String) -> Result<Range<String>, XlsbError> {
        let mut iter = RecordIter::from_zip(&mut self.zip, &path)?;
        let mut buf = vec![0; 1024];

        // BrtWsDim
        let _ = iter.next_skip_blocks(
            0x0094,
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
        let (start, end) = parse_dimensions(&buf[..16]);
        let len = (end.0 - start.0 + 1) * (end.1 - start.1 + 1);
        let mut cells = if len < 1_000_000 {
            Vec::with_capacity(len as usize)
        } else {
            Vec::new()
        };

        // BrtBeginSheetData
        let _ = iter.next_skip_blocks(
            0x0091,
            &[
                (0x0085, Some(0x0086)), // Views
                (0x0025, Some(0x0026)), // AC blocks
                (0x01E5, None),         // BrtWsFmtInfo
                (0x0186, Some(0x0187)), // Col Infos
            ],
            &mut buf,
        )?;

        // Initialization: first BrtRowHdr
        let mut typ: u16;
        let mut row = 0u32;

        // loop until end of sheet
        loop {
            typ = iter.read_type()?;
            let _ = iter.fill_buffer(&mut buf)?;

            let value = match typ {
                // 0x0001 => continue, // DataType::Empty, // BrtCellBlank
                0x0008 => {
                    // BrtFmlaString
                    let cch = read_u32(&buf[8..]) as usize;
                    let formula = &buf[14 + cch * 2..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.metadata.names)?
                }
                0x0009 => {
                    // BrtFmlaNum
                    let formula = &buf[18..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.metadata.names)?
                }
                0x000A | 0x000B => {
                    // BrtFmlaBool | BrtFmlaError
                    let formula = &buf[11..];
                    let cce = read_u32(formula) as usize;
                    let rgce = &formula[4..4 + cce];
                    parse_formula(rgce, &self.extern_sheets, &self.metadata.names)?
                }
                0x0000 => {
                    // BrtRowHdr
                    row = read_u32(&buf);
                    if row > 0x0010_0000 {
                        return Ok(Range::from_sparse(cells)); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(Range::from_sparse(cells)), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };

            let col = read_u32(&buf);
            cells.push(Cell::new((row, col), value));
        }
    }
}

impl<RS: Read + Seek> Reader for Xlsb<RS> {
    type RS = RS;
    type Error = XlsbError;

    fn new(reader: RS) -> Result<Self, XlsbError>
    where
        RS: Read + Seek,
    {
        let mut xlsb = Xlsb {
            zip: ZipArchive::new(reader)?,
            sheets: Vec::new(),
            strings: Vec::new(),
            extern_sheets: Vec::new(),
            metadata: Metadata::default(),
        };
        xlsb.read_shared_strings()?;
        let relationships = xlsb.read_relationships()?;
        xlsb.read_workbook(&relationships)?;

        Ok(xlsb)
    }

    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, XlsbError>> {
        self.zip.by_name("xl/vbaProject.bin").ok().map(|mut f| {
            let len = f.size() as usize;
            VbaProject::new(&mut f, len)
                .map(Cow::Owned)
                .map_err(XlsbError::Vba)
        })
    }

    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// MS-XLSB 2.1.7.62
    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, XlsbError>> {
        let path = match self.sheets.iter().find(|&&(ref n, _)| n == name) {
            Some(&(_, ref path)) => path.clone(),
            None => return None,
        };
        Some(self.worksheet_range_from_path(path))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, XlsbError>> {
        let path = match self.sheets.iter().find(|&&(ref n, _)| n == name) {
            Some(&(_, ref path)) => path.clone(),
            None => return None,
        };
        Some(self.worksheet_formula_from_path(path))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheets(&mut self) -> Vec<(String, Range<DataType>)> {
        let sheets = self.sheets.clone();
        sheets
            .into_iter()
            .filter_map(|(name, path)| {
                let ws = self.worksheet_range_from_path(path).ok()?;
                Some((name, ws))
            })
            .collect()
    }
}

struct RecordIter<'a> {
    b: [u8; 1],
    r: BufReader<ZipFile<'a>>,
}

impl<'a> RecordIter<'a> {
    fn from_zip<RS: Read + Seek>(
        zip: &'a mut ZipArchive<RS>,
        path: &str,
    ) -> Result<RecordIter<'a>, XlsbError> {
        match zip.by_name(path) {
            Ok(f) => Ok(RecordIter {
                r: BufReader::new(f),
                b: [0],
            }),
            Err(ZipError::FileNotFound) => Err(XlsbError::FileNotFound(path.into())),
            Err(e) => Err(XlsbError::Zip(e)),
        }
    }

    fn read_u8(&mut self) -> Result<u8, std::io::Error> {
        self.r.read_exact(&mut self.b)?;
        Ok(self.b[0])
    }

    /// Read next type, until we have no future record
    fn read_type(&mut self) -> Result<u16, std::io::Error> {
        let b = self.read_u8()?;
        let typ = if (b & 0x80) == 0x80 {
            (b & 0x7F) as u16 + (((self.read_u8()? & 0x7F) as u16) << 7)
        } else {
            b as u16
        };
        Ok(typ)
    }

    fn fill_buffer(&mut self, buf: &mut Vec<u8>) -> Result<usize, std::io::Error> {
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

        self.r.read_exact(&mut buf[..len])?;
        Ok(len)
    }

    /// Reads next type, and discard blocks between `start` and `end`
    fn next_skip_blocks(
        &mut self,
        record_type: u16,
        bounds: &[(u16, Option<u16>)],
        buf: &mut Vec<u8>,
    ) -> Result<usize, XlsbError> {
        loop {
            let typ = self.read_type()?;
            let len = self.fill_buffer(buf)?;
            if typ == record_type {
                return Ok(len);
            }
            if let Some(end) = bounds.iter().find(|b| b.0 == typ).and_then(|b| b.1) {
                while self.read_type()? != end {
                    let _ = self.fill_buffer(buf)?;
                }
                let _ = self.fill_buffer(buf)?;
            }
        }
    }
}

fn wide_str<'a, 'b>(buf: &'a [u8], str_len: &'b mut usize) -> Result<Cow<'a, str>, XlsbError> {
    let len = read_u32(buf) as usize;
    if buf.len() < 4 + len * 2 {
        return Err(XlsbError::WideStr {
            ws_len: 4 + len * 2,
            buf_len: buf.len(),
        });
    }
    *str_len = 4 + len * 2;
    let s = &buf[4..*str_len];
    Ok(UTF_16LE.decode(s).0)
}

fn parse_dimensions(buf: &[u8]) -> ((u32, u32), (u32, u32)) {
    (
        (read_u32(&buf[0..4]), read_u32(&buf[8..12])),
        (read_u32(&buf[4..8]), read_u32(&buf[12..16])),
    )
}

/// Formula parsing
///
/// [MS-XLSB 2.2.2]
/// [MS-XLSB 2.5.97]
///
/// See Ptg [2.5.97.16]
fn parse_formula(
    mut rgce: &[u8],
    sheets: &[String],
    names: &[(String, String)],
) -> Result<String, XlsbError> {
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
            0x01 => {
                // PtgExp: array/shared formula, ignore
                debug!("ignoring PtgExp array/shared formula");
                stack.push(formula.len());
                rgce = &rgce[4..];
            }
            0x03..=0x11 => {
                // binary operation
                let e2 = stack.pop().ok_or(XlsbError::StackLen)?;
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
                let e = stack.last().ok_or(XlsbError::StackLen)?;
                formula.insert(*e, '+');
            }
            0x13 => {
                let e = stack.last().ok_or(XlsbError::StackLen)?;
                formula.insert(*e, '-');
            }
            0x14 => {
                formula.push('%');
            }
            0x15 => {
                let e = stack.last().ok_or(XlsbError::StackLen)?;
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
            0x18 => {
                stack.push(formula.len());
                let eptg = rgce[0];
                rgce = &rgce[1..];
                match eptg {
                    0x19 => rgce = &rgce[12..],
                    0x1D => rgce = &rgce[4..],
                    e => return Err(XlsbError::Etpg(e)),
                }
            }
            0x19 => {
                let eptg = rgce[0];
                rgce = &rgce[1..];
                match eptg {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80 => rgce = &rgce[2..],
                    0x04 => rgce = &rgce[10..],
                    0x10 => {
                        rgce = &rgce[2..];
                        let e = stack.last().ok_or(XlsbError::StackLen)?;
                        let e = formula.split_off(*e);
                        formula.push_str("SUM(");
                        formula.push_str(&e);
                        formula.push(')');
                    }
                    e => return Err(XlsbError::Etpg(e)),
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
                    e => return Err(XlsbError::BErr(e)),
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
                formula.push_str(&format!("{}", read_f64(rgce)));
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
                        if iftab > crate::utils::FTAB_LEN {
                            return Err(XlsbError::IfTab(iftab));
                        }
                        rgce = &rgce[2..];
                        let argc = crate::utils::FTAB_ARGC[iftab] as usize;
                        (iftab, argc)
                    }
                };
                if stack.len() < argc {
                    return Err(XlsbError::StackLen);
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
                    formula.push_str(crate::utils::FTAB[iftab]);
                    formula.push('(');
                    for w in args.windows(2) {
                        formula.push_str(&fargs[w[0]..w[1]]);
                        formula.push(',');
                    }
                    formula.pop();
                    formula.push(')');
                } else {
                    stack.push(formula.len());
                    formula.push_str(crate::utils::FTAB[iftab]);
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
            0x29 | 0x49 | 0x69 => {
                let cce = read_u16(rgce) as usize;
                rgce = &rgce[2..];
                let f = parse_formula(&rgce[..cce], sheets, names)?;
                stack.push(formula.len());
                formula.push_str(&f);
                rgce = &rgce[cce..];
            }
            0x39 | 0x59 | 0x79 => {
                // TODO: external workbook ... ignore this formula ...
                stack.push(formula.len());
                formula.push_str("EXTERNAL_WB_NAME");
                rgce = &rgce[6..];
            }
            _ => return Err(XlsbError::Ptg(ptg)),
        }
    }

    if stack.len() != 1 {
        Err(XlsbError::StackLen)
    } else {
        Ok(formula)
    }
}
