use std::fs::File;
use std::collections::HashMap;
use std::io::BufReader;
use std::borrow::Cow;

use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;

use errors::*;
use {ExcelReader, Range, DataType, CellErrorType};
use vba::VbaProject;
use cfb::Cfb;
use utils::{read_u16, read_u32, read_slice};

enum CfbWrap {
    Wb(Cfb, BufReader<File>),
    Vba(VbaProject),
}

/// A struct representing an old xls format file (CFB)
pub struct Xls {
    codepage: u16,
    sheets: HashMap<String, Range>,
    strings: Vec<String>,
    cfb: Option<CfbWrap>,
}

impl ExcelReader for Xls {

    fn new(r: File) -> Result<Self> {
        let len = try!(r.metadata()).len() as usize;
        let mut r = BufReader::new(r);
        let mut cfb = try!(Cfb::new(&mut r, len ));
        let wb = try!(cfb.get_stream("Workbook", &mut r));

        let mut xls = Xls { 
            codepage: 1200,
            sheets: HashMap::new(), 
            strings: Vec::new(), 
            cfb: Some(CfbWrap::Wb(cfb, r)),
        };
        try!(xls.parse_workbook(&wb));
        Ok(xls)
    }

    fn has_vba(&mut self) -> bool {
        match self.cfb {
            Some(CfbWrap::Wb(ref cfb, _)) => cfb.has_directory("_VBA_PROJECT_CUR"),
            Some(CfbWrap::Vba(_)) => true,
            None => unreachable!(), // option is used to transfer ownership only
        }
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        if let Some(CfbWrap::Wb(..)) = self.cfb {
            match self.cfb.take() {
                Some(CfbWrap::Wb(cfb, mut r)) => {
                    let vba = try!(VbaProject::from_cfb(&mut r, cfb));
                    self.cfb = Some(CfbWrap::Vba(vba));
                },
                _ => unreachable!(),
            }
        }

        match self.cfb {
            Some(CfbWrap::Vba(ref v)) => Ok(Cow::Borrowed(v)),
            _ => unreachable!(),
        }
    }

    /// Parses Workbook stream, no need for the relationships variable
    fn read_sheets_names(&mut self, _: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>> {
        Ok(self.sheets.keys().map(|k| (k.to_string(), k.to_string())).collect())
    }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        Ok(self.strings.clone())
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        Ok(HashMap::new())
    }

    fn read_worksheet_range(&mut self, name: &str, _: &[String]) -> Result<Range> {
        match self.sheets.get(name) {
            None => Err(format!("Sheet '{}' does not exist", name).into()),
            Some(r) => Ok(r.clone()),
        }
    }
}

impl Xls {
    fn parse_workbook(&mut self, stream: &[u8]) -> Result<()> {

        let mut sheets = Vec::new(); 
        {
            let mut wb = stream;
            let records = RecordIter { stream: &mut wb };
            for record in records {
                let r = try!(record);
                match r.typ {
                    0x0009 => if read_u16(&r.data[2..]) != 0x0005 {
                        return Err("Expecting Workbook BOF".into());
                    }, // BOF,
                    0x0042 => self.codepage = read_u16(&r.data), // CodePage (defines encoding)
                    0x013D => {
                        let sheet_len = r.data.len() / 2;
                        sheets.reserve(sheet_len);
                    }, // RRTabId
                    0x0085 => sheets.push(try!(parse_sheet_name(&r.data))), // BoundSheet8
                    0x00FC => self.strings = try!(parse_sst(&r.data)), // SST
                    0x000A => break, // EOF,
                    _ => (),
                }
            }
        }

        'sh: for (pos, name) in sheets.into_iter() {
            let mut sh = &stream[pos..];
            let records = RecordIter { stream: &mut sh };
            let mut range = Range::default();
            for record in records {
                let r = try!(record);
                match r.typ {
                    0x0009 => if read_u16(&r.data[2..]) != 0x0010 { continue 'sh; }, // BOF, worksheet
                    0x0200 => range = try!(parse_dimensions(&r.data)), // Dimensions
                    0x0203 => try!(parse_number(&r.data, &mut range)), // Number 
                    0x0205 => try!(parse_bool_err(&r.data, &mut range)), // BoolErr
                    0x027E => try!(parse_rk(&r.data, &mut range)), // RK
                    0x00FD => try!(parse_label_sst(&r.data, &self.strings, &mut range)), // LabelSst
                    0x000A => break, // EOF,
                    _ => (),
                }
            }
            self.sheets.insert(name, range);
        }
        Ok(())
    }
}

fn parse_sheet_name(r: &[u8]) -> Result<(usize, String)> {
    let pos = read_u32(r) as usize;
    let sheet = try!(parse_short_string(&r[6..]));
    Ok((pos, sheet))
}

fn parse_sst(r: &[u8]) -> Result<Vec<String>> {
    if r.len() < 8 {
        return Err("Invalid sst length".into());
    }
    let len = read_slice::<i32>(&r[4..]) as usize;
    let mut sst = Vec::with_capacity(len);
    let mut read = &mut &r[8..];
    for _ in 0..len {
        sst.push(try!(read_rich_extended_string(read)));
    }
    Ok(sst)
}

fn parse_number(r: &[u8], range: &mut Range) -> Result<()> {
    if r.len() < 14 {
        return Err("Invalid number length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let v = read_slice::<f64>(&r[6..]);
    range.set_value((row as u32, col as u32), DataType::Float(v))
}

fn parse_bool_err(r: &[u8], range: &mut Range) -> Result<()> {
    if r.len() < 8 {
        return Err("Invalid BoolErr length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let v = match r[7] {
        0x00 => DataType::Bool(r[6] != 0),
        0x01 => match r[6] {
            0x00 => DataType::Error(CellErrorType::Null),
            0x07 => DataType::Error(CellErrorType::Div0),
            0x0F => DataType::Error(CellErrorType::Value),
            0x17 => DataType::Error(CellErrorType::Ref),
            0x1D => DataType::Error(CellErrorType::Name),
            0x24 => DataType::Error(CellErrorType::Num),
            0x2A => DataType::Error(CellErrorType::NA),
            0x2B => DataType::Error(CellErrorType::GettingData),
            e => return Err(format!("Unrecognized error {:x}", e).into()),
        },
        e => return Err(format!("Unrecognized fError {:x}", e).into()),
    };
    range.set_value((row as u32, col as u32), v)
}
 
fn parse_rk(r: &[u8], range: &mut Range) -> Result<()> {
    if r.len() < 10 {
        return Err("Invalid rk length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);

    let d100 = (r[6] & 1) != 0;
    let is_int = (r[6] & 2) != 0;

    let mut v = [0u8; 8];
    v[4..].copy_from_slice(&r[6..10]);
    v[0] &= 0xFC;
    let v = if is_int { 
        let v = (read_slice::<i32>(&v[4..]) >> 2) as i64;
        DataType::Int( if d100 { v/100 } else { v })
    } else {
        let v = read_slice(&v);
        DataType::Float( if d100 { v/100.0 } else { v })
    };

    range.set_value((row as u32, col as u32), v)
}

fn parse_short_string(r: &[u8]) -> Result<String> {
    if r.len() < 2 {
        return Err("Invalid short string length".into());
    }
    let mut len = r[0] as usize;
    let zero_high_byte = r[1] == 0;
    let bytes = if zero_high_byte {
        // add 0x00 high bytes to unicodes
        let mut bytes = vec![0; len * 2];
        for (i, sce) in r[2..2 + len].iter().enumerate() {
            bytes[2 * i] = *sce;
        }
        Cow::Owned(bytes)
    } else {
        len *= 2;
        Cow::Borrowed(&r[2..2 + len])
    };
    UTF_16LE.decode(&bytes, DecoderTrap::Ignore).map_err(|e| e.to_string().into())
}

fn parse_label_sst(r: &[u8], strings: &[String], range: &mut Range) -> Result<()> {
    if r.len() < 10 {
        return Err("Invalid short string length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let i = read_u32(&r[6..]) as usize;
    range.set_value((row as u32, col as u32), DataType::String(strings[i].clone()))
}

fn parse_dimensions(r: &[u8]) -> Result<Range> {
    if r.len() != 14 {
        return Err("Invalid dimensions lengths".into());
    }
    let rw_first = read_u32(&r[0..4]);
    let rw_last = read_u32(&r[4..8]);
    let col_first = read_u16(&r[8..10]);
    let col_last = read_u16(&r[10..12]);

    let start = (rw_first, col_first as u32);
    let size = ((rw_last - rw_first) as usize, (col_last - col_first) as usize);

    Ok(Range::new(start, size))
}

fn read_rich_extended_string(r: &mut &[u8]) -> Result<String> {
    if r.len() < 3 {
        return Err("Invalid rich extended string length".into());
    }

    let mut len = read_u16(r) as usize;
    let mut start = 3;
    let flags = r[2];
    let high_byte = flags & 0x1;
    let ext_st = flags & 0x4;
    let rich_st = flags & 0x8;
    let str_len = len;

    if rich_st != 0 { 
        len += 4 * read_u16(&r[start..]) as usize;
        start += 2;
    }
    if ext_st != 0 { 
        len += read_slice::<i32>(&r[start..]) as usize;
        start += 4;
    }
    
    let bytes = if high_byte == 0 {
        // add 0x00 high bytes to unicodes
        let mut bytes = vec![0; str_len * 2];
        for (i, sce) in r[start..start + str_len].iter().enumerate() {
            bytes[2 * i] = *sce;
        }
        Cow::Owned(bytes)
    } else {
        len += str_len;
        Cow::Borrowed(&r[start..start + 2 * str_len])
    };

    if r.len() < start + len {
        return Err(format!("Invalid rich extended string \
                            length: {} < {}", r.len(), start + len).into());
    }

    *r = &r[start + len..];
    UTF_16LE.decode(&bytes, DecoderTrap::Ignore).map_err(|e| e.to_string().into())
}

struct Record<'a> {
    typ: u16,
    data: Cow<'a, [u8]>,
}

struct RecordIter<'a> {
    stream: &'a[u8],
}

impl<'a> Iterator for RecordIter<'a> {
    type Item=Result<Record<'a>>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.stream.len() < 4 {
            return if self.stream.is_empty() {
                None
            } else {
                Some(Err("Expecting record type and length, found end of stream".into()))
            }
        }
        let t = read_u16(self.stream);
        let mut len = read_u16(&self.stream[2..]) as usize;
        if self.stream.len() < len + 4 {
            return Some(Err("Expecting record length, found end of stream".into()));
        }
        let (mut data, mut next) = self.stream.split_at(len + 4);
        self.stream = next;

        // Append next record data if it is a Continue record
        let cow = if next.len() > 4 && read_u16(next) == 0x003C {
            let mut c = data[4..].to_vec();
            while next.len() > 4 && read_u16(next) == 0x003C {
                len = read_u16(&self.stream[2..]) as usize;
                if self.stream.len() < len + 4 {
                    return Some(Err("Expecting continue record length, found end of stream".into()));
                }
                let sp = self.stream.split_at(len + 4);
                data = sp.0;
                next = sp.1;
                c.extend_from_slice(&data[4..]);
                self.stream = next;
            }
            Cow::Owned(c)
        } else {
            Cow::Borrowed(&data[4..])
        };

        Some(Ok(Record { typ: t, data: cow }))
    }
}
