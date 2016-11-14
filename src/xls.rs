use std::fs::File;
use std::collections::HashMap;
use std::io::BufReader;
use std::borrow::Cow;
use std::cmp::min;

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

enum XlsEncoding {
    HighByte(bool), // stores high_byte for the stream (unicode ....)
    Other, // Windows-1252 ...
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
        let wb = try!(cfb.get_stream("Workbook", &mut r)
                      .or_else(|_| cfb.get_stream("Book", &mut r)));

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
            let mut encoding = XlsEncoding::HighByte(false);
            let records = RecordIter { stream: &mut wb };
            for record in records {
                let mut r = try!(record);
                match r.typ {
                    0x0009 => if read_u16(&r.data[2..]) != 0x0005 {
                        return Err("Expecting Workbook BOF".into());
                    }, // BOF,
                    0x0042 => {
                        self.codepage = read_u16(&r.data);
                        if self.codepage != 1200 {
                            encoding = XlsEncoding::Other;
                        }
                    }, // CodePage (defines encoding)
                    0x013D => {
                        let sheet_len = r.data.len() / 2;
                        sheets.reserve(sheet_len);
                    }, // RRTabId
                    0x0085 => sheets.push(try!(parse_sheet_name(&mut r, &mut encoding))), // BoundSheet8
                    0x00FC => self.strings = try!(parse_sst(&mut r, &mut encoding)), // SST
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
                    0x0200 => {
                        range = try!(parse_dimensions(&r.data));
                        if range.get_size().0 == 0 || range.get_size().1 == 0 { continue 'sh; }
                    }, // Dimensions
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

fn parse_sheet_name(r: &mut Record, encoding: &mut XlsEncoding) -> Result<(usize, String)> {
    let pos = read_u32(r.data) as usize;
    r.data = &r.data[6..];
    let sheet = try!(parse_short_string(r, encoding));
    Ok((pos, sheet))
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

fn parse_short_string(r: &mut Record, encoding: &mut XlsEncoding) -> Result<String> {
    if r.data.len() < 2 {
        return Err("Invalid short string length".into());
    }
    let len = r.data[0] as usize;
    match encoding {
        &mut XlsEncoding::HighByte(ref mut b) => {
            *b = r.data[1] != 0;
            r.data = &r.data[2..];
        },
        &mut XlsEncoding::Other => r.data = &r.data[1..],
    }
    read_dbcs(encoding, len, r)
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
    let (rf, rl, cf, cl) = match r.len() {
        10 => (read_u16(&r[0..2]) as u32,
               read_u16(&r[2..4]) as u32,
               read_u16(&r[4..6]) as u32,
               read_u16(&r[6..8]) as u32),
        14 => (read_u32(&r[0..4]),
               read_u32(&r[4..8]),
               read_u16(&r[8..10]) as u32,
               read_u16(&r[10..12]) as u32),
        _ => return Err("Invalid dimensions lengths".into()),
    };

    if rl == 0 || cl == 0 {
        Ok(Range::new((0, 0), (0, 0)))
    } else {
        Ok(Range::new((rf, cf), (rl - 1, cl - 1)))
    }
}

fn parse_sst(r: &mut Record, encoding: &mut XlsEncoding) -> Result<Vec<String>> {
    if r.data.len() < 8 {
        return Err("Invalid sst length".into());
    }
    let len = read_slice::<i32>(&r.data[4..]) as usize;
    let mut sst = Vec::with_capacity(len);
    r.data = &r.data[8..];
    for _ in 0..len {
        sst.push(try!(read_rich_extended_string(r, encoding)));
    }
    Ok(sst)
}

fn read_rich_extended_string(r: &mut Record, encoding: &mut XlsEncoding) -> Result<String> {
    if r.data.is_empty() && !r.continue_record() || r.data.len() < 3 {
        return Err("Invalid rich extended string length".into());
    }

    let str_len = read_u16(r.data) as usize;
    let flags = r.data[2];
    r.data = &r.data[3..];
    let ext_st = flags & 0x4;
    let rich_st = flags & 0x8;

    if let &mut XlsEncoding::HighByte(ref mut b) = encoding {
        *b = flags & 0x1 != 0;
    }

    let mut unused_len = if rich_st != 0 { 
        let l = 4 * read_u16(&r.data) as usize;
        r.data = &r.data[2..];
        l
    } else {
        0
    };
    if ext_st != 0 { 
        unused_len += read_slice::<i32>(&r.data) as usize;
        r.data = &r.data[4..];
    };

    let s = try!(read_dbcs(encoding, str_len, r));

    while unused_len > 0 {
        if r.data.is_empty() && !r.continue_record() {
            return Err("continued record too short while reading extended string".into());
        }
        let l = min(unused_len, r.data.len());
        let (_, next) = r.data.split_at(l);
        r.data = next;
        unused_len -= l;
    }

    Ok(s)
}

fn read_dbcs<'a>(encoding: &mut XlsEncoding, mut len: usize, r: &mut Record) -> Result<String>
{
    let mut s = String::with_capacity(len);
    while len > 0 {
        let (l, bytes) = match encoding {
            &mut XlsEncoding::Other | &mut XlsEncoding::HighByte(false) => {
                let l = min(r.data.len(), len);
                let (data, next) = r.data.split_at(l);
                r.data = next;

                // add 0x00 high bytes to unicodes
                let mut bytes = vec![0; l * 2];
                for (i, sce) in data.iter().enumerate() {
                    bytes[2 * i] = *sce;
                }
                (l, Cow::Owned(bytes))
            },
            &mut XlsEncoding::HighByte(true) => {
                let l = min(r.data.len() / 2, len);
                let (data, next) = r.data.split_at(2 * l);
                r.data = next;
                (l, Cow::Borrowed(data))
            }
        };

        let s_rec: Result<String> = UTF_16LE.decode(&bytes, DecoderTrap::Ignore)
            .map_err(|e| e.to_string().into());
        s.push_str(&try!(s_rec));

        len -= l;
        if len > 0 {
           if r.continue_record() {
               match encoding {
                   &mut XlsEncoding::HighByte(ref mut b) => {
                       *b = r.data[0] & 0x1 != 0;
                       r.data = &r.data[1..];
                   },
                   &mut XlsEncoding::Other => {},
               }
           } else {
               return Err("Cannot decode entire dbcs stream".into());
           }
        }
    }
    Ok(s)
}

struct Record<'a> {
    typ: u16,
    data: &'a[u8],
    cont: Option<Vec<&'a[u8]>>,
}

impl<'a> Record<'a> {
    fn continue_record(&mut self) -> bool {
        match self.cont {
            None => false,
            Some(ref mut v) => {
                if v.is_empty() {
                    false
                } else {
                    self.data = v.remove(0);
                    true
                }
            }
        }
    }
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
        let (data, next) = self.stream.split_at(len + 4);
        self.stream = next;
        let d = &data[4..];

        // Append next record data if it is a Continue record
        let cont = if next.len() > 4 && read_u16(next) == 0x003C {
            let mut cont = Vec::new();
            while self.stream.len() > 4 && read_u16(self.stream) == 0x003C {
                len = read_u16(&self.stream[2..]) as usize;
                if self.stream.len() < len + 4 {
                    return Some(Err("Expecting continue record length, found end of stream".into()));
                }
                let sp = self.stream.split_at(len + 4);
                cont.push(&sp.0[4..]);
                self.stream = sp.1;
            }
            Some(cont)
        } else {
            None
        };

        Some(Ok(Record { typ: t, data: d, cont: cont }))
    }
}
