use std::fs::File;
use std::collections::HashMap;
use std::io::BufReader;

use encoding::{Encoding, DecoderTrap};
use encoding::all::UTF_16LE;

use errors::*;
use {ExcelReader, Range, DataType, CellErrorType};
use vba::VbaProject;
use cfb::Cfb;
use utils::{read_u16, read_slice};

/// A struct representing an old xls format file (CFB)
pub struct Xls {
    r: BufReader<File>,
    cfb: Cfb,
}

impl ExcelReader for Xls {

    fn new(r: File) -> Result<Self> {
        let len = try!(r.metadata()).len() as usize;
        let mut r = BufReader::new(r);
        let cfb = try!(Cfb::new(&mut r, len));
        Ok(Xls { r: r, cfb: cfb })
    }

    fn has_vba(&mut self) -> bool {
        self.cfb.has_directory("_VBA_PROJECT_CUR")
    }

    fn vba_project(&mut self) -> Result<VbaProject> {
//         let len = try!(self.file.get_ref().metadata()).len() as usize;
//         VbaProject::new(&mut self.file, len)
        unimplemented!()        
    }

    /// Parses Workbook stream, no need for the relationships variable
    fn read_sheets_names(&mut self, _: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>> {
        let sheets = HashMap::new();
        let wb = try!(self.cfb.get_stream("Workbook", &mut self.r));
        try!(parse_workbook(&wb));
        Ok(sheets)
    }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        Ok(Vec::new())
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        unimplemented!()
    }

    fn read_worksheet_range(&mut self, _: &str, _: &[String]) -> Result<Range> {
        unimplemented!()
    }
}

struct Cell {
    row: u16,
    col: u16,
    val: DataType,
}

struct Record<'a> {
    typ: u16,
    data: &'a [u8],
}

struct RecordIter<'a> {
    stream: &'a[u8],
}

fn parse_workbook(mut wb: &[u8]) -> Result<()> {
    let records = RecordIter { stream: &mut wb };

    let mut sheet_names = Vec::new();
    let mut biff = 0;
    let mut depth = 0;
    let mut ignore = false;
    let mut cells = Vec::new();
    let mut strings = Vec::new();

    for record in records {
        let r = try!(record);
        if ignore && r.typ != 0x0009 { // within an unsupported substream
            continue;
        }
        match r.typ {
            0x0009 => { // BOF
                depth += 1;
                biff = read_u16(r.data);
                let dt = read_u16(&r.data[2..]);
                ignore = !(dt == 0x0005 || dt == 0x0010);
            },
            0x000A => { // EOF
                if depth == 1 { break; }
                depth -= 1;
            },
            0x00FC => strings = try!(parse_sst(r.data)), // SST
            0x0085 => sheet_names.push(r.data.to_vec()), // BoundSheet8
            0x0203 => cells.push(try!(parse_number(r.data))), // Number 
            0x0205 => cells.push(try!(parse_bool_err(r.data))),// BoolErr
            0x027E => cells.push(try!(parse_rk(r.data))),// RK
            _ => (),
        }
    }
    Ok(())
}

fn parse_sst(r: &[u8]) -> Result<Vec<String>> {
    if r.len() < 8 {
        return Err("Invalid sst length".into());
    }
    let len = read_slice::<i32>(&r[4..]) as usize;
    let mut sst = Vec::with_capacity(len);
    let mut read = &mut &*r;
    for _ in 0..len {
        sst.push(try!(read_rich_extended_string(read)));
    }
    Ok(sst)
}

fn parse_number(r: &[u8]) -> Result<Cell> {
    if r.len() < 14 {
        return Err("Invalid number length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let v = read_slice::<f64>(&r[6..]);
    Ok(Cell { row: row, col: col, val: DataType::Float(v), })
}

fn parse_bool_err(r: &[u8]) -> Result<Cell> {
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
    Ok(Cell { row: row, col: col, val: v, })
}
 
fn parse_rk(r: &[u8]) -> Result<Cell> {
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

    Ok(Cell { row: row, col: col, val: v })
}

fn parse_string(r: &[u8]) -> Result<String> {
    if r.len() < 2 {
        return Err("Invalid string length".into());
    }
    let mut len = read_u16(r) as usize;
    let zero_high_byte = r[3] == 0;
    if zero_high_byte {
        if r.len() < 3 + len {
            return Err("Invalid string length".into());
        }
        ::std::str::from_utf8(&r[3..3 + len]).map(|s| s.to_string()).map_err(|e| e.into())
    } else {
        len *= 2;
        if r.len() < 3 + len {
            return Err("Invalid string length".into());
        }
        UTF_16LE.decode(&r[3..3 + len], DecoderTrap::Ignore).map_err(|e| e.to_string().into())
    }
}

fn read_rich_extended_string(r: &mut &[u8]) -> Result<String> {
    if r.len() < 2 {
        return Err("Invalid string length".into());
    }
    let mut len = read_u16(r) as usize;
    let zero_high_byte = (r[3] & 0x1) == 0;
    let s = if zero_high_byte {
        if r.len() < 9 + len {
            return Err("Invalid string length".into());
        }
        ::std::str::from_utf8(&r[9..9 + len]).map(|s| s.to_string()).map_err(|e| e.into())
    } else {
        len *= 2;
        if r.len() < 9 + len {
            return Err("Invalid string length".into());
        }
        UTF_16LE.decode(&r[9..9 + len], DecoderTrap::Ignore).map_err(|e| e.to_string().into())
    };
    *r = &r[9 + len..];
    s
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
        let len = read_u16(&self.stream[2..]) as usize;
        if self.stream.len() < len + 4 {
            return Some(Err("Expecting record length, found end of stream".into()));
        }
        let (data, next) = self.stream.split_at(len + 4);
        self.stream = next;
        Some(Ok(Record { typ: t, data: &data[4..] }))
    }
}
