use std::fs::File;
use std::collections::HashMap;
use std::io::BufReader;
use std::borrow::Cow;
use std::cmp::min;

use errors::*;
use {Metadata, Reader, Range, Cell, DataType, CellErrorType};
use vba::VbaProject;
use cfb::{Cfb, XlsEncoding};
use utils::{read_u16, read_u32, read_slice, push_column};

enum SheetsState {
    NotParsed(BufReader<File>, Cfb),
    Parsed(HashMap<String, (Range<DataType>, Range<String>)>),
}

/// A struct representing an old xls format file (CFB)
pub struct Xls {
    sheets: SheetsState,
    vba: Option<VbaProject>,
}

impl Reader for Xls {
    fn new(r: File) -> Result<Self> {

        let len = r.metadata()?.len() as usize;
        let mut r = BufReader::new(r);
        let mut cfb = Cfb::new(&mut r, len)?;

        // Reads vba once for all (better than reading all worksheets once for all)
        let vba = if cfb.has_directory("_VBA_PROJECT_CUR") {
            Some(VbaProject::from_cfb(&mut r, &mut cfb)?)
        } else {
            None
        };

        Ok(Xls {
               sheets: SheetsState::NotParsed(r, cfb),
               vba: vba,
           })
    }

    fn has_vba(&mut self) -> bool {
        self.vba.is_some()
    }

    fn vba_project(&mut self) -> Result<Cow<VbaProject>> {
        self.vba
            .as_ref()
            .map(|vba| Cow::Borrowed(vba))
            .ok_or("No vba project".into())
    }

    /// Parses Workbook stream, no need for the relationships variable
    fn initialize(&mut self) -> Result<Metadata> {
        let defined_names = self.parse_workbook()?;
        let sheets = match self.sheets {
            SheetsState::NotParsed(_, _) => unreachable!(),
            SheetsState::Parsed(ref shs) => shs.keys().map(|k| k.to_string()).collect(),
        };
        Ok(Metadata {
               sheets: sheets,
               defined_names: defined_names,
           })
    }

    fn read_worksheet_range(&mut self, name: &str) -> Result<Range<DataType>> {
        let _ = self.parse_workbook()?;
        match self.sheets {
            SheetsState::NotParsed(_, _) => unreachable!(),
            SheetsState::Parsed(ref shs) => {
                shs.get(name)
                    .ok_or_else(|| format!("Sheet '{}' does not exist", name).into())
                    .map(|r| r.0.clone())
            }
        }
    }

    fn read_worksheet_formula(&mut self, name: &str) -> Result<Range<String>> {
        let _ = self.parse_workbook()?;
        match self.sheets {
            SheetsState::NotParsed(_, _) => unreachable!(),
            SheetsState::Parsed(ref shs) => {
                shs.get(name)
                    .ok_or_else(|| format!("Sheet '{}' does not exist", name).into())
                    .map(|r| r.1.clone())
            }
        }
    }
}

impl Xls {
    fn parse_workbook(&mut self) -> Result<Vec<(String, String)>> {

        // gets workbook and worksheets stream, or early exit
        let stream = match self.sheets {
            SheetsState::NotParsed(ref mut reader, ref mut cfb) => {
                cfb.get_stream("Workbook", reader)
                    .or_else(|_| cfb.get_stream("Book", reader))?
            }
            SheetsState::Parsed(_) => return Ok(Vec::new()),
        };

        let mut sheet_names = Vec::new();
        let mut strings = Vec::new();
        let mut defined_names = Vec::new();
        let mut xtis = Vec::new();
        let mut encoding = XlsEncoding::from_codepage(1200)?;
        {
            let mut wb = &stream;
            let records = RecordIter { stream: &mut wb };
            for record in records {
                let mut r = record?;
                match r.typ {
                    0x0012 => {
                        if read_u16(r.data) != 0 {
                            return Err("Workbook is password protected".into());
                        }
                    }
                    0x0042 => encoding = XlsEncoding::from_codepage(read_u16(r.data))?, // CodePage
                    0x013D => {
                        let sheet_len = r.data.len() / 2;
                        sheet_names.reserve(sheet_len);
                    } // RRTabId
                    0x0085 => {
                        let name = parse_sheet_name(&mut r, &mut encoding)?;
                        sheet_names.push(name); // BoundSheet8
                    }
                    0x0018 => {
                        // Lbl for defined_names
                        let mut cch = r.data[3] as usize;
                        let cce = read_u16(&r.data[4..]) as usize;
                        let name =
                            read_unicode_string_no_cch(&mut encoding, &r.data[14..], &mut cch)?;
                        let rgce = &r.data[r.data.len() - cce..];
                        let formula = parse_defined_names(rgce)?;
                        defined_names.push((name, formula));
                    }
                    0x0017 => {
                        // ExternSheet
                        let len = read_u16(r.data) as usize;
                        xtis.reserve(len);
                        let mut start = 4;
                        for _ in 0..len {
                            let i = read_u16(&r.data[start..]) as usize;
                            xtis.push(i);
                            start += 6;
                        }
                    }
                    0x00FC => strings = parse_sst(&mut r, &mut encoding)?, // SST
                    0x000A => break, // EOF,
                    _ => (),
                }
            }
        }

        let defined_names = defined_names
            .into_iter()
            .map(|(name, (i, f))| if let Some(i) = i {
                     (name, format!("{}!{}", sheet_names[xtis[i]].1, f))
                 } else {
                     (name, f)
                 })
            .collect::<Vec<_>>();

        let mut sheets = HashMap::with_capacity(sheet_names.len());
        let fmla_sheet_names = sheet_names
            .iter()
            .map(|&(_, ref n)| n.clone())
            .collect::<Vec<_>>();
        'sh: for (pos, name) in sheet_names.into_iter() {
            let mut sh = &stream[pos..];
            let records = RecordIter { stream: &mut sh };
            let mut cells = Vec::new();
            let mut formulas = Vec::new();
            for record in records {
                let r = record?;
                match r.typ {
                    0x0200 => {
                        let (start, end) = parse_dimensions(&r.data)?;
                        cells.reserve(((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize);
                    } // Dimensions
                    0x0203 => cells.push(parse_number(&r.data)?), // Number
                    0x0205 => cells.push(parse_bool_err(&r.data)?), // BoolErr
                    0x027E => cells.push(parse_rk(&r.data)?), // RK
                    0x00FD => cells.push(parse_label_sst(&r.data, &strings)?), // LabelSst
                    0x000A => break, // EOF,
                    0x0006 => {
                        let row = read_u16(&r.data);
                        let col = read_u16(&r.data[2..]);

                        // Formula
                        let cce = read_u16(&r.data[20..]) as usize;
                        let rgce = &r.data[22..22 + cce];
                        let fmla =
                            parse_formula(rgce, &fmla_sheet_names, &defined_names, &mut encoding)
                                .unwrap_or_else(|e| {
                                                    format!("Unrecognised formula \
                                for cell ({}, {}): {:?}",
                                                            row,
                                                            col,
                                                            e)
                                                });
                        formulas.push(Cell::new((row as u32, col as u32), fmla));
                    }
                    _ => (),
                }
            }
            let range = Range::from_sparse(cells);
            let formula = Range::from_sparse(formulas);
            sheets.insert(name, (range, formula));
        }

        self.sheets = SheetsState::Parsed(sheets);

        Ok(defined_names)
    }
}

fn parse_sheet_name(r: &mut Record, encoding: &mut XlsEncoding) -> Result<(usize, String)> {
    let pos = read_u32(r.data) as usize;
    r.data = &r.data[6..];
    let sheet = parse_short_string(r, encoding)?;
    Ok((pos, sheet))
}

fn parse_number(r: &[u8]) -> Result<Cell<DataType>> {
    if r.len() < 14 {
        return Err("Invalid number length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let v = read_slice::<f64>(&r[6..]);
    Ok(Cell::new((row as u32, col as u32), DataType::Float(v)))
}

fn parse_bool_err(r: &[u8]) -> Result<Cell<DataType>> {
    if r.len() < 8 {
        return Err("Invalid BoolErr length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let v = match r[7] {
        0x00 => DataType::Bool(r[6] != 0),
        0x01 => {
            match r[6] {
                0x00 => DataType::Error(CellErrorType::Null),
                0x07 => DataType::Error(CellErrorType::Div0),
                0x0F => DataType::Error(CellErrorType::Value),
                0x17 => DataType::Error(CellErrorType::Ref),
                0x1D => DataType::Error(CellErrorType::Name),
                0x24 => DataType::Error(CellErrorType::Num),
                0x2A => DataType::Error(CellErrorType::NA),
                0x2B => DataType::Error(CellErrorType::GettingData),
                e => return Err(format!("Unrecognized error {:x}", e).into()),
            }
        }
        e => return Err(format!("Unrecognized fError {:x}", e).into()),
    };
    Ok(Cell::new((row as u32, col as u32), v))
}

fn parse_rk(r: &[u8]) -> Result<Cell<DataType>> {
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
        DataType::Int(if d100 { v / 100 } else { v })
    } else {
        let v = read_slice(&v);
        DataType::Float(if d100 { v / 100.0 } else { v })
    };
    Ok(Cell::new((row as u32, col as u32), v))
}

fn parse_short_string(r: &mut Record, encoding: &mut XlsEncoding) -> Result<String> {
    if r.data.len() < 2 {
        return Err("Invalid short string length".into());
    }
    let len = r.data[0] as usize;
    if let Some(ref mut b) = encoding.high_byte {
        *b = r.data[1] != 0;
    }
    r.data = &r.data[2..];
    read_dbcs(encoding, len, r)
}

fn parse_label_sst(r: &[u8], strings: &[String]) -> Result<Cell<DataType>> {
    if r.len() < 10 {
        return Err("Invalid short string length".into());
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let i = read_u32(&r[6..]) as usize;
    Ok(Cell::new((row as u32, col as u32),
                 DataType::String(strings[i].clone())))
}

fn parse_dimensions(r: &[u8]) -> Result<((u32, u32), (u32, u32))> {
    let (rf, rl, cf, cl) = match r.len() {
        10 => {
            (read_u16(&r[0..2]) as u32,
             read_u16(&r[2..4]) as u32,
             read_u16(&r[4..6]) as u32,
             read_u16(&r[6..8]) as u32)
        }
        14 => {
            (read_u32(&r[0..4]),
             read_u32(&r[4..8]),
             read_u16(&r[8..10]) as u32,
             read_u16(&r[10..12]) as u32)
        }
        _ => return Err("Invalid dimensions lengths".into()),
    };
    if (1, 1) <= (rl, cl) {
        Ok(((rf, cf), (rl - 1, cl - 1)))
    } else {
        Ok(((rf, cf), (rf, cf)))
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
        sst.push(read_rich_extended_string(r, encoding)?);
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

    if let Some(ref mut b) = encoding.high_byte {
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

    let s = read_dbcs(encoding, str_len, r)?;

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

fn read_dbcs<'a>(encoding: &mut XlsEncoding, mut len: usize, r: &mut Record) -> Result<String> {
    let mut s = String::with_capacity(len);
    while len > 0 {
        let (l, at) = encoding.decode_to(r.data, len, &mut s)?;
        r.data = &r.data[at..];
        len -= l;
        if len > 0 {
            if r.continue_record() {
                if let Some(ref mut b) = encoding.high_byte {
                    *b = r.data[0] & 0x1 != 0;
                }
                r.data = &r.data[1..];
            } else {
                return Err("Cannot decode entire dbcs stream".into());
            }
        }
    }
    Ok(s)
}

fn read_unicode_string_no_cch(encoding: &mut XlsEncoding,
                              buf: &[u8],
                              len: &mut usize)
                              -> Result<String> {
    let mut s = String::new();
    if let Some(ref mut b) = encoding.high_byte {
        *b = buf[0] & 0x1 != 0;
        if *b {
            *len *= 2;
        }
    }
    let _ = encoding.decode_to(&buf[1..*len + 1], *len, &mut s)?;
    Ok(s)
}

struct Record<'a> {
    typ: u16,
    data: &'a [u8],
    cont: Option<Vec<&'a [u8]>>,
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
    stream: &'a [u8],
}

impl<'a> Iterator for RecordIter<'a> {
    type Item = Result<Record<'a>>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.stream.len() < 4 {
            return if self.stream.is_empty() {
                       None
                   } else {
                       Some(Err("Expecting record type and length, found end of stream".into()))
                   };
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
                    return Some(Err("Expecting continue record length, found end of stream"
                                        .into()));
                }
                let sp = self.stream.split_at(len + 4);
                cont.push(&sp.0[4..]);
                self.stream = sp.1;
            }
            Some(cont)
        } else {
            None
        };

        Some(Ok(Record {
                    typ: t,
                    data: d,
                    cont: cont,
                }))
    }
}

/// Formula parsing
///
/// Does not implement ALL possibilities, only Area are parsed
fn parse_defined_names(rgce: &[u8]) -> Result<(Option<usize>, String)> {
    let ptg = rgce[0];
    let res = match ptg {
        0x3a | 0x5a | 0x7a => {
            // PtgRef3d
            let ixti = read_u16(&rgce[1..3]) as usize;
            let mut f = String::new();
            // TODO: check with relative columns
            f.push('$');
            push_column(read_u16(&rgce[5..7]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u16(&rgce[3..5]) + 1));
            (Some(ixti), f)
        }
        0x3b | 0x5b | 0x7b => {
            // PtgArea3d
            let ixti = read_u16(&rgce[1..3]) as usize;
            let mut f = String::new();
            // TODO: check with relative columns
            f.push('$');
            push_column(read_u16(&rgce[7..9]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u16(&rgce[3..5]) + 1));
            f.push(':');
            f.push('$');
            push_column(read_u16(&rgce[9..11]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u16(&rgce[5..7]) + 1));
            (Some(ixti), f)
        }
        0x3c | 0x5c | 0x7c | 0x3d | 0x5d | 0x7d => {
            // PtgAreaErr3d or PtfRefErr3d
            let ixti = read_u16(&rgce[1..3]) as usize;
            (Some(ixti), "#REF!".to_string())
        }
        _ => (None, format!("Unsupported ptg: {:x}", ptg)),
    };
    Ok(res)
}

/// Formula parsing
///
/// [MS-XLSB 2.2.2]
/// [MS-XLSB 2.5.97]
fn parse_formula(mut rgce: &[u8],
                 sheets: &[String],
                 names: &[(String, String)],
                 encoding: &mut XlsEncoding)
                 -> Result<String> {
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
                push_column(read_u16(&rgce[4..6]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) + 1));
                rgce = &rgce[6..];
            }
            0x3b | 0x5b | 0x7b => {
                // PtgArea3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                // TODO: check with relative columns
                formula.push('$');
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[8..10]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[4..6]) + 1));
                rgce = &rgce[10..];
            }
            0x3c | 0x5c | 0x7c => {
                // PtfRefErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[6..];
            }
            0x3d | 0x5d | 0x7d => {
                // PtgAreaErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&sheets[ixti as usize]);
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[10..];
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
                let mut cch = rgce[0] as usize;
                formula.push_str(&read_unicode_string_no_cch(encoding, &rgce[1..], &mut cch)?);
                formula.push('\"');
                rgce = &rgce[1 + cch..];
            }
            0x18 => {
                rgce = &rgce[5..];
            }
            0x19 => {
                // ignore most of these ptgs ...
                let etpg = rgce[0];
                rgce = &rgce[1..];
                match etpg {
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
                rgce = &rgce[8..];
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
                let row = read_u16(rgce) + 1;
                let col = [rgce[2], rgce[3] & 0x3F];
                let col = read_u16(&col);
                stack.push(formula.len());
                if rgce[3] & 0x80 != 0x80 {
                    formula.push('$');
                }
                push_column(col as u32, &mut formula);
                if rgce[3] & 0x40 != 0x40 {
                    formula.push('$');
                }
                formula.push_str(&format!("{}", row));
                rgce = &rgce[4..];
            }
            0x25 | 0x45 | 0x65 => {
                stack.push(formula.len());
                formula.push('$');
                push_column(read_u16(&rgce[4..6]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[0..2]) + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) + 1));
                rgce = &rgce[8..];
            }
            0x2A | 0x4A | 0x6A => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[4..];
            }
            0x2B | 0x4B | 0x6B => {
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[8..];
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
