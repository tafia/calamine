use std::borrow::Cow;
use std::cmp::min;
use std::collections::HashMap;
use std::io::SeekFrom;
use std::io::{Read, Seek};
use std::marker::PhantomData;

use cfb::{Cfb, XlsEncoding};
use utils::{push_column, read_slice, read_u16, read_u32};
use vba::VbaProject;
use {Cell, CellErrorType, DataType, Metadata, Range, Reader};

#[derive(Fail, Debug)]
/// An enum to handle Xls specific errors
pub enum XlsError {
    /// Io error
    #[fail(display = "{}", _0)]
    Io(#[cause] ::std::io::Error),
    /// Cfb error
    #[fail(display = "{}", _0)]
    Cfb(#[cause] ::cfb::CfbError),
    /// Vba error
    #[fail(display = "{}", _0)]
    Vba(#[cause] ::vba::VbaError),

    /// Cannot parse formula, stack is too short
    #[fail(display = "Invalid stack length")]
    StackLen,
    /// Unrecognized data
    #[fail(display = "Unrecognized {}: 0x{:0X}", typ, val)]
    Unrecognized {
        /// data type
        typ: &'static str,
        /// value found
        val: u8,
    },
    /// Workook is password protected
    #[fail(display = "Workbook is password protected")]
    Password,
    /// Invalid length
    #[fail(display = "Invalid {} length, expected {} maximum, found {}", typ, expected, found)]
    Len {
        /// expected length
        expected: usize,
        /// found length
        found: usize,
        /// length type
        typ: &'static str,
    },
    /// Continue Record is too short
    #[fail(display = "Continued record too short while reading extended string")]
    ContinueRecordTooShort,
    /// End of stream
    #[fail(display = "End of stream for {}", _0)]
    EoStream(&'static str),

    /// Invalid Formula
    #[fail(display = "Invalid formula (stack size: {})", stack_size)]
    InvalidFormula {
        /// stack size
        stack_size: usize,
    },
    /// Invalid or unknown iftab
    #[fail(display = "Invalid iftab {:X}", _0)]
    IfTab(usize),
    /// Invalid etpg
    #[fail(display = "Invalid etpg {:X}", _0)]
    Etpg(u8),
    /// No vba project
    #[fail(display = "No VBA project")]
    NoVba,
}

from_err!(::std::io::Error, XlsError, Io);
from_err!(::cfb::CfbError, XlsError, Cfb);
from_err!(::vba::VbaError, XlsError, Vba);

/// A struct representing an old xls format file (CFB)
pub struct Xls<RS> {
    sheets: HashMap<String, (Range<DataType>, Range<String>)>,
    vba: Option<VbaProject>,
    metadata: Metadata,
    marker: PhantomData<RS>,
}

impl<RS: Read + Seek> Reader for Xls<RS> {
    type Error = XlsError;
    type RS = RS;

    fn new(mut reader: RS) -> Result<Self, XlsError>
    where
        RS: Read + Seek,
    {
        let mut cfb = {
            let offset_end = reader.seek(SeekFrom::End(0))? as usize;
            reader.seek(SeekFrom::Start(0))?;
            Cfb::new(&mut reader, offset_end)?
        };

        // Reads vba once for all (better than reading all worksheets once for all)
        let vba = if cfb.has_directory("_VBA_PROJECT_CUR") {
            Some(VbaProject::from_cfb(&mut reader, &mut cfb)?)
        } else {
            None
        };

        let mut xls = Xls {
            sheets: HashMap::new(),
            vba: vba,
            marker: PhantomData,
            metadata: Metadata::default(),
        };

        xls.parse_workbook(reader, cfb)?;
        Ok(xls)
    }

    fn vba_project(&mut self) -> Option<Result<Cow<VbaProject>, XlsError>> {
        self.vba.as_ref().map(|vba| Ok(Cow::Borrowed(vba)))
    }

    /// Parses Workbook stream, no need for the relationships variable
    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    fn worksheet_range(&mut self, name: &str) -> Option<Result<Range<DataType>, XlsError>> {
        self.sheets.get(name).map(|r| Ok(r.0.clone()))
    }

    fn worksheet_formula(&mut self, name: &str) -> Option<Result<Range<String>, XlsError>> {
        self.sheets.get(name).map(|r| Ok(r.1.clone()))
    }
}

impl<RS: Read + Seek> Xls<RS> {
    fn parse_workbook(&mut self, mut reader: RS, mut cfb: Cfb) -> Result<(), XlsError> {
        // gets workbook and worksheets stream, or early exit
        let stream = cfb
            .get_stream("Workbook", &mut reader)
            .or_else(|_| cfb.get_stream("Book", &mut reader))?;

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
                    0x0012 if read_u16(r.data) != 0 => return Err(XlsError::Password),
                    0x0042 => encoding = XlsEncoding::from_codepage(read_u16(r.data))?, // CodePage
                    0x013D => {
                        let sheet_len = r.data.len() / 2;
                        sheet_names.reserve(sheet_len);
                    }
                    // RRTabId
                    0x0085 => {
                        let name = parse_sheet_name(&mut r, &mut encoding)?;
                        sheet_names.push(name); // BoundSheet8
                    }
                    0x0018 => {
                        // Lbl for defined_names
                        let mut cch = r.data[3] as usize;
                        let cce = read_u16(&r.data[4..]) as usize;
                        let name =
                            read_unicode_string_no_cch(&mut encoding, &r.data[14..], &mut cch);
                        let rgce = &r.data[r.data.len() - cce..];
                        let formula = parse_defined_names(rgce)?;
                        defined_names.push((name, formula));
                    }
                    0x0017 => {
                        // ExternSheet
                        let cxti = read_u16(r.data) as usize;
                        xtis.extend(
                            r.data[2..]
                                .chunks(6)
                                .take(cxti)
                                .map(|xti| read_u16(&xti[2..]) as usize),
                        );
                    }
                    0x00FC => strings = parse_sst(&mut r, &mut encoding)?, // SST
                    0x000A => break,                                       // EOF,
                    _ => (),
                }
            }
        }

        let defined_names = defined_names
            .into_iter()
            .map(|(name, (i, f))| {
                if let Some(i) = i {
                    if i >= xtis.len() || xtis[i] >= sheet_names.len() {
                        (name, format!("#REF!{}", f))
                    } else {
                        (name, format!("{}!{}", sheet_names[xtis[i]].1, f))
                    }
                } else {
                    (name, f)
                }
            })
            .collect::<Vec<_>>();

        let mut sheets = HashMap::with_capacity(sheet_names.len());
        let fmla_sheet_names = sheet_names
            .iter()
            .map(|&(_, ref n)| n.clone())
            .collect::<Vec<_>>();
        for (pos, name) in sheet_names {
            let mut sh = &stream[pos..];
            let records = RecordIter { stream: &mut sh };
            let mut cells = Vec::new();
            let mut formulas = Vec::new();
            for record in records {
                let r = record?;
                match r.typ {
                    0x0200 => {
                        let (start, end) = parse_dimensions(r.data)?;
                        cells.reserve(((end.0 - start.0 + 1) * (end.1 - start.1 + 1)) as usize);
                    }
                    // Dimensions
                    0x0203 => cells.push(parse_number(r.data)?), // Number
                    0x0205 => cells.push(parse_bool_err(r.data)?), // BoolErr
                    0x027E => cells.push(parse_rk(r.data)?),     // RK
                    0x00FD => cells.push(parse_label_sst(r.data, &strings)?), // LabelSst
                    0x000A => break,                             // EOF,
                    0x0006 => {
                        // Formula
                        let row = read_u16(r.data);
                        let col = read_u16(&r.data[2..]);

                        // Formula
                        let fmla = parse_formula(
                            &r.data[20..],
                            &fmla_sheet_names,
                            &defined_names,
                            &mut encoding,
                        ).unwrap_or_else(|e| {
                            format!(
                                "Unrecognised formula \
                                 for cell ({}, {}): {:?}",
                                row, col, e
                            )
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

        self.sheets = sheets;
        self.metadata.names = defined_names;
        self.metadata.sheets = self.sheets.keys().map(|k| k.to_string()).collect();

        Ok(())
    }
}

/// BoundSheet8 [MS-XLS 2.4.28]
fn parse_sheet_name(
    r: &mut Record,
    encoding: &mut XlsEncoding,
) -> Result<(usize, String), XlsError> {
    let pos = read_u32(r.data) as usize;
    r.data = &r.data[6..];
    parse_short_string(r, encoding).map(|s| (pos, s))
}

fn parse_number(r: &[u8]) -> Result<Cell<DataType>, XlsError> {
    if r.len() < 14 {
        return Err(XlsError::Len {
            typ: "number",
            expected: 14,
            found: r.len(),
        });
    }
    let row = read_u16(r) as u32;
    let col = read_u16(&r[2..]) as u32;
    let v = read_slice::<f64>(&r[6..]);
    Ok(Cell::new((row, col), DataType::Float(v)))
}

fn parse_bool_err(r: &[u8]) -> Result<Cell<DataType>, XlsError> {
    if r.len() < 8 {
        return Err(XlsError::Len {
            typ: "BoolErr",
            expected: 8,
            found: r.len(),
        });
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
            e => {
                return Err(XlsError::Unrecognized {
                    typ: "error",
                    val: e,
                })
            }
        },
        e => {
            return Err(XlsError::Unrecognized {
                typ: "fError",
                val: e,
            })
        }
    };
    Ok(Cell::new((row as u32, col as u32), v))
}

fn parse_rk(r: &[u8]) -> Result<Cell<DataType>, XlsError> {
    if r.len() < 10 {
        return Err(XlsError::Len {
            typ: "rk",
            expected: 10,
            found: r.len(),
        });
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

/// ShortXLUnicodeString [MS-XLS 2.5.240]
fn parse_short_string(r: &mut Record, encoding: &mut XlsEncoding) -> Result<String, XlsError> {
    if r.data.len() < 2 {
        return Err(XlsError::Len {
            typ: "short string",
            expected: 2,
            found: r.data.len(),
        });
    }
    let cch = r.data[0] as usize;
    if let Some(ref mut b) = encoding.high_byte {
        *b = r.data[1] != 0;
    }
    r.data = &r.data[2..];
    let mut s = String::with_capacity(cch);
    let _ = encoding.decode_to(r.data, cch, &mut s);
    Ok(s)
}

fn parse_label_sst(r: &[u8], strings: &[String]) -> Result<Cell<DataType>, XlsError> {
    if r.len() < 10 {
        return Err(XlsError::Len {
            typ: "label sst",
            expected: 10,
            found: r.len(),
        });
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let i = read_u32(&r[6..]) as usize;
    Ok(Cell::new(
        (row as u32, col as u32),
        DataType::String(strings[i].clone()),
    ))
}

fn parse_dimensions(r: &[u8]) -> Result<((u32, u32), (u32, u32)), XlsError> {
    let (rf, rl, cf, cl) = match r.len() {
        10 => (
            read_u16(&r[0..2]) as u32,
            read_u16(&r[2..4]) as u32,
            read_u16(&r[4..6]) as u32,
            read_u16(&r[6..8]) as u32,
        ),
        14 => (
            read_u32(&r[0..4]),
            read_u32(&r[4..8]),
            read_u16(&r[8..10]) as u32,
            read_u16(&r[10..12]) as u32,
        ),
        _ => {
            return Err(XlsError::Len {
                typ: "dimensions",
                expected: 14,
                found: r.len(),
            })
        }
    };
    if (1, 1) <= (rl, cl) {
        Ok(((rf, cf), (rl - 1, cl - 1)))
    } else {
        Ok(((rf, cf), (rf, cf)))
    }
}

fn parse_sst(r: &mut Record, encoding: &mut XlsEncoding) -> Result<Vec<String>, XlsError> {
    if r.data.len() < 8 {
        return Err(XlsError::Len {
            typ: "sst",
            expected: 8,
            found: r.data.len(),
        });
    }
    let len = read_slice::<i32>(&r.data[4..]) as usize;
    let mut sst = Vec::with_capacity(len);
    r.data = &r.data[8..];
    for _ in 0..len {
        sst.push(read_rich_extended_string(r, encoding)?);
    }
    Ok(sst)
}

fn read_rich_extended_string(
    r: &mut Record,
    encoding: &mut XlsEncoding,
) -> Result<String, XlsError> {
    if r.data.is_empty() && !r.continue_record() || r.data.len() < 3 {
        return Err(XlsError::Len {
            typ: "rick extended string",
            expected: 3,
            found: r.data.len(),
        });
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
        let l = 4 * read_u16(r.data) as usize;
        r.data = &r.data[2..];
        l
    } else {
        0
    };
    if ext_st != 0 {
        unused_len += read_slice::<i32>(r.data) as usize;
        r.data = &r.data[4..];
    };

    let s = read_dbcs(encoding, str_len, r)?;

    while unused_len > 0 {
        if r.data.is_empty() && !r.continue_record() {
            return Err(XlsError::ContinueRecordTooShort);
        }
        let l = min(unused_len, r.data.len());
        let (_, next) = r.data.split_at(l);
        r.data = next;
        unused_len -= l;
    }

    Ok(s)
}

fn read_dbcs(
    encoding: &mut XlsEncoding,
    mut len: usize,
    r: &mut Record,
) -> Result<String, XlsError> {
    let mut s = String::with_capacity(len);
    while len > 0 {
        let (l, at) = encoding.decode_to(r.data, len, &mut s);
        r.data = &r.data[at..];
        len -= l;
        if len > 0 {
            if r.continue_record() {
                if let Some(ref mut b) = encoding.high_byte {
                    *b = r.data[0] & 0x1 != 0;
                }
                r.data = &r.data[1..];
            } else {
                return Err(XlsError::EoStream("dbcs"));
            }
        }
    }
    Ok(s)
}

fn read_unicode_string_no_cch(encoding: &mut XlsEncoding, buf: &[u8], len: &mut usize) -> String {
    let mut s = String::new();
    if let Some(ref mut b) = encoding.high_byte {
        *b = buf[0] & 0x1 != 0;
        if *b {
            *len *= 2;
        }
    }
    let _ = encoding.decode_to(&buf[1..*len + 1], *len, &mut s);
    s
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
            Some(ref mut v) => if v.is_empty() {
                false
            } else {
                self.data = v.remove(0);
                true
            },
        }
    }
}

struct RecordIter<'a> {
    stream: &'a [u8],
}

impl<'a> Iterator for RecordIter<'a> {
    type Item = Result<Record<'a>, XlsError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.stream.len() < 4 {
            return if self.stream.is_empty() {
                None
            } else {
                Some(Err(XlsError::EoStream("record type and length")))
            };
        }
        let t = read_u16(self.stream);
        let mut len = read_u16(&self.stream[2..]) as usize;
        if self.stream.len() < len + 4 {
            return Some(Err(XlsError::EoStream("record length")));
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
                    return Some(Err(XlsError::EoStream("continue record length")));
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
fn parse_defined_names(rgce: &[u8]) -> Result<(Option<usize>, String), XlsError> {
    if rgce.is_empty() {
        // TODO: do something better here ...
        return Ok((None, "empty rgce".to_string()));
    }
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
            f.push_str(&format!("{}", read_u16(&rgce[3..5]) as u32 + 1));
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
            f.push_str(&format!("{}", read_u16(&rgce[3..5]) as u32 + 1));
            f.push(':');
            f.push('$');
            push_column(read_u16(&rgce[9..11]) as u32, &mut f);
            f.push('$');
            f.push_str(&format!("{}", read_u16(&rgce[5..7]) as u32 + 1));
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
/// CellParsedForumula [MS-XLS 2.5.198.3]
fn parse_formula(
    mut rgce: &[u8],
    sheets: &[String],
    names: &[(String, String)],
    encoding: &mut XlsEncoding,
) -> Result<String, XlsError> {
    let mut stack = Vec::new();
    let mut formula = String::with_capacity(rgce.len());
    let cce = read_u16(rgce) as usize;
    rgce = &rgce[2..2 + cce];
    while !rgce.is_empty() {
        let ptg = rgce[0];
        rgce = &rgce[1..];
        match ptg {
            0x3a | 0x5a | 0x7a => {
                // PtgRef3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(sheets.get(ixti as usize).map_or("#REF", |s| &**s));
                formula.push('!');
                // TODO: check with relative columns
                formula.push('$');
                push_column(read_u16(&rgce[4..6]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) as u32 + 1));
                rgce = &rgce[6..];
            }
            0x3b | 0x5b | 0x7b => {
                // PtgArea3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(sheets.get(ixti as usize).map_or("#REF", |s| &**s));
                formula.push('!');
                // TODO: check with relative columns
                formula.push('$');
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) as u32 + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[8..10]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[4..6]) as u32 + 1));
                rgce = &rgce[10..];
            }
            0x3c | 0x5c | 0x7c => {
                // PtfRefErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(sheets.get(ixti as usize).map_or("#REF", |s| &**s));
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[6..];
            }
            0x3d | 0x5d | 0x7d => {
                // PtgAreaErr3d
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(sheets.get(ixti as usize).map_or("#REF", |s| &**s));
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[10..];
            }
            0x01 => {
                // PtgExp: array/shared formula, ignore
                debug!("ignoring PtgExp array/shared formula");
                stack.push(formula.len());
                rgce = &rgce[4..];
            }
            0x03...0x11 => {
                // binary operation
                let e2 = stack.pop().ok_or(XlsError::StackLen)?;
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
                let e = stack.last().ok_or(XlsError::StackLen)?;
                formula.insert(*e, '+');
            }
            0x13 => {
                let e = stack.last().ok_or(XlsError::StackLen)?;
                formula.insert(*e, '-');
            }
            0x14 => {
                formula.push('%');
            }
            0x15 => {
                let e = stack.last().ok_or(XlsError::StackLen)?;
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
                formula.push_str(&read_unicode_string_no_cch(encoding, &rgce[1..], &mut cch));
                formula.push('\"');
                rgce = &rgce[2 + cch..];
            }
            0x18 => {
                rgce = &rgce[5..];
            }
            0x19 => {
                // ignore most of these ptgs ...
                let etpg = rgce[0];
                rgce = &rgce[1..];
                match etpg {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x40 | 0x41 => rgce = &rgce[2..],
                    0x04 => rgce = &rgce[10..],
                    0x10 => {
                        rgce = &rgce[2..];
                        let e = *stack.last().ok_or(XlsError::StackLen)?;
                        let e = formula.split_off(e);
                        formula.push_str("SUM(");
                        formula.push_str(&e);
                        formula.push(')');
                    }
                    e => return Err(XlsError::Etpg(e)),
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
                    e => {
                        return Err(XlsError::Unrecognized {
                            typ: "BErr",
                            val: e,
                        })
                    }
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
                rgce = &rgce[7..];
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
                            return Err(XlsError::IfTab(iftab));
                        }
                        rgce = &rgce[2..];
                        let argc = ::utils::FTAB_ARGC[iftab] as usize;
                        (iftab, argc)
                    }
                };
                if stack.len() < argc {
                    return Err(XlsError::StackLen);
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
                formula.push_str(names.get(iname).map_or("#REF!", |n| &*n.0));
                rgce = &rgce[4..];
            }
            0x24 | 0x44 | 0x64 => {
                stack.push(formula.len());
                let row = read_u16(rgce) + 1;
                let col = read_u16(&[rgce[2], rgce[3] & 0x3F]);
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
                formula.push_str(&format!("{}", read_u16(&rgce[0..2]) as u32 + 1));
                formula.push(':');
                formula.push('$');
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                formula.push('$');
                formula.push_str(&format!("{}", read_u16(&rgce[2..4]) as u32 + 1));
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
            _ => {
                return Err(XlsError::Unrecognized {
                    typ: "ptg",
                    val: ptg,
                })
            }
        }
    }
    if stack.len() != 1 {
        Err(XlsError::InvalidFormula {
            stack_size: stack.len(),
        })
    } else {
        Ok(formula)
    }
}
