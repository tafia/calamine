use std::borrow::Cow;
use std::cmp::min;
use std::collections::BTreeMap;
use std::convert::TryInto;
use std::fmt::Write;
use std::io::{Read, Seek, SeekFrom};
use std::marker::PhantomData;

use tracing::debug;

use crate::cfb::{Cfb, XlsEncoding};
use crate::formats::{
    builtin_format_by_code, detect_custom_number_format, format_excel_f64, format_excel_i64,
    CellFormat,
};
#[cfg(feature = "picture")]
use crate::utils::read_usize;
use crate::utils::{push_column, read_f64, read_i16, read_i32, read_u16, read_u32};
use crate::vba::VbaProject;
use crate::{Cell, CellErrorType, Data, Metadata, Range, Reader, Sheet, SheetType, SheetVisible};

#[derive(Debug)]
/// An enum to handle Xls specific errors
pub enum XlsError {
    /// Io error
    Io(std::io::Error),
    /// Cfb error
    Cfb(crate::cfb::CfbError),
    /// Vba error
    Vba(crate::vba::VbaError),

    /// Cannot parse formula, stack is too short
    StackLen,
    /// Unrecognized data
    Unrecognized {
        /// data type
        typ: &'static str,
        /// value found
        val: u8,
    },
    /// Workbook is password protected
    Password,
    /// Invalid length
    Len {
        /// expected length
        expected: usize,
        /// found length
        found: usize,
        /// length type
        typ: &'static str,
    },
    /// Continue Record is too short
    ContinueRecordTooShort,
    /// End of stream
    EoStream(&'static str),

    /// Invalid Formula
    InvalidFormula {
        /// stack size
        stack_size: usize,
    },
    /// Invalid or unknown iftab
    IfTab(usize),
    /// Invalid etpg
    Etpg(u8),
    /// No vba project
    NoVba,
    /// Invalid OfficeArt Record
    #[cfg(feature = "picture")]
    Art(&'static str),
    /// Worksheet not found
    WorksheetNotFound(String),
}

from_err!(std::io::Error, XlsError, Io);
from_err!(crate::cfb::CfbError, XlsError, Cfb);
from_err!(crate::vba::VbaError, XlsError, Vba);

impl std::fmt::Display for XlsError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsError::Io(e) => write!(f, "I/O error: {e}"),
            XlsError::Cfb(e) => write!(f, "Cfb error: {e}"),
            XlsError::Vba(e) => write!(f, "Vba error: {e}"),
            XlsError::StackLen => write!(f, "Invalid stack length"),
            XlsError::Unrecognized { typ, val } => write!(f, "Unrecognized {typ}: 0x{val:0X}"),
            XlsError::Password => write!(f, "Workbook is password protected"),
            XlsError::Len {
                expected,
                found,
                typ,
            } => write!(
                f,
                "Invalid {typ} length, expected {expected} maximum, found {found}",
            ),
            XlsError::ContinueRecordTooShort => write!(
                f,
                "Continued record too short while reading extended string"
            ),
            XlsError::EoStream(s) => write!(f, "End of stream '{s}'"),
            XlsError::InvalidFormula { stack_size } => {
                write!(f, "Invalid formula (stack size: {stack_size})")
            }
            XlsError::IfTab(iftab) => write!(f, "Invalid iftab {iftab:X}"),
            XlsError::Etpg(etpg) => write!(f, "Invalid etpg {etpg:X}"),
            XlsError::NoVba => write!(f, "No VBA project"),
            #[cfg(feature = "picture")]
            XlsError::Art(s) => write!(f, "Invalid art record '{s}'"),
            XlsError::WorksheetNotFound(name) => write!(f, "Worksheet '{name}' not found"),
        }
    }
}

impl std::error::Error for XlsError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            XlsError::Io(e) => Some(e),
            XlsError::Cfb(e) => Some(e),
            XlsError::Vba(e) => Some(e),
            _ => None,
        }
    }
}

/// Options to perform specialized parsing.
#[derive(Debug, Clone, Default)]
#[non_exhaustive]
pub struct XlsOptions {
    /// Force a spreadsheet to be interpreted using a particular code page.
    ///
    /// XLS files can contain [code page] identifiers. If this identifier is missing or incorrect,
    /// strings in the parsed spreadsheet may be decoded incorrectly. Setting this field causes
    /// `calamine::Xls` to interpret strings using the specified code page, which may allow such
    /// spreadsheets to be decoded properly.
    ///
    /// [code page]: https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
    pub force_codepage: Option<u16>,
}

/// A struct representing an old xls format file (CFB)
pub struct Xls<RS> {
    sheets: BTreeMap<String, (Range<Data>, Range<String>)>,
    vba: Option<VbaProject>,
    metadata: Metadata,
    marker: PhantomData<RS>,
    options: XlsOptions,
    formats: Vec<CellFormat>,
    is_1904: bool,
    #[cfg(feature = "picture")]
    pictures: Option<Vec<(String, Vec<u8>)>>,
}

impl<RS: Read + Seek> Xls<RS> {
    /// Creates a new instance using `Options` to inform parsing.
    ///
    /// ```
    /// use calamine::{Xls,XlsOptions};
    /// # use std::io::Cursor;
    /// # const BYTES: &'static [u8] = b"";
    ///
    /// # fn run() -> Result<Xls<Cursor<&'static [u8]>>, calamine::XlsError> {
    /// # let reader = std::io::Cursor::new(BYTES);
    /// let mut options = XlsOptions::default();
    /// // ...set options...
    /// let workbook = Xls::new_with_options(reader, options)?;
    /// # Ok(workbook) }
    /// # fn main() { assert!(run().is_err()); }
    /// ```
    pub fn new_with_options(mut reader: RS, options: XlsOptions) -> Result<Self, XlsError> {
        let mut cfb = {
            let offset_end = reader.seek(SeekFrom::End(0))? as usize;
            reader.seek(SeekFrom::Start(0))?;
            Cfb::new(&mut reader, offset_end)?
        };

        debug!("cfb loaded");

        // Reads vba once for all (better than reading all worksheets once for all)
        let vba = if cfb.has_directory("_VBA_PROJECT_CUR") {
            Some(VbaProject::from_cfb(&mut reader, &mut cfb)?)
        } else {
            None
        };

        debug!("vba ok");

        let mut xls = Xls {
            sheets: BTreeMap::new(),
            vba,
            marker: PhantomData,
            metadata: Metadata::default(),
            options,
            is_1904: false,
            formats: Vec::new(),
            #[cfg(feature = "picture")]
            pictures: None,
        };

        xls.parse_workbook(reader, cfb)?;

        debug!("xls parsed");

        Ok(xls)
    }
}

impl<RS: Read + Seek> Reader<RS> for Xls<RS> {
    type Error = XlsError;

    fn new(reader: RS) -> Result<Self, XlsError> {
        Self::new_with_options(reader, XlsOptions::default())
    }

    fn vba_project(&mut self) -> Option<Result<Cow<'_, VbaProject>, XlsError>> {
        self.vba.as_ref().map(|vba| Ok(Cow::Borrowed(vba)))
    }

    /// Parses Workbook stream, no need for the relationships variable
    fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    fn worksheet_range(&mut self, name: &str) -> Result<Range<Data>, XlsError> {
        self.sheets
            .get(name)
            .map(|r| r.0.clone())
            .ok_or_else(|| XlsError::WorksheetNotFound(name.into()))
    }

    fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
        self.sheets
            .iter()
            .map(|(name, (data, _))| (name.to_owned(), data.clone()))
            .collect()
    }

    fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, XlsError> {
        self.sheets
            .get(name)
            .ok_or_else(|| XlsError::WorksheetNotFound(name.into()))
            .map(|r| r.1.clone())
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        self.pictures.to_owned()
    }
}

#[derive(Debug, Clone, Copy)]
struct Xti {
    _isup_book: u16,
    itab_first: i16,
    _itab_last: i16,
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
        let mut formats = BTreeMap::new();
        let mut xfs = Vec::new();
        let mut biff = Biff::Biff8; // Binary Interchange File Format (BIFF) version
        let codepage = self.options.force_codepage.unwrap_or(1200);
        let mut encoding = XlsEncoding::from_codepage(codepage)?;
        #[cfg(feature = "picture")]
        let mut draw_group: Vec<u8> = Vec::new();
        {
            let wb = &stream;
            let records = RecordIter { stream: wb };
            for record in records {
                let mut r = record?;
                match r.typ {
                    // 2.4.117 FilePass
                    0x002F if read_u16(r.data) != 0 => return Err(XlsError::Password),
                    // CodePage
                    0x0042 => {
                        if self.options.force_codepage.is_none() {
                            encoding = XlsEncoding::from_codepage(read_u16(r.data))?
                        }
                    }
                    0x013D => {
                        let sheet_len = r.data.len() / 2;
                        sheet_names.reserve(sheet_len);
                        self.metadata.sheets.reserve(sheet_len);
                    }
                    // Date1904
                    0x0022 => {
                        if read_u16(r.data) == 1 {
                            self.is_1904 = true
                        }
                    }
                    // FORMATTING
                    0x041E => {
                        let (idx, format) = parse_format(&mut r, &encoding)?;
                        formats.insert(idx, format);
                    }
                    // XFS
                    0x00E0 => {
                        xfs.push(parse_xf(&r)?);
                    }
                    // RRTabId
                    0x0085 => {
                        let (pos, sheet) = parse_sheet_metadata(&mut r, &encoding, biff)?;
                        self.metadata.sheets.push(sheet.clone());
                        sheet_names.push((pos, sheet.name)); // BoundSheet8
                    }
                    // BOF
                    0x0809 => {
                        let bof = parse_bof(&mut r)?;
                        biff = bof.biff;
                    }
                    0x0018 => {
                        // Lbl for defined_names
                        let cch = r.data[3] as usize;
                        let cce = read_u16(&r.data[4..]) as usize;
                        let mut name = String::new();
                        read_unicode_string_no_cch(&encoding, &r.data[14..], &cch, &mut name);
                        let rgce = &r.data[r.data.len() - cce..];
                        let formula = parse_defined_names(rgce)?;
                        defined_names.push((name, formula));
                    }
                    0x0017 => {
                        // ExternSheet
                        let cxti = read_u16(r.data) as usize;
                        xtis.extend(r.data[2..].chunks_exact(6).take(cxti).map(|xti| Xti {
                            _isup_book: read_u16(&xti[..2]),
                            itab_first: read_i16(&xti[2..4]),
                            _itab_last: read_i16(&xti[4..]),
                        }));
                    }
                    0x00FC => strings = parse_sst(&mut r, &encoding)?, // SST
                    #[cfg(feature = "picture")]
                    0x00EB => {
                        // MsoDrawingGroup
                        draw_group.extend(r.data);
                        if let Some(cont) = r.cont {
                            draw_group.extend(cont.iter().flat_map(|v| *v));
                        }
                    }
                    0x000A => break, // EOF,
                    _ => (),
                }
            }
        }

        self.formats = xfs
            .into_iter()
            .map(|fmt| match formats.get(&fmt) {
                Some(s) => *s,
                _ => builtin_format_by_code(fmt),
            })
            .collect();

        debug!("formats: {:?}", self.formats);

        let defined_names = defined_names
            .into_iter()
            .map(|(name, (i, mut f))| {
                if let Some(i) = i {
                    let sh = xtis
                        .get(i)
                        .and_then(|xti| sheet_names.get(xti.itab_first as usize))
                        .map_or("#REF", |sh| &sh.1);
                    f = format!("{sh}!{f}");
                }
                (name, f)
            })
            .collect::<Vec<_>>();

        debug!("defined_names: {:?}", defined_names);

        let mut sheets = BTreeMap::new();
        let fmla_sheet_names = sheet_names
            .iter()
            .map(|(_, n)| n.clone())
            .collect::<Vec<_>>();
        for (pos, name) in sheet_names {
            let sh = &stream[pos..];
            let records = RecordIter { stream: sh };
            let mut cells = Vec::new();
            let mut formulas = Vec::new();
            let mut fmla_pos = (0, 0);
            for record in records {
                let r = record?;
                match r.typ {
                    // 512: Dimensions
                    0x0200 => {
                        let Dimensions { start, end } = parse_dimensions(r.data)?;
                        let rows = (end.0 - start.0 + 1) as usize;
                        let cols = (end.1 - start.1 + 1) as usize;
                        cells.reserve(rows.saturating_mul(cols));
                    }
                    //0x0201 => cells.push(parse_blank(r.data)?), // 513: Blank
                    0x0203 => cells.push(parse_number(r.data, &self.formats, self.is_1904)?), // 515: Number
                    0x0204 => cells.extend(parse_label(r.data, &encoding, biff)?), // 516: Label [MS-XLS 2.4.148]
                    0x0205 => cells.push(parse_bool_err(r.data)?),                 // 517: BoolErr
                    0x0207 => {
                        // 519 String (formula value)
                        let val = Data::String(parse_string(r.data, &encoding, biff)?);
                        cells.push(Cell::new(fmla_pos, val))
                    }
                    0x027E => cells.push(parse_rk(r.data, &self.formats, self.is_1904)?), // 638: Rk
                    0x00FD => cells.extend(parse_label_sst(r.data, &strings)?), // LabelSst
                    0x00BD => parse_mul_rk(r.data, &mut cells, &self.formats, self.is_1904)?, // 189: MulRk
                    0x000A => break, // 10: EOF,
                    0x0006 => {
                        // 6: Formula
                        if r.data.len() < 20 {
                            return Err(XlsError::Len {
                                expected: 20,
                                found: r.data.len(),
                                typ: "Formula",
                            });
                        }
                        let row = read_u16(r.data);
                        let col = read_u16(&r.data[2..]);
                        fmla_pos = (row as u32, col as u32);
                        if let Some(val) = parse_formula_value(&r.data[6..14])? {
                            // If the value is a string
                            // it will appear in 0x0207 record coming next
                            cells.push(Cell::new(fmla_pos, val));
                        }
                        let fmla = parse_formula(
                            &r.data[20..],
                            &fmla_sheet_names,
                            &defined_names,
                            &xtis,
                            &encoding,
                        )
                        .unwrap_or_else(|e| {
                            debug!("{}", e);
                            format!(
                                "Unrecognised formula \
                                 for cell ({}, {}): {:?}",
                                row, col, e
                            )
                        });
                        formulas.push(Cell::new(fmla_pos, fmla));
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

        #[cfg(feature = "picture")]
        if !draw_group.is_empty() {
            let pics = parse_pictures(&draw_group)?;
            if !pics.is_empty() {
                self.pictures = Some(pics);
            }
        }

        Ok(())
    }
}

/// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/4d6a3d1e-d7c5-405f-bbae-d01e9cb79366
struct Bof {
    /// Binary Interchange File Format
    biff: Biff,
}

/// https://www.loc.gov/preservation/digital/formats/fdd/fdd000510.shtml#notes
#[derive(Clone, Copy)]
enum Biff {
    Biff2,
    Biff3,
    Biff4,
    Biff5,
    Biff8,
    // Used by MS-XLSB Workbook(2.1.7.61) or Worksheet(2.1.7.61) which are not supported yet.
    // Biff12,
}

/// BOF [MS-XLS] 2.4.21
fn parse_bof(r: &mut Record<'_>) -> Result<Bof, XlsError> {
    let mut dt = 0;
    let biff_version = read_u16(&r.data[..2]);

    if r.data.len() >= 4 {
        dt = read_u16(&r.data[2..]);
    };

    let biff = match biff_version {
        0x0200 | 0x0002 | 0x0007 => Biff::Biff2,
        0x0300 => Biff::Biff3,
        0x0400 => Biff::Biff4,
        0x0500 => Biff::Biff5,
        0x0600 => Biff::Biff8,
        0 => {
            if dt == 0x1000 {
                Biff::Biff5
            } else {
                Biff::Biff8
            }
        }
        _ => Biff::Biff8,
    };

    Ok(Bof { biff })
}

/// BoundSheet8 [MS-XLS 2.4.28]
fn parse_sheet_metadata(
    r: &mut Record<'_>,
    encoding: &XlsEncoding,
    biff: Biff,
) -> Result<(usize, Sheet), XlsError> {
    let pos = read_u32(r.data) as usize;
    let visible = match r.data[4] & 0b0011_1111 {
        0x00 => SheetVisible::Visible,
        0x01 => SheetVisible::Hidden,
        0x02 => SheetVisible::VeryHidden,
        e => {
            return Err(XlsError::Unrecognized {
                typ: "BoundSheet8:hsState",
                val: e,
            });
        }
    };
    let typ = match r.data[5] {
        0x00 => SheetType::WorkSheet,
        0x01 => SheetType::MacroSheet,
        0x02 => SheetType::ChartSheet,
        0x06 => SheetType::Vba,
        e => {
            return Err(XlsError::Unrecognized {
                typ: "BoundSheet8:dt",
                val: e,
            });
        }
    };
    r.data = &r.data[6..];
    let name = parse_short_string(r, encoding, biff)?;
    let sheet_name = name
        .as_bytes()
        .iter()
        .cloned()
        .filter(|b| *b != 0)
        .collect::<Vec<_>>();
    let name = String::from_utf8(sheet_name).unwrap();
    Ok((pos, Sheet { name, visible, typ }))
}

fn parse_number(r: &[u8], formats: &[CellFormat], is_1904: bool) -> Result<Cell<Data>, XlsError> {
    if r.len() < 14 {
        return Err(XlsError::Len {
            typ: "number",
            expected: 14,
            found: r.len(),
        });
    }
    let row = read_u16(r) as u32;
    let col = read_u16(&r[2..]) as u32;
    let v = read_f64(&r[6..]);
    let format = formats.get(read_u16(&r[4..]) as usize);

    Ok(Cell::new((row, col), format_excel_f64(v, format, is_1904)))
}

fn parse_bool_err(r: &[u8]) -> Result<Cell<Data>, XlsError> {
    if r.len() < 8 {
        return Err(XlsError::Len {
            typ: "BoolErr",
            expected: 8,
            found: r.len(),
        });
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let pos = (row as u32, col as u32);
    match r[7] {
        0x00 => Ok(Cell::new(pos, Data::Bool(r[6] != 0))),
        0x01 => Ok(Cell::new(pos, parse_err(r[6])?)),
        e => Err(XlsError::Unrecognized {
            typ: "fError",
            val: e,
        }),
    }
}

fn parse_err(e: u8) -> Result<Data, XlsError> {
    match e {
        0x00 => Ok(Data::Error(CellErrorType::Null)),
        0x07 => Ok(Data::Error(CellErrorType::Div0)),
        0x0F => Ok(Data::Error(CellErrorType::Value)),
        0x17 => Ok(Data::Error(CellErrorType::Ref)),
        0x1D => Ok(Data::Error(CellErrorType::Name)),
        0x24 => Ok(Data::Error(CellErrorType::Num)),
        0x2A => Ok(Data::Error(CellErrorType::NA)),
        0x2B => Ok(Data::Error(CellErrorType::GettingData)),
        e => Err(XlsError::Unrecognized {
            typ: "error",
            val: e,
        }),
    }
}

fn parse_rk(r: &[u8], formats: &[CellFormat], is_1904: bool) -> Result<Cell<Data>, XlsError> {
    if r.len() < 10 {
        return Err(XlsError::Len {
            typ: "rk",
            expected: 10,
            found: r.len(),
        });
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);

    Ok(Cell::new(
        (row as u32, col as u32),
        rk_num(&r[4..10], formats, is_1904),
    ))
}

fn parse_mul_rk(
    r: &[u8],
    cells: &mut Vec<Cell<Data>>,
    formats: &[CellFormat],
    is_1904: bool,
) -> Result<(), XlsError> {
    if r.len() < 6 {
        return Err(XlsError::Len {
            typ: "rk",
            expected: 6,
            found: r.len(),
        });
    }

    let row = read_u16(r);
    let col_first = read_u16(&r[2..]);
    let col_last = read_u16(&r[r.len() - 2..]);

    if r.len() != 6 + 6 * (col_last - col_first + 1) as usize {
        return Err(XlsError::Len {
            typ: "rk",
            expected: 6 + 6 * (col_last - col_first + 1) as usize,
            found: r.len(),
        });
    }

    let mut col = col_first as u32;

    for rk in r[4..r.len() - 2].chunks(6) {
        cells.push(Cell::new((row as u32, col), rk_num(rk, formats, is_1904)));
        col += 1;
    }
    Ok(())
}

fn rk_num(rk: &[u8], formats: &[CellFormat], is_1904: bool) -> Data {
    let d100 = (rk[2] & 1) != 0;
    let is_int = (rk[2] & 2) != 0;
    let format = formats.get(read_u16(rk) as usize);

    let mut v = [0u8; 8];
    v[4..].copy_from_slice(&rk[2..]);
    v[4] &= 0xFC;
    if is_int {
        let v = (read_i32(&v[4..8]) >> 2) as i64;
        if d100 && v % 100 != 0 {
            format_excel_f64(v as f64 / 100.0, format, is_1904)
        } else {
            format_excel_i64(if d100 { v / 100 } else { v }, format, is_1904)
        }
    } else {
        let v = read_f64(&v);
        format_excel_f64(if d100 { v / 100.0 } else { v }, format, is_1904)
    }
}

/// ShortXLUnicodeString [MS-XLS 2.5.240]
fn parse_short_string(
    r: &mut Record<'_>,
    encoding: &XlsEncoding,
    biff: Biff,
) -> Result<String, XlsError> {
    if r.data.len() < 2 {
        return Err(XlsError::Len {
            typ: "short string",
            expected: 2,
            found: r.data.len(),
        });
    }

    let cch = r.data[0] as usize;
    r.data = &r.data[1..];
    let mut high_byte = None;

    if matches!(biff, Biff::Biff8) {
        high_byte = Some(r.data[0] & 0x1 != 0);
        r.data = &r.data[1..];
    }

    let mut s = String::with_capacity(cch);
    let _ = encoding.decode_to(r.data, cch, &mut s, high_byte);
    Ok(s)
}

/// XLUnicodeString [MS-XLS 2.5.294]
fn parse_string(r: &[u8], encoding: &XlsEncoding, biff: Biff) -> Result<String, XlsError> {
    if r.len() < 4 {
        return Err(XlsError::Len {
            typ: "string",
            expected: 4,
            found: r.len(),
        });
    }
    let cch = read_u16(r) as usize;

    let (high_byte, start) = match biff {
        Biff::Biff2 | Biff::Biff3 | Biff::Biff4 | Biff::Biff5 => (None, 2),
        _ => (Some(r[2] & 0x1 != 0), 3),
    };

    let mut s = String::with_capacity(cch);
    let _ = encoding.decode_to(&r[start..], cch, &mut s, high_byte);
    Ok(s)
}

fn parse_label(
    r: &[u8],
    encoding: &XlsEncoding,
    biff: Biff,
) -> Result<Option<Cell<Data>>, XlsError> {
    if r.len() < 6 {
        return Err(XlsError::Len {
            typ: "label",
            expected: 6,
            found: r.len(),
        });
    }
    let row = read_u16(r);
    let col = read_u16(&r[2..]);
    let _ixfe = read_u16(&r[4..]);
    return Ok(Some(Cell::new(
        (row as u32, col as u32),
        Data::String(parse_string(&r[6..], encoding, biff)?),
    )));
}

fn parse_label_sst(r: &[u8], strings: &[String]) -> Result<Option<Cell<Data>>, XlsError> {
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
    if let Some(s) = strings.get(i) {
        if !s.is_empty() {
            return Ok(Some(Cell::new(
                (row as u32, col as u32),
                Data::String(s.clone()),
            )));
        }
    }
    Ok(None)
}

struct Dimensions {
    start: (u32, u32),
    end: (u32, u32),
}

fn parse_dimensions(r: &[u8]) -> Result<Dimensions, XlsError> {
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
            });
        }
    };
    if 1 <= rl && 1 <= cl {
        Ok(Dimensions {
            start: (rf, cf),
            end: (rl - 1, cl - 1),
        })
    } else {
        Ok(Dimensions {
            start: (rf, cf),
            end: (rf, cf),
        })
    }
}

fn parse_sst(r: &mut Record<'_>, encoding: &XlsEncoding) -> Result<Vec<String>, XlsError> {
    if r.data.len() < 8 {
        return Err(XlsError::Len {
            typ: "sst",
            expected: 8,
            found: r.data.len(),
        });
    }
    let len: usize = read_i32(&r.data[4..8]).try_into().unwrap();
    let mut sst = Vec::with_capacity(len);
    r.data = &r.data[8..];

    for _ in 0..len {
        sst.push(read_rich_extended_string(r, encoding)?);
    }
    Ok(sst)
}

/// Decode XF (extract only ifmt - Format identifier)
///
/// See: https://learn.microsoft.com/ru-ru/openspecs/office_file_formats/ms-xls/993d15c4-ec04-43e9-ba36-594dfb336c6d
fn parse_xf(r: &Record<'_>) -> Result<u16, XlsError> {
    if r.data.len() < 4 {
        return Err(XlsError::Len {
            typ: "xf",
            expected: 4,
            found: r.data.len(),
        });
    }

    Ok(read_u16(&r.data[2..]))
}

/// Decode Format
///
/// See: https://learn.microsoft.com/ru-ru/openspecs/office_file_formats/ms-xls/300280fd-e4fe-4675-a924-4d383af48d3b
fn parse_format(r: &mut Record<'_>, encoding: &XlsEncoding) -> Result<(u16, CellFormat), XlsError> {
    if r.data.len() < 4 {
        return Err(XlsError::Len {
            typ: "format",
            expected: 4,
            found: r.data.len(),
        });
    }

    let idx = read_u16(r.data);

    let cch = read_u16(&r.data[2..]) as usize;
    let high_byte = r.data[4] & 0x1 != 0;
    r.data = &r.data[5..];
    let mut s = String::with_capacity(cch);
    encoding.decode_to(r.data, cch, &mut s, Some(high_byte));

    Ok((idx, detect_custom_number_format(&s)))
}

/// Decode XLUnicodeRichExtendedString.
///
/// See: <https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/173d9f51-e5d3-43da-8de2-be7f22e119b9>
fn read_rich_extended_string(
    r: &mut Record<'_>,
    encoding: &XlsEncoding,
) -> Result<String, XlsError> {
    if r.data.is_empty() && !r.continue_record() || r.data.len() < 3 {
        return Err(XlsError::Len {
            typ: "rich extended string",
            expected: 3,
            found: r.data.len(),
        });
    }

    let cch = read_u16(r.data) as usize;
    let flags = r.data[2];

    r.data = &r.data[3..];

    let high_byte = flags & 0x1 != 0;

    // how many FormatRun in rgRun data block
    let mut c_run = 0;

    // how many bytes in ExtRst data block
    let mut cb_ext_rst = 0;

    // if flag fRichSt exists, read cRun and forward.
    if flags & 0x8 != 0 {
        c_run = read_u16(r.data) as usize;
        r.data = &r.data[2..];
    }

    // if flag fExtSt exists, read cbExtRst and forward.
    if flags & 0x4 != 0 {
        cb_ext_rst = read_i32(r.data) as usize;
        r.data = &r.data[4..];
    }

    // read rgb data block for the string we want
    let s = read_dbcs(encoding, cch, r, high_byte)?;

    // skip rgRun data block. Note: each FormatRun contain 4 bytes.
    r.skip(c_run * 4)?;

    // skip ExtRst data block.
    r.skip(cb_ext_rst)?;

    Ok(s)
}

fn read_dbcs(
    encoding: &XlsEncoding,
    mut len: usize,
    r: &mut Record<'_>,
    mut high_byte: bool,
) -> Result<String, XlsError> {
    let mut s = String::with_capacity(len);
    while len > 0 {
        let (l, at) = encoding.decode_to(r.data, len, &mut s, Some(high_byte));
        r.data = &r.data[at..];
        len -= l;
        if len > 0 {
            if r.continue_record() {
                high_byte = r.data[0] & 0x1 != 0;
                r.data = &r.data[1..];
            } else {
                return Err(XlsError::EoStream("dbcs"));
            }
        }
    }
    Ok(s)
}

fn read_unicode_string_no_cch(encoding: &XlsEncoding, buf: &[u8], len: &usize, s: &mut String) {
    encoding.decode_to(&buf[1..=*len], *len, s, Some(buf[0] & 0x1 != 0));
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

    fn skip(&mut self, mut len: usize) -> Result<(), XlsError> {
        while len > 0 {
            if self.data.is_empty() && !self.continue_record() {
                return Err(XlsError::ContinueRecordTooShort);
            }
            let l = min(len, self.data.len());
            let (_, next) = self.data.split_at(l);
            self.data = next;
            len -= l;
        }
        Ok(())
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
            cont,
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
/// CellParsedFormula [MS-XLS 2.5.198.3]
fn parse_formula(
    mut rgce: &[u8],
    sheets: &[String],
    names: &[(String, String)],
    xtis: &[Xti],
    encoding: &XlsEncoding,
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
                let rowu = read_u16(&rgce[2..]);
                let colu = read_u16(&rgce[4..]);
                let sh = xtis
                    .get(ixti as usize)
                    .and_then(|xti| sheets.get(xti.itab_first as usize))
                    .map_or("#REF", |sh| sh);
                stack.push(formula.len());
                formula.push_str(sh);
                formula.push('!');
                let col = colu << 2; // first 14 bits only
                if colu & 2 != 0 {
                    formula.push('$');
                }
                push_column(col as u32, &mut formula);
                if colu & 1 != 0 {
                    formula.push('$');
                }
                write!(&mut formula, "{}", rowu + 1).unwrap();
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
                write!(&mut formula, "${}:$", read_u16(&rgce[2..4]) as u32 + 1).unwrap();
                push_column(read_u16(&rgce[8..10]) as u32, &mut formula);
                write!(&mut formula, "${}", read_u16(&rgce[4..6]) as u32 + 1).unwrap();
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
            0x03..=0x11 => {
                // binary operation
                let e2 = stack.pop().ok_or(XlsError::StackLen)?;
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
                let e2 = formula.split_off(e2);
                write!(&mut formula, "{}{}", op, e2).unwrap();
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
                let cch = rgce[0] as usize;
                read_unicode_string_no_cch(encoding, &rgce[1..], &cch, &mut formula);
                formula.push('\"');
                rgce = &rgce[2 + cch..];
            }
            0x18 => {
                rgce = &rgce[5..];
            }
            0x19 => {
                let etpg = rgce[0];
                rgce = &rgce[1..];
                match etpg {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 => rgce = &rgce[2..],
                    0x04 => {
                        // PtgAttrChoose
                        let n = read_u16(&rgce[..2]) as usize + 1;
                        rgce = &rgce[2 + 2 * n..]; // ignore
                    }
                    0x10 => {
                        rgce = &rgce[2..];
                        let e = *stack.last().ok_or(XlsError::StackLen)?;
                        let e = formula.split_off(e);
                        write!(&mut formula, "SUM({})", e).unwrap();
                    }
                    0x40 | 0x41 => {
                        // PtfAttrSpace
                        let e = *stack.last().ok_or(XlsError::StackLen)?;
                        let space = match rgce[0] {
                            0x00 | 0x02 | 0x04 | 0x06 => ' ',
                            0x01 | 0x03 | 0x05 => '\r',
                            val => {
                                return Err(XlsError::Unrecognized {
                                    typ: "PtgAttrSpaceType",
                                    val,
                                });
                            }
                        };
                        let cch = rgce[1];
                        for _ in 0..cch {
                            formula.insert(e, space);
                        }
                        rgce = &rgce[2..];
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
                        });
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
                write!(&mut formula, "{}", read_u16(rgce)).unwrap();
                rgce = &rgce[2..];
            }
            0x1F => {
                stack.push(formula.len());
                write!(&mut formula, "{}", read_f64(rgce)).unwrap();
                rgce = &rgce[8..];
            }
            0x20 | 0x40 | 0x60 => {
                // PtgArray: ignore
                stack.push(formula.len());
                formula.push_str("{PtgArray}");
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
                        if iftab > crate::utils::FTAB_LEN {
                            return Err(XlsError::IfTab(iftab));
                        }
                        rgce = &rgce[2..];
                        let argc = crate::utils::FTAB_ARGC[iftab] as usize;
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
                    formula.push_str(
                        crate::utils::FTAB
                            .get(iftab)
                            .ok_or(XlsError::IfTab(iftab))?,
                    );
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
                write!(&mut formula, "${}:$", read_u16(&rgce[0..2]) as u32 + 1).unwrap();
                push_column(read_u16(&rgce[6..8]) as u32, &mut formula);
                write!(&mut formula, "${}", read_u16(&rgce[2..4]) as u32 + 1).unwrap();
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
            0x39 | 0x59 => {
                // PfgNameX
                stack.push(formula.len());
                formula.push_str("[PtgNameX]");
                rgce = &rgce[6..];
            }
            _ => {
                return Err(XlsError::Unrecognized {
                    typ: "ptg",
                    val: ptg,
                });
            }
        }
    }
    if stack.len() == 1 {
        Ok(formula)
    } else {
        Err(XlsError::InvalidFormula {
            stack_size: stack.len(),
        })
    }
}

/// FormulaValue [MS-XLS 2.5.133]
fn parse_formula_value(r: &[u8]) -> Result<Option<Data>, XlsError> {
    match *r {
        // String, value should be in next record
        [0x00, .., 0xFF, 0xFF] => Ok(None),
        [0x01, _, b, .., 0xFF, 0xFF] => Ok(Some(Data::Bool(b != 0))),
        [0x02, _, e, .., 0xFF, 0xFF] => parse_err(e).map(Some),
        // ignore, return blank string value
        [0x03, _, .., 0xFF, 0xFF] => Ok(Some(Data::String("".to_string()))),
        [e, .., 0xFF, 0xFF] => Err(XlsError::Unrecognized {
            typ: "error",
            val: e,
        }),
        _ => Ok(Some(Data::Float(read_f64(r)))),
    }
}

/// OfficeArtRecord [MS-ODRAW 1.3.1]
#[cfg(feature = "picture")]
struct ArtRecord<'a> {
    instance: u16,
    typ: u16,
    data: &'a [u8],
}

#[cfg(feature = "picture")]
struct ArtRecordIter<'a> {
    stream: &'a [u8],
}

#[cfg(feature = "picture")]
impl<'a> Iterator for ArtRecordIter<'a> {
    type Item = Result<ArtRecord<'a>, XlsError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.stream.len() < 8 {
            return if self.stream.is_empty() {
                None
            } else {
                Some(Err(XlsError::EoStream("art record header")))
            };
        }
        let ver_ins = read_u16(self.stream);
        let instance = ver_ins >> 4;
        let typ = read_u16(&self.stream[2..]);
        if typ < 0xF000 {
            return Some(Err(XlsError::Art("type range 0xF000 - 0xFFFF")));
        }
        let len = read_usize(&self.stream[4..]);
        if self.stream.len() < len + 8 {
            return Some(Err(XlsError::EoStream("art record length")));
        }
        let (d, next) = self.stream.split_at(len + 8);
        self.stream = next;
        let data = &d[8..];

        Some(Ok(ArtRecord {
            instance,
            typ,
            data,
        }))
    }
}

/// Parsing pictures
#[cfg(feature = "picture")]
fn parse_pictures(stream: &[u8]) -> Result<Vec<(String, Vec<u8>)>, XlsError> {
    let mut pics = Vec::new();
    let records = ArtRecordIter { stream };
    for record in records {
        let r = record?;
        match r.typ {
            // OfficeArtDggContainer [MS-ODRAW 2.2.12]
            // OfficeArtBStoreContainer [MS-ODRAW 2.2.20]
            0xF000 | 0xF001 => pics.extend(parse_pictures(r.data)?),
            // OfficeArtFBSE [MS-ODRAW 2.2.32]
            0xF007 => {
                let skip = 36 + r.data[33] as usize;
                pics.extend(parse_pictures(&r.data[skip..])?);
            }
            // OfficeArtBlip [MS-ODRAW 2.2.23]
            0xF01A | 0xF01B | 0xF01C | 0xF01D | 0xF01E | 0xF01F | 0xF029 | 0xF02A => {
                let ext_skip = match r.typ {
                    // OfficeArtBlipEMF [MS-ODRAW 2.2.24]
                    0xF01A => {
                        let skip = match r.instance {
                            0x3D4 => 50usize,
                            0x3D5 => 66,
                            _ => unreachable!(),
                        };
                        Ok(("emf", skip))
                    }
                    // OfficeArtBlipWMF [MS-ODRAW 2.2.25]
                    0xF01B => {
                        let skip = match r.instance {
                            0x216 => 50usize,
                            0x217 => 66,
                            _ => unreachable!(),
                        };
                        Ok(("wmf", skip))
                    }
                    // OfficeArtBlipPICT [MS-ODRAW 2.2.26]
                    0xF01C => {
                        let skip = match r.instance {
                            0x542 => 50usize,
                            0x543 => 66,
                            _ => unreachable!(),
                        };
                        Ok(("pict", skip))
                    }
                    // OfficeArtBlipJPEG [MS-ODRAW 2.2.27]
                    0xF01D | 0xF02A => {
                        let skip = match r.instance {
                            0x46A | 0x6E2 => 17usize,
                            0x46B | 0x6E3 => 33,
                            _ => unreachable!(),
                        };
                        Ok(("jpg", skip))
                    }
                    // OfficeArtBlipPNG [MS-ODRAW 2.2.28]
                    0xF01E => {
                        let skip = match r.instance {
                            0x6E0 => 17usize,
                            0x6E1 => 33,
                            _ => unreachable!(),
                        };
                        Ok(("png", skip))
                    }
                    // OfficeArtBlipDIB [MS-ODRAW 2.2.29]
                    0xF01F => {
                        let skip = match r.instance {
                            0x7A8 => 17usize,
                            0x7A9 => 33,
                            _ => unreachable!(),
                        };
                        Ok(("dib", skip))
                    }
                    // OfficeArtBlipTIFF [MS-ODRAW 2.2.30]
                    0xF029 => {
                        let skip = match r.instance {
                            0x6E4 => 17usize,
                            0x6E5 => 33,
                            _ => unreachable!(),
                        };
                        Ok(("tiff", skip))
                    }
                    _ => Err(XlsError::Art("picture type not support")),
                };
                let ext_skip = ext_skip?;
                pics.push((ext_skip.0.to_string(), Vec::from(&r.data[ext_skip.1..])));
            }
            _ => {}
        }
    }
    Ok(pics)
}
