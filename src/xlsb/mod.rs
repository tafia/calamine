mod cells_reader;

pub use cells_reader::XlsbCellsReader;

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{BufReader, Read, Seek};
use std::string::String;

use log::debug;

use encoding_rs::UTF_16LE;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::datatype::DataRef;
use crate::formats::{builtin_format_by_code, detect_custom_number_format, CellFormat};
use crate::utils::{push_column, read_f64, read_i32, read_u16, read_u32, read_usize};
use crate::vba::VbaProject;
use crate::{Cell, Data, Metadata, Range, Reader, RichText, Sheet, SheetType, SheetVisible};

/// A Xlsb specific error
#[derive(Debug)]
pub enum XlsbError {
    /// Io error
    Io(std::io::Error),
    /// Zip error
    Zip(zip::result::ZipError),
    /// Xml error
    Xml(quick_xml::Error),
    /// Xml attribute error
    XmlAttr(quick_xml::events::attributes::AttrError),
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
    /// Unrecognized data
    Unrecognized {
        /// data type
        typ: &'static str,
        /// value found
        val: String,
    },
    /// Workbook is password protected
    Password,
    /// Worksheet not found
    WorksheetNotFound(String),
}

from_err!(std::io::Error, XlsbError, Io);
from_err!(zip::result::ZipError, XlsbError, Zip);
from_err!(quick_xml::Error, XlsbError, Xml);

impl std::fmt::Display for XlsbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            XlsbError::Io(e) => write!(f, "I/O error: {e}"),
            XlsbError::Zip(e) => write!(f, "Zip error: {e}"),
            XlsbError::Xml(e) => write!(f, "Xml error: {e}"),
            XlsbError::XmlAttr(e) => write!(f, "Xml attribute error: {e}"),
            XlsbError::Vba(e) => write!(f, "Vba error: {e}"),
            XlsbError::Mismatch { expected, found } => {
                write!(f, "Expecting {expected}, got {found:X}")
            }
            XlsbError::FileNotFound(file) => write!(f, "File not found: '{file}'"),
            XlsbError::StackLen => write!(f, "Invalid stack length"),
            XlsbError::UnsupportedType(t) => write!(f, "Unsupported type {t:X}"),
            XlsbError::Etpg(t) => write!(f, "Unsupported etpg {t:X}"),
            XlsbError::IfTab(t) => write!(f, "Unsupported iftab {t:X}"),
            XlsbError::BErr(t) => write!(f, "Unsupported BErr {t:X}"),
            XlsbError::Ptg(t) => write!(f, "Unsupported Ptf {t:X}"),
            XlsbError::CellError(t) => write!(f, "Unsupported Cell Error code {t:X}"),
            XlsbError::WideStr { ws_len, buf_len } => write!(
                f,
                "Wide str length exceeds buffer length ({ws_len} > {buf_len})",
            ),
            XlsbError::Unrecognized { typ, val } => {
                write!(f, "Unrecognized {typ}: {val}")
            }
            XlsbError::Password => write!(f, "Workbook is password protected"),
            XlsbError::WorksheetNotFound(name) => write!(f, "Worksheet '{name}' not found"),
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
pub struct Xlsb<RS> {
    zip: ZipArchive<RS>,
    extern_sheets: Vec<String>,
    sheets: Vec<(String, String)>,
    strings: Vec<RichText>,
    /// Cell (number) formats
    formats: Vec<CellFormat>,
    is_1904: bool,
    metadata: Metadata,
    #[cfg(feature = "picture")]
    pictures: Option<Vec<(String, Vec<u8>)>>,
}

impl<RS: Read + Seek> Xlsb<RS> {
    /// MS-XLSB
    fn read_relationships(&mut self) -> Result<BTreeMap<Vec<u8>, String>, XlsbError> {
        let mut relationships = BTreeMap::new();
        match self.zip.by_name("xl/_rels/workbook.bin.rels") {
            Ok(f) => {
                let mut xml = XmlReader::from_reader(BufReader::new(f));
                xml.check_end_names(false)
                    .trim_text(false)
                    .check_comments(false)
                    .expand_empty_elements(true);
                let mut buf: Vec<u8> = Vec::with_capacity(64);

                loop {
                    match xml.read_event_into(&mut buf) {
                        Ok(Event::Start(ref e)) if e.name() == QName(b"Relationship") => {
                            let mut id = None;
                            let mut target = None;
                            for a in e.attributes() {
                                match a.map_err(XlsbError::XmlAttr)? {
                                    Attribute {
                                        key: QName(b"Id"),
                                        value: v,
                                    } => {
                                        id = Some(v.to_vec());
                                    }
                                    Attribute {
                                        key: QName(b"Target"),
                                        value: v,
                                    } => {
                                        target = Some(xml.decoder().decode(&v)?.into_owned());
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

    /// MS-XLSB 2.1.7.50 Styles
    fn read_styles(&mut self) -> Result<(), XlsbError> {
        let mut iter = match RecordIter::from_zip(&mut self.zip, "xl/styles.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(()), // it is fine if path does not exists
        };
        let mut buf = Vec::with_capacity(1024);
        let mut number_formats = BTreeMap::new();

        loop {
            match iter.read_type()? {
                0x0267 => {
                    // BrtBeginFmts
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);

                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002C, &[], &mut buf)?; // BrtFmt
                        let fmt_code = read_u16(&buf);
                        let fmt_str = wide_str(&buf[2..], &mut 0)?;
                        number_formats
                            .insert(fmt_code, detect_custom_number_format(fmt_str.as_ref()));
                    }
                }
                0x0269 => {
                    // BrtBeginCellXFs
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);
                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002F, &[], &mut buf)?; // BrtXF
                        let fmt_code = read_u16(&buf[2..4]);
                        match builtin_format_by_code(fmt_code) {
                            CellFormat::DateTime => self.formats.push(CellFormat::DateTime),
                            CellFormat::TimeDelta => self.formats.push(CellFormat::TimeDelta),
                            CellFormat::Other => {
                                self.formats.push(
                                    number_formats
                                        .get(&fmt_code)
                                        .copied()
                                        .unwrap_or(CellFormat::Other),
                                );
                            }
                        }
                    }
                    // BrtBeginCellXFs is always present and always after BrtBeginFmts
                    break;
                }
                _ => (),
            }
            buf.clear();
        }

        Ok(())
    }

    /// MS-XLSB 2.1.7.45
    fn read_shared_strings(&mut self) -> Result<(), XlsbError> {
        let mut iter = match RecordIter::from_zip(&mut self.zip, "xl/sharedStrings.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(()), // it is fine if path does not exists
        };
        let mut buf = Vec::with_capacity(1024);

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
            self.strings
                .push(RichText::plain(wide_str(&buf[1..], &mut 0)?.into_owned()));
        }
        Ok(())
    }

    /// MS-XLSB 2.1.7.61
    fn read_workbook(
        &mut self,
        relationships: &BTreeMap<Vec<u8>, String>,
    ) -> Result<(), XlsbError> {
        let mut iter = RecordIter::from_zip(&mut self.zip, "xl/workbook.bin")?;
        let mut buf = Vec::with_capacity(1024);

        loop {
            match iter.read_type()? {
                0x0099 => {
                    let _ = iter.fill_buffer(&mut buf)?;
                    self.is_1904 = &buf[0] & 0x1 != 0;
                } // BrtWbProp
                0x009C => {
                    // BrtBundleSh
                    let len = iter.fill_buffer(&mut buf)?;
                    let rel_len = read_u32(&buf[8..len]);
                    if rel_len != 0xFFFF_FFFF {
                        let rel_len = rel_len as usize * 2;
                        let relid = &buf[12..12 + rel_len];
                        // converts utf16le to utf8 for BTreeMap search
                        let relid = UTF_16LE.decode(relid).0;
                        let path = format!("xl/{}", relationships[relid.as_bytes()]);
                        // ST_SheetState
                        let visible = match read_u32(&buf) {
                            0 => SheetVisible::Visible,
                            1 => SheetVisible::Hidden,
                            2 => SheetVisible::VeryHidden,
                            v => {
                                return Err(XlsbError::Unrecognized {
                                    typ: "BoundSheet8:hsState",
                                    val: v.to_string(),
                                })
                            }
                        };
                        let typ = match path.split('/').nth(1) {
                            Some("worksheets") => SheetType::WorkSheet,
                            Some("chartsheets") => SheetType::ChartSheet,
                            Some("dialogsheets") => SheetType::DialogSheet,
                            _ => {
                                return Err(XlsbError::Unrecognized {
                                    typ: "BoundSheet8:dt",
                                    val: path.to_string(),
                                })
                            }
                        };
                        let name = wide_str(&buf[12 + rel_len..len], &mut 0)?;
                        self.metadata.sheets.push(Sheet {
                            name: name.to_string(),
                            typ,
                            visible,
                        });
                        self.sheets.push((name.into_owned(), path));
                    };
                }
                0x0090 => break, // BrtEndBundleShs
                _ => (),
            }
            buf.clear();
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

    /// Get a cells reader for a given worksheet
    pub fn worksheet_cells_reader<'a>(
        &'a mut self,
        name: &str,
    ) -> Result<XlsbCellsReader<'a>, XlsbError> {
        let path = match self.sheets.iter().find(|&(n, _)| n == name) {
            Some((_, path)) => path.clone(),
            None => return Err(XlsbError::WorksheetNotFound(name.into())),
        };
        let iter = RecordIter::from_zip(&mut self.zip, &path)?;
        XlsbCellsReader::new(
            iter,
            &self.formats,
            &self.strings,
            &self.extern_sheets,
            &self.metadata.names,
            self.is_1904,
        )
    }

    #[cfg(feature = "picture")]
    fn read_pictures(&mut self) -> Result<(), XlsbError> {
        let mut pics = Vec::new();
        for i in 0..self.zip.len() {
            let mut zfile = self.zip.by_index(i)?;
            let zname = zfile.name().to_owned();
            if zname.starts_with("xl/media") {
                let name_ext: Vec<&str> = zname.split(".").collect();
                if let Some(ext) = name_ext.last() {
                    if [
                        "emf", "wmf", "pict", "jpeg", "jpg", "png", "dib", "gif", "tiff", "eps",
                        "bmp", "wpg",
                    ]
                    .contains(ext)
                    {
                        let mut buf: Vec<u8> = Vec::new();
                        zfile.read_to_end(&mut buf)?;
                        pics.push((ext.to_string(), buf));
                    }
                }
            }
        }
        if !pics.is_empty() {
            self.pictures = Some(pics);
        }
        Ok(())
    }
}

impl<RS: Read + Seek> Reader<RS> for Xlsb<RS> {
    type Error = XlsbError;

    fn new(mut reader: RS) -> Result<Self, XlsbError> {
        check_for_password_protected(&mut reader)?;

        let mut xlsb = Xlsb {
            zip: ZipArchive::new(reader)?,
            sheets: Vec::new(),
            strings: Vec::new(),
            extern_sheets: Vec::new(),
            formats: Vec::new(),
            is_1904: false,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
        };
        xlsb.read_shared_strings()?;
        xlsb.read_styles()?;
        let relationships = xlsb.read_relationships()?;
        xlsb.read_workbook(&relationships)?;
        #[cfg(feature = "picture")]
        xlsb.read_pictures()?;

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
    fn worksheet_range(&mut self, name: &str) -> Result<Range<Data>, XlsbError> {
        let mut cells_reader = self.worksheet_cells_reader(name)?;
        let mut cells = Vec::with_capacity(cells_reader.dimensions().len().min(1_000_000) as _);
        while let Some(cell) = cells_reader.next_cell()? {
            if cell.val != DataRef::Empty {
                cells.push(Cell::new(cell.pos, Data::from(cell.val)));
            }
        }
        Ok(Range::from_sparse(cells))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheet_formula(&mut self, name: &str) -> Result<Range<String>, XlsbError> {
        let mut cells_reader = self.worksheet_cells_reader(name)?;
        let mut cells = Vec::with_capacity(cells_reader.dimensions().len().min(1_000_000) as _);
        while let Some(cell) = cells_reader.next_formula()? {
            if !cell.val.is_empty() {
                cells.push(cell);
            }
        }
        Ok(Range::from_sparse(cells))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheets(&mut self) -> Vec<(String, Range<Data>)> {
        let sheets = self
            .sheets
            .iter()
            .map(|(name, _)| name.clone())
            .collect::<Vec<_>>();
        sheets
            .into_iter()
            .filter_map(|name| {
                let ws = self.worksheet_range(&name).ok()?;
                Some((name, ws))
            })
            .collect()
    }

    #[cfg(feature = "picture")]
    fn pictures(&self) -> Option<Vec<(String, Vec<u8>)>> {
        self.pictures.to_owned()
    }
}

pub(crate) struct RecordIter<'a> {
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

fn wide_str<'a>(buf: &'a [u8], str_len: &mut usize) -> Result<Cow<'a, str>, XlsbError> {
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
                formula.push_str(&UTF_16LE.decode(&rgce[2..2 + 2 * cch]).0);
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
                if let Some(name) = names.get(iname) {
                    formula.push_str(&name.0);
                }
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

    if stack.len() == 1 {
        Ok(formula)
    } else {
        Err(XlsbError::StackLen)
    }
}

fn cell_format<'a>(formats: &'a [CellFormat], buf: &[u8]) -> Option<&'a CellFormat> {
    // Parses a Cell (MS-XLSB 2.5.9) and determines if it references a Date format

    // iStyleRef is stored as a 24bit integer starting at the fifth byte
    let style_ref = u32::from_le_bytes([buf[4], buf[5], buf[6], 0]);

    formats.get(style_ref as usize)
}

fn check_for_password_protected<RS: Read + Seek>(reader: &mut RS) -> Result<(), XlsbError> {
    let offset_end = reader.seek(std::io::SeekFrom::End(0))? as usize;
    reader.seek(std::io::SeekFrom::Start(0))?;

    if let Ok(cfb) = crate::cfb::Cfb::new(reader, offset_end) {
        if cfb.has_directory("EncryptedPackage") {
            return Err(XlsbError::Password);
        }
    };

    Ok(())
}
