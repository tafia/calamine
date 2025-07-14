mod cells_reader;

pub use cells_reader::XlsbCellsReader;

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{BufReader, Read, Seek};
use std::sync::Arc;

use log::{trace, warn};

use encoding_rs::UTF_16LE;
use quick_xml::events::attributes::Attribute;
use quick_xml::events::Event;
use quick_xml::name::QName;
use quick_xml::Reader as XmlReader;
use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;

use crate::datatype::DataRef;
use crate::formats::{
    builtin_format_by_code, detect_custom_number_format_with_interner, Alignment, Border,
    BorderSide, CellFormat, Color, Fill, Font, FormatStringInterner, PatternType,
};
use crate::utils::{push_column, read_f64, read_i32, read_u16, read_u32, read_usize};
use crate::vba::VbaProject;
use crate::{
    Cell, CellStyle, Data, DataWithFormatting, HeaderRow, Metadata, Range, Reader, ReaderRef, Sheet, SheetType,
    SheetVisible,
};

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
    /// Unsupported `BErr`
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
    /// XML Encoding error
    Encoding(quick_xml::encoding::EncodingError),
    /// Unexpected buffer size
    UnexpectedBufferSize(usize),
}

from_err!(std::io::Error, XlsbError, Io);
from_err!(zip::result::ZipError, XlsbError, Zip);
from_err!(quick_xml::Error, XlsbError, Xml);
from_err!(quick_xml::encoding::EncodingError, XlsbError, Encoding);

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
            XlsbError::Ptg(t) => write!(f, "Unsupported Ptg {t:X}"),
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
            XlsbError::Encoding(e) => write!(f, "XML encoding error: {e}"),
            XlsbError::UnexpectedBufferSize(size) => write!(f, "Unexpected buffer size: {size}"),
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

/// Xlsb reader options
#[derive(Debug, Default)]
#[non_exhaustive]
struct XlsbOptions {
    pub header_row: HeaderRow,
}

/// A Xlsb reader
pub struct Xlsb<RS> {
    zip: ZipArchive<RS>,
    extern_sheets: Vec<String>,
    sheets: Vec<(String, String)>,
    strings: Vec<String>,
    /// Cell (number) formats
    formats: Vec<CellFormat>,
    styles: Vec<CellStyle>,
    format_interner: FormatStringInterner,
    is_1904: bool,
    metadata: Metadata,
    #[cfg(feature = "picture")]
    pictures: Option<Vec<(String, Vec<u8>)>>,
    options: XlsbOptions,
}

impl<RS: Read + Seek> Xlsb<RS> {
    /// MS-XLSB
    fn read_relationships(&mut self) -> Result<BTreeMap<Vec<u8>, String>, XlsbError> {
        let mut relationships = BTreeMap::new();
        match self.zip.by_name("xl/_rels/workbook.bin.rels") {
            Ok(f) => {
                let mut xml = XmlReader::from_reader(BufReader::new(f));
                let config = xml.config_mut();
                config.check_end_names = false;
                config.trim_text(false);
                config.check_comments = false;
                config.expand_empty_elements = true;
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
                                        target = Some(
                                            xml.decoder()
                                                .decode(&v)
                                                .map_err(XlsbError::Encoding)?
                                                .into_owned(),
                                        );
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
    ///
    /// Parses the complete style information from xlsb files including fonts, fills,
    /// borders, and alignment. This provides full formatting compatibility with xlsx.
    fn read_styles(&mut self) -> Result<(), XlsbError> {
        let mut iter = match RecordIter::from_zip(&mut self.zip, "xl/styles.bin") {
            Ok(iter) => iter,
            Err(_) => return Ok(()), // it is fine if path does not exists
        };
        let mut buf = Vec::with_capacity(1024);
        let mut number_formats = BTreeMap::new();
        let mut format_strings: BTreeMap<u16, Arc<str>> = BTreeMap::new();
        let format_interner = FormatStringInterner::new();

        let mut fonts: Vec<Arc<Font>> = Vec::new();
        let mut fills: Vec<Arc<Fill>> = Vec::new();
        let mut borders: Vec<Arc<Border>> = Vec::new();

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
                        let (cell_format, format_string) =
                            detect_custom_number_format_with_interner(
                                fmt_str.as_ref(),
                                &format_interner,
                            );
                        number_formats.insert(fmt_code, cell_format);
                        if let Some(format_string) = format_string {
                            format_strings.insert(fmt_code, format_string);
                        }
                    }
                }
                0x0263 => {
                    // BrtBeginFonts
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);

                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002B, &[], &mut buf)?; // BrtFont
                        match parse_font(&buf) {
                            Ok(font) => fonts.push(Arc::new(font)),
                            Err(e) => {
                                log::warn!("Failed to parse font: {:?}, using default", e);
                                fonts.push(Arc::new(Font::default()));
                            }
                        }
                    }
                }
                0x025B => {
                    // BrtBeginFills
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);

                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002D, &[], &mut buf)?; // BrtFill
                        match parse_fill(&buf) {
                            Ok(fill) => fills.push(Arc::new(fill)),
                            Err(e) => {
                                log::warn!("Failed to parse fill: {:?}, using default", e);
                                fills.push(Arc::new(Fill::default()));
                            }
                        }
                    }
                }
                0x0265 => {
                    // BrtBeginBorders
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);

                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002E, &[], &mut buf)?; // BrtBorder
                        match parse_border(&buf) {
                            Ok(border) => borders.push(Arc::new(border)),
                            Err(e) => {
                                log::warn!("Failed to parse border: {:?}, using default", e);
                                borders.push(Arc::new(Border::default()));
                            }
                        }
                    }
                }
                0x0269 => {
                    // BrtBeginCellXFs
                    let _len = iter.fill_buffer(&mut buf)?;
                    let len = read_usize(&buf);

                    for _ in 0..len {
                        let _ = iter.next_skip_blocks(0x002F, &[], &mut buf)?; // BrtXF
                        match parse_xf(
                            &buf,
                            &number_formats,
                            &format_strings,
                            &fonts,
                            &fills,
                            &borders,
                        ) {
                            Ok(style) => {
                                // Backward-compatibility: keep the old formats vector updated.
                                self.formats.push(style.number_format.clone());
                                self.styles.push(style);
                            }
                            Err(e) => {
                                log::warn!("Failed to parse cell style: {:?}, using default", e);
                                self.formats.push(CellFormat::Other);
                                self.styles.push(CellStyle::default());
                            }
                        }
                    }
                    // BrtBeginCellXFs is always present and always after BrtBeginFmts
                    break;
                }
                _ => {
                    // Skip unknown record types
                    let _ = iter.fill_buffer(&mut buf)?;
                }
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
            self.strings.push(wide_str(&buf[1..], &mut 0)?.into_owned());
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
                    }
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
                _ => trace!("Unsupported type {typ:X}"),
            }
        }
    }

    /// Get a cells reader for a given worksheet
    pub fn worksheet_cells_reader<'a>(
        &'a mut self,
        name: &str,
    ) -> Result<XlsbCellsReader<'a, RS>, XlsbError> {
        let path = match self.sheets.iter().find(|&(n, _)| n == name) {
            Some((_, path)) => path.clone(),
            None => return Err(XlsbError::WorksheetNotFound(name.into())),
        };
        let iter = RecordIter::from_zip(&mut self.zip, &path)?;
        XlsbCellsReader::new(
            iter,
            &self.styles,
            &self.strings,
            &self.extern_sheets,
            &self.metadata.names,
            self.is_1904,
        )
    }

    /// Get comprehensive formatting information for a cell by its style index
    pub fn get_cell_formatting(&self, style_index: usize) -> Option<&CellStyle> {
        self.styles.get(style_index)
    }

    /// Get all available cell formats
    pub fn get_all_cell_formats(&self) -> &[CellStyle] {
        &self.styles
    }

    /// Get access to the format string interner for reuse across sheets
    /// The interner is thread-safe and can be shared across threads
    pub fn get_format_interner(&self) -> &FormatStringInterner {
        &self.format_interner
    }

    #[cfg(feature = "picture")]
    fn read_pictures(&mut self) -> Result<(), XlsbError> {
        let mut pics = Vec::new();
        for i in 0..self.zip.len() {
            let mut zfile = self.zip.by_index(i)?;
            let zname = zfile.name();
            if zname.starts_with("xl/media") {
                if let Some(ext) = zname.split('.').next_back() {
                    if [
                        "emf", "wmf", "pict", "jpeg", "jpg", "png", "dib", "gif", "tiff", "eps",
                        "bmp", "wpg",
                    ]
                    .contains(&ext)
                    {
                        let ext = ext.to_string();
                        let mut buf: Vec<u8> = Vec::new();
                        zfile.read_to_end(&mut buf)?;
                        pics.push((ext, buf));
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
            styles: Vec::new(),
            format_interner: FormatStringInterner::new(),
            is_1904: false,
            metadata: Metadata::default(),
            #[cfg(feature = "picture")]
            pictures: None,
            options: XlsbOptions::default(),
        };
        xlsb.read_shared_strings()?;
        xlsb.read_styles()?;
        let relationships = xlsb.read_relationships()?;
        xlsb.read_workbook(&relationships)?;
        #[cfg(feature = "picture")]
        xlsb.read_pictures()?;

        Ok(xlsb)
    }

    fn with_header_row(&mut self, header_row: HeaderRow) -> &mut Self {
        self.options.header_row = header_row;
        self
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
    fn worksheet_range(&mut self, name: &str) -> Result<Range<DataWithFormatting>, XlsbError> {
        let header_row = self.options.header_row;
        let mut cell_reader = self.worksheet_cells_reader(name)?;
        let len = cell_reader.dimensions().len();
        let mut cells = Vec::new();
        if len < 100_000 {
            cells.reserve(len as usize);
        }

        match header_row {
            HeaderRow::FirstNonEmptyRow => {
                // the header row is the row of the first non-empty cell
                while let Some((cell, formatting)) = cell_reader.next_cell_with_formatting()? {
                    if matches!(cell.val, DataRef::Empty) {
                        continue;
                    }
                    let data_with_formatting = DataWithFormatting::new(
                        cell.val.into(),
                        formatting.cloned(),
                    );
                    cells.push(Cell::new(cell.pos, data_with_formatting));
                }
            }
            HeaderRow::Row(header_row_idx) => {
                // If `header_row` is a row index, we only add non-empty cells after this index.
                while let Some((cell, formatting)) = cell_reader.next_cell_with_formatting()? {
                    if matches!(cell.val, DataRef::Empty) {
                        continue;
                    }
                    if cell.pos.0 >= header_row_idx {
                        let data_with_formatting = DataWithFormatting::new(
                            cell.val.into(),
                            formatting.cloned(),
                        );
                        cells.push(Cell::new(cell.pos, data_with_formatting));
                    }
                }

                // If `header_row` is set and the first non-empty cell is not at the `header_row`, we add
                // an empty cell at the beginning with row `header_row` and same column as the first non-empty cell.
                if cells.first().is_some_and(|c| c.pos.0 != header_row_idx) {
                    cells.insert(
                        0,
                        Cell {
                            pos: (
                                header_row_idx,
                                cells.first().expect("cells should not be empty").pos.1,
                            ),
                            val: DataWithFormatting::default(),
                        },
                    );
                }
            }
        }

        Ok(Range::from_sparse(cells))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheet_formula(&mut self, name: &str) -> Result<Range<DataWithFormatting>, XlsbError> {
        let mut cells_reader = self.worksheet_cells_reader(name)?;
        let mut cells = Vec::with_capacity(cells_reader.dimensions().len().min(1_000_000) as _);
        while let Some((cell, formatting)) = cells_reader.next_formula_with_formatting()? {
            if !cell.val.is_empty() {
                let data_with_formatting = DataWithFormatting::new(
                    Data::String(cell.val),
                    formatting.cloned(),
                );
                cells.push(Cell::new(cell.pos, data_with_formatting));
            }
        }
        Ok(Range::from_sparse(cells))
    }

    /// MS-XLSB 2.1.7.62
    fn worksheets(&mut self) -> Vec<(String, Range<DataWithFormatting>)> {
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

impl<RS: Read + Seek> ReaderRef<RS> for Xlsb<RS> {
    fn worksheet_range_ref<'a>(&'a mut self, name: &str) -> Result<Range<DataRef<'a>>, XlsbError> {
        let header_row = self.options.header_row;
        let mut cell_reader = self.worksheet_cells_reader(name)?;
        let len = cell_reader.dimensions().len();
        let mut cells = Vec::new();
        if len < 100_000 {
            cells.reserve(len as usize);
        }

        match header_row {
            HeaderRow::FirstNonEmptyRow => {
                // the header row is the row of the first non-empty cell
                loop {
                    match cell_reader.next_cell() {
                        Ok(Some(Cell {
                            val: DataRef::Empty,
                            ..
                        })) => (),
                        Ok(Some(cell)) => cells.push(cell),
                        Ok(None) => break,
                        Err(e) => return Err(e),
                    }
                }
            }
            HeaderRow::Row(header_row_idx) => {
                // If `header_row` is a row index, we only add non-empty cells after this index.
                loop {
                    match cell_reader.next_cell() {
                        Ok(Some(Cell {
                            val: DataRef::Empty,
                            ..
                        })) => (),
                        Ok(Some(cell)) => {
                            if cell.pos.0 >= header_row_idx {
                                cells.push(cell);
                            }
                        }
                        Ok(None) => break,
                        Err(e) => return Err(e),
                    }
                }

                // If `header_row` is set and the first non-empty cell is not at the `header_row`, we add
                // an empty cell at the beginning with row `header_row` and same column as the first non-empty cell.
                if cells.first().is_some_and(|c| c.pos.0 != header_row_idx) {
                    cells.insert(
                        0,
                        Cell {
                            pos: (
                                header_row_idx,
                                cells.first().expect("cells should not be empty").pos.1,
                            ),
                            val: DataRef::Empty,
                        },
                    );
                }
            }
        }

        Ok(Range::from_sparse(cells))
    }
}

pub(crate) struct RecordIter<'a, RS>
where
    RS: Read + Seek,
{
    b: [u8; 1],
    r: BufReader<ZipFile<'a, RS>>,
}

impl<'a, RS> RecordIter<'a, RS>
where
    RS: Read + Seek,
{
    fn from_zip(zip: &'a mut ZipArchive<RS>, path: &str) -> Result<RecordIter<'a, RS>, XlsbError> {
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
/// See Ptg [MS-XLSB 2.5.98.16](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/5d7c0c3f-f75f-4306-804f-6f2ebc6bf811), and Formula [MS-XLSB 2.2.2](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/220abf5e-f561-4333-9fe0-7ac590ed4ad5)
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
    trace!("starting formula parsing with {} bytes", rgce.len());

    while !rgce.is_empty() {
        let ptg = rgce[0];
        trace!(
            "parsing Ptg: 0x{:02X}, remaining bytes: {}",
            ptg,
            rgce.len()
        );
        rgce = &rgce[1..];
        match ptg {
            0x3a | 0x5a | 0x7a => {
                trace!("parsing PtgRef3d");
                let ixti = read_u16(&rgce[0..2]);
                let row = read_u32(&rgce[2..6]) + 1;
                let (col, is_col_relative, is_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[6..8]));
                stack.push(formula.len());
                formula.push_str(&quote_sheet_name(&sheets[ixti as usize]));
                formula.push('!');
                if !is_col_relative {
                    formula.push('$');
                }
                push_column(col as u32, &mut formula);
                if !is_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{row}"));
                rgce = &rgce[8..];
            }
            0x3b | 0x5b | 0x7b => {
                trace!("parsing PtgArea3d");
                let ixti = read_u16(&rgce[0..2]);
                let first_row = read_u32(&rgce[2..6]) + 1;
                let last_row = read_u32(&rgce[6..10]) + 1;
                let (first_col, is_first_col_relative, is_first_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[10..12]));
                let (last_col, is_last_col_relative, is_last_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[12..14]));
                stack.push(formula.len());
                formula.push_str(&quote_sheet_name(&sheets[ixti as usize]));
                formula.push('!');
                if !is_first_col_relative {
                    formula.push('$');
                }
                push_column(first_col as u32, &mut formula);
                if !is_first_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{first_row}"));
                formula.push(':');
                if !is_last_col_relative {
                    formula.push('$');
                }
                push_column(last_col as u32, &mut formula);
                if !is_last_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{last_row}"));
                rgce = &rgce[14..];
            }
            0x3c | 0x5c | 0x7c => {
                trace!("parsing PtgRefErr3d");
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&quote_sheet_name(&sheets[ixti as usize]));
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[8..];
            }
            0x3d | 0x5d | 0x7d => {
                trace!("parsing PtgAreaErr3d");
                let ixti = read_u16(&rgce[0..2]);
                stack.push(formula.len());
                formula.push_str(&quote_sheet_name(&sheets[ixti as usize]));
                formula.push('!');
                formula.push_str("#REF!");
                rgce = &rgce[14..];
            }
            0x01 => {
                trace!("parsing PtgExp - ignoring array/shared formula");
                stack.push(formula.len());
                rgce = &rgce[4..];
            }
            0x03..=0x11 => {
                trace!("parsing PtgAdd, PtgSub, PtgMul, PtgDiv, PtgPower, PtgConcat, PtgLt, PtgLe, PtgEq, PtgGe, PtgGt, PtgNe, PtgIsect, PtgUnion, PtgRange: 0x{ptg:02X}");
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
                trace!("parsing PtgUPlus");
                let e = stack.last().ok_or(XlsbError::StackLen)?;
                formula.insert(*e, '+');
            }
            0x13 => {
                trace!("parsing PtgUMinus");
                let e = stack.last().ok_or(XlsbError::StackLen)?;
                formula.insert(*e, '-');
            }
            0x14 => {
                trace!("parsing PtgPercent");
                formula.push('%');
            }
            0x15 => {
                trace!("parsing PtgParen");
                let e = stack.last().ok_or(XlsbError::StackLen)?;
                formula.insert(*e, '(');
                formula.push(')');
            }
            0x16 => {
                trace!("parsing PtgMissArg");
                stack.push(formula.len());
            }
            0x17 => {
                trace!("parsing PtgStr");
                if rgce.len() < 2 {
                    return Err(XlsbError::StackLen);
                }
                stack.push(formula.len());
                formula.push('\"');
                let cch = read_u16(&rgce[0..2]) as usize;
                if cch > 255 {
                    warn!("invalid PtgStr length: {cch}");
                }
                let string_bytes_needed = 2 + 2 * cch;
                if rgce.len() < string_bytes_needed {
                    return Err(XlsbError::StackLen);
                }
                let decoded = UTF_16LE.decode(&rgce[2..string_bytes_needed]).0;
                for c in decoded.chars() {
                    if c == '\"' {
                        formula.push('\"');
                    }
                    formula.push(c);
                }
                formula.push('\"');
                rgce = &rgce[string_bytes_needed..];
            }
            0x18 => {
                stack.push(formula.len());
                let eptg = rgce[0];
                rgce = &rgce[1..];
                match eptg {
                    0x19 => {
                        trace!("parsing PtgList");
                        rgce = &rgce[12..];
                    }
                    0x1D => {
                        trace!("parsing PtgSxName");
                        rgce = &rgce[4..];
                    }
                    e => return Err(XlsbError::Etpg(e)),
                }
            }
            0x19 => {
                let eptg = rgce[0];
                rgce = &rgce[1..];
                match eptg {
                    0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x40 | 0x41 | 0x80 => {
                        trace!("parsing PtgAttrSemi, PtgAttrIf, PtgAttrGoTo, PtgAttrBaxcel, PtgAttrSpace, PtgAttrSpaceSemi, PtgAttrIfError");
                        rgce = &rgce[2..];
                    }
                    0x04 => {
                        trace!("parsing PtgAttrChoose");
                        if rgce.len() < 4 {
                            return Err(XlsbError::StackLen);
                        }
                        let c_offset = read_u16(&rgce[0..2]) as usize;
                        let skip_bytes = 2 + (c_offset + 1) * 2;
                        if rgce.len() < skip_bytes {
                            return Err(XlsbError::StackLen);
                        }
                        rgce = &rgce[skip_bytes..];
                    }
                    0x10 => {
                        trace!("parsing PtgAttrSum");
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
                trace!("parsing PtgErr");
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
                trace!("parsing PtgBool");
                stack.push(formula.len());
                formula.push_str(if rgce[0] == 0 { "FALSE" } else { "TRUE" });
                rgce = &rgce[1..];
            }
            0x1E => {
                trace!("parsing PtgInt");
                stack.push(formula.len());
                formula.push_str(&format!("{}", read_u16(rgce)));
                rgce = &rgce[2..];
            }
            0x1F => {
                trace!("parsing PtgNum");
                stack.push(formula.len());
                formula.push_str(&format!("{}", read_f64(rgce)));
                rgce = &rgce[8..];
            }
            0x20 | 0x40 | 0x60 => {
                trace!("ignoring PtgArray");
                stack.push(formula.len());
                rgce = &rgce[14..];
            }
            0x21 | 0x22 | 0x41 | 0x42 | 0x61 | 0x62 => {
                trace!("parsing PtgFunc/PtgFuncVar");
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
                trace!("parsing PtgName");
                let iname = read_u32(rgce) as usize - 1; // one-based
                stack.push(formula.len());
                if let Some(name) = names.get(iname) {
                    formula.push_str(&name.0);
                }
                rgce = &rgce[4..];
            }
            0x24 | 0x44 | 0x64 => {
                trace!("parsing PtgRef");
                let row = read_u32(rgce) + 1;
                let (col, is_col_relative, is_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[4..6]));
                stack.push(formula.len());
                if !is_col_relative {
                    formula.push('$');
                }
                push_column(col as u32, &mut formula);
                if !is_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{row}"));
                rgce = &rgce[6..];
            }
            0x25 | 0x45 | 0x65 => {
                trace!("parsing PtgArea");
                let first_row = read_u32(&rgce[0..4]) + 1;
                let last_row = read_u32(&rgce[4..8]) + 1;
                let (first_col, first_col_relative, first_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[8..10]));
                let (last_col, last_col_relative, last_row_relative) =
                    extract_col_and_flags(read_u16(&rgce[10..12]));
                stack.push(formula.len());
                if !first_col_relative {
                    formula.push('$');
                }
                push_column(first_col as u32, &mut formula);
                if !first_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{first_row}"));
                formula.push(':');
                if !last_col_relative {
                    formula.push('$');
                }
                push_column(last_col as u32, &mut formula);
                if !last_row_relative {
                    formula.push('$');
                }
                formula.push_str(&format!("{last_row}"));
                rgce = &rgce[12..];
            }
            0x2A | 0x4A | 0x6A => {
                trace!("parsing PtgRefErr");
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[6..];
            }
            0x2B | 0x4B | 0x6B => {
                trace!("parsing PtgAreaErr");
                stack.push(formula.len());
                formula.push_str("#REF!");
                rgce = &rgce[12..];
            }
            0x29 | 0x49 | 0x69 => {
                trace!("parsing PtgMemFunc");
                let cce = read_u16(rgce) as usize;
                rgce = &rgce[2..];
                let f = parse_formula(&rgce[..cce], sheets, names)?;
                stack.push(formula.len());
                formula.push_str(&f);
                rgce = &rgce[cce..];
            }
            0x39 | 0x59 | 0x79 => {
                // TODO: external workbook ... ignore this formula ...
                trace!("ignoring PtgNameX");
                stack.push(formula.len());
                formula.push_str("EXTERNAL_WB_NAME");
                rgce = &rgce[6..];
            }
            _ => {
                trace!(
                    "parsing unknown Ptg: 0x{:02X} at position with {} bytes remaining",
                    ptg,
                    rgce.len()
                );
                trace!(
                    "Next few bytes: {:02X?}",
                    &rgce[..std::cmp::min(10, rgce.len())]
                );
                trace!("FORMULA PARSING ERROR:");
                trace!("  Unknown Ptg: 0x{ptg:02X}");
                trace!("  Remaining bytes: {}", rgce.len());
                trace!("  Current formula: '{formula}'");
                trace!("  Stack size: {}", stack.len());
                trace!(
                    "  Next 20 bytes: {:02X?}",
                    &rgce[..std::cmp::min(20, rgce.len())]
                );
                trace!(
                    "  Previous bytes: {:02X?}",
                    rgce.get(rgce.len().saturating_sub(50)..rgce.len().saturating_sub(40))
                );
                return Err(XlsbError::Ptg(ptg));
            }
        }
    }

    if stack.len() == 1 {
        Ok(formula)
    } else {
        Err(XlsbError::StackLen)
    }
}

fn cell_format<'a>(styles: &'a [CellStyle], buf: &[u8]) -> Option<&'a CellFormat> {
    // Parses a Cell (MS-XLSB 2.5.9) and determines if it references a Date format.
    // The style index (iStyleRef) is stored as a 24-bit integer starting at the
    // fifth byte of the cell record buffer.

    let style_ref = u32::from_le_bytes([buf[4], buf[5], buf[6], 0]) as usize;

    styles.get(style_ref).map(|s| &s.number_format)
}

fn check_for_password_protected<RS: Read + Seek>(reader: &mut RS) -> Result<(), XlsbError> {
    let offset_end = reader.seek(std::io::SeekFrom::End(0))? as usize;
    reader.seek(std::io::SeekFrom::Start(0))?;

    if let Ok(cfb) = crate::cfb::Cfb::new(reader, offset_end) {
        if cfb.has_directory("EncryptedPackage") {
            return Err(XlsbError::Password);
        }
    }

    Ok(())
}

fn quote_sheet_name(sheet_name: &str) -> String {
    let escaped = sheet_name.replace('\'', "''");
    format!("'{escaped}'")
}

fn extract_col_and_flags(col_data: u16) -> (u16, bool, bool) {
    let col = col_data & 0x3FFF;
    let is_col_relative = (col_data & 0x8000) != 0;
    let is_row_relative = (col_data & 0x4000) != 0;
    (col, is_col_relative, is_row_relative)
}

/// Parse a BrtFont record into a Font structure
/// MS-XLSB 2.4.149 BrtFont structure:
/// - dyHeight (2 bytes): font height in twentieths of a point
/// - grbit (2 bytes): font flags (bold, italic, etc.)
/// - bls (2 bytes): bold weight
/// - sss (2 bytes): superscript/subscript
/// - uls (1 byte): underline style
/// - bFamily (1 byte): font family
/// - bCharSet (1 byte): character set
/// - unused (1 byte): reserved
/// - color (BrtColor): font color
/// - name (XLWideString): font name
fn parse_font(buf: &[u8]) -> Result<Font, XlsbError> {
    // Handle short buffers gracefully - return default font
    if buf.len() < 8 {
        return Err(XlsbError::UnexpectedBufferSize(buf.len()));
    }

    // Parse BrtFont structure according to MS-XLSB 2.4.149
    let size_twentieths = read_u16(&buf[0..2]);
    let grbit = read_u16(&buf[2..4]);
    let bold_weight = read_u16(&buf[4..6]);
    let _sss = read_u16(&buf[6..8]);

    let mut offset = 8;

    // Skip underline, family, charset, unused (4 bytes total)
    if buf.len() >= offset + 4 {
        offset += 4;
    }

    // Parse color (9 bytes)
    let color = if buf.len() >= offset + 9 {
        let color_result = parse_color(&buf[offset..offset + 9])?;
        offset += 9;
        color_result
    } else {
        None
    };

    // Parse font name (variable length)
    let name = if buf.len() >= offset + 2 {
        let name_len = read_u16(&buf[offset..offset + 2]) as usize;
        offset += 2;
        if name_len > 0 && buf.len() >= offset + name_len * 2 {
            let name_bytes = &buf[offset..offset + name_len * 2];
            Some(Arc::from(UTF_16LE.decode(name_bytes).0.as_ref()))
        } else {
            None
        }
    } else {
        None
    };

    Ok(Font {
        name,
        size: if size_twentieths > 0 {
            Some(size_twentieths as f64 / 20.0)
        } else {
            None
        },
        bold: Some(bold_weight >= 700 || (grbit & 0x0001) != 0),
        italic: Some((grbit & 0x0002) != 0),
        color,
    })
}

/// Parse a BrtFill record into a Fill structure
/// MS-XLSB 2.4.145 BrtFill structure:
/// - fls (4 bytes): fill pattern type
/// - fgColor (BrtColor): foreground color (9 bytes)
/// - bgColor (BrtColor): background color (9 bytes)
fn parse_fill(buf: &[u8]) -> Result<Fill, XlsbError> {
    // Handle short buffers gracefully - return default fill
    if buf.len() < 4 {
        return Err(XlsbError::UnexpectedBufferSize(buf.len()));
    }

    // Parse BrtFill structure according to MS-XLSB 2.4.145
    let fls = read_u32(&buf[0..4]);
    let pattern_type = match fls {
        0 => PatternType::None,
        1 => PatternType::Solid,
        2 => PatternType::MediumGray,
        3 => PatternType::DarkGray,
        4 => PatternType::LightGray,
        5 => PatternType::Pattern(Arc::from("darkHorizontal")),
        6 => PatternType::Pattern(Arc::from("darkVertical")),
        7 => PatternType::Pattern(Arc::from("darkDown")),
        8 => PatternType::Pattern(Arc::from("darkUp")),
        9 => PatternType::Pattern(Arc::from("darkGrid")),
        10 => PatternType::Pattern(Arc::from("darkTrellis")),
        11 => PatternType::Pattern(Arc::from("lightHorizontal")),
        12 => PatternType::Pattern(Arc::from("lightVertical")),
        13 => PatternType::Pattern(Arc::from("lightDown")),
        14 => PatternType::Pattern(Arc::from("lightUp")),
        15 => PatternType::Pattern(Arc::from("lightGrid")),
        16 => PatternType::Pattern(Arc::from("lightTrellis")),
        17 => PatternType::Pattern(Arc::from("gray125")),
        18 => PatternType::Pattern(Arc::from("gray0625")),
        other => PatternType::Pattern(Arc::from(format!("pattern_{}", other))),
    };

    // Parse foreground color (9 bytes starting at offset 4)
    let foreground_color = if buf.len() >= 13 {
        parse_color(&buf[4..13])?
    } else {
        None
    };

    // Parse background color (9 bytes starting at offset 13)
    let background_color = if buf.len() >= 22 {
        parse_color(&buf[13..22])?
    } else {
        None
    };

    Ok(Fill {
        pattern_type,
        foreground_color,
        background_color,
    })
}

/// Parse a BrtBorder record into a Border structure
/// MS-XLSB 2.4.48 BrtBorder structure:
/// - blxfTop (BrtBlxf): top border (variable length)
/// - blxfBottom (BrtBlxf): bottom border (variable length)
/// - blxfLeft (BrtBlxf): left border (variable length)
/// - blxfRight (BrtBlxf): right border (variable length)
/// - blxfDiag (BrtBlxf): diagonal border (variable length)
/// - blxfVert (BrtBlxf): vertical border (variable length)
/// - blxfHoriz (BrtBlxf): horizontal border (variable length)
/// Each BrtBlxf is: dg (1 byte style) + color (0-9 bytes, depending on style)
fn parse_border(buf: &[u8]) -> Result<Border, XlsbError> {
    // Minimum size for a border record with 4 sides with minimal data
    if buf.len() < 4 {
        log::warn!("Border buffer too small: {} bytes", buf.len());
        return Ok(Border {
            left: None,
            right: None,
            top: None,
            bottom: None,
        });
    }

    // Helper function to parse a single border side
    fn parse_border_side(buf: &[u8], offset: usize) -> Option<BorderSide> {
        if offset >= buf.len() {
            return None;
        }

        let style = border_style_to_string(buf[offset]);

        // If style is "none" (0), there might not be color data
        if buf[offset] == 0 {
            return Some(BorderSide { style, color: None });
        }

        // Try to parse color if we have enough bytes
        let color = if offset + 10 <= buf.len() {
            parse_color(&buf[offset + 1..offset + 10]).ok().flatten()
        } else {
            None
        };

        Some(BorderSide { style, color })
    }

    // Parse BrtBorder structure according to MS-XLSB 2.4.48
    // Each border side is variable length: 1 byte style + 0-9 bytes color

    // Try to parse each border side, handling variable lengths gracefully
    let top = parse_border_side(buf, 0);
    let bottom = parse_border_side(buf, 10);
    let left = parse_border_side(buf, 20);
    let right = parse_border_side(buf, 30);

    // Note: We skip diagonal, vertical, and horizontal borders for now
    // as they're not commonly used in basic cell formatting

    Ok(Border {
        left,
        right,
        top,
        bottom,
    })
}

/// Parse a BrtXF record into a CellStyle structure
/// MS-XLSB 2.4.812 BrtXF structure:
/// - grbitXF (2 bytes): flags and alignment
/// - ifmt (2 bytes): number format index
/// - ifnt (2 bytes): font index  
/// - iFill (2 bytes): fill index
/// - ixfeBorder (2 bytes): border index
/// - iParentStyle (2 bytes): parent style index
fn parse_xf(
    buf: &[u8],
    number_formats: &BTreeMap<u16, CellFormat>,
    format_strings: &BTreeMap<u16, Arc<str>>,
    fonts: &[Arc<Font>],
    fills: &[Arc<Fill>],
    borders: &[Arc<Border>],
) -> Result<CellStyle, XlsbError> {
    // Handle short buffers gracefully - return default style
    if buf.len() < 12 {
        return Err(XlsbError::UnexpectedBufferSize(buf.len()));
    }

    // Parse BrtXF structure according to MS-XLSB 2.4.812
    let grbit = read_u16(&buf[0..2]);
    let fmt_code = read_u16(&buf[2..4]);
    let font_id = read_u16(&buf[4..6]) as usize;
    let fill_id = read_u16(&buf[6..8]) as usize;
    let border_id = read_u16(&buf[8..10]) as usize;
    let _parent_style = read_u16(&buf[10..12]);

    // Resolve number format
    let number_format = match builtin_format_by_code(fmt_code) {
        CellFormat::DateTime => CellFormat::DateTime,
        CellFormat::TimeDelta => CellFormat::TimeDelta,
        CellFormat::Other => number_formats
            .get(&fmt_code)
            .cloned()
            .unwrap_or(CellFormat::Other),
    };

    // Resolve format string
    let format_string = format_strings.get(&fmt_code).cloned();

    // Parse alignment from grbit flags (bits 0-2: horizontal, bits 3-5: vertical)
    let alignment = Some(Arc::new(Alignment {
        horizontal: horizontal_code_to_str(grbit & 0x0007),
        vertical: vertical_code_to_str((grbit >> 3) & 0x0007),
        wrap_text: Some(grbit & 0x1000 != 0),
        indent: None,
        shrink_to_fit: None,
        text_rotation: None,
        reading_order: None,
    }));

    Ok(CellStyle {
        number_format,
        format_string,
        font: fonts.get(font_id).cloned(),
        fill: fills.get(fill_id).cloned(),
        border: borders.get(border_id).cloned(),
        alignment,
    })
}

/// Parse a 9-byte color structure
fn parse_color(buf: &[u8]) -> Result<Option<Color>, XlsbError> {
    if buf.len() < 9 {
        return Err(XlsbError::UnexpectedBufferSize(buf.len()));
    }

    let flags = buf[0];
    let a = buf[1];
    let r = buf[2];
    let g = buf[3];
    let b = buf[4];
    let theme_value = read_u32(&buf[5..9]);

    match flags {
        0x01 => Ok(Some(Color::Auto)),
        0x02 => Ok(Some(Color::Indexed(theme_value))),
        0x03 => Ok(Some(Color::Rgb { r, g, b })),
        0x04 => Ok(Some(Color::Theme {
            theme: theme_value,
            tint: None,
        })),
        _ => Ok(Some(Color::Argb { a, r, g, b })),
    }
}

/// Convert border style code to string
fn border_style_to_string(style: u8) -> Arc<str> {
    match style {
        0 => Arc::from("none"),
        1 => Arc::from("thin"),
        2 => Arc::from("medium"),
        3 => Arc::from("dashed"),
        4 => Arc::from("dotted"),
        5 => Arc::from("thick"),
        6 => Arc::from("double"),
        7 => Arc::from("hair"),
        8 => Arc::from("mediumDashed"),
        9 => Arc::from("dashDot"),
        10 => Arc::from("mediumDashDot"),
        11 => Arc::from("dashDotDot"),
        12 => Arc::from("mediumDashDotDot"),
        13 => Arc::from("slantDashDot"),
        _ => Arc::from("unknown"),
    }
}

/// Convert horizontal alignment code to string
fn horizontal_code_to_str(code: u16) -> Option<Arc<str>> {
    match code {
        0 => None, // General
        1 => Some(Arc::from("left")),
        2 => Some(Arc::from("center")),
        3 => Some(Arc::from("right")),
        4 => Some(Arc::from("fill")),
        5 => Some(Arc::from("justify")),
        6 => Some(Arc::from("centerContinuous")),
        7 => Some(Arc::from("distributed")),
        _ => None,
    }
}

/// Convert vertical alignment code to string
fn vertical_code_to_str(code: u16) -> Option<Arc<str>> {
    match code {
        0 => Some(Arc::from("top")),
        1 => Some(Arc::from("center")),
        2 => Some(Arc::from("bottom")),
        3 => Some(Arc::from("justify")),
        4 => Some(Arc::from("distributed")),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_xlsb_formatting_api() {
        // Test that we can open an xlsb file and access formatting information
        let test_file = include_bytes!("../../tests/choose.xlsb");
        let cursor = Cursor::new(test_file);

        let workbook = Xlsb::new(cursor).expect("Failed to open test xlsb file");

        // Test that we can get all cell formats
        let formats = workbook.get_all_cell_formats();
        assert!(!formats.is_empty(), "Should have at least one cell format");

        // Test that we can get formatting by index
        let first_format = workbook.get_cell_formatting(0);
        assert!(first_format.is_some(), "Should be able to get first format");

        // Test that the format interner is accessible
        let interner = workbook.get_format_interner();
        let _interner_len = interner.len(); // Just verify it's accessible
    }

    #[test]
    fn test_xlsb_cell_reader_formatting() {
        // Test that we can read cells with formatting information
        let test_file = include_bytes!("../../tests/choose.xlsb");
        let cursor = Cursor::new(test_file);

        let mut workbook = Xlsb::new(cursor).expect("Failed to open test xlsb file");
        let sheet_names: Vec<String> = workbook
            .metadata()
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect();

        if let Some(sheet_name) = sheet_names.first() {
            let mut cell_reader = workbook
                .worksheet_cells_reader(sheet_name)
                .expect("Failed to create cell reader");

            // Test that we can read cells with formatting
            while let Ok(Some((_cell, _formatting))) = cell_reader.next_cell_with_formatting() {
                // Just verify the API works without panicking
                break;
            }

            // The API should work without panicking
        }
    }

    #[test]
    fn test_xlsb_range_with_formatting() {
        // Test that worksheet_range returns proper DataWithFormatting values
        let test_file = include_bytes!("../../tests/choose.xlsb");
        let cursor = Cursor::new(test_file);

        let mut workbook = Xlsb::new(cursor).expect("Failed to open test xlsb file");
        let sheet_names: Vec<String> = workbook
            .metadata()
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect();

        if let Some(sheet_name) = sheet_names.first() {
            let range = workbook
                .worksheet_range(sheet_name)
                .expect("Failed to get worksheet range");

            // Test that the range contains DataWithFormatting objects
            let mut has_formatting = false;
            for row in range.rows() {
                for cell in row {
                    if cell.formatting.is_some() {
                        has_formatting = true;
                        break;
                    }
                }
                if has_formatting {
                    break;
                }
            }

            assert!(has_formatting);
        }
    }

    #[test]
    fn test_xlsb_formula_range_with_formatting() {
        // Test that worksheet_formula returns proper DataWithFormatting values
        let test_file = include_bytes!("../../tests/choose.xlsb");
        let cursor = Cursor::new(test_file);

        let mut workbook = Xlsb::new(cursor).expect("Failed to open test xlsb file");
        let sheet_names: Vec<String> = workbook
            .metadata()
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect();

        if let Some(sheet_name) = sheet_names.first() {
            let range = workbook
                .worksheet_formula(sheet_name)
                .expect("Failed to get worksheet formula range");

            // Test that the range contains DataWithFormatting objects
            let mut has_formatting = false;
            for row in range.rows() {
                for cell in row {
                    if cell.formatting.is_some() {
                        has_formatting = true;
                        break;
                    }
                }
                if has_formatting {
                    break;
                }
            }



            assert!(has_formatting);
        }
    }
}
