// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};
use std::{
    borrow::Cow,
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string_with_bufs,
    translate_formula, Dimensions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::unescape_entity_to_buffer,
    Cell, XlsxError,
};

#[derive(Clone, Debug)]
struct SharedFormulaTemplate {
    formula: String,
    anchor: (u32, u32),
}

/// Workbook-level context used when reading cell values.
struct WorkbookContext<'a> {
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
}

/// Reusable scratch buffers for cell value parsing (avoid per-cell allocations).
struct ValueBufs {
    xml: Vec<u8>,
    value: String,
    str_inner: Vec<u8>,
}

impl ValueBufs {
    fn new() -> Self {
        Self {
            xml: Vec::with_capacity(1024),
            value: String::with_capacity(64),
            str_inner: Vec::with_capacity(1024),
        }
    }
}

/// Formula metadata attached to an XLSX cell record.
#[non_exhaustive]
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum XlsxFormulaMetadata {
    /// Ordinary, non-shared formula.
    Normal {
        /// Formula text.
        formula: String,
    },
    /// Shared formula anchor/template cell.
    SharedAnchor {
        /// Shared formula index (`si`). Shared formula indices are worksheet-local.
        shared_index: usize,
        /// Shared formula reference range, when present.
        reference: Option<Dimensions>,
        /// Template formula text as stored on the anchor cell.
        formula: String,
    },
    /// Shared formula derived cell.
    SharedDerived {
        /// Shared formula index (`si`). Shared formula indices are worksheet-local.
        shared_index: usize,
    },
}

impl XlsxFormulaMetadata {
    /// Return the shared formula index when this record belongs to a shared group.
    pub fn shared_index(&self) -> Option<usize> {
        match self {
            Self::Normal { .. } => None,
            Self::SharedAnchor { shared_index, .. } | Self::SharedDerived { shared_index } => {
                Some(*shared_index)
            }
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum FormulaRecord {
    Normal {
        formula: String,
    },
    SharedAnchor {
        shared_index: usize,
        reference: Option<Dimensions>,
        formula: String,
    },
    SharedDerived {
        shared_index: usize,
        translated_formula: Option<String>,
    },
}

impl FormulaRecord {
    fn formula_text(&self) -> Option<&str> {
        match self {
            Self::Normal { formula } | Self::SharedAnchor { formula, .. } => Some(formula),
            Self::SharedDerived {
                translated_formula, ..
            } => translated_formula.as_deref(),
        }
    }

    fn into_metadata(self) -> XlsxFormulaMetadata {
        match self {
            Self::Normal { formula } => XlsxFormulaMetadata::Normal { formula },
            Self::SharedAnchor {
                shared_index,
                reference,
                formula,
            } => XlsxFormulaMetadata::SharedAnchor {
                shared_index,
                reference,
                formula,
            },
            Self::SharedDerived { shared_index, .. } => {
                XlsxFormulaMetadata::SharedDerived { shared_index }
            }
        }
    }
}

/// A single XLSX cell record containing both cached/literal value and expanded formula text.
#[derive(Clone, Debug, PartialEq)]
pub struct XlsxCellFormulaRecord<'a> {
    /// Zero-based `(row, column)` cell position.
    pub pos: (u32, u32),
    /// Literal or cached value associated with the cell.
    pub value: DataRef<'a>,
    /// Formula text, expanded for shared formulas when the shared-formula anchor
    /// has already been observed in stream order.
    pub formula: Option<String>,
}

/// A single XLSX cell record containing cached/literal value and formula metadata.
#[derive(Clone, Debug, PartialEq)]
pub struct XlsxCellFormulaMetadataRecord<'a> {
    /// Zero-based `(row, column)` cell position.
    pub pos: (u32, u32),
    /// Literal or cached value associated with the cell.
    pub value: DataRef<'a>,
    /// Formula metadata, when the cell contains a formula.
    pub formula: Option<XlsxFormulaMetadata>,
}

struct XlsxCellFormulaMetadataRecordInternal<'a> {
    pos: (u32, u32),
    value: DataRef<'a>,
    formula: Option<FormulaRecord>,
}

/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    xml: XlReader<'a, RS>,
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    value_bufs: ValueBufs,
    formulas: Vec<Option<SharedFormulaTemplate>>,
}

impl<'a, RS> XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    /// Create a new XLSX cell reader over a worksheet XML stream.
    pub fn new(
        mut xml: XlReader<'a, RS>,
        strings: &'a [String],
        formats: &'a [CellFormat],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        let mut dimensions = Dimensions::default();
        let mut sh_type = None;
        'xml: loop {
            buf.clear();
            match xml.read_event_into(&mut buf).map_err(XlsxError::Xml)? {
                Event::Start(e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: QName(b"ref"),
                                value: rdim,
                            } = a?
                            {
                                dimensions = get_dimension(&rdim)?;
                                continue 'xml;
                            }
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
                    b"sheetData" => break,
                    typ => {
                        if sh_type.is_none() {
                            sh_type = Some(xml.decoder().decode(typ)?.to_string());
                        }
                    }
                },
                Event::Eof => {
                    if let Some(typ) = sh_type {
                        return Err(XlsxError::NotAWorksheet(typ));
                    } else {
                        return Err(XlsxError::XmlEof("worksheet"));
                    }
                }
                _ => (),
            }
        }
        Ok(Self {
            xml,
            strings,
            formats,
            is_1904,
            dimensions,
            row_index: 0,
            col_index: 0,
            buf: Vec::with_capacity(1024),
            cell_buf: Vec::with_capacity(1024),
            value_bufs: ValueBufs::new(),
            formulas: Vec::with_capacity(1024),
        })
    }

    /// Return the worksheet dimensions declared by the sheet XML.
    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    /// Return the next cell value in XML stream order.
    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    // Extract all needed attributes in one pass (avoids calling
                    // `get_attribute` multiple times as each re-iterates).
                    let mut pos_attr = None;
                    let mut style_attr = None;
                    let mut type_attr = None;
                    for a in c_element.attributes() {
                        let a = a.map_err(XlsxError::XmlAttr)?;
                        let Cow::Borrowed(val) = a.value else {
                            continue;
                        };
                        match a.key {
                            QName(b"r") => pos_attr = Some(val),
                            QName(b"s") => style_attr = Some(val),
                            QName(b"t") => type_attr = Some(val),
                            _ => {}
                        }
                    }
                    let pos = if let Some(range) = pos_attr {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = DataRef::Empty;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                let ctx = WorkbookContext {
                                    strings: self.strings,
                                    formats: self.formats,
                                    is_1904: self.is_1904,
                                };
                                value = read_value(
                                    &ctx,
                                    &mut self.xml,
                                    &e,
                                    style_attr,
                                    type_attr,
                                    &mut self.value_bufs,
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value)));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    fn read_formula_record(
        xml: &mut XlReader<'_, RS>,
        formulas: &mut Vec<Option<SharedFormulaTemplate>>,
        e: &BytesStart<'_>,
        pos: (u32, u32),
        expand_shared_derived: bool,
    ) -> Result<Option<FormulaRecord>, XlsxError> {
        let formula = read_formula(xml, e)?;

        if let Ok(Some(b"shared")) = get_attribute(e.attributes(), QName(b"t")) {
            let shared_index = match get_attribute(e.attributes(), QName(b"si"))? {
                Some(res) => match atoi_simd::parse::<usize>(res) {
                    Ok(res) => res,
                    Err(_) => return Err(XlsxError::Unexpected("si attribute must be a number")),
                },
                None => {
                    return Err(XlsxError::Unexpected(
                        "si attribute is mandatory if it is shared",
                    ));
                }
            };

            return match get_attribute(e.attributes(), QName(b"ref"))? {
                Some(res) => {
                    let reference = get_dimension(res)?;
                    if let Some(f) = formula {
                        if expand_shared_derived {
                            if formulas.len() <= shared_index {
                                formulas.resize(shared_index + 1, None);
                            }
                            formulas[shared_index] = Some(SharedFormulaTemplate {
                                formula: f.clone(),
                                anchor: pos,
                            });
                        }
                        Ok(Some(FormulaRecord::SharedAnchor {
                            shared_index,
                            reference: Some(reference),
                            formula: f,
                        }))
                    } else {
                        Ok(Some(FormulaRecord::SharedDerived {
                            shared_index,
                            translated_formula: None,
                        }))
                    }
                }
                None => {
                    let translated_formula = if expand_shared_derived {
                        formulas
                            .get(shared_index)
                            .and_then(|template| template.as_ref())
                            .map(|template| {
                                translate_formula(&template.formula, template.anchor, pos)
                            })
                            .transpose()?
                    } else {
                        None
                    };
                    Ok(Some(FormulaRecord::SharedDerived {
                        shared_index,
                        translated_formula,
                    }))
                }
            };
        }

        Ok(formula.map(|formula| FormulaRecord::Normal { formula }))
    }

    /// Return the next cell record, exposing cached/literal value plus expanded
    /// per-cell formula text. Shared-formula metadata is intentionally not
    /// exposed through this compatibility-oriented one-pass API; use
    /// [`Self::next_cell_with_formula_metadata`] when shared-formula metadata is needed.
    pub fn next_cell_with_formula(
        &mut self,
    ) -> Result<Option<XlsxCellFormulaRecord<'a>>, XlsxError> {
        Ok(self
            .next_cell_formula_record_impl(true)?
            .map(|record| XlsxCellFormulaRecord {
                pos: record.pos,
                value: record.value,
                formula: record
                    .formula
                    .and_then(|formula| formula.formula_text().map(str::to_string)),
            }))
    }

    /// Return the next cell record, exposing cached/literal value plus formula
    /// metadata. Shared formulas are reported semantically as anchors/derived
    /// placements instead of only as expanded text. Derived shared formulas carry
    /// only their shared index, avoiding per-cell formula expansion/allocation.
    pub fn next_cell_with_formula_metadata(
        &mut self,
    ) -> Result<Option<XlsxCellFormulaMetadataRecord<'a>>, XlsxError> {
        Ok(self
            .next_cell_formula_record_impl(false)?
            .map(|record| XlsxCellFormulaMetadataRecord {
                pos: record.pos,
                value: record.value,
                formula: record.formula.map(FormulaRecord::into_metadata),
            }))
    }

    fn next_cell_formula_record_impl(
        &mut self,
        expand_shared_derived: bool,
    ) -> Result<Option<XlsxCellFormulaMetadataRecordInternal<'a>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let mut pos_attr = None;
                    let mut style_attr = None;
                    let mut type_attr = None;
                    for a in c_element.attributes() {
                        let a = a.map_err(XlsxError::XmlAttr)?;
                        let Cow::Borrowed(val) = a.value else {
                            continue;
                        };
                        match a.key {
                            QName(b"r") => pos_attr = Some(val),
                            QName(b"s") => style_attr = Some(val),
                            QName(b"t") => type_attr = Some(val),
                            _ => {}
                        }
                    }
                    let pos = if let Some(range) = pos_attr {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = DataRef::Empty;
                    let mut formula = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) if e.local_name().as_ref() == b"f" => {
                                formula = Self::read_formula_record(
                                    &mut self.xml,
                                    &mut self.formulas,
                                    &e,
                                    pos,
                                    expand_shared_derived,
                                )?;
                            }
                            Ok(Event::Start(e)) => {
                                let ctx = WorkbookContext {
                                    strings: self.strings,
                                    formats: self.formats,
                                    is_1904: self.is_1904,
                                };
                                value = read_value(
                                    &ctx,
                                    &mut self.xml,
                                    &e,
                                    style_attr,
                                    type_attr,
                                    &mut self.value_bufs,
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(XlsxCellFormulaMetadataRecordInternal {
                        pos,
                        value,
                        formula,
                    }));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Return the next formula in XML stream order, expanding shared formulas.
    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(e)) => {
                                let formula_record = Self::read_formula_record(
                                    &mut self.xml,
                                    &mut self.formulas,
                                    &e,
                                    pos,
                                    true,
                                )?;
                                if let Some(formula_text) = formula_record
                                    .and_then(|record| record.formula_text().map(str::to_string))
                                {
                                    value = Some(formula_text);
                                }
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
                }
                Ok(Event::End(e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }
}

/// Reads a cell value using pre-extracted `s` and `t` attributes
/// (avoids repeating attribute iteration on the `<c>` element).
fn read_value<'s, RS>(
    ctx: &WorkbookContext<'s>,
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    style_attr: Option<&[u8]>,
    type_attr: Option<&[u8]>,
    bufs: &mut ValueBufs,
) -> Result<DataRef<'s>, XlsxError>
where
    RS: Read + Seek,
{
    Ok(match e.local_name().as_ref() {
        b"is" => {
            // inlineStr
            read_string_with_bufs(xml, e.name(), &mut bufs.xml, &mut bufs.str_inner)?
                .map_or(DataRef::Empty, DataRef::String)
        }
        // Ignore <v> for inlineStr cells since it is redundant. The value is in
        // the <is> element, which is handled above.
        b"v" if matches!(type_attr, Some(b"inlineStr") | Some(b"is")) => {
            bufs.xml.clear();
            xml.read_to_end_into(e.name(), &mut bufs.xml)?;
            DataRef::Empty
        }
        b"v" => match type_attr {
            Some(b"n") | Some(b"s") | Some(b"b") | Some(b"e") | None => {
                // These types are always plain ASCII (no CR/LF or entities), so we can
                // parse directly from raw bytes, skipping `xml10_content()` + String
                bufs.xml.clear();
                let val = match xml.read_event_into(&mut bufs.xml)? {
                    Event::Text(t) => read_v(ctx, &t, style_attr, type_attr)?,
                    Event::End(end) if end.name() == e.name() => return Ok(DataRef::Empty),
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => DataRef::Empty,
                };
                bufs.xml.clear();
                xml.read_to_end_into(e.name(), &mut bufs.xml)?;
                val
            }
            _ => {
                // Types that may contain entities, or need owned Strings (eg: "str", "d")
                bufs.value.clear();
                loop {
                    bufs.xml.clear();
                    match xml.read_event_into(&mut bufs.xml)? {
                        Event::Text(t) => bufs.value.push_str(&t.xml10_content()?),
                        Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut bufs.value)?,
                        Event::End(end) if end.name() == e.name() => break,
                        Event::Eof => return Err(XlsxError::XmlEof("v")),
                        _ => (),
                    }
                }
                read_v(ctx, bufs.value.as_bytes(), style_attr, type_attr)?
            }
        },
        b"f" => {
            bufs.xml.clear();
            xml.read_to_end_into(e.name(), &mut bufs.xml)?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// Convert raw `<v>` bytes to a `&str`, returning an error on invalid UTF-8.
fn v_as_str(v: &[u8]) -> Result<&str, XlsxError> {
    std::str::from_utf8(v).map_err(|_| XlsxError::Unexpected("invalid UTF-8 in cell value"))
}

/// Parse a `<v>` cell value from raw bytes with pre-extracted
/// `s` (style) and `t` (type) attributes.
fn read_v<'s>(
    ctx: &WorkbookContext<'s>,
    v: &[u8],
    style_attr: Option<&[u8]>,
    type_attr: Option<&[u8]>,
) -> Result<DataRef<'s>, XlsxError> {
    let cell_format = match style_attr {
        Some(style) => {
            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
            ctx.formats.get(id)
        }
        None => Some(&CellFormat::Other),
    };
    match type_attr {
        Some(b"s") => {
            if v.is_empty() {
                return Ok(DataRef::Empty);
            }
            let idx = atoi_simd::parse::<usize>(v).unwrap_or(0);
            ctx.strings
                .get(idx)
                .map(|s| DataRef::SharedString(s))
                .ok_or(XlsxError::Unexpected(
                    "Cell string index not found in shared strings table",
                ))
        }
        Some(b"b") => Ok(DataRef::Bool(v != b"0")),
        Some(b"d") => Ok(DataRef::DateTimeIso(v_as_str(v)?.to_string())),
        Some(b"e") => Ok(DataRef::Error(v_as_str(v)?.parse()?)),
        Some(b"str") => Ok(DataRef::String(v_as_str(v)?.to_string())),
        Some(b"n") | None => {
            if v.is_empty() {
                return Ok(DataRef::Empty);
            }
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            fast_float2::parse::<f64, _>(v)
                .map(|n| format_excel_f64_ref(n, cell_format, ctx.is_1904))
                .or_else(|_| {
                    if type_attr.is_none() {
                        // No explicit type: fall back to String if not a valid float
                        Ok(DataRef::String(v_as_str(v)?.to_string()))
                    } else {
                        Err(XlsxError::ParseFloat(
                            v_as_str(v)?.parse::<f64>().unwrap_err(),
                        ))
                    }
                })
        }
        Some(t) => {
            let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
            Err(XlsxError::CellTAttribute(t))
        }
    }
}

fn read_formula<RS>(xml: &mut XlReader<RS>, e: &BytesStart) -> Result<Option<String>, XlsxError>
where
    RS: Read + Seek,
{
    match e.local_name().as_ref() {
        b"is" | b"v" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            Ok(None)
        }
        b"f" => {
            let mut f_buf = Vec::with_capacity(512);
            let mut f = String::new();
            loop {
                match xml.read_event_into(&mut f_buf)? {
                    Event::Text(t) => f.push_str(&t.xml10_content()?),
                    Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut f)?,
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                    _ => (),
                }
                f_buf.clear();
            }
            Ok(Some(f))
        }
        _ => Err(XlsxError::UnexpectedNode("v, f, or is")),
    }
}
