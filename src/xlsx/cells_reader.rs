// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use hashbrown::HashMap;
use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};
use std::{
    borrow::{Borrow, Cow},
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string_with_bufs,
    replace_cell_names, Dimensions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::unescape_entity_to_buffer,
    Cell, XlsxError,
};

type FormulaMap = HashMap<(u32, u32), (i64, i64)>;

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
    formulas: Vec<Option<(String, FormulaMap)>>,
}

impl<'a, RS> XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
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

    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

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
                                value = read_value(
                                    self.strings,
                                    self.formats,
                                    self.is_1904,
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
                                let formula = read_formula(&mut self.xml, &e)?;
                                if let Some(f) = formula.borrow() {
                                    value = Some(f.clone());
                                }
                                if let Ok(Some(b"shared")) =
                                    get_attribute(e.attributes(), QName(b"t"))
                                {
                                    // shared formula
                                    let mut offset_map: HashMap<(u32, u32), (i64, i64)> =
                                        HashMap::new();
                                    // shared index
                                    let shared_index =
                                        match get_attribute(e.attributes(), QName(b"si"))? {
                                            Some(res) => match atoi_simd::parse::<usize>(res) {
                                                Ok(res) => res,
                                                Err(_) => {
                                                    return Err(XlsxError::Unexpected(
                                                        "si attribute must be a number",
                                                    ));
                                                }
                                            },
                                            None => {
                                                return Err(XlsxError::Unexpected(
                                                    "si attribute is mandatory if it is shared",
                                                ));
                                            }
                                        };
                                    // shared reference
                                    match get_attribute(e.attributes(), QName(b"ref"))? {
                                        Some(res) => {
                                            // original reference formula
                                            let reference = get_dimension(res)?;

                                            for row in reference.start.0..=reference.end.0 {
                                                for column in reference.start.1..=reference.end.1 {
                                                    offset_map.insert(
                                                        (row, column),
                                                        (
                                                            row as i64 - pos.0 as i64,
                                                            column as i64 - pos.1 as i64,
                                                        ),
                                                    );
                                                }
                                            }

                                            if let Some(f) = formula.borrow() {
                                                if self.formulas.len() <= shared_index {
                                                    self.formulas.resize(shared_index + 1, None);
                                                }
                                                self.formulas[shared_index] =
                                                    Some((f.clone(), offset_map));
                                            }
                                            value = formula;
                                        }
                                        None => {
                                            // calculated formula
                                            if let Some(Some((f, offset_map))) =
                                                self.formulas.get(shared_index)
                                            {
                                                if let Some(offset) = offset_map.get(&pos) {
                                                    value = Some(replace_cell_names(f, *offset)?);
                                                }
                                            }
                                        }
                                    }
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
    strings: &'s [String],
    formats: &[CellFormat],
    is_1904: bool,
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
        b"v" => {
            // value
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
            read_v(
                &mut bufs.value,
                strings,
                formats,
                style_attr,
                type_attr,
                is_1904,
            )?
        }
        b"f" => {
            bufs.xml.clear();
            xml.read_to_end_into(e.name(), &mut bufs.xml)?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// Read the contents of a `<v>` cell using pre-extracted `s` and `t` attributes.
/// Takes `v` by mutable reference to allow for buffer reuse across cells.
fn read_v<'s>(
    v: &mut String,
    strings: &'s [String],
    formats: &[CellFormat],
    style_attr: Option<&[u8]>,
    type_attr: Option<&[u8]>,
    is_1904: bool,
) -> Result<DataRef<'s>, XlsxError> {
    let cell_format = match style_attr {
        Some(style) => {
            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
            formats.get(id)
        }
        _ => Some(&CellFormat::Other),
    };
    match type_attr {
        Some(b"s") => {
            if v.is_empty() {
                return Ok(DataRef::Empty);
            }
            // Cell value is an index into the shared string table.
            let idx = atoi_simd::parse::<usize>(v.as_bytes()).unwrap_or(0);
            match strings.get(idx) {
                Some(shared_string) => Ok(DataRef::SharedString(shared_string)),
                None => Err(XlsxError::Unexpected(
                    "Cell string index not found in shared strings table",
                )),
            }
        }
        Some(b"b") => {
            // boolean
            Ok(DataRef::Bool(v.as_str() != "0"))
        }
        Some(b"e") => {
            // error
            Ok(DataRef::Error(v.parse()?))
        }
        Some(b"d") => {
            // date (needs owned String)
            Ok(DataRef::DateTimeIso(std::mem::take(v)))
        }
        Some(b"str") => {
            // string (needs owned String)
            Ok(DataRef::String(std::mem::take(v)))
        }
        Some(b"n") => {
            // n - number
            if v.is_empty() {
                Ok(DataRef::Empty)
            } else {
                fast_float2::parse::<f64, _>(v.as_bytes())
                    .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                    .map_err(|_| XlsxError::ParseFloat(v.parse::<f64>().unwrap_err()))
            }
        }
        None => {
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            fast_float2::parse::<f64, _>(v.as_bytes())
                .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                .or(Ok(DataRef::String(std::mem::take(v))))
        }
        Some(b"is") => {
            // this case should be handled in outer loop over cell elements, in which
            // case read_inline_str is called instead. Case included here for completeness.
            Err(XlsxError::Unexpected(
                "called read_value on a cell of type inlineStr",
            ))
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
