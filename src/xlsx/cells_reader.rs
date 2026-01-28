// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{BytesStart, Event},
    name::QName,
};
use std::{
    borrow::Borrow,
    collections::HashMap,
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string, Dimensions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::unescape_entity_to_buffer,
    Cell, Style, XlsxError,
};

type FormulaMap = HashMap<(u32, u32), (i64, i64)>;

/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    xml: XlReader<'a, RS>,
    strings: &'a [String],
    formats: &'a [CellFormat],
    styles: &'a [Style],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
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
        styles: &'a [Style],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut dimensions = Dimensions {
            start: (0, 0),
            end: (0, 0),
        };
        let mut buf = Vec::with_capacity(1024);
        let mut sheet_type: Option<String> = None;
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.local_name().as_ref() {
                        b"dimension" => {
                            let attribute = get_attribute(e.attributes(), QName(b"ref"))?;
                            if let Some(range) = attribute {
                                dimensions = get_dimension(range)?;
                            }
                        }
                        b"sheetData" => {
                            return Ok(Self {
                                xml,
                                strings,
                                formats,
                                styles,
                                is_1904,
                                dimensions,
                                row_index: 0,
                                col_index: 0,
                                buf: Vec::with_capacity(1024),
                                cell_buf: Vec::with_capacity(1024),
                                formulas: Vec::with_capacity(1024),
                            });
                        }
                        typ => {
                            // Track the type of element we found (for non-worksheet detection)
                            if sheet_type.is_none() {
                                sheet_type = xml.decoder().decode(typ).ok().map(|s| s.to_string());
                            }
                        }
                    }
                }
                Ok(Event::Eof) => {
                    // If we reached EOF without finding sheetData, check if this is a non-worksheet
                    if let Some(typ) = sheet_type {
                        return Err(XlsxError::NotAWorksheet(typ));
                    } else {
                        return Err(XlsxError::XmlEof("sheetData"));
                    }
                }
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
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
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    let mut value = DataRef::Empty;
                    let mut style = None;

                    // Extract style ID if present, default to 0 if not present
                    let style_id = if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        atoi_simd::parse::<usize>(style_id_str).unwrap_or(0)
                    } else {
                        0 // Default to style ID 0 when not present
                    };

                    if style_id < self.styles.len() {
                        let mut s = self.styles[style_id].clone();
                        s.style_id = Some(style_id as u32);
                        style = Some(s);
                    }

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
                                    &c_element,
                                )?;
                            }
                            Ok(Event::End(e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;

                    if let Some(cell_style) = style {
                        return Ok(Some(Cell::with_style(pos, value, cell_style)));
                    } else {
                        return Ok(Some(Cell::new(pos, value)));
                    }
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
                    let mut style = None;

                    // Extract style ID if present, default to 0 if not present
                    let style_id = if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        atoi_simd::parse::<usize>(style_id_str).unwrap_or(0)
                    } else {
                        0 // Default to style ID 0 when not present
                    };

                    if style_id < self.styles.len() {
                        let mut s = self.styles[style_id].clone();
                        s.style_id = Some(style_id as u32);
                        style = Some(s);
                    }

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
                                            // This cell uses an existing shared formula - look it up and apply offset
                                            if let Some(Some((base_formula, offset_map))) =
                                                self.formulas.get(shared_index)
                                            {
                                                if let Some(offset) = offset_map.get(&pos) {
                                                    if let Ok(offset_formula) =
                                                        super::replace_cell_names(
                                                            base_formula,
                                                            *offset,
                                                        )
                                                    {
                                                        value = Some(offset_formula);
                                                    }
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

                    if let Some(cell_style) = style {
                        return Ok(Some(Cell::with_style(
                            pos,
                            value.unwrap_or_default(),
                            cell_style,
                        )));
                    } else {
                        return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    pub fn next_style(&mut self) -> Result<Option<Cell<Style>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };

                    // Extract style ID if present
                    let style = if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        if let Ok(style_id) = atoi_simd::parse::<usize>(style_id_str) {
                            if style_id < self.styles.len() {
                                let mut s = self.styles[style_id].clone();
                                s.style_id = Some(style_id as u32);
                                s
                            } else {
                                Style::new()
                            }
                        } else {
                            Style::new()
                        }
                    } else {
                        Style::new()
                    };

                    // Skip the cell content since we only care about the style
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, style)));
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

    /// Iterate over cells, returning just position and style_id (no clone).
    ///
    /// Returns `(row, col, style_id)` where `style_id` is an index into the styles palette.
    /// This is more efficient than `next_style()` when building compressed style storage.
    pub fn next_style_id(&mut self) -> Result<Option<(u32, u32, usize)>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    if let Some(row_index) = get_attribute(row_element.attributes(), QName(b"r"))? {
                        self.row_index = atoi_simd::parse::<u32>(row_index)
                            .unwrap_or(1)
                            .saturating_sub(1);
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };

                    // Extract style ID if present (no clone needed!)
                    let style_id = if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        atoi_simd::parse::<usize>(style_id_str).unwrap_or(0)
                    } else {
                        0
                    };

                    // Skip the cell content since we only care about the style ID
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;

                    // Only return cells with actual styles
                    if style_id > 0 && style_id < self.styles.len() {
                        return Ok(Some((pos.0, pos.1, style_id)));
                    }
                    // Continue to next cell if no style
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

    /// Get the styles palette (reference to avoid clones)
    pub fn styles(&self) -> &[Style] {
        self.styles
    }
}

fn read_value<'s, RS>(
    strings: &'s [String],
    formats: &[CellFormat],
    is_1904: bool,
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    c_element: &BytesStart<'_>,
) -> Result<DataRef<'s>, XlsxError>
where
    RS: Read + Seek,
{
    Ok(match e.local_name().as_ref() {
        b"is" => {
            // inlineStr
            read_string(xml, e.name())?.map_or(DataRef::Empty, DataRef::String)
        }
        b"v" => {
            // value
            let mut v = String::new();
            let mut v_buf = Vec::new();
            loop {
                v_buf.clear();
                match xml.read_event_into(&mut v_buf)? {
                    Event::Text(t) => v.push_str(&t.xml10_content()?),
                    Event::GeneralRef(e) => unescape_entity_to_buffer(&e, &mut v)?,
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => (),
                }
            }
            read_v(v, strings, formats, c_element, is_1904)?
        }
        b"f" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    })
}

/// read the contents of a <v> cell
fn read_v<'s>(
    v: String,
    strings: &'s [String],
    formats: &[CellFormat],
    c_element: &BytesStart<'_>,
    is_1904: bool,
) -> Result<DataRef<'s>, XlsxError> {
    let cell_format = match get_attribute(c_element.attributes(), QName(b"s")) {
        Ok(Some(style)) => {
            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
            formats.get(id)
        }
        _ => Some(&CellFormat::Other),
    };
    match get_attribute(c_element.attributes(), QName(b"t"))? {
        Some(b"s") => {
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
            Ok(DataRef::Bool(v != "0"))
        }
        Some(b"e") => {
            // error
            Ok(DataRef::Error(v.parse()?))
        }
        Some(b"d") => {
            // date
            Ok(DataRef::DateTimeIso(v))
        }
        Some(b"str") => {
            // string
            Ok(DataRef::String(v))
        }
        Some(b"n") => {
            // n - number
            if v.is_empty() {
                Ok(DataRef::Empty)
            } else {
                v.parse()
                    .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                    .map_err(XlsxError::ParseFloat)
            }
        }
        None => {
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            v.parse()
                .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                .or(Ok(DataRef::String(v)))
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
