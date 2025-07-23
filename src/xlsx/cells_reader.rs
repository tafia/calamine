// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};
use std::{
    borrow::Borrow,
    collections::HashMap,
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string, replace_cell_names,
    Dimensions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
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
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dimension" => {
                    let attribute = get_attribute(e.attributes(), QName(b"ref"))?;
                    if let Some(range) = attribute {
                        dimensions = get_dimension(range)?;
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetData" => {
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
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
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
                    let mut value = DataRef::Empty;
                    let mut style = None;

                    // Extract style ID if present
                    if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        if let Ok(style_id) = atoi_simd::parse::<usize>(style_id_str) {
                            if style_id < self.styles.len() {
                                style = Some(self.styles[style_id].clone());
                            }
                        }
                    }

                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                value = read_value(
                                    self.strings,
                                    self.formats,
                                    self.is_1904,
                                    &mut self.xml,
                                    e,
                                    c_element,
                                )?;
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
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
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
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
                    let mut value = None;
                    let mut style = None;

                    // Extract style ID if present
                    if let Ok(Some(style_id_str)) =
                        get_attribute(c_element.attributes(), QName(b"s"))
                    {
                        if let Ok(style_id) = atoi_simd::parse::<usize>(style_id_str) {
                            if style_id < self.styles.len() {
                                style = Some(self.styles[style_id].clone());
                            }
                        }
                    }

                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                let formula = read_formula(&mut self.xml, e)?;
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
                                            // orignal reference formula
                                            let reference = get_dimension(res)?;
                                            if reference.start.0 != reference.end.0 {
                                                for i in 0..=(reference.end.0 - reference.start.0) {
                                                    offset_map.insert(
                                                        (reference.start.0 + i, reference.start.1),
                                                        (
                                                            (reference.start.0 as i64
                                                                - pos.0 as i64
                                                                + i as i64),
                                                            0,
                                                        ),
                                                    );
                                                }
                                            } else if reference.start.1 != reference.end.1 {
                                                for i in 0..=(reference.end.1 - reference.start.1) {
                                                    offset_map.insert(
                                                        (reference.start.0, reference.start.1 + i),
                                                        (
                                                            0,
                                                            (reference.start.1 as i64
                                                                - pos.1 as i64
                                                                + i as i64),
                                                        ),
                                                    );
                                                }
                                            }
                                            self.formulas.push(Some((
                                                value.clone().unwrap_or_default(),
                                                offset_map,
                                            )));
                                        }
                                        None => {
                                            self.formulas.push(Some((
                                                value.clone().unwrap_or_default(),
                                                HashMap::new(),
                                            )));
                                        }
                                    }
                                } else {
                                    self.formulas.push(None);
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
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
                    Event::Text(t) => v.push_str(&t.unescape()?),
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
            // shared string
            let idx = atoi_simd::parse::<usize>(v.as_bytes()).unwrap_or(0);
            Ok(DataRef::SharedString(&strings[idx]))
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
                    Event::Text(t) => f.push_str(&t.unescape()?),
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
