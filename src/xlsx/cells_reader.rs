use std::collections::HashMap;

use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};
use regex::{Captures, Regex};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, position_to_title, read_string,
    Dimensions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    Cell, XlsxError,
};

fn replace_all<E>(
    re: &Regex,
    haystack: &str,
    replacement: impl Fn(&Captures) -> Result<String, E>,
) -> Result<String, E> {
    let mut new = String::with_capacity(haystack.len());
    let mut last_match = 0;
    for caps in re.captures_iter(haystack) {
        let m = caps.get(0).unwrap();
        new.push_str(&haystack[last_match..m.start()]);
        new.push_str(&replacement(&caps)?);
        last_match = m.end();
    }
    new.push_str(&haystack[last_match..]);
    Ok(new)
}
/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a> {
    xml: XlReader<'a>,
    strings: &'a [String],
    formats: &'a [CellFormat],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    formulas: Vec<Option<(String, HashMap<String, (i32, i32)>)>>,
}

impl<'a> XlsxCellReader<'a> {
    pub fn new(
        mut xml: XlReader<'a>,
        strings: &'a [String],
        formats: &'a [CellFormat],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        let mut dimensions = Dimensions::default();
        'xml: loop {
            buf.clear();
            match xml.read_event_into(&mut buf).map_err(XlsxError::Xml)? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: QName(b"ref"),
                                value: rdim,
                            } = a.map_err(XlsxError::XmlAttr)?
                            {
                                dimensions = get_dimension(&rdim)?;
                                continue 'xml;
                            }
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
                    b"sheetData" => break,
                    _ => (),
                },
                Event::Eof => return Err(XlsxError::XmlEof("sheetData")),
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
            formulas: Vec::with_capacity(1024),
        })
    }

    pub(crate) fn dimensions(&self) -> Dimensions {
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
                                )?
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value)));
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
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                let mut offset_map: HashMap<String, (i32, i32)> = HashMap::new();
                                let mut shared_index = None;
                                let mut shared_ref = None;
                                let shared =
                                    get_attribute(e.attributes(), QName(b"t")).unwrap_or(None);
                                match shared {
                                    Some(b"shared") => {
                                        shared_index = Some(
                                            String::from_utf8(
                                                get_attribute(e.attributes(), QName(b"si"))?
                                                    .unwrap()
                                                    .to_vec(),
                                            )
                                            .unwrap()
                                            .parse::<u32>()?,
                                        );
                                        match get_attribute(e.attributes(), QName(b"ref"))? {
                                            Some(res) => {
                                                let reference = get_dimension(res)?;
                                                if reference.start.0 != reference.end.0 {
                                                    for i in
                                                        0..=(reference.end.0 - reference.start.0)
                                                    {
                                                        offset_map.insert(
                                                            position_to_title((
                                                                reference.start.0 + i,
                                                                reference.start.1,
                                                            ))?,
                                                            (
                                                                (reference.start.0 as i64
                                                                    - pos.0 as i64
                                                                    + i as i64)
                                                                    as i32,
                                                                0,
                                                            ),
                                                        );
                                                    }
                                                } else if reference.start.1 != reference.end.1 {
                                                    for i in
                                                        0..=(reference.end.1 - reference.start.1)
                                                    {
                                                        offset_map.insert(
                                                            position_to_title((
                                                                reference.start.0,
                                                                reference.start.1 + i,
                                                            ))?,
                                                            (
                                                                0,
                                                                (reference.start.1 as i64
                                                                    - pos.1 as i64
                                                                    + i as i64)
                                                                    as i32,
                                                            ),
                                                        );
                                                    }
                                                }
                                                shared_ref = Some(reference);
                                            }
                                            None => {}
                                        }
                                    }
                                    _ => {}
                                }
                                if let Some(f) = read_formula(&mut self.xml, e)? {
                                    value = Some(f.clone());
                                    if shared_index.is_some() && shared_ref.is_some() {
                                        // original shared formula
                                        while self.formulas.len() < shared_index.unwrap() as usize {
                                            self.formulas.push(None);
                                        }
                                        self.formulas.push(Some((f, offset_map)));
                                    }
                                }
                                if shared_index.is_some() && shared_ref.is_none() {
                                    // shared formula
                                    let cell_regex = Regex::new(r"[A-Z]+[0-9]+").unwrap();
                                    if let Some((f, offset)) =
                                        self.formulas[shared_index.unwrap() as usize].clone()
                                    {
                                        let (row, col) =
                                            offset.get(&position_to_title(pos)?).unwrap();
                                        let replacement =
                                            |caps: &Captures| -> Result<String, &'static str> {
                                                match get_row_column(caps[0].as_bytes()) {
                                                    Ok(cell) => {
                                                        match position_to_title((
                                                            (cell.0 as i64 + *col as i64) as u32,
                                                            (cell.1 as i64 + *row as i64 + 1)
                                                                as u32,
                                                        )) {
                                                            Ok(name) => Ok(name),
                                                            Err(_) => Err("invalid cell reference"),
                                                        }
                                                    }
                                                    Err(_) => Err("invalid cell reference"),
                                                }
                                            };
                                        match replace_all(&cell_regex, f.as_str(), &replacement) {
                                            Ok(s) => {
                                                value = Some(s);
                                            }
                                            Err(_) => {}
                                        }
                                    };
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some(Cell::new(pos, value.unwrap_or_default())));
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

fn read_value<'s>(
    strings: &'s [String],
    formats: &[CellFormat],
    is_1904: bool,
    xml: &mut XlReader<'_>,
    e: &BytesStart<'_>,
    c_element: &BytesStart<'_>,
) -> Result<DataRef<'s>, XlsxError> {
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
            let id: usize = std::str::from_utf8(style).unwrap_or("0").parse()?;
            formats.get(id)
        }
        _ => Some(&CellFormat::Other),
    };
    match get_attribute(c_element.attributes(), QName(b"t"))? {
        Some(b"s") => {
            // shared string
            let idx: usize = v.parse()?;
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
            // see http://officeopenxml.com/SScontentOverview.php
            // str - refers to formula cells
            // * <c .. t='v' .. > indicates calculated value (this case)
            // * <c .. t='f' .. > to the formula string (ignored case
            // TODO: Fully support a Data::Formula representing both Formula string &
            // last calculated value?
            //
            // NB: the result of a formula may not be a numeric value (=A3&" "&A4).
            // We do try an initial parse as Float for utility, but fall back to a string
            // representation if that fails
            v.parse().map(DataRef::Float).or(Ok(DataRef::String(v)))
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

fn read_formula<'s>(
    xml: &mut XlReader<'_>,
    e: &BytesStart<'_>,
) -> Result<Option<String>, XlsxError> {
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
