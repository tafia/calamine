use std::fs::File;
use std::io::BufReader;
use std::collections::HashMap;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;
use quick_xml::{XmlReader, Event, AsStr};

use {DataType, ExcelReader, Range};
use vba::VbaProject;
use utils;
use errors::*;

/// A struct representing xml zipped excel file
/// Xlsx, Xlsm, Xlam
pub struct Xlsx {
    zip: ZipArchive<File>,
}

impl Xlsx {
    fn xml_reader<'a>(&'a mut self, path: &str) 
        -> Option<Result<XmlReader<BufReader<ZipFile<'a>>>>> 
    {
        match self.zip.by_name(path) {
            Ok(f) => Some(Ok(XmlReader::from_reader(BufReader::new(f))
                             .with_check(false).trim_text(false))),
            Err(ZipError::FileNotFound) => None,
            Err(e) => return Some(Err(e.into())),
        }
    }
}

impl ExcelReader for Xlsx {

    fn new(f: File) -> Result<Self> {
        Ok(Xlsx { zip: try!(ZipArchive::new(f)) })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

//     fn vba_project(&mut self) -> Result<VbaProject<ZipFile>> {
//         let mut f = try!(self.zip.by_name("xl/vbaProject.bin"));
//         let len = f.size() as usize;
//         VbaProject::new(&mut f, len)
//     }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        let mut xml = match self.xml_reader("xl/sharedStrings.xml") {
            None => return Ok(Vec::new()),
            Some(x) => try!(x),
        };
        let mut strings = Vec::new();
        let mut rich_buffer: Option<String> = None;
        while let Some(res_event) = xml.next() {
            match res_event {
                Ok(Event::Start(ref e)) if e.name() == b"r" => {
                    if let None = rich_buffer {
                        // use a buffer since richtext has multiples <r> and <t> for the same cell
                        rich_buffer = Some(String::new());
                    }
                },
                Ok(Event::End(ref e)) if e.name() == b"si" => {
                    if let Some(s) = rich_buffer {
                        strings.push(s);
                        rich_buffer = None;
                    }
                },
                Ok(Event::Start(ref e)) if e.name() == b"t" => {
                    let value = try!(xml.read_text_unescaped(b"t"));
                    if let Some(ref mut s) = rich_buffer {
                        s.push_str(&value);
                    } else {
                        strings.push(value);
                    }
                }
                Err(e) => return Err(e.into()),
                _ => (),
            }
        }
        Ok(strings)
    }

    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>>
    {
        let xml = match self.xml_reader("xl/workbook.xml") {
            None => return Ok(HashMap::new()),
            Some(x) => try!(x),
        };
        let mut sheets = HashMap::new();
        for res_event in xml {
            match res_event {
                Ok(Event::Start(ref e)) if e.name() == b"sheet" => {
                    let mut name = String::new();
                    let mut path = String::new();
                    for a in e.unescaped_attributes() {
                        match try!(a) {
                            (b"name", v) => name = try!(v.as_str()).to_string(),
                            (b"r:id", v) => path = format!("xl/{}", relationships[&*v]),
                            _ => (),
                        }
                    }
                    sheets.insert(name, path);
                }
                Err(e) => return Err(e.into()),
                _ => (),
            }
        }
        Ok(sheets)
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let xml = match self.xml_reader("xl/_rels/workbook.xml.rels") {
            None => return Err("Cannot find relationships file".into()),
            Some(x) => try!(x),
        };
        let mut relationships = HashMap::new();
        for res_event in xml {
            match res_event {
                Ok(Event::Start(ref e)) if e.name() == b"Relationship" => {
                    let mut id = Vec::new();
                    let mut target = String::new();
                    for a in e.attributes() {
                        match try!(a) {
                            (b"Id", v) => id.extend_from_slice(v),
                            (b"Target", v) => target = try!(v.as_str()).to_string(),
                            _ => (),
                        }
                    }
                    relationships.insert(id, target);
                }
                Err(e) => return Err(e.into()),
                _ => (),
            }
        }
        Ok(relationships)
    }

    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        let mut xml = match self.xml_reader(path) {
            None => return Err(format!("Cannot find {} path", path).into()),
            Some(x) => try!(x),
        };
        let mut data = Range::default();
        'xml: while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(e.into()),
                Ok(Event::Start(ref e)) => {
                    match e.name() {
                        b"dimension" => {
                            for a in e.attributes() {
                                if let (b"ref", rdim) = try!(a) {
                                    let (position, size) = try!(utils::get_dimension(rdim));
                                    data = Range::new(position, (size.0 as usize, size.1 as usize));
                                    continue 'xml;
                                }
                            }
                            return Err(format!("Expecting dimension, got {:?}", e).into());
                        },
                        b"sheetData" => try!(read_sheet_data(&mut xml, strings, &mut data)),
                        _ => (),
                    }
                },
                _ => (),
            }
        }
        data.inner.shrink_to_fit();
        Ok(data)
    }
}

/// read sheetData node
fn read_sheet_data(xml: &mut XmlReader<BufReader<ZipFile>>, 
                   strings: &[String], range: &mut Range) -> Result<()> {
    while let Some(res_event) = xml.next() {
        match res_event {
            Err(e) => return Err(e.into()),
            Ok(Event::Start(ref c_element)) if c_element.name() == b"c" => {
                let pos = match c_element.attributes().filter_map(|a| match a {
                    Err(e) => Some(Err(e.into())),
                    Ok((b"r", v)) => Some(utils::get_row_column(v)),
                    _ => None,
                }).next() {
                    Some(v) => try!(v),
                    None => return Err("Cell without a 'r' reference tag".into()),
                };
                loop {
                    match xml.next() {
                        Some(Err(e)) => return Err(e.into()),
                        Some(Ok(Event::Start(ref e))) => match e.name() {
                            b"v" => {
                                // value
                                let v = try!(xml.read_text_unescaped(b"v"));
                                let value = match c_element.attributes()
                                    .filter_map(|a| a.ok())
                                    .find(|&(k, _)| k == b"t") {
                                        Some((_, b"s")) => {
                                            // shared string
                                            let idx: usize = try!(v.parse());
                                            DataType::String(strings[idx].clone())
                                        },
                                        Some((_, b"str")) => {
                                            // regular string
                                            DataType::String(v)
                                        },
                                        Some((_, b"b")) => {
                                            // boolean
                                            DataType::Bool(v != "0")
                                        },
                                        Some((_, b"e")) => {
                                            // error
                                            DataType::Error(try!(v.parse()))
                                        },
                                        _ => try!(v.parse().map(DataType::Float)),
                                    };
                                range.set_value(pos, value);
                                break;
                            },
                            b"f" => (), // formula, ignore
                            _name => return Err("not v or f node".into()),
                        },
                        Some(Ok(Event::End(ref e))) if e.name() == b"c" => break,
                        None => return Err("End of xml".into()),
                        _ => (),
                    }
                }
            },
            Ok(Event::End(ref e)) if e.name() == b"sheetData" => return Ok(()),
            _ => (),
        }
    }
    Err("Could not find </sheetData>".into())
}
