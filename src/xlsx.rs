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
    pub fn new(f: File) -> Result<Self> {
        let zip = try!(ZipArchive::new(f));
        Ok(Xlsx {
            zip: zip,
        })
    }

}

impl ExcelReader for Xlsx {

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<VbaProject> {
        let mut f = try!(self.zip.by_name("xl/vbaProject.bin"));
        let len = f.size() as usize;
        VbaProject::new(&mut f, len)
    }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        let mut strings = Vec::new();
        match self.zip.by_name("xl/sharedStrings.xml") {
            Ok(f) => {
                let mut xml = XmlReader::from_reader(BufReader::new(f))
                    .with_check(false)
                    .trim_text(false);

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
                            let value = try!(xml.read_text(b"t"));
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
            },
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(strings)
    }

    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>>
    {
        let mut sheets = HashMap::new();
        match self.zip.by_name("xl/workbook.xml") {
            Ok(f) => {
                let xml = XmlReader::from_reader(BufReader::new(f))
                    .with_check(false)
                    .trim_text(false);

                for res_event in xml {
                    match res_event {
                        Ok(Event::Start(ref e)) if e.name() == b"sheet" => {
                            let mut name = String::new();
                            let mut path = String::new();
                            for a in e.attributes() {
                                match try!(a) {
                                    (b"name", v) => name = try!(v.as_str()).to_string(),
                                    (b"r:id", v) => path = format!("xl/{}", relationships[v]),
                                    _ => (),
                                }
                            }
                            sheets.insert(name, path);
                        }
                        Err(e) => return Err(e.into()),
                        _ => (),
                    }
                }
            },
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(sheets)
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        let mut relationships = HashMap::new();
        match self.zip.by_name("xl/_rels/workbook.xml.rels") {
            Ok(f) => {
                let xml = XmlReader::from_reader(BufReader::new(f))
                    .with_check(false)
                    .trim_text(false);

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
            },
            Err(ZipError::FileNotFound) => (),
            Err(e) => return Err(e.into()),
        }
        Ok(relationships)
    }

    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        let xml = try!(self.zip.by_name(path));
        let mut xml = XmlReader::from_reader(BufReader::new(xml))
            .with_check(false)
            .trim_text(false);
        let mut data = Range::default();
        while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(e.into()),
                Ok(Event::Start(ref e)) => {
                    match e.name() {
                        b"dimension" => {
                            let mut dim = None;
                            for a in e.attributes() {
                                if let (b"ref", rdim) = try!(a) {
                                    dim = Some(rdim);
                                    break;
                                }
                            }
                            match dim {
                                None => return Err(format!("Expecting dimension, got {:?}", e).into()),
                                Some(dim) => {
                                    let (position, size) = try!(utils::get_dimension(dim));
                                    data.position = position;
                                    data.size = (size.0 as usize, size.1 as usize);
                                }
                            }
                        },
                        b"sheetData" => {
                            data.inner = try!(read_sheet_data(&mut xml, strings, 
                                                              data.size.0 * data.size.1))
                        },
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
fn read_sheet_data(xml: &mut XmlReader<BufReader<ZipFile>>, strings: &[String], len: usize) 
    -> Result<Vec<DataType>>
{
    let mut data = Vec::with_capacity(len);
    while let Some(res_event) = xml.next() {
        match res_event {
            Err(e) => return Err(e.into()),
            Ok(Event::Start(ref c_element)) => {
                if c_element.name() == b"c" {
                    loop {
                        match xml.next() {
                            Some(Err(e)) => return Err(e.into()),
                            Some(Ok(Event::Start(ref e))) => match e.name() {
                                b"v" => {
                                    // value
                                    let v = try!(xml.read_text(b"v"));
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
                                            _ => match v.parse() {
                                                // TODO: check in styles to know which type is
                                                // supposed to be used
                                                Ok(i) => DataType::Int(i),
                                                Err(_) => try!(v.parse().map(DataType::Float)),
                                            },
                                        };
                                    data.push(value);
                                    break;
                                },
                                b"f" => (), // formula, ignore
                                _name => return Err("not v or f node".into()),
                            },
                            Some(Ok(Event::End(ref e))) => {
                                if e.name() == b"c" {
                                    data.push(DataType::Empty);
                                    break;
                                }
                            }
                            None => return Err("End of xml".into()),
                            _ => (),
                        }
                    }
                }
            },
            Ok(Event::End(ref e)) if e.name() == b"sheetData" => return Ok(data),
            _ => (),
        }
    }
    Err("Could not find </sheetData>".into())
}
