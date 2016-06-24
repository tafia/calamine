extern crate zip;
extern crate quick_xml;

mod error;

use std::path::Path;
use std::fs::File;
use std::io::BufReader;

use error::{ExcelError, ExcelResult};

use zip::read::{ZipFile, ZipArchive};
use quick_xml::{XmlReader, Event, AsStr};

type SharedStringIndex = usize;

#[derive(Debug)]
pub enum DataType {
    Int(i64),
    Float(f64),
    String(String),
    Empty,
}

pub struct Excel {
    zip: ZipArchive<File>,
    strings: Vec<String>,
}

#[derive(Debug, Default)]
pub struct WorksheetData {
    top_left: (u32, u32),
    size: (usize, usize),
    inner: Vec<DataType>,
}

impl Excel {

    /// Opens a new workbook
    pub fn open<P: AsRef<Path>>(path: P) -> ExcelResult<Excel> {
        let f = try!(File::open(path));
        let mut zip = try!(ZipArchive::new(f));
        let strings = {
            let xml = try!(zip.by_name("xl/sharedStrings.xml"));
            try!(Excel::read_shared_strings(xml))
        };
        Ok(Excel{ zip: zip, strings: strings })
    }

    /// Get all data from `Worksheet`
    pub fn worksheet_data(&mut self, name: &str) -> ExcelResult<WorksheetData> {
        let strings = &self.strings;
        let ws = match self.zip.by_name(&format!("xl/worksheets/{}.xml", name)) {
            Ok(f) => f,
            Err(e) => return Err(ExcelError::Zip(e)),
        };
        WorksheetData::from_xml(ws, strings)
    }

    /// Read shared string list
    fn read_shared_strings(xml: ZipFile) -> ExcelResult<Vec<String>> {
        let mut xml = XmlReader::from_reader(BufReader::new(xml))
            .with_check(false)
            .trim_text(false);

        let mut strings = Vec::new();
        while let Some(res_event) = xml.next() {
            match res_event {
                Ok(Event::Start(ref e)) if e.name() == b"t" => {
                    strings.push(try!(xml.read_text(b"t")));
                }
                Err(e) => return Err(ExcelError::Xml(e)),
                _ => (),
            }
        }
        Ok(strings)
    }

}

impl WorksheetData {

    /// open a xml `ZipFile` reader and read content of *sheetData* and *dimension* node
    fn from_xml(xml: ZipFile, strings: &[String]) -> ExcelResult<WorksheetData> {
        let mut xml = XmlReader::from_reader(BufReader::new(xml))
            .with_check(false)
            .trim_text(false);
        let mut data = WorksheetData::default();
        while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(ExcelError::Xml(e)),
                Ok(Event::Start(ref e)) => {
                    match e.name() {
                        b"dimension" => match e.attributes().filter_map(|a| a.ok())
                                .find(|&(key, _)| key == b"ref") {
                            Some((_, dim)) => {
                                let (top_left, size) = try!(get_dimension(try!(dim.as_str())));
                                data.top_left = top_left;
                                data.size = (size.0 as usize, size.1 as usize);
                                data.inner.reserve_exact(data.size.0 * data.size.1);
                            },
                            None => return Err(ExcelError::Unexpected(
                                    format!("Expecting dimension, got {:?}", e))),
                        },
                        b"sheetData" => {
                            let _ = try!(data.read_sheet_data(&mut xml, strings));
                        }
                        _ => (),
                    }
                },
                _ => (),
            }
        }
        data.inner.shrink_to_fit();
        Ok(data)
    }
    
    /// get worksheet position (row, column)
    pub fn get_position(&self) -> (u32, u32) {
        self.top_left
    }

    pub fn get_size(&self) -> (usize, usize) {
        self.size
    }

    pub fn get_value(&self, i: usize, j: usize) -> &DataType {
        let idx = i * self.size.0 + j;
        &self.inner[idx]
    }

    /// read sheetData node
    fn read_sheet_data(&mut self, xml: &mut XmlReader<BufReader<ZipFile>>, strings: &[String]) 
        -> ExcelResult<()> 
    {
        while let Some(res_event) = xml.next() {
            match res_event {
                Err(e) => return Err(ExcelError::Xml(e)),
                Ok(Event::Start(ref c_element)) => {
                    if c_element.name() == b"c" {
                        loop {
                            match xml.next() {
                                Some(Err(e)) => return Err(ExcelError::Xml(e)),
                                Some(Ok(Event::Start(ref e))) => {
                                    if e.name() == b"v" {
                                        let v = try!(xml.read_text(b"v"));
                                        let value = match c_element.attributes()
                                            .filter_map(|a| a.ok())
                                            .find(|&(k, _)| k == b"t") {
                                                Some((_, b"s")) => {
                                                    let idx: usize = try!(v.parse());
                                                    DataType::String(strings[idx].clone())
                                                },
                                                _ => DataType::Float(try!(v.parse()))
                                            };
                                        self.inner.push(value);
                                        break;
                                    } else {
                                        return Err(ExcelError::Unexpected("not v node".to_string()));
                                    }
                                },
                                Some(Ok(Event::End(ref e))) => {
                                    if e.name() == b"c" {
                                        self.inner.push(DataType::Empty);
                                        break;
                                    }
                                }
                                None => {
                                    return Err(ExcelError::Unexpected("End of xml".to_string()));
                                }
                                _ => (),
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) if e.name() == b"sheetData" => return Ok(()),
                _ => (),
            }
        }
        Err(ExcelError::Unexpected("Reached end of file, expecting </sheetData>".to_string()))
    }

}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column), 
/// - size (width, height)
fn get_dimension(dimension: &str) -> ExcelResult<((u32, u32), (u32, u32))> {
    match dimension.chars().position(|c| c == ':') {
        None => {
            get_row_column(dimension).map(|top_left| (top_left, (1, 1)))
        }, 
        Some(p) => {
            let top_left = try!(get_row_column(&dimension[..p]));
            let bottom_right = try!(get_row_column(&dimension[p + 1..]));
            Ok((top_left, (bottom_right.0 - top_left.0 + 1, bottom_right.1 - top_left.1 + 1)))
        }
    }
}

/// converts a text range name into its position (row, column)
fn get_row_column(range: &str) -> ExcelResult<(u32, u32)> {
    let mut col = 0;
    let mut pow = 1;
    let mut rowpos = range.len();
    let mut readrow = true;
    for c in range.chars().rev() {
        match c {
            '0'...'9' => {
                if readrow {
                    rowpos -= 1;
                } else {
                    return Err(ExcelError::Unexpected(
                        format!("Numeric character are only allowed at the end of the range: {}", c)));
                }
            }
            c @ 'A'...'Z' => {
                readrow = false;
                col += ((c as u8 - b'A') as u32 + 1) * pow;
                pow *= 26;
            },
            c @ 'a'...'z' => {
                readrow = false;
                col += ((c as u8 - b'a') as u32 + 1) * pow;
                pow *= 26;
            },
            _ => return Err(ExcelError::Unexpected(
                    format!("Expecting alphanumeric character, got {:?}", c))),
        }
    }
    let row = try!(range[rowpos..].parse());
    Ok((row, col))
}

#[cfg(test)]
mod tests {
    use super::Excel;
    #[test]
    fn it_works() {
        let mut xl = Excel::open("/home/jtuffe/download/DailyValo_FX_Rates_Credit_05 25 16.xlsm")
            .expect("cannot open excel file");
        let data = xl.worksheet_data("sheet1");
        println!("{:?}", data);
        assert!(data.is_ok());
    }
}
