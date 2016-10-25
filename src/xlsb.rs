use std::string::String;
use std::fs::File;
use std::io::BufReader;
use std::collections::HashMap;

use zip::read::{ZipFile, ZipArchive};
use zip::result::ZipError;

use {DataType, ExcelReader, Range};
use vba::VbaProject;
use utils;
use errors::*;

pub struct Xlsb {
    zip: ZipArchive<File>,
}

impl Xlsb {

//     /// Creates a new `BinarySheet`
//     pub fn new<R: Read>(mut f: R, len: usize) -> Result<BinarySheet> {
//         debug!("new binary sheet");
// 
//         let mut data = Vec::with_capacity(len);
//         try!(f.read_to_end(&mut data));
//         let data = &mut &*data;
// 
//         try!(utils::check_record(0x0081, data)); // BrtBeginSheet
//         let name = try!(read_ws_prop(data));
// 
//         Ok(BinarySheet {
//             name: name,
//         })
//     }

}

// /// Parses BrtWsProp record and returns sheet name
// /// MS-XLSB 2.4.820
// fn read_ws_prop(data: &mut &[u8]) -> Result<String> {
//     *data = &data[19 * 8..]; // discard first 19 bytes
//     let name = try!(utils::read_variable_record(data, 2));
//     let name = try!(String::from_utf8(name.to_vec()));
//     Ok(name)
// }

impl ExcelReader for Xlsb {

    fn new(f: File) -> Result<Self> {
        Ok(Xlsb { zip: try!(ZipArchive::new(f)) })
    }

    fn has_vba(&mut self) -> bool {
        self.zip.by_name("xl/vbaProject.bin").is_ok()
    }

    fn vba_project(&mut self) -> Result<VbaProject> {
        let mut f = try!(self.zip.by_name("xl/vbaProject.bin"));
        let len = f.size() as usize;
        VbaProject::new(&mut f, len)
    }

    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        unimplemented!()
    }

    fn read_sheets_names(&mut self, relationships: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>>
    {
        unimplemented!()
    }

    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        unimplemented!()
    }

    fn read_worksheet_range(&mut self, path: &str, strings: &[String]) -> Result<Range> {
        unimplemented!()
    }
}
