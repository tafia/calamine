use std::fs::File;
use std::collections::HashMap;

use errors::*;

use {ExcelReader, Range};
use vba::VbaProject;

/// A struct representing an old xls format file (CFB)
pub struct Xls {
    file: File,
}

impl ExcelReader for Xls {
    fn new(f: File) -> Result<Self> {
        Ok(Xls { file: f, })
    }
    fn has_vba(&mut self) -> bool {
        true
    }
    fn vba_project(&mut self) -> Result<VbaProject> {
        let len = try!(self.file.metadata()).len() as usize;
        VbaProject::new(&mut self.file, len)
    }
    fn read_sheets_names(&mut self, _: &HashMap<Vec<u8>, String>) 
        -> Result<HashMap<String, String>> {
        unimplemented!()
    }
    fn read_shared_strings(&mut self) -> Result<Vec<String>> {
        unimplemented!()
    }
    fn read_relationships(&mut self) -> Result<HashMap<Vec<u8>, String>> {
        unimplemented!()
    }
    fn read_worksheet_range(&mut self, _: &str, _: &[String]) -> Result<Range> {
        unimplemented!()
    }
}
