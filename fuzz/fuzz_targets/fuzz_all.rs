#![no_main]
use libfuzzer_sys::fuzz_target;
use calamine::{Xlsx, open_workbook, Reader};
use std::io::Write;
use std::fs::File;

fuzz_target!(|data: &[u8]| {
    let file_name = "fuzz.xlsx";
    if let Ok(mut file) = File::create(file_name) {
        if file.write_all(data).is_err() {
            return
        }
    }
    let mut workbook: Xlsx<_> = match open_workbook(file_name) {
        Ok(excel) => excel,
        Err(_) => return,
    };
    for worksheet in workbook.worksheets() {
        if let Ok(range) = workbook.worksheet_range(&worksheet.0) {
           let _ = range.get_size().0 * range.get_size().1;
           range.used_cells().count();
        }
    }
    if let Some(Ok(mut vba)) = workbook.vba_project() {
        let vba = vba.to_mut();
        for module_name in vba.get_module_names() {
            if vba.get_module(module_name).is_ok() {
                for r in vba.get_references() {
                    r.is_missing();
                }
            }
        }
    }
    let sheets = workbook.sheet_names().to_owned();
    for s in sheets {
        if let Ok(formula) = workbook.worksheet_formula(&s) {
            formula.rows().flat_map(|r| r.iter().filter(|f| !f.is_empty())).count();
        }
    }
});
