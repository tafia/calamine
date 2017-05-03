#![feature(test)]

extern crate calamine;
extern crate test;

use test::Bencher;
use calamine::Sheets;

#[bench]
fn bench_xls(b: &mut Bencher) {
    b.iter(|| {
        let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
        let mut excel = Sheets::open(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().unwrap();
        let mut count = 0;
        for s in sheets {
            count += excel.worksheet_range(&s).unwrap().rows().flat_map(|r| r.iter()).count();
            count += excel.worksheet_formula(&s).unwrap().rows().flat_map(|r| r.iter()).count();
        }
        count
    })
}


#[bench]
fn bench_xlsx(b: &mut Bencher) {
    b.iter(|| {
        let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
        let mut excel = Sheets::open(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().unwrap();
        let mut count = 0;
        for s in sheets {
            count += excel.worksheet_range(&s).unwrap().rows().flat_map(|r| r.iter()).count();
            count += excel.worksheet_formula(&s).unwrap().rows().flat_map(|r| r.iter()).count();
        }
        count
    })
}

#[bench]
fn bench_xlsb(b: &mut Bencher) {
    b.iter(|| {
        let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
        let mut excel = Sheets::open(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().unwrap();
        let mut count = 0;
        for s in sheets {
            count += excel.worksheet_range(&s).unwrap().rows().flat_map(|r| r.iter()).count();
            count += excel.worksheet_formula(&s).unwrap().rows().flat_map(|r| r.iter()).count();
        }
        count
    })
}

#[bench]
fn bench_ods(b: &mut Bencher) {
    b.iter(|| {
        let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
        let mut excel = Sheets::open(&path).expect("cannot open excel file");

        let sheets = excel.sheet_names().unwrap();
        let mut count = 0;
        for s in sheets {
            count += excel.worksheet_range(&s).unwrap().rows().flat_map(|r| r.iter()).count();
            count += excel.worksheet_formula(&s).unwrap().rows().flat_map(|r| r.iter()).count();
        }
        count
    })
}
