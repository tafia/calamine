extern crate calamine;

use calamine::Sheets;
use calamine::DataType::{String, Empty, Float, Bool, Error};
use calamine::CellErrorType::*;

macro_rules! range_eq {
    ($range:expr, $right:expr) => {
        assert_eq!($range.get_size(), ($right.len(), $right[0].len()), "Size mismatch");
        for (i, (rl, rr)) in $range.rows().zip($right.iter()).enumerate() {
            for (j, (cl, cr)) in rl.iter().zip(rr.iter()).enumerate() {
                assert_eq!(cl, cr, "Mismatch at position ({}, {})", i, j);
            }
        }
    };
}

#[test]
fn issue_2() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(range,
              [[Float(1.), String("a".to_string())],
               [Float(2.), String("b".to_string())],
               [Float(3.), String("c".to_string())]]);
}

#[test]
fn issue_3() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(range, [[Float(1.), String("a".to_string())]]);
}

#[test]
fn issue_4() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue5").unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn issue_6() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue6").unwrap();
    range_eq!(range,
              [[Float(1.)],
               [Float(2.)],
               [String("ab".to_string())],
               [Bool(false)]]);
}

#[test]
fn error_file() {
    let path = format!("{}/tests/errors.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Feuil1").unwrap();
    range_eq!(range,
              [[Error(Div0)],
               [Error(Name)],
               [Error(Value)],
               [Error(Null)],
               [Error(Ref)],
               [Error(Num)],
               [Error(NA)]]);
}

#[test]
fn issue_9() {
    let path = format!("{}/tests/issue9.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Feuil1").unwrap();
    range_eq!(range,
              [[String("test1".to_string())],
               [String("test2 other".to_string())],
               [String("test3 aaa".to_string())],
               [String("test4".to_string())]]);
}

#[test]
fn vba() {
    let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let mut vba = excel.vba_project().unwrap();
    assert_eq!(vba.to_mut().get_module("testVBA").unwrap(),
               "Attribute VB_Name = \"testVBA\"\r\nPublic Sub test()\r\n    MsgBox \"Hello from \
                vba!\"\r\nEnd Sub\r\n");
}

#[test]
fn xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(range,
              [[Float(1.), String("a".to_string())],
               [Float(2.), String("b".to_string())],
               [Float(3.), String("c".to_string())]]);
}

#[test]
fn xls() {
    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(range,
              [[Float(1.), String("a".to_string())],
               [Float(2.), String("b".to_string())],
               [Float(3.), String("c".to_string())]]);
}

#[test]
fn ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("datatypes").unwrap();
    range_eq!(range,
              [[Float(1.)],
               [Float(1.5)],
               [String("ab".to_string())],
               [Bool(false)],
               [String("test".to_string())],
               [String("2016-10-20T00:00:00".to_string())]]);

    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(range,
              [[Float(1.), String("a".to_string())],
               [Float(2.), String("b".to_string())],
               [Float(3.), String("c".to_string())]]);

    let range = excel.worksheet_range("issue5").unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn special_chrs_xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("spc_chrs").unwrap();
    range_eq!(range,
              [[String("&".to_string())],
               [String("<".to_string())],
               [String(">".to_string())],
               [String("aaa ' aaa".to_string())],
               [String("\"".to_string())],
               [String("☺".to_string())],
               [String("֍".to_string())],
               [String("àâéêèçöïî«»".to_string())]]);
}

#[test]
fn special_chrs_xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("spc_chrs").unwrap();
    range_eq!(range,
              [[String("&".to_string())],
               [String("<".to_string())],
               [String(">".to_string())],
               [String("aaa ' aaa".to_string())],
               [String("\"".to_string())],
               [String("☺".to_string())],
               [String("֍".to_string())],
               [String("àâéêèçöïî«»".to_string())]]);
}

#[test]
fn special_chrs_ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("spc_chrs").unwrap();
    range_eq!(range,
              [[String("&".to_string())],
               [String("<".to_string())],
               [String(">".to_string())],
               [String("aaa ' aaa".to_string())],
               [String("\"".to_string())],
               [String("☺".to_string())],
               [String("֍".to_string())],
               [String("àâéêèçöïî«»".to_string())]]);
}

#[test]
fn richtext_namespaced() {
    let path = format!("{}/tests/richtext-namespaced.xlsx",
                       env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(range,
              [[String("inline string\r\nLine 2\r\nLine 3".to_string()),
                Empty,
                Empty,
                Empty,
                Empty,
                Empty,
                Empty,
                String("shared string\r\nLine 2\r\nLine 3".to_string())]]);
}

#[test]
fn defined_names_xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let mut defined_names = excel.defined_names().unwrap().to_vec();
    defined_names.sort();
    assert_eq!(defined_names,
               vec![("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
                    ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string())]);
}

#[test]
fn defined_names_xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let mut defined_names = excel.defined_names().unwrap().to_vec();
    defined_names.sort();
    assert_eq!(defined_names,
               vec![("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
                    ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string())]);
}

#[test]
fn defined_names_xls() {
    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Sheets::open(&path).expect("cannot open excel file");

    let mut defined_names = excel.defined_names().unwrap().to_vec();
    defined_names.sort();
    assert_eq!(defined_names,
               vec![("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
                    ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string())]);
}
