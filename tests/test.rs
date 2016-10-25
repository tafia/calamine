extern crate office;

use office::Excel;
use office::DataType::{self, Int, String, Float, Bool, Error, Empty};
use office::CellErrorType::*;

#[test]
fn issue_2() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue2").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[Int(1), String("a".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[Int(2), String("b".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[Int(3), String("c".to_string())] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn issue_3() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Sheet1").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[Int(1), String("a".to_string())] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn issue_4() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue5").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[Float(0.5)] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn issue_6() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("issue6").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[Int(1)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Int(2)] as &[DataType]));
    assert_eq!(r.next(), Some(&[String("ab".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[Bool(false)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Empty] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn error_file() {
    let path = format!("{}/tests/errors.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Feuil1").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[Error(Div0)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(Name)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(Value)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(Null)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(Ref)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(Num)] as &[DataType]));
    assert_eq!(r.next(), Some(&[Error(NA)] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn issue_9() {
    let path = format!("{}/tests/issue9.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let range = excel.worksheet_range("Feuil1").unwrap();
    let mut r = range.rows();
    assert_eq!(r.next(), Some(&[String("test1".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[String("test2 other".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[String("test3 aaa".to_string())] as &[DataType]));
    assert_eq!(r.next(), Some(&[String("test4".to_string())] as &[DataType]));
    assert_eq!(r.next(), None);
}

#[test]
fn vba() {
    let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel = Excel::open(&path).expect("cannot open excel file");

    let vba = excel.vba_project().unwrap();
    let modules = vba.read_vba().unwrap().1;
    let test_vba = modules.into_iter().find(|m| &*m.name == "testVBA").unwrap();
    assert_eq!(vba.read_module(&test_vba).unwrap(), "Attribute VB_Name = \"testVBA\"\
    \r\nPublic Sub test()\
    \r\n    MsgBox \"Hello from vba!\"\
    \r\nEnd Sub\
    \r\n");
}
