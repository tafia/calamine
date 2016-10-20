extern crate office;

use office::Excel;
use office::DataType::{self, Int, String, Float, Bool};

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
    assert_eq!(r.next(), None);
}
