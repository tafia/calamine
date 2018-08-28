extern crate calamine;

use calamine::CellErrorType::*;
use calamine::DataType::{Bool, Empty, Error, Float, String};
use calamine::{open_workbook, Ods, Reader, Xls, Xlsb, Xlsx};
use std::io::Cursor;

macro_rules! range_eq {
    ($range:expr, $right:expr) => {
        assert_eq!(
            $range.get_size(),
            ($right.len(), $right[0].len()),
            "Size mismatch"
        );
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
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );
}

#[test]
fn issue_3() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap().unwrap();
    range_eq!(range, [[Float(1.), String("a".to_string())]]);
}

#[test]
fn issue_4() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue5").unwrap().unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn issue_6() {
    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue6").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.)],
            [Float(2.)],
            [String("ab".to_string())],
            [Bool(false)]
        ]
    );
}

#[test]
fn error_file() {
    let path = format!("{}/tests/errors.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Feuil1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Error(Div0)],
            [Error(Name)],
            [Error(Value)],
            [Error(Null)],
            [Error(Ref)],
            [Error(Num)],
            [Error(NA)]
        ]
    );
}

#[test]
fn issue_9() {
    let path = format!("{}/tests/issue9.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Feuil1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("test1".to_string())],
            [String("test2 other".to_string())],
            [String("test3 aaa".to_string())],
            [String("test4".to_string())]
        ]
    );
}

#[test]
fn vba() {
    let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let mut vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.to_mut().get_module("testVBA").unwrap(),
        "Attribute VB_Name = \"testVBA\"\r\nPublic Sub test()\r\n    MsgBox \"Hello from \
         vba!\"\r\nEnd Sub\r\n"
    );
}

#[test]
fn xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );
}

#[test]
fn xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );
}

#[test]
fn xls() {
    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xls<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );
}

#[test]
fn ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("datatypes").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.)],
            [Float(1.5)],
            [String("ab".to_string())],
            [Bool(false)],
            [String("test".to_string())],
            [String("2016-10-20T00:00:00".to_string())]
        ]
    );

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );

    let range = excel.worksheet_range("issue5").unwrap().unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn ods_covered() {
    let path = format!("{}/tests/covered.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("sheet1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("a1".to_string())],
            [String("a2".to_string())],
            [String("a3".to_string())],
        ]
    );
}

#[test]
fn special_chrs_xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("spc_chrs").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("&".to_string())],
            [String("<".to_string())],
            [String(">".to_string())],
            [String("aaa ' aaa".to_string())],
            [String("\"".to_string())],
            [String("☺".to_string())],
            [String("֍".to_string())],
            [String("àâéêèçöïî«»".to_string())]
        ]
    );
}

#[test]
fn special_chrs_xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("spc_chrs").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("&".to_string())],
            [String("<".to_string())],
            [String(">".to_string())],
            [String("aaa ' aaa".to_string())],
            [String("\"".to_string())],
            [String("☺".to_string())],
            [String("֍".to_string())],
            [String("àâéêèçöïî«»".to_string())]
        ]
    );
}

#[test]
fn special_chrs_ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("spc_chrs").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("&".to_string())],
            [String("<".to_string())],
            [String(">".to_string())],
            [String("aaa ' aaa".to_string())],
            [String("\"".to_string())],
            [String("☺".to_string())],
            [String("֍".to_string())],
            [String("àâéêèçöïî«»".to_string())]
        ]
    );
}

#[test]
fn partial_richtext_ods() {
    let path = format!("{}/tests/richtext_issue.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("datatypes").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("abc".to_string())]
        ]
    );
}

#[test]
fn xlsx_richtext_namespaced() {
    let path = format!(
        "{}/tests/richtext-namespaced.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap().unwrap();
    range_eq!(
        range,
        [[
            String("inline string\r\nLine 2\r\nLine 3".to_string()),
            Empty,
            Empty,
            Empty,
            Empty,
            Empty,
            Empty,
            String("shared string\r\nLine 2\r\nLine 3".to_string())
        ]]
    );
}

#[test]
fn defined_names_xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let excel: Xlsx<_> = open_workbook(&path).unwrap();

    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        vec![
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
}

#[test]
fn defined_names_xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let excel: Xlsb<_> = open_workbook(&path).unwrap();

    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        vec![
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
}

#[test]
fn defined_names_xls() {
    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let excel: Xls<_> = open_workbook(&path).unwrap();

    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        vec![
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
}

#[test]
fn defined_names_ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let excel: Ods<_> = open_workbook(&path).unwrap();

    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        vec![
            (
                "MyBrokenRange".to_string(),
                "of:=[Sheet1.#REF!]".to_string(),
            ),
            (
                "MyDataTypes".to_string(),
                "datatypes.$A$1:datatypes.$A$6".to_string(),
            ),
            ("OneRange".to_string(), "Sheet1.$A$1".to_string()),
        ]
    );
}

#[test]
fn parse_sheet_names_in_xls() {
    let path = format!(
        "{}/tests/sheet_name_parsing.xls",
        env!("CARGO_MANIFEST_DIR")
    );
    let excel: Xls<_> = open_workbook(&path).unwrap();
    assert_eq!(excel.sheet_names(), &["Sheet1"]);
}

#[test]
fn read_xls_from_memory() {
    const DATA_XLS: &[u8] = include_bytes!("sheet_name_parsing.xls");
    let reader = Cursor::new(DATA_XLS);
    let excel = Xls::new(reader).unwrap();
    assert_eq!(excel.sheet_names(), &["Sheet1"]);
}

#[test]
fn search_references() {
    let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    let names = references.iter().map(|r| &*r.name).collect::<Vec<&str>>();
    assert_eq!(names, vec!["stdole", "Office"]);
}

#[test]
fn formula_xlsx() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap().unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap().unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_xlsb() {
    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap().unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap().unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_xls() {
    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xls<_> = open_workbook(&path).unwrap();

    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap().unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap().unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_ods() {
    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    for s in excel.sheet_names().to_owned() {
        let _ = excel.worksheet_formula(&s).unwrap().unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap().unwrap();
    range_eq!(formula, [["of:=[.B1]+$$OneRange".to_string()]]);
}

#[test]
fn empty_sheet() {
    let path = format!("{}/tests/empty_sheet.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    for s in excel.sheet_names().to_owned() {
        let range = excel.worksheet_range(&s).unwrap().unwrap();
        assert_eq!(range.start(), None, "wrong start");
        assert_eq!(range.end(), None, "wrong end");
        assert_eq!(range.get_size(), (0, 0), "wrong size");
    }
}

#[test]
fn issue_120() {
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue2").unwrap().unwrap();
    let end = range.end().unwrap();

    let a = range.get_value((0, end.1 + 1));
    assert_eq!(None, a);

    let b = range.get_value((0, 0));
    assert_eq!(Some(&Float(1.)), b);
}

#[test]
fn issue_127() {
    //Ods::sheet_names should now return sheet names in their correct order
    let path = format!("{}/tests/issue127.ods", env!("CARGO_MANIFEST_DIR"));
    let ods: Ods<_> = open_workbook(&path).unwrap();

    let ordered_names: Vec<std::string::String> = vec![
        "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6", "Sheet7", "Sheet8",
    ].iter()
        .map(|&s| s.to_owned())
        .collect();

    assert_eq!(ods.sheet_names(), &ordered_names[..]);
}
