use calamine::CellErrorType::*;
use calamine::DataType::{Bool, DateTime, Empty, Error, Float, String};
use calamine::{open_workbook, open_workbook_auto, Ods, Reader, Xls, Xlsb, Xlsx};
use std::io::Cursor;
use std::sync::Once;

static INIT: Once = Once::new();

/// Setup function that is only run once, even if called multiple times.
fn setup() {
    INIT.call_once(|| {
        env_logger::init();
    });
}

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
    setup();

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
    setup();

    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap().unwrap();
    range_eq!(range, [[Float(1.), String("a".to_string())]]);
}

#[test]
fn issue_4() {
    setup();

    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue5").unwrap().unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn issue_6() {
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

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

// test ignored because the file is too large to be committed and tested
#[ignore]
#[test]
fn issue_195() {
    setup();

    let path = format!(
        "{}/JLCPCB SMT Parts Library(20210204).xls",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut excel: Xls<_> = open_workbook(&path).expect("can't open wb");
    let range = excel
        .worksheet_range("JLCPCB SMT Parts Library")
        .expect("error in wks range")
        .expect("sheet not found");
    assert_eq!(range.get_size(), (52046, 12));
}

#[test]
fn ods() {
    setup();

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
    setup();

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
fn special_cells() {
    let path = format!("{}/tests/special_cells.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("sheet1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("Split\nLine".to_string())],
            [String("Value With spaces".to_string())],
        ]
    );
}

#[test]
fn special_chrs_xlsx() {
    setup();

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
    setup();

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
    setup();

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
    setup();

    let path = format!("{}/tests/richtext_issue.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("datatypes").unwrap().unwrap();
    range_eq!(range, [[String("abc".to_string())]]);
}

#[test]
fn xlsx_richtext_namespaced() {
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

    let path = format!(
        "{}/tests/sheet_name_parsing.xls",
        env!("CARGO_MANIFEST_DIR")
    );
    let excel: Xls<_> = open_workbook(&path).unwrap();
    assert_eq!(excel.sheet_names(), &["Sheet1"]);
}

#[test]
fn read_xls_from_memory() {
    setup();

    const DATA_XLS: &[u8] = include_bytes!("sheet_name_parsing.xls");
    let reader = Cursor::new(DATA_XLS);
    let excel = Xls::new(reader).unwrap();
    assert_eq!(excel.sheet_names(), &["Sheet1"]);
}

#[test]
fn search_references() {
    setup();

    let path = format!("{}/tests/vba.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    let names = references.iter().map(|r| &*r.name).collect::<Vec<&str>>();
    assert_eq!(names, vec!["stdole", "Office"]);
}

#[test]
fn formula_xlsx() {
    setup();

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
    setup();

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
fn formula_vals_xlsb() {
    setup();

    let path = format!("{}/tests/issue_182.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("formula_vals").unwrap().unwrap();
    range_eq!(
        range,
        [[Float(3.)], [String("Ab".to_string())], [Bool(false)]]
    );
}

#[test]
fn float_vals_xlsb() {
    setup();

    let path = format!("{}/tests/issue_186.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [Float(1.23)],
            [Float(12.34)],
            [Float(123.45)],
            [Float(1234.56)],
            [Float(12345.67)],
        ]
    );
}

#[test]
fn formula_xls() {
    setup();

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
    setup();

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
    setup();

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
    setup();

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
    setup();

    let root = env!("CARGO_MANIFEST_DIR");
    let ordered_names: Vec<std::string::String> = vec![
        "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6", "Sheet7", "Sheet8",
    ]
    .iter()
    .map(|&s| s.to_owned())
    .collect();

    for ext in &["ods", "xls", "xlsx", "xlsb"] {
        let p = format!("{}/tests/issue127.{}", root, ext);
        let workbook = open_workbook_auto(&p).expect(&p);
        assert_eq!(
            workbook.sheet_names(),
            &ordered_names[..],
            "{} sheets should be ordered",
            ext
        );
    }
}

#[test]
fn mul_rk() {
    setup();

    let path = format!(
        "{}/tests/adhocallbabynames1996to2016.xls",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut xls: Xls<_> = open_workbook(&path).unwrap();
    let range = xls.worksheet_range("Boys").unwrap().unwrap();
    assert_eq!(range.get_value((6, 2)), Some(&Float(9.)));
}

#[test]
fn skip_phonetic_text() {
    setup();

    let path = format!("{}/tests/rph.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
    let range = xls.worksheet_range("Sheet1").unwrap().unwrap();
    assert_eq!(
        range.get_value((0, 0)),
        Some(&String("課きく　毛こ".to_string()))
    );
}

#[test]
fn issue_174() {
    setup();

    let path = format!("{}/tests/issue_174.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
    xls.worksheet_range_at(0).unwrap().unwrap();
}

#[test]
fn table() {
    setup();
    let path = format!(
        "{}/tests/temperature-table.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
    xls.load_tables().unwrap();
    let table_names = xls.table_names();
    assert_eq!(table_names[0], "Temperature");
    assert_eq!(table_names[1], "OtherTable");
    let table = xls
        .table_by_name("Temperature")
        .expect("Table should exist")
        .expect("Parsing table's sheet should not error");
    assert_eq!(table.name(), "Temperature");
    assert_eq!(table.columns()[0], "label");
    assert_eq!(table.columns()[1], "value");
    let data = table.data();
    assert_eq!(data.get((0, 0)), Some(&String("celsius".to_owned())));
    assert_eq!(data.get((1, 0)), Some(&String("fahrenheit".to_owned())));
    assert_eq!(data.get((0, 1)), Some(&Float(22.2222)));
    assert_eq!(data.get((1, 1)), Some(&Float(72.0)));
    // Check the second table
    let table = xls
        .table_by_name("OtherTable")
        .expect("Table should exist")
        .expect("Parsing table's sheet should not error");
    assert_eq!(table.name(), "OtherTable");
    assert_eq!(table.columns()[0], "label2");
    assert_eq!(table.columns()[1], "value2");
    let data = table.data();
    assert_eq!(data.get((0, 0)), Some(&String("something".to_owned())));
    assert_eq!(data.get((1, 0)), Some(&String("else".to_owned())));
    assert_eq!(data.get((0, 1)), Some(&Float(12.5)));
    assert_eq!(data.get((1, 1)), Some(&Float(64.0)));
    xls.worksheet_range_at(0).unwrap().unwrap();
}

#[test]
fn date() {
    setup();

    let path = format!("{}/tests/date.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(range.get_value((0, 0)), Some(&DateTime(44197.0)));

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd(2021, 01, 01);
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));
    }
}

#[test]
fn issue_219() {
    setup();

    let path = format!("{}/tests/issue219.xls", env!("CARGO_MANIFEST_DIR"));

    // should not panic
    let _: Xls<_> = open_workbook(&path).unwrap();
}

#[test]
fn issue_221() {
    setup();

    let path = format!("{}/tests/issue221.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("Cell_A1".to_string()), String("Cell_B1".to_string())],
            [String("Cell_A2".to_string()), String("Cell_B2".to_string())]
        ]
    );
}
