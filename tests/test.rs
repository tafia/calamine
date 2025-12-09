// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use calamine::vba::Reference;
use calamine::Data::{Bool, DateTime, DateTimeIso, DurationIso, Empty, Error, Float, Int, String};
use calamine::{
    open_workbook, open_workbook_auto, DataRef, DataType, Dimensions, ExcelDateTime,
    ExcelDateTimeType, HeaderRow, Ods, Range, Reader, ReaderRef, Sheet, SheetType, SheetVisible,
    Xls, Xlsb, Xlsx,
};
use calamine::{CellErrorType::*, Data};
use rstest::rstest;
use std::collections::BTreeSet;
use std::fs::File;
use std::io::{BufReader, Cursor};
use std::sync::Once;

static INIT: Once = Once::new();

fn test_path(name: &str) -> std::string::String {
    format!("tests/{name}")
}

/// Setup function that is only run once, even if called multiple times.
fn wb<R: Reader<BufReader<File>>>(name: &str) -> R {
    INIT.call_once(|| {
        env_logger::init();
    });
    let path = test_path(name);
    open_workbook(&path).expect(&path)
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
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let range = excel.worksheet_range("issue2").unwrap();
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
    let mut excel: Xlsx<_> = wb("issue3.xlsm");
    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(range, [[Float(1.), String("a".to_string())]]);
}

#[test]
fn issue_4() {
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let range = excel.worksheet_range("issue5").unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn issue_6() {
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let range = excel.worksheet_range("issue6").unwrap();
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
    let mut excel: Xlsx<_> = wb("errors.xlsx");
    let range = excel.worksheet_range("Feuil1").unwrap();
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
    let mut excel: Xlsx<_> = wb("issue9.xlsx");
    let range = excel.worksheet_range("Feuil1").unwrap();
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
    let mut excel: Xlsx<_> = wb("vba.xlsm");
    let vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.get_module("testVBA").unwrap(),
        "Attribute VB_Name = \"testVBA\"\r\nPublic Sub test()\r\n    MsgBox \"Hello from \
         vba!\"\r\nEnd Sub\r\n"
    );
}

#[test]
fn xlsb() {
    let mut excel: Xlsb<_> = wb("issues.xlsb");
    let range = excel.worksheet_range("issue2").unwrap();
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
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let range = excel.worksheet_range("issue2").unwrap();
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
    let mut excel: Xls<_> = wb("issues.xls");
    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    assert_eq!(
        references,
        [
            Reference {
                name: "stdole".to_string(),
                description: "OLE Automation".to_string(),
                path: "C:\\Windows\\SysWOW64\\stdole2.tlb".into(),
            },
            Reference {
                name: "Office".to_string(),
                description: "Microsoft Office 16.0 Object Library".to_string(),
                path: "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL"
                    .into(),
            },
        ]
    );
    assert_eq!(
        vba.get_module_names(),
        [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet4",
            "ThisWorkbook",
            "testVBA",
        ]
    );
}

// test ignored because the file is too large to be committed and tested
#[ignore]
#[test]
fn issue_195() {
    let mut excel: Xls<_> = wb("JLCPCB SMT Parts Library(20210204).xls");
    let range = excel
        .worksheet_range("JLCPCB SMT Parts Library")
        .expect("error in wks range");
    assert_eq!(range.get_size(), (52046, 12));
}

#[test]
fn ods() {
    let mut excel: Ods<_> = wb("issues.ods");
    let range = excel.worksheet_range("datatypes").unwrap();
    range_eq!(
        range,
        [
            [Float(1.)],
            [Float(1.5)],
            [String("ab".to_string())],
            [Bool(false)],
            [String("test".to_string())],
            [DateTimeIso("2016-10-20T00:00:00".to_string())]
        ]
    );

    let range = excel.worksheet_range("issue2").unwrap();
    range_eq!(
        range,
        [
            [Float(1.), String("a".to_string())],
            [Float(2.), String("b".to_string())],
            [Float(3.), String("c".to_string())]
        ]
    );

    let range = excel.worksheet_range("issue5").unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn ods_covered() {
    let mut excel: Ods<_> = wb("covered.ods");
    let range = excel.worksheet_range("sheet1").unwrap();
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
    let mut excel: Ods<_> = wb("special_cells.ods");
    let range = excel.worksheet_range("sheet1").unwrap();
    range_eq!(
        range,
        [
            [String("Split\nLine".to_string())],
            [String("Value  With spaces".to_string())],
            [String("Value   With 3 spaces".to_string())],
            [String(" Value   With spaces before and after ".to_string())],
            [String(
                "  Value   With 2 spaces before and after  ".to_string()
            )],
        ]
    );
}

#[test]
fn special_chrs_xlsx() {
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let range = excel.worksheet_range("spc_chrs").unwrap();
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
    let mut excel: Xlsb<_> = wb("issues.xlsb");
    let range = excel.worksheet_range("spc_chrs").unwrap();
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
    let mut excel: Ods<_> = wb("issues.ods");
    let range = excel.worksheet_range("spc_chrs").unwrap();
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

// Test for decoding/unescaping simple XML entities and also for numeric
// entities like "&#xA;" for new line.
#[test]
fn decode_xml_entities() {
    let mut excel: Xlsx<_> = wb("encoded_entities.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();

    assert_eq!(range.get_value((0, 0)), Some(&String("&".to_string())));
    assert_eq!(range.get_value((1, 0)), Some(&String("\n".to_string())));
}

// Test for unescaping Excel XML escapes in a cell string. Excel encodes a
// character like "\r" as "_x000D_". In turn it escapes the literal string
// "_x000D_" as "_x005F_x000D_".
//
// See https://github.com/tafia/calamine/issues/469
#[test]
fn unescape_excel_xml() {
    let mut excel: Xlsx<_> = wb("has_x000D_.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&String("ABC\r\nDEF".to_string()))
    );

    // Test a file with an inline string.
    let mut excel: Xlsx<_> = wb("has_x000D_inline.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&String("ABC\r\nDEF".to_string()))
    );
}

#[test]
fn partial_richtext_ods() {
    let mut excel: Ods<_> = wb("richtext_issue.ods");
    let range = excel.worksheet_range("datatypes").unwrap();
    range_eq!(range, [[String("abc".to_string())]]);
}

#[test]
fn xlsx_richtext_namespaced() {
    let mut excel: Xlsx<_> = wb("richtext-namespaced.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(
        range,
        [[
            String("inline string\nLine 2\nLine 3".to_string()),
            Empty,
            Empty,
            Empty,
            Empty,
            Empty,
            Empty,
            String("shared string\nLine 2\nLine 3".to_string())
        ]]
    );
}

#[test]
fn defined_names_xlsx() {
    let excel: Xlsx<_> = wb("issues.xlsx");
    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        [
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
}

#[test]
fn defined_names_xlsb() {
    let excel: Xlsb<_> = wb("issues.xlsb");
    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        [
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
}

#[test]
fn defined_names_xls() {
    let mut excel: Xls<_> = wb("issues.xls");
    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        [
            ("MyBrokenRange".to_string(), "Sheet1!#REF!".to_string()),
            ("MyDataTypes".to_string(), "datatypes!$A$1:$A$6".to_string()),
            ("OneRange".to_string(), "Sheet1!$A$1".to_string()),
        ]
    );
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    assert_eq!(
        references,
        [
            Reference {
                name: "stdole".to_string(),
                description: "OLE Automation".to_string(),
                path: "C:\\Windows\\SysWOW64\\stdole2.tlb".into(),
            },
            Reference {
                name: "Office".to_string(),
                description: "Microsoft Office 16.0 Object Library".to_string(),
                path: "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL"
                    .into(),
            },
        ]
    );
    assert_eq!(
        vba.get_module_names(),
        [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet4",
            "ThisWorkbook",
            "testVBA",
        ],
    );
}

#[test]
fn defined_names_ods() {
    let excel: Ods<_> = wb("issues.ods");
    let mut defined_names = excel.defined_names().to_vec();
    defined_names.sort();
    assert_eq!(
        defined_names,
        [
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
    let excel: Xls<_> = wb("sheet_name_parsing.xls");
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
    let mut excel: Xlsx<_> = wb("vba.xlsm");
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    let names = references.iter().map(|r| &*r.name).collect::<Vec<&str>>();
    assert_eq!(names, ["stdole", "Office"]);
}

#[test]
fn formula_xlsx() {
    let mut excel: Xlsx<_> = wb("issues.xlsx");
    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_xlsb() {
    let mut excel: Xlsb<_> = wb("issues.xlsb");
    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_vals_xlsb() {
    let mut excel: Xlsb<_> = wb("issue_182.xlsb");
    let range = excel.worksheet_range("formula_vals").unwrap();
    range_eq!(
        range,
        [[Float(3.)], [String("Ab".to_string())], [Bool(false)]]
    );
}

#[test]
fn float_vals_xlsb() {
    let mut excel: Xlsb<_> = wb("issue_186.xlsb");
    let range = excel.worksheet_range("Sheet1").unwrap();
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
    let mut excel: Xls<_> = wb("issues.xls");
    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
    let vba = excel.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    assert_eq!(
        references,
        [
            Reference {
                name: "stdole".to_string(),
                description: "OLE Automation".to_string(),
                path: "C:\\Windows\\SysWOW64\\stdole2.tlb".into(),
            },
            Reference {
                name: "Office".to_string(),
                description: "Microsoft Office 16.0 Object Library".to_string(),
                path: "C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL"
                    .into(),
            },
        ]
    );
    assert_eq!(
        vba.get_module_names(),
        [
            "Sheet1",
            "Sheet2",
            "Sheet3",
            "Sheet4",
            "ThisWorkbook",
            "testVBA",
        ]
    );
}

#[test]
fn formula_ods() {
    let mut excel: Ods<_> = wb("issues.ods");
    for s in excel.sheet_names() {
        let _ = excel.worksheet_formula(&s).unwrap();
    }
    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["of:=[.B1]+$$OneRange".to_string()]]);
}

#[test]
fn empty_sheet() {
    let mut excel: Xlsx<_> = wb("empty_sheet.xlsx");
    for s in excel.sheet_names() {
        let range = excel.worksheet_range(&s).unwrap();
        assert_eq!(range.start(), None, "wrong start");
        assert_eq!(range.end(), None, "wrong end");
        assert_eq!(range.get_size(), (0, 0), "wrong size");
    }
}

#[test]
fn issue_120() {
    let mut excel: Xlsx<_> = wb("issues.xlsx");

    let range = excel.worksheet_range("issue2").unwrap();
    let end = range.end().unwrap();

    let a = range.get_value((0, end.1 + 1));
    assert_eq!(None, a);

    let b = range.get_value((0, 0));
    assert_eq!(Some(&Float(1.)), b);
}

#[test]
fn issue_127() {
    let ordered_names: Vec<_> = [
        "Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6", "Sheet7", "Sheet8",
    ]
    .iter()
    .map(|&s| s.to_string())
    .collect();

    for ext in &["ods", "xls", "xlsx", "xlsb"] {
        let p = test_path(&format!("issue127.{ext}"));
        let workbook = open_workbook_auto(&p).expect(&p);
        assert_eq!(
            workbook.sheet_names(),
            &ordered_names[..],
            "{ext} sheets should be ordered"
        );
    }
}

#[test]
fn mul_rk() {
    let mut xls: Xls<_> = wb("adhocallbabynames1996to2016.xls");
    let range = xls.worksheet_range("Boys").unwrap();
    assert_eq!(range.get_value((6, 2)), Some(&Float(9.)));
}

#[test]
fn skip_phonetic_text() {
    let mut xls: Xlsx<_> = wb("rph.xlsx");
    let range = xls.worksheet_range("Sheet1").unwrap();
    assert_eq!(
        range.get_value((0, 0)),
        Some(&String("課きく　毛こ".to_string()))
    );
}

#[test]
fn issue_174() {
    let mut xls: Xlsx<_> = wb("issue_174.xlsx");
    xls.worksheet_range_at(0).unwrap().unwrap();
}

#[test]
fn table() {
    let mut xls: Xlsx<_> = wb("temperature-table.xlsx");
    xls.load_tables().unwrap();
    let table_names = xls.table_names();
    assert_eq!(table_names[0], "Temperature");
    assert_eq!(table_names[1], "OtherTable");
    let table = xls
        .table_by_name("Temperature")
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

    // Check if owned data works
    let owned_data: Range<Data> = table.into();

    assert_eq!(
        owned_data.get((0, 0)),
        Some(&String("something".to_owned()))
    );
    assert_eq!(owned_data.get((1, 0)), Some(&String("else".to_owned())));
    assert_eq!(owned_data.get((0, 1)), Some(&Float(12.5)));
    assert_eq!(owned_data.get((1, 1)), Some(&Float(64.0)));
}

#[test]
fn table_by_ref() {
    let mut xls: Xlsx<_> = wb("temperature-table.xlsx");
    xls.load_tables().unwrap();
    let table_names = xls.table_names();
    assert_eq!(table_names[0], "Temperature");
    assert_eq!(table_names[1], "OtherTable");
    let table = xls
        .table_by_name_ref("Temperature")
        .expect("Parsing table's sheet should not error");
    assert_eq!(table.name(), "Temperature");
    assert_eq!(table.columns()[0], "label");
    assert_eq!(table.columns()[1], "value");
    let data = table.data();
    assert_eq!(
        data.get((0, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("celsius")
    );
    assert_eq!(
        data.get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("fahrenheit")
    );
    assert_eq!(
        data.get((0, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(22.2222)
    );
    assert_eq!(
        data.get((1, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(72.0)
    );
    // Check the second table
    let table = xls
        .table_by_name_ref("OtherTable")
        .expect("Parsing table's sheet should not error");
    assert_eq!(table.name(), "OtherTable");
    assert_eq!(table.columns()[0], "label2");
    assert_eq!(table.columns()[1], "value2");
    let data = table.data();
    assert_eq!(
        data.get((0, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("something")
    );
    assert_eq!(
        data.get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("else")
    );
    assert_eq!(
        data.get((0, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(12.5)
    );
    assert_eq!(
        data.get((1, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(64.0)
    );

    // Check if owned data works
    let owned_data: Range<DataRef> = table.into();

    assert_eq!(
        owned_data
            .get((0, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("something")
    );
    assert_eq!(
        owned_data
            .get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString("else")
    );
    assert_eq!(
        owned_data
            .get((0, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(12.5)
    );
    assert_eq!(
        owned_data
            .get((1, 1))
            .expect("Could not get data from table ref."),
        &DataRef::Float(64.0)
    );
}

#[test]
fn date_xls() {
    let mut xls: Xls<_> = wb("date.xls");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.632060185185185,
            ExcelDateTimeType::TimeDelta,
            false
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xls_1904() {
    let mut xls: Xls<_> = wb("date_1904.xls");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.632060185185185,
            ExcelDateTimeType::TimeDelta,
            true
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xlsx() {
    let mut xls: Xlsx<_> = wb("date.xlsx");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            false
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xlsx_1904() {
    let mut xls: Xlsx<_> = wb("date_1904.xlsx");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            true
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xlsx_iso() {
    let mut xls: Xlsx<_> = wb("date_iso.xlsx");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTimeIso("2021-01-01".to_string()))
    );
    assert_eq!(
        range.get_value((1, 0)),
        Some(&DateTimeIso("2021-01-01T10:10:10".to_string()))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTimeIso("10:10:10".to_string()))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));
        assert_eq!(range.get_value((0, 0)).unwrap().as_time(), None);
        assert_eq!(range.get_value((0, 0)).unwrap().as_datetime(), None);

        let time = chrono::NaiveTime::from_hms_opt(10, 10, 10).unwrap();
        assert_eq!(range.get_value((2, 0)).unwrap().as_time(), Some(time));
        assert_eq!(range.get_value((2, 0)).unwrap().as_date(), None);
        assert_eq!(range.get_value((2, 0)).unwrap().as_datetime(), None);

        let datetime = chrono::NaiveDateTime::new(date, time);
        assert_eq!(
            range.get_value((1, 0)).unwrap().as_datetime(),
            Some(datetime)
        );
        assert_eq!(range.get_value((1, 0)).unwrap().as_time(), Some(time));
        assert_eq!(range.get_value((1, 0)).unwrap().as_date(), Some(date));
    }
}

#[test]
fn date_ods() {
    let mut ods: Ods<_> = wb("date.ods");
    let range = ods.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTimeIso("2021-01-01".to_string()))
    );
    assert_eq!(
        range.get_value((1, 0)),
        Some(&DateTimeIso("2021-01-01T10:10:10".to_string()))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DurationIso("PT10H10M10S".to_string()))
    );
    assert_eq!(
        range.get_value((3, 0)),
        Some(&DurationIso("PT10H10M10.123456S".to_string()))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let time = chrono::NaiveTime::from_hms_opt(10, 10, 10).unwrap();
        assert_eq!(range.get_value((2, 0)).unwrap().as_time(), Some(time));

        let datetime = chrono::NaiveDateTime::new(date, time);
        assert_eq!(
            range.get_value((1, 0)).unwrap().as_datetime(),
            Some(datetime)
        );

        let time = chrono::NaiveTime::from_hms_micro_opt(10, 10, 10, 123456).unwrap();
        assert_eq!(range.get_value((3, 0)).unwrap().as_time(), Some(time));

        let duration =
            chrono::Duration::microseconds((10 * 60 * 60 + 10 * 60 + 10) * 1_000_000 + 123456);
        assert_eq!(
            range.get_value((3, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xlsb() {
    let mut xls: Xlsb<_> = wb("date.xlsb");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            false
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn date_xlsb_1904() {
    let mut xls: Xlsb<_> = wb("date_1904.xlsb");
    let range = xls.worksheet_range_at(0).unwrap().unwrap();

    assert_eq!(
        range.get_value((0, 0)),
        Some(&DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        )))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            true
        )))
    );

    #[cfg(feature = "chrono")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(range.get_value((0, 0)).unwrap().as_date(), Some(date));

        let duration = chrono::Duration::seconds(255 * 60 * 60 + 10 * 60 + 10);
        assert_eq!(
            range.get_value((2, 0)).unwrap().as_duration(),
            Some(duration)
        );
    }
}

#[test]
fn issue_219() {
    // should not panic
    let _: Xls<_> = wb("issue219.xls");
}

#[test]
fn issue_221() {
    let mut excel: Xlsx<_> = wb("issue221.xlsm");

    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(
        range,
        [
            [String("Cell_A1".to_string()), String("Cell_B1".to_string())],
            [String("Cell_A2".to_string()), String("Cell_B2".to_string())]
        ]
    );
}

#[test]
fn merged_regions_xlsx() {
    use calamine::Dimensions;
    use std::string::String;
    let mut excel: Xlsx<_> = wb("merged_range.xlsx");
    excel.load_merged_regions().unwrap();
    assert_eq!(
        excel
            .merged_regions()
            .iter()
            .map(|(o1, o2, o3)| (o1.to_string(), o2.to_string(), *o3))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        [
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 0), (1, 0))
            ), // A1:A2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 1), (1, 1))
            ), // B1:B2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 2), (1, 3))
            ), // C1:D2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((2, 2), (2, 3))
            ), // C3:D3
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((3, 2), (3, 3))
            ), // C4:D4
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 4), (1, 4))
            ), // E1:E2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 5), (1, 5))
            ), // F1:F2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 6), (1, 6))
            ), // G1:G2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 7), (1, 7))
            ), // H1:H2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 0), (3, 0))
            ), // A1:A4
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 1), (1, 1))
            ), // B1:B2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 2), (1, 3))
            ), // C1:D2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((2, 2), (3, 3))
            ), // C3:D4
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 4), (1, 4))
            ), // E1:E2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 5), (3, 7))
            ), // F1:H4
        ]
        .into_iter()
        .collect::<BTreeSet<(String, String, Dimensions)>>(),
    );
    assert_eq!(
        excel
            .merged_regions_by_sheet("Sheet1")
            .iter()
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), *o3))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        [
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 0), (1, 0))
            ), // A1:A2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 1), (1, 1))
            ), // B1:B2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 2), (1, 3))
            ), // C1:D2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((2, 2), (2, 3))
            ), // C3:D3
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((3, 2), (3, 3))
            ), // C4:D4
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 4), (1, 4))
            ), // E1:E2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 5), (1, 5))
            ), // F1:F2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 6), (1, 6))
            ), // G1:G2
            (
                "Sheet1".to_string(),
                "xl/worksheets/sheet1.xml".to_string(),
                Dimensions::new((0, 7), (1, 7))
            ), // H1:H2
        ]
        .into_iter()
        .collect::<BTreeSet<(String, String, Dimensions)>>(),
    );
    assert_eq!(
        excel
            .merged_regions_by_sheet("Sheet2")
            .iter()
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), *o3))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        [
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 0), (3, 0))
            ), // A1:A4
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 1), (1, 1))
            ), // B1:B2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 2), (1, 3))
            ), // C1:D2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((2, 2), (3, 3))
            ), // C3:D4
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 4), (1, 4))
            ), // E1:E2
            (
                "Sheet2".to_string(),
                "xl/worksheets/sheet2.xml".to_string(),
                Dimensions::new((0, 5), (3, 7))
            ), // F1:H4
        ]
        .into_iter()
        .collect::<BTreeSet<(String, String, Dimensions)>>(),
    );
}

#[test]
fn issue_252() {
    let path = "issue252.xlsx";

    // should err, not panic
    assert!(open_workbook::<Xls<_>, _>(&path).is_err());
}

#[test]
fn issue_261() {
    let mut workbook_with_missing_r_attributes: Xlsx<_> = wb("issue_261.xlsx");
    let mut workbook_fixed_by_excel: Xlsx<_> = wb("issue_261_fixed_by_excel.xlsx");

    let range_a = workbook_fixed_by_excel
        .worksheet_range("Some Sheet")
        .unwrap();

    let range_b = workbook_with_missing_r_attributes
        .worksheet_range("Some Sheet")
        .unwrap();

    assert_eq!(range_a.cells().count(), 462);
    assert_eq!(range_a.cells().count(), 462);
    assert_eq!(range_a.rows().count(), 66);
    assert_eq!(range_b.rows().count(), 66);

    assert_eq!(
        range_b.get_value((0, 0)).unwrap(),
        &String("String Value 32".into())
    );
    range_b
        .rows()
        .nth(4)
        .unwrap()
        .iter()
        .for_each(|cell| assert!(cell.is_empty()));

    assert_eq!(range_b.get_value((60, 6)).unwrap(), &Float(939.));
    assert_eq!(
        range_b.get_value((65, 0)).unwrap(),
        &String("String Value 42".into())
    );

    assert_eq!(
        range_b.get_value((65, 3)).unwrap(),
        &String("String Value 8".into())
    );

    range_a
        .rows()
        .zip(range_b.rows().filter(|r| !r.is_empty()))
        .enumerate()
        .for_each(|(i, (lhs, rhs))| {
            assert_eq!(
                lhs,
                rhs,
                "Expected row {} to be {:?}, but found {:?}",
                i + 1,
                lhs,
                rhs
            )
        });
}

#[test]
fn test_values_xls() {
    let mut excel: Xls<_> = wb("xls_wrong_decimals.xls");
    let range = excel
        .worksheet_range_at(0)
        .unwrap()
        .unwrap()
        .range((0, 0), (0, 0));
    range_eq!(range, [[0.525625],]);
}

#[test]
fn issue_271() -> Result<(), calamine::Error> {
    let mut count = 0;
    let mut values = Vec::new();
    loop {
        let mut workbook: Xls<_> = wb("issue_271.xls");
        let v = workbook.worksheets();
        let (sheetname, range) = v.first().expect("bad format");
        assert_eq!(sheetname, "sheet1");
        let value = range.get((0, 1)).map(|s| s.to_string());
        values.push(value);
        count += 1;
        if count > 20 {
            break;
        }
    }

    values.sort_unstable();
    values.dedup();

    assert_eq!(&values, &[Some("yyy_name".to_string())]);

    Ok(())
}

#[test]
fn issue_305_merge_cells() {
    let mut excel: Xlsx<_> = wb("merge_cells.xlsx");
    let merge_cells = excel.worksheet_merge_cells_at(0).unwrap().unwrap();

    assert_eq!(
        merge_cells,
        [
            Dimensions::new((0, 0), (0, 1)),
            Dimensions::new((1, 0), (3, 0)),
            Dimensions::new((1, 1), (3, 3))
        ]
    );
}

#[test]
fn issue_305_merge_cells_xls() {
    let excel: Xls<_> = wb("merge_cells.xls");
    let merge_cells = excel.worksheet_merge_cells_at(0).unwrap();

    assert_eq!(
        merge_cells,
        [
            Dimensions::new((0, 0), (0, 1)),
            Dimensions::new((1, 0), (3, 0)),
            Dimensions::new((1, 1), (3, 3))
        ]
    );
}

#[cfg(feature = "picture")]
fn digest(data: &[u8]) -> [u8; 32] {
    use sha2::digest::Digest;
    let mut hasher = sha2::Sha256::new();
    hasher.update(data);
    hasher.finalize().into()
}

// cargo test --features picture
#[test]
#[cfg(feature = "picture")]
fn pictures() -> Result<(), calamine::Error> {
    let jpg_path = test_path("picture.jpg");
    let png_path = test_path("picture.png");

    let xlsx_path = "picture.xlsx";
    let xlsb_path = "picture.xlsb";
    let xls_path = "picture.xls";
    let ods_path = "picture.ods";

    let jpg_hash = digest(&std::fs::read(jpg_path)?);
    let png_hash = digest(&std::fs::read(png_path)?);

    let xlsx: Xlsx<_> = wb(xlsx_path);
    let xlsb: Xlsb<_> = wb(xlsb_path);
    let xls: Xls<_> = wb(xls_path);
    let ods: Ods<_> = wb(ods_path);

    let mut pictures = Vec::with_capacity(8);
    let mut pass = 0;

    if let Some(pics) = xlsx.pictures() {
        pictures.extend(pics);
    }
    if let Some(pics) = xlsb.pictures() {
        pictures.extend(pics);
    }
    if let Some(pics) = xls.pictures() {
        pictures.extend(pics);
    }
    if let Some(pics) = ods.pictures() {
        pictures.extend(pics);
    }
    for (ext, data) in pictures {
        let pic_hash = digest(&data);
        if ext == "jpg" || ext == "jpeg" {
            assert_eq!(jpg_hash, pic_hash);
        } else if ext == "png" {
            assert_eq!(png_hash, pic_hash);
        }
        pass += 1;
    }
    assert_eq!(pass, 8);

    Ok(())
}

#[test]
fn ods_merged_cells() {
    let mut ods: Ods<_> = wb("merged_cells.ods");
    let range = ods.worksheet_range_at(0).unwrap().unwrap();

    range_eq!(
        range,
        [
            [
                String("A".to_string()),
                String("B".to_string()),
                String("C".to_string())
            ],
            [
                String("A".to_string()),
                String("B".to_string()),
                String("C".to_string())
            ],
            [Empty, Empty, String("C".to_string())],
        ]
    );
}

#[test]
fn ods_number_rows_repeated() {
    let mut ods: Ods<_> = wb("number_rows_repeated.ods");
    let test_cropped_range = [
        [String("A".to_string()), String("B".to_string())],
        [String("C".to_string()), String("D".to_string())],
        [String("C".to_string()), String("D".to_string())],
        [Empty, Empty],
        [Empty, Empty],
        [String("C".to_string()), String("D".to_string())],
        [Empty, Empty],
        [String("C".to_string()), String("D".to_string())],
    ];

    let range = ods.worksheet_range_at(0).unwrap().unwrap();
    range_eq!(range, test_cropped_range);

    let range = range.range((0, 0), range.end().unwrap());
    range_eq!(
        range,
        [
            [String("A".to_string()), String("B".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
        ]
    );

    let range = ods.worksheet_range_at(1).unwrap().unwrap();
    range_eq!(range, test_cropped_range);

    let range = range.range((0, 0), range.end().unwrap());
    range_eq!(
        range,
        [
            [Empty, Empty],
            [String("A".to_string()), String("B".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
        ]
    );

    let range = ods.worksheet_range_at(2).unwrap().unwrap();
    range_eq!(range, test_cropped_range);

    let range = range.range((0, 0), range.end().unwrap());

    range_eq!(
        range,
        [
            [Empty, Empty],
            [Empty, Empty],
            [String("A".to_string()), String("B".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
            [Empty, Empty],
            [String("C".to_string()), String("D".to_string())],
        ]
    );
}

#[test]
fn issue304_xls_formula() {
    let mut wb: Xls<_> = wb("xls_formula.xls");
    let formula = wb.worksheet_formula("Sheet1").unwrap();
    let mut rows = formula.rows();
    assert_eq!(rows.next(), Some(&["A1*2".to_owned()][..]));
    assert_eq!(rows.next(), Some(&["2*Sheet2!A1".to_owned()][..]));
    assert_eq!(rows.next(), Some(&["A1+Sheet2!A1".to_owned()][..]));
    assert_eq!(rows.next(), None);
}

#[test]
fn issue304_xls_values() {
    let mut wb: Xls<_> = wb("xls_formula.xls");
    let rge = wb.worksheet_range("Sheet1").unwrap();
    let mut rows = rge.rows();
    assert_eq!(rows.next(), Some(&[Data::Float(10.)][..]));
    assert_eq!(rows.next(), Some(&[Data::Float(20.)][..]));
    assert_eq!(rows.next(), Some(&[Data::Float(110.)][..]));
    assert_eq!(rows.next(), Some(&[Data::Float(65.)][..]));
    assert_eq!(rows.next(), None);
}

#[test]
fn issue334_xls_values_string() {
    let mut wb: Xls<_> = wb("xls_ref_String.xls");
    let rge = wb.worksheet_range("Sheet1").unwrap();
    let mut rows = rge.rows();
    assert_eq!(rows.next(), Some(&[Data::String("aa".into())][..]));
    assert_eq!(rows.next(), Some(&[Data::String("bb".into())][..]));
    assert_eq!(rows.next(), Some(&[Data::String("aa".into())][..]));
    assert_eq!(rows.next(), Some(&[Data::String("bb".into())][..]));
    assert_eq!(rows.next(), None);
}

#[test]
fn issue281_vba() {
    let mut excel: Xlsx<_> = wb("issue281.xlsm");

    let vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.get_module("testVBA").unwrap(),
        "Attribute VB_Name = \"testVBA\"\r\nPublic Sub test()\r\n    MsgBox \"Hello from \
         vba!\"\r\nEnd Sub\r\n"
    );
}

#[test]
fn issue343() {
    // should not panic
    let _: Xls<_> = wb("issue343.xls");
}

#[test]
fn any_sheets_xlsx() {
    let workbook: Xlsx<_> = wb("any_sheets.xlsx");

    assert_eq!(
        workbook.sheets_metadata(),
        &[
            Sheet {
                name: "Visible".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Visible
            },
            Sheet {
                name: "Hidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Hidden
            },
            Sheet {
                name: "VeryHidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::VeryHidden
            },
            Sheet {
                name: "Chart".to_string(),
                typ: SheetType::ChartSheet,
                visible: SheetVisible::Visible
            },
        ]
    );
}

#[test]
fn any_sheets_xlsb() {
    let workbook: Xlsb<_> = wb("any_sheets.xlsb");

    assert_eq!(
        workbook.sheets_metadata(),
        &[
            Sheet {
                name: "Visible".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Visible
            },
            Sheet {
                name: "Hidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Hidden
            },
            Sheet {
                name: "VeryHidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::VeryHidden
            },
            Sheet {
                name: "Chart".to_string(),
                typ: SheetType::ChartSheet,
                visible: SheetVisible::Visible
            },
        ]
    );
}

#[test]
fn any_sheets_xls() {
    let mut workbook: Xls<_> = wb("any_sheets.xls");

    assert_eq!(
        workbook.sheets_metadata(),
        &[
            Sheet {
                name: "Visible".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Visible
            },
            Sheet {
                name: "Hidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Hidden
            },
            Sheet {
                name: "VeryHidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::VeryHidden
            },
            Sheet {
                name: "Chart".to_string(),
                typ: SheetType::ChartSheet,
                visible: SheetVisible::Visible
            },
        ]
    );
    let vba = workbook.vba_project().unwrap().unwrap();
    let references = vba.get_references();
    assert_eq!(
        references,
        [
            Reference {
                name: "stdole".to_string(),
                description: "OLE Automation".to_string(),
                path: "C:\\Windows\\System32\\stdole2.tlb".into(),
            },
            Reference {
                name: "Office".to_string(),
                description: "Microsoft Office 16.0 Object Library".to_string(),
                path: "C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL".into(),
            },
        ],
    );
    assert_eq!(
        vba.get_module_names(),
        ["Диаграмма4", "Лист1", "Лист2", "Лист3",],
    );
}

#[test]
fn any_sheets_ods() {
    let workbook: Ods<_> = wb("any_sheets.ods");

    assert_eq!(
        workbook.sheets_metadata(),
        &[
            Sheet {
                name: "Visible".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Visible
            },
            Sheet {
                name: "Hidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Hidden
            },
            // ODS doesn't support Very Hidden
            Sheet {
                name: "VeryHidden".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Hidden
            },
            // ODS doesn't support chartsheet
            Sheet {
                name: "Chart".to_string(),
                typ: SheetType::WorkSheet,
                visible: SheetVisible::Visible
            },
        ]
    );
}

#[test]
fn issue_102() {
    let path = test_path("pass_protected.xlsx");
    assert!(
        matches!(
            open_workbook::<Xlsx<_>, _>(path),
            Err(calamine::XlsxError::Password)
        ),
        "Is expected to return XlsxError::Password error"
    );
}

#[test]
fn issue_374() {
    let mut workbook: Xls<_> = wb("biff5_write.xls");

    let first_sheet_name = workbook.sheet_names().first().unwrap().to_owned();

    assert_eq!("SheetJS", first_sheet_name);

    let range = workbook.worksheet_range(&first_sheet_name).unwrap();
    let second_row = range.rows().nth(1).unwrap();
    let cell_text = second_row.get(3).unwrap().to_string();

    assert_eq!("sheetjs", cell_text);
}

#[test]
fn issue_385() {
    let path = test_path("issue_385.xls");
    assert!(
        matches!(
            open_workbook::<Xls<_>, _>(path),
            Err(calamine::XlsError::Password)
        ),
        "Is expected to return XlsError::Password error"
    );
}

#[test]
fn pass_protected_with_readable_text() {
    let path = test_path("pass_protected_with_readable_text.xls");
    assert!(
        matches!(
            open_workbook::<Xls<_>, _>(path),
            Err(calamine::XlsError::Password)
        ),
        "Is expected to return XlsError::Password error"
    );
}

#[test]
fn pass_protected_xlsb() {
    let path = test_path("pass_protected.xlsb");
    assert!(
        matches!(
            open_workbook::<Xlsb<_>, _>(path),
            Err(calamine::XlsbError::Password)
        ),
        "Is expected to return XlsbError::Password error"
    );
}

#[test]
fn pass_protected_ods() {
    let path = test_path("pass_protected.ods");
    assert!(
        matches!(
            open_workbook::<Ods<_>, _>(path),
            Err(calamine::OdsError::Password)
        ),
        "Is expected to return OdsError::Password error"
    );
}

#[test]
fn issue_384_multiple_formula() {
    let mut workbook: Xlsx<_> = wb("formula.issue.xlsx");

    // first check values
    let range = workbook.worksheet_range("Sheet1").unwrap();
    let expected = [
        (0, 0, Data::Float(23.)),
        (0, 2, Data::Float(23.)),
        (12, 6, Data::Float(2.)),
        (13, 9, Data::String("US".into())),
    ];
    let expected = expected
        .iter()
        .map(|(r, c, v)| (*r, *c, v))
        .collect::<Vec<_>>();
    assert_eq!(range.used_cells().collect::<Vec<_>>(), expected);

    // check formula
    let formula = workbook.worksheet_formula("Sheet1").unwrap();
    let formula = formula
        .used_cells()
        .map(|(r, c, v)| (r, c, v.as_str()))
        .collect::<Vec<_>>();
    let expected = [
        (0, 0, "C1+E5"),
        // (0, 2, Data::Float(23.)),
        (12, 6, "SUM(1+1)"),
        (
            13,
            9,
            "IF(OR(Q22=\"\",Q22=\"United States\"),\"US\",\"Foreign\")",
        ),
    ];
    assert_eq!(formula, expected)
}

#[test]
fn issue_401_empty_tables() {
    let mut excel: Xlsx<_> = wb("date.xlsx");
    excel.load_tables().unwrap();
    let tables = excel.table_names();
    assert!(tables.is_empty());
}

#[test]
fn issue_391_shared_formula() {
    let mut excel: Xlsx<_> = wb("issue_391.xlsx");
    let mut expect = Range::<std::string::String>::new((1, 0), (6, 0));
    for (i, cell) in ["A1+1", "A2+1", "A3+1", "A4+1", "A5+1", "A6+1"]
        .iter()
        .enumerate()
    {
        expect.set_value((1 + i as u32, 0), cell.to_string());
    }
    let res = excel.worksheet_formula("Sheet1").unwrap();
    assert_eq!(expect.start(), res.start());
    assert_eq!(expect.end(), res.end());
    assert!(expect.cells().eq(res.cells()));
}

#[test]
fn issue_553_non_ascii_shared_formula() {
    // Test incrementing a shared formula in an xlsx file where the formula
    // contains utf-8 characters.
    let mut excel: Xlsx<_> = wb("issue_553.xlsx");
    let formula = excel.worksheet_formula("Sheet1").unwrap();
    assert!(formula
        .cells()
        .all(|(_, _, x)| x == r#"IF(ROW()>5,"한글","영어")"#))
}

#[test]
fn non_monotonic_si_shared_formula() {
    // This excel has been manually edited so that the si numbers do not monotonically increase (si
    // 0 swapped with 1)
    let mut excel: Xlsx<_> = wb("non_monotonic_si.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();
    let formula = excel.worksheet_formula("Sheet1").unwrap();

    let expected_values = [
        [Float(1.), Float(2.), Float(3.)],
        [Float(2.), Float(4.), Float(6.)],
        [Float(3.), Float(6.), Float(9.)],
        [Float(4.), Float(8.), Float(12.)],
        [Float(5.), Float(10.), Float(15.)],
        [Float(6.), Float(12.), Float(18.)],
    ];
    range_eq!(range, expected_values);

    let expected_formulas = [
        ["A1+1", "B1+2", "C1+3"],
        ["A2+1", "B2+2", "C2+3"],
        ["A3+1", "B3+2", "C3+3"],
        ["A4+1", "B4+2", "C4+3"],
        ["A5+1", "B5+2", "C5+3"],
    ];

    for (row_idx, row) in expected_formulas.iter().enumerate() {
        for (col_idx, expected_formula) in row.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32 + 1, col_idx as u32)),
                Some(&expected_formula.to_string())
            );
        }
    }
}

#[test]
fn issue_565_multi_axis_shared_formula() {
    // B1:D2 contains a shared formula that expands in 2 dimnensions
    let mut excel: Xlsx<_> = wb("issue_565_multi_axis_shared.xlsx");
    let formula = excel.worksheet_formula("Sheet1").unwrap();

    let expected_formulas = [["A1", "B1", "C1", "D1"], ["A2", "B2", "C2", "D2"]];

    for (row_idx, row) in expected_formulas.iter().enumerate() {
        for (col_idx, expected_formula) in row.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32, col_idx as u32)),
                Some(&expected_formula.to_string())
            );
        }
    }
}

#[test]
fn shared_formula_reversed() {
    // One sheet has a shared formula created by dragging downwards, the other has one created
    // by dragging upwards.
    let mut excel: Xlsx<_> = wb("shared_formula_reversed.xlsx");

    for sheet_name in &["Sheet1", "Sheet2"] {
        let range = excel.worksheet_range(sheet_name).unwrap();
        let formula = excel.worksheet_formula(sheet_name).unwrap();

        let expected_values = [Float(1.), Float(2.), Float(3.), Float(4.), Float(5.)];

        for (row_idx, expected_value) in expected_values.iter().enumerate() {
            assert_eq!(range.get_value((row_idx as u32, 1)), Some(expected_value));
        }

        let expected_formulas = ["A1", "A2", "A3", "A4", "A5"];

        for (row_idx, expected_formula) in expected_formulas.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32, 1)),
                Some(&expected_formula.to_string())
            );
        }
    }
}

#[test]
fn issue_567_absolute_shared_formula() {
    // Test absolute references in shared formulas. B$1 dragged to E3 should increment
    // the column (B, C, D, E) but keep the row fixed at 1.
    let mut excel: Xlsx<_> = wb("issue_567_absolute_shared.xlsx");
    let formula = excel.worksheet_formula("Sheet1").unwrap();

    let expected_formulas = [
        ["A$1", "B$1", "C$1", "D$1", "E$1"],
        ["A$1", "B$1", "C$1", "D$1", "E$1"],
        ["A$1", "B$1", "C$1", "D$1", "E$1"],
    ];

    for (row_idx, row) in expected_formulas.iter().enumerate() {
        for (col_idx, expected_formula) in row.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32, col_idx as u32)),
                Some(&expected_formula.to_string())
            );
        }
    }
}

#[test]
fn column_row_ranges() {
    // Test column and row ranges in formulas (e.g., E:F, 5:6)
    let mut excel: Xlsx<_> = wb("column_row_ranges.xlsx");

    // Test column ranges (E:F, F:G, G:H)
    let formula = excel.worksheet_formula("Column ranges").unwrap();

    let expected_formulas = [
        ["SUM(E:F)", "SUM(F:G)", "SUM(G:H)"],
        ["SUM(E:F)", "SUM(F:G)", "SUM(G:H)"],
        ["SUM(E:F)", "SUM(F:G)", "SUM(G:H)"],
        ["SUM(E:F)", "SUM(F:G)", "SUM(G:H)"],
        ["SUM(E:F)", "SUM(F:G)", "SUM(G:H)"],
    ];

    for (row_idx, row) in expected_formulas.iter().enumerate() {
        for (col_idx, expected_formula) in row.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32, col_idx as u32)),
                Some(&expected_formula.to_string()),
                "Column ranges mismatch at ({}, {})",
                row_idx,
                col_idx
            );
        }
    }

    // Test row ranges (5:6, 6:7, 7:8)
    let formula = excel.worksheet_formula("Row ranges").unwrap();

    let expected_formulas = [
        ["SUM(5:6)", "SUM(5:6)", "SUM(5:6)", "SUM(5:6)", "SUM(5:6)"],
        ["SUM(6:7)", "SUM(6:7)", "SUM(6:7)", "SUM(6:7)", "SUM(6:7)"],
        ["SUM(7:8)", "SUM(7:8)", "SUM(7:8)", "SUM(7:8)", "SUM(7:8)"],
    ];

    for (row_idx, row) in expected_formulas.iter().enumerate() {
        for (col_idx, expected_formula) in row.iter().enumerate() {
            assert_eq!(
                formula.get_value((row_idx as u32, col_idx as u32)),
                Some(&expected_formula.to_string()),
                "Row ranges mismatch at ({}, {})",
                row_idx,
                col_idx
            );
        }
    }
}

#[test]
fn issue_420_empty_s_attribute() {
    let mut excel: Xlsx<_> = wb("empty_s_attribute.xlsx");

    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(
        range,
        [
            [String("Name".to_string()), String("Value".to_string())],
            [String("John".to_string()), Float(1.)],
            [String("Sophia".to_string()), Float(2.)],
            [String("Peter".to_string()), Float(3.)],
            [String("Sam".to_string()), Float(4.)],
        ]
    );
}

#[test]
fn issue_438_charts() {
    let mut excel: Xlsx<_> = wb("issue438.xlsx");
    let _range = excel
        .worksheet_range("Chart1")
        .expect("could not open worksheet range");
}

#[test]
fn issue_444_memory_allocation() {
    let mut excel: Xls<_> = wb("issue444.xls"); // should not fail
    let range = excel
        .worksheet_range("Sheet1")
        .expect("could not open worksheet range");
    assert_eq!(range.get_size(), (10, 8));
}

#[test]
fn issue_446_formulas() {
    let mut excel: Xlsx<_> = wb("issue446.xlsx");
    let _ = excel.worksheet_formula("Sheet1").unwrap(); // should not fail
}

#[test]
fn test_ref_xlsx() {
    let mut excel: Xlsx<_> = wb("date.xlsx");
    let range = excel.worksheet_range_at_ref(0).unwrap().unwrap();

    range_eq!(
        range,
        [
            [
                DataRef::DateTime(ExcelDateTime::new(
                    44197.0,
                    ExcelDateTimeType::DateTime,
                    false
                )),
                DataRef::Float(15.0)
            ],
            [
                DataRef::DateTime(ExcelDateTime::new(
                    44198.0,
                    ExcelDateTimeType::DateTime,
                    false
                )),
                DataRef::Float(16.0)
            ],
            [
                DataRef::DateTime(ExcelDateTime::new(
                    10.6320601851852,
                    ExcelDateTimeType::TimeDelta,
                    false
                )),
                DataRef::Float(17.0)
            ]
        ]
    );
}

#[test]
fn test_ref_xlsb() {
    let mut excel: Xlsb<_> = wb("date.xlsb");
    let range = excel.worksheet_range_at_ref(0).unwrap().unwrap();

    range_eq!(
        range,
        [
            [
                DataRef::DateTime(ExcelDateTime::new(
                    44197.0,
                    ExcelDateTimeType::DateTime,
                    false
                )),
                DataRef::Float(15.0)
            ],
            [
                DataRef::DateTime(ExcelDateTime::new(
                    44198.0,
                    ExcelDateTimeType::DateTime,
                    false
                )),
                DataRef::Float(16.0)
            ],
            [
                DataRef::DateTime(ExcelDateTime::new(
                    10.6320601851852,
                    ExcelDateTimeType::TimeDelta,
                    false
                )),
                DataRef::Float(17.0)
            ]
        ]
    );
}

#[rstest]
#[case("header-row.xlsx", HeaderRow::FirstNonEmptyRow, (2, 0), (9, 3), &[Empty, Empty, String("Note 1".to_string()), Empty], 32)]
#[case("header-row.xlsx", HeaderRow::Row(0), (0, 0), (9, 3), &[Empty, Empty, Empty, Empty], 40)]
#[case("header-row.xlsx", HeaderRow::Row(8), (8, 0), (9, 3), &[String("Columns".to_string()), String("Column A".to_string()), String("Column B".to_string()), String("Column C".to_string())], 8)]
#[case("temperature.xlsx", HeaderRow::FirstNonEmptyRow, (0, 0), (2, 1), &[String("label".to_string()), String("value".to_string())], 6)]
#[case("temperature.xlsx", HeaderRow::Row(0), (0, 0), (2, 1), &[String("label".to_string()), String("value".to_string())], 6)]
#[case("temperature-in-middle.xlsx", HeaderRow::FirstNonEmptyRow, (3, 1), (5, 2), &[String("label".to_string()), String("value".to_string())], 6)]
#[case("temperature-in-middle.xlsx", HeaderRow::Row(0), (0, 1), (5, 2), &[Empty, Empty], 12)]
fn test_header_row_xlsx(
    #[case] fixture_path: &str,
    #[case] header_row: HeaderRow,
    #[case] expected_start: (u32, u32),
    #[case] expected_end: (u32, u32),
    #[case] expected_first_row: &[Data],
    #[case] expected_total_cells: usize,
) {
    let mut excel: Xlsx<_> = wb(fixture_path);
    assert_eq!(
        excel.sheets_metadata(),
        &[Sheet {
            name: "Sheet1".to_string(),
            typ: SheetType::WorkSheet,
            visible: SheetVisible::Visible
        },]
    );

    let range = excel
        .with_header_row(header_row)
        .worksheet_range("Sheet1")
        .unwrap();
    assert_eq!(range.start(), Some(expected_start));
    assert_eq!(range.end(), Some(expected_end));
    assert_eq!(range.rows().next().unwrap(), expected_first_row);
    assert_eq!(range.cells().count(), expected_total_cells);
}

#[test]
fn test_read_twice_with_different_header_rows() {
    let mut xlsx: Xlsx<_> = wb("any_sheets.xlsx");
    let _ = xlsx
        .with_header_row(HeaderRow::Row(2))
        .worksheet_range("Visible")
        .unwrap();
    let _ = xlsx
        .with_header_row(HeaderRow::Row(1))
        .worksheet_range("Visible")
        .unwrap();
}

#[test]
fn test_header_row_xlsb() {
    let mut xlsb: Xlsb<_> = wb("date.xlsb");
    assert_eq!(
        xlsb.sheets_metadata(),
        &[Sheet {
            name: "Sheet1".to_string(),
            typ: SheetType::WorkSheet,
            visible: SheetVisible::Visible
        }]
    );

    let first_line = [
        DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false,
        )),
        Float(15.0),
    ];
    let second_line = [
        DateTime(ExcelDateTime::new(
            44198.0,
            ExcelDateTimeType::DateTime,
            false,
        )),
        Float(16.0),
    ];

    let range = xlsb.worksheet_range("Sheet1").unwrap();
    assert_eq!(range.start(), Some((0, 0)));
    assert_eq!(range.end(), Some((2, 1)));
    assert_eq!(range.rows().next().unwrap(), &first_line);
    assert_eq!(range.rows().nth(1).unwrap(), &second_line);

    let range = xlsb
        .with_header_row(HeaderRow::Row(1))
        .worksheet_range("Sheet1")
        .unwrap();
    assert_eq!(range.start(), Some((1, 0)));
    assert_eq!(range.end(), Some((2, 1)));
    assert_eq!(range.rows().next().unwrap(), &second_line);
}

#[test]
fn test_header_row_xls() {
    let mut xls: Xls<_> = wb("date.xls");
    assert_eq!(
        xls.sheets_metadata(),
        &[Sheet {
            name: "Sheet1".to_string(),
            typ: SheetType::WorkSheet,
            visible: SheetVisible::Visible
        }]
    );

    let first_line = [
        DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false,
        )),
        Int(15),
    ];
    let second_line = [
        DateTime(ExcelDateTime::new(
            44198.0,
            ExcelDateTimeType::DateTime,
            false,
        )),
        Int(16),
    ];

    let range = xls.worksheet_range("Sheet1").unwrap();
    assert_eq!(range.start(), Some((0, 0)));
    assert_eq!(range.end(), Some((2, 1)));
    assert_eq!(range.rows().next().unwrap(), &first_line);
    assert_eq!(range.rows().nth(1).unwrap(), &second_line);

    let range = xls
        .with_header_row(HeaderRow::Row(1))
        .worksheet_range("Sheet1")
        .unwrap();
    assert_eq!(range.start(), Some((1, 0)));
    assert_eq!(range.end(), Some((2, 1)));
    assert_eq!(range.rows().next().unwrap(), &second_line);
}

#[test]
fn test_header_row_ods() {
    let mut ods: Ods<_> = wb("date.ods");
    assert_eq!(
        ods.sheets_metadata(),
        &[Sheet {
            name: "Sheet1".to_string(),
            typ: SheetType::WorkSheet,
            visible: SheetVisible::Visible
        }]
    );

    let first_line = [DateTimeIso("2021-01-01".to_string()), Float(15.0)];
    let third_line = [DurationIso("PT10H10M10S".to_string()), Float(17.0)];

    let range = ods.worksheet_range("Sheet1").unwrap();
    assert_eq!(range.start(), Some((0, 0)));
    assert_eq!(range.end(), Some((3, 1)));
    assert_eq!(range.rows().next().unwrap(), &first_line);
    assert_eq!(range.rows().nth(2).unwrap(), &third_line);

    let range = ods
        .with_header_row(HeaderRow::Row(2))
        .worksheet_range("Sheet1")
        .unwrap();
    assert_eq!(range.start(), Some((2, 0)));
    assert_eq!(range.end(), Some((3, 1)));
    assert_eq!(range.rows().next().unwrap(), &third_line);
}

#[rstest]
#[case("single-empty.ods")]
#[case("multi-empty.ods")]
fn issue_repeated_empty(#[case] fixture_path: &str) {
    let mut ods: Ods<_> = wb(fixture_path);
    let range = ods.worksheet_range_at(0).unwrap().unwrap();
    range_eq!(
        range,
        [
            [String("StringCol".to_string())],
            [String("bbb".to_string())],
            [String("ccc".to_string())],
            [String("ddd".to_string())],
            [String("eee".to_string())],
            [Empty],
            [Empty],
            [Empty],
            [Empty],
            [Empty],
            [Empty],
            [Empty],
            [String("zzz".to_string())],
        ]
    );
}

#[test]
fn ods_with_annotations() {
    let mut ods: Ods<_> = wb("with-annotation.ods");
    let range = ods.worksheet_range("table1").unwrap();
    range_eq!(range, [[String("cell a.1".to_string())],]);
}

#[rstest]
#[case(HeaderRow::Row(0), &[
    [Empty, Empty],
    [Empty, Empty],
    [String("a".to_string()), Float(0.0)],
    [String("b".to_string()), Float(1.0)]
])]
#[case(HeaderRow::Row(1), &[
    [Empty, Empty],
    [String("a".to_string()), Float(0.0)],
    [String("b".to_string()), Float(1.0)]
])]
#[case(HeaderRow::Row(2), &[
    [String("a".to_string()), Float(0.0)],
    [String("b".to_string()), Float(1.0)]
])]
fn test_no_header(#[case] header_row: HeaderRow, #[case] expected: &[[Data; 2]]) {
    let mut excel: Xlsx<_> = wb("no-header.xlsx");
    let range = excel
        .with_header_row(header_row)
        .worksheet_range_at(0)
        .unwrap()
        .unwrap();
    range_eq!(range, expected);
}

#[test]
fn test_string_ref() {
    let mut xlsx: Xlsx<_> = wb("string-ref.xlsx");
    let expected_range = [
        [String("col1".to_string())],
        [String("-8086931554011838357".to_string())],
    ];
    // first sheet
    range_eq!(xlsx.worksheet_range_at(0).unwrap().unwrap(), expected_range);
    // second sheet is the same with a cell reference to the first sheet
    range_eq!(xlsx.worksheet_range_at(1).unwrap().unwrap(), expected_range);
}

#[test]
fn test_malformed_format() {
    let _xls: Xls<_> = wb("malformed_format.xls");
}

#[test]
fn test_oom_allocation() {
    let _xls: Xls<_> = wb("OOM_alloc.xls");
    let mut xls: Xls<_> = wb("OOM_alloc2.xls");
    let ws = xls.worksheets();
    assert_eq!(ws.len(), 1);
    assert_eq!(ws[0].0, "Colsale (Aug".to_string());

    let path = test_path("OOM_alloc3.xls");
    assert!(
        matches!(
            open_workbook::<Xls<_>, _>(path),
            Err(calamine::XlsError::Cfb(_))
        ),
        "Is expected to return XlsError::Cfb error"
    );
}

// Test for issue #548. The SST table in the test file has an incorrect unique
// string count.
#[test]
fn test_incorrect_sst_unique_count() {
    // Check for the string that appears last in the SST table: "11th May 2023".
    // This appears in cell C10 of each worksheet in the workbook.
    let mut xls: Xls<_> = wb("gh548_incorrect_sst_unique_count.xls");
    let range = xls.worksheet_range("System Level Data").unwrap();

    assert_eq!(
        range.get_value((9, 2)).unwrap(),
        &String("11th May 2023".into())
    );
}

// Test the parsing of an SST table that finishes with a complete, untruncated,
// string at the end of the block, and where the next string is in a CONTINUE
// block. This is related to the previous test for gh548 to ensure that the edge
// condition is met.
#[test]
fn test_sst_continue() {
    let mut xls: Xls<_> = wb("sst_continue.xls");
    let range = xls.worksheet_range("Sheet1").unwrap();

    // Check for the string that appears last in the SST table.
    assert_eq!(
        range.get_value((135, 0)).unwrap(),
        &String("New CONTINUE block".into())
    );
}

// Test for issue #419 where the part name is sentence case instead of camel
// case. The test file contains a sub-file called "xl/SharedStrings.xml" (note
// the uppercase S in Shared). This is allowed by "Office Open XML File Formats
// — Open Packaging Conventions" 6.2.2.3.
#[test]
fn test_xlsx_case_insensitive_part_name() {
    let mut xlsx: Xlsx<_> = wb("issue_419.xlsx");

    let range = xlsx.worksheet_range("Sheet1").unwrap();
    let expected_range = [[String("Hello".to_string())]];

    range_eq!(range, expected_range);
}

// Test for issue #419 in Xlsb file. See the previous test for the details.
#[test]
fn test_xlsb_case_insensitive_part_name() {
    let mut xlsb: Xlsb<_> = wb("issue_419.xlsb");

    let range = xlsb.worksheet_range("Sheet1").unwrap();
    let expected_range = [[String("Hello".to_string())]];

    range_eq!(range, expected_range);
}

// Test for issue #530 where the part names in the xlsx file use a Windows-style
// backslash. For example "xl\_rels\workbook.xml.rels" instead of
// "xl/_rels/workbook.xml.rels".
#[test]
fn test_xlsx_backward_slash_part_name() {
    let _: Xlsx<_> = wb("issue_530.xlsx");
}

// Test for issue #573 where a shared string doesn't have a <t> sub-element.
// This is unusual but valid according to the xlsx specification.
#[test]
fn test_xlsx_empty_shared_string() {
    let mut excel: Xlsx<_> = wb("empty_shared_string.xlsx");
    let range = excel.worksheet_range("Sheet1").unwrap();

    range_eq!(
        range,
        [[String("abc".to_string())], [String("".to_string())]]
    );
}
