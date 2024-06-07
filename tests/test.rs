use calamine::Data::{
    Bool, DateTime, DateTimeIso, DurationIso, Empty, Error, Float, RichText, String,
};
use calamine::{
    open_workbook, open_workbook_auto, Color, DataType, Dimensions, ExcelDateTime,
    ExcelDateTimeType, FontFormat, Ods, Range, Reader, RichTextPart, Sheet, SheetType,
    SheetVisible, Xls, Xlsb, Xlsx,
};
use calamine::{CellErrorType::*, Data};
use std::borrow::Cow;
use std::collections::BTreeSet;
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
    setup();

    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issue3.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("Sheet1").unwrap();
    range_eq!(range, [[Float(1.), String("a".to_string())]]);
}

#[test]
fn issue_4() {
    setup();

    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("issue5").unwrap();
    range_eq!(range, [[Float(0.5)]]);
}

#[test]
fn issue_6() {
    setup();

    // test if sheet is resolved with only one row
    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/errors.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/issue9.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    let range = excel.worksheet_range("Feuil1").unwrap();

    let format1 = FontFormat {
        bold: true,
        color: Color::Theme(1, 0.0),
        name: Some("Calibri".to_owned()),
        ..Default::default()
    };
    let format2 = FontFormat {
        color: Color::Theme(1, 0.0),
        name: Some("Calibri".to_owned()),
        ..Default::default()
    };
    let format3 = FontFormat {
        underlined: true,
        color: Color::Theme(1, 0.0),
        name: Some("Calibri".to_owned()),
        ..Default::default()
    };

    let mut cell2 = calamine::RichText::new();
    cell2.push(RichTextPart {
        text: "test2",
        format: Cow::Borrowed(&format1),
    });
    cell2.push(RichTextPart {
        text: " o",
        format: Cow::Borrowed(&format2),
    });
    cell2.push(RichTextPart {
        text: "ther",
        format: Cow::Borrowed(&format3),
    });

    let format4 = FontFormat {
        color: Color::ARGB(255, 0, 176, 80),
        name: Some("Calibri".to_owned()),
        ..Default::default()
    };
    let format5 = FontFormat {
        color: Color::ARGB(255, 0, 112, 192),
        name: Some("Calibri".to_owned()),
        ..Default::default()
    };

    let mut cell3 = calamine::RichText::new();
    cell3.push(RichTextPart {
        text: "test3",
        format: Cow::Borrowed(&format4),
    });
    cell3.push(RichTextPart {
        text: " ",
        format: Cow::Borrowed(&format2),
    });
    cell3.push(RichTextPart {
        text: "aaa",
        format: Cow::Borrowed(&format5),
    });

    range_eq!(
        range,
        [
            [String("test1".to_string())],
            [RichText(cell2.clone())],
            [RichText(cell3.clone())],
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
    setup();

    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xls<_> = open_workbook(&path).unwrap();

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
        .expect("error in wks range");
    assert_eq!(range.get_size(), (52046, 12));
}

#[test]
fn ods() {
    setup();

    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/covered.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

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
    let path = format!("{}/tests/special_cells.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/issues.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

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
    setup();

    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

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
fn partial_richtext_ods() {
    setup();

    let path = format!("{}/tests/richtext_issue.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("datatypes").unwrap();
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

    let range = excel.worksheet_range("Sheet1").unwrap();
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
fn rich_text_support() {
    setup();

    let format = FontFormat {
        color: Color::Theme(1, 0.0),
        name: Some("Aptos Narrow".to_owned()),
        ..Default::default()
    };

    // TODO: Add XLSB, XLS, ODS once supported.
    #[allow(clippy::single_element_loop)]
    for file in ["rich_text_support.xlsx"] {
        let path = format!("{}/tests/{}", env!("CARGO_MANIFEST_DIR"), file);
        let mut excel = open_workbook_auto(path).unwrap();
        let range = excel.worksheet_range_at(0).unwrap().unwrap();

        let cell = range.get((0, 0)).unwrap();
        if let Data::RichText(rich_text) = cell {
            let elements = rich_text.elements().collect::<Vec<_>>();
            assert_eq!(elements.len(), 7);
            assert_eq!(elements[0].text, "N ");
            assert!(elements[0].format.is_default());
            assert_eq!(elements[1].text, "F");
            assert!(elements[1].format.bold);
            assert_eq!(elements[2].text, " ");
            assert_eq!(elements[2].format, Cow::Borrowed(&format));
            assert_eq!(elements[3].text, "U");
            assert!(elements[3].format.underlined);
            assert_eq!(elements[4].text, " ");
            assert_eq!(elements[4].format, Cow::Borrowed(&format));
            assert_eq!(elements[5].text, "I");
            assert!(elements[5].format.italic);
            assert_eq!(elements[6].text, " N");
            assert_eq!(elements[6].format, Cow::Borrowed(&format));
        } else {
            panic!("Cell was not parsed as RichText");
        }

        let cell = range.get((0, 1)).unwrap();
        if let Data::RichText(rich_text) = cell {
            let elements = rich_text.elements().collect::<Vec<_>>();
            assert_eq!(elements.len(), 3);
            assert_eq!(elements[0].text, "small ");
            assert_eq!(elements[0].format.size, 8);
            assert_eq!(elements[1].text, "normal ");
            assert_eq!(elements[1].format.size, 11);
            assert_eq!(elements[2].text, "big");
            assert_eq!(elements[2].format.size, 14);
        } else {
            panic!("Cell was not parsed as RichText");
        }

        let cell = range.get((1, 0)).unwrap();
        if let Data::RichText(rich_text) = cell {
            let elements = rich_text.elements().collect::<Vec<_>>();
            assert_eq!(elements.len(), 2);
            assert_eq!(elements[0].text, "black ");
            assert_eq!(elements[0].format, Cow::Owned(FontFormat::default()));
            assert_eq!(elements[1].text, "green");
            assert_eq!(elements[1].format.color, Color::Theme(6, 0.0));
        } else {
            panic!("Cell was not parsed as RichText");
        }

        let cell = range.get((1, 1)).unwrap();
        if let Data::RichText(rich_text) = cell {
            let elements = rich_text.elements().collect::<Vec<_>>();
            assert_eq!(elements.len(), 2);
            assert_eq!(elements[0].text, "aptos ");
            assert_eq!(elements[0].format.name, None);
            assert_eq!(elements[1].text, "calibri");
            assert_eq!(elements[1].format.name.as_deref(), Some("Calibri"));
        } else {
            panic!("Cell was not parsed as RichText");
        }
    }
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
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_xlsb() {
    setup();

    let path = format!("{}/tests/issues.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_vals_xlsb() {
    setup();

    let path = format!("{}/tests/issue_182.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsb<_> = open_workbook(&path).unwrap();

    let range = excel.worksheet_range("formula_vals").unwrap();
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
    setup();

    let path = format!("{}/tests/issues.xls", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xls<_> = open_workbook(&path).unwrap();

    let sheets = excel.sheet_names().to_owned();
    for s in sheets {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["B1+OneRange".to_string()]]);
}

#[test]
fn formula_ods() {
    setup();

    let path = format!("{}/tests/issues.ods", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Ods<_> = open_workbook(&path).unwrap();

    for s in excel.sheet_names().to_owned() {
        let _ = excel.worksheet_formula(&s).unwrap();
    }

    let formula = excel.worksheet_formula("Sheet1").unwrap();
    range_eq!(formula, [["of:=[.B1]+$$OneRange".to_string()]]);
}

#[test]
fn empty_sheet() {
    setup();

    let path = format!("{}/tests/empty_sheet.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    for s in excel.sheet_names().to_owned() {
        let range = excel.worksheet_range(&s).unwrap();
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

    let range = excel.worksheet_range("issue2").unwrap();
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
    let range = xls.worksheet_range("Boys").unwrap();
    assert_eq!(range.get_value((6, 2)), Some(&Float(9.)));
}

#[test]
fn skip_phonetic_text() {
    setup();

    let path = format!("{}/tests/rph.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
    let range = xls.worksheet_range("Sheet1").unwrap();
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
}

#[test]
fn date_xls() {
    setup();

    let path = format!("{}/tests/date.xls", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xls<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date_1904.xls", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xls<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date_1904.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date_iso.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsx<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date.ods", env!("CARGO_MANIFEST_DIR"));
    let mut ods: Ods<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsb<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    setup();

    let path = format!("{}/tests/date_1904.xlsb", env!("CARGO_MANIFEST_DIR"));
    let mut xls: Xlsb<_> = open_workbook(&path).unwrap();
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

    #[cfg(feature = "dates")]
    {
        let date = chrono::NaiveDate::from_ymd_opt(2021, 01, 01).unwrap();
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
    let path = format!("{}/tests/merged_range.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    excel.load_merged_regions().unwrap();
    assert_eq!(
        excel
            .merged_regions()
            .iter()
            .map(|(o1, o2, o3)| (o1.to_string(), o2.to_string(), o3.clone()))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        vec![
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
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), o3.clone()))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        vec![
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
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), o3.clone()))
            .collect::<BTreeSet<(String, String, Dimensions)>>(),
        vec![
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
    setup();

    let path = format!("{}/tests/issue252.xlsx", env!("CARGO_MANIFEST_DIR"));

    // should err, not panic
    assert!(open_workbook::<Xls<_>, _>(&path).is_err());
}

#[test]
fn issue_261() {
    setup();

    let mut workbook_with_missing_r_attributes = {
        let path = format!("{}/tests/issue_261.xlsx", env!("CARGO_MANIFEST_DIR"));
        open_workbook::<Xlsx<_>, _>(&path).unwrap()
    };

    let mut workbook_fixed_by_excel = {
        let path = format!(
            "{}/tests/issue_261_fixed_by_excel.xlsx",
            env!("CARGO_MANIFEST_DIR")
        );
        open_workbook::<Xlsx<_>, _>(&path).unwrap()
    };

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
    let path = format!(
        "{}/tests/xls_wrong_decimals.xls",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut excel: Xls<_> = open_workbook(&path).unwrap();
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
        let path = format!("{}/tests/issue_271.xls", env!("CARGO_MANIFEST_DIR"));
        let mut workbook: Xls<_> = open_workbook(path)?;
        let v = workbook.worksheets();
        let (_sheetname, range) = v.first().expect("bad format");
        dbg!(_sheetname);
        let value = range.get((0, 1)).map(|s| s.to_string());
        values.push(value);
        count += 1;
        if count > 20 {
            break;
        }
    }

    dbg!(&values);

    values.sort_unstable();
    values.dedup();

    assert_eq!(values.len(), 1);

    Ok(())
}

#[test]
fn issue_305_merge_cells() {
    let path = format!("{}/tests/merge_cells.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    let merge_cells = excel.worksheet_merge_cells_at(0).unwrap().unwrap();

    assert_eq!(
        merge_cells,
        vec![
            Dimensions::new((0, 0), (0, 1)),
            Dimensions::new((1, 0), (3, 0)),
            Dimensions::new((1, 1), (3, 3))
        ]
    );
}

#[test]
fn issue_305_merge_cells_xls() {
    let path = format!("{}/tests/merge_cells.xls", env!("CARGO_MANIFEST_DIR"));
    let excel: Xls<_> = open_workbook(&path).unwrap();
    let merge_cells = excel.worksheet_merge_cells_at(0).unwrap();

    assert_eq!(
        merge_cells,
        vec![
            Dimensions::new((0, 0), (0, 1)),
            Dimensions::new((1, 0), (3, 0)),
            Dimensions::new((1, 1), (3, 3))
        ]
    );
}

// cargo test --features picture
#[test]
#[cfg(feature = "picture")]
fn pictures() -> Result<(), calamine::Error> {
    let jpg_path = format!("{}/tests/picture.jpg", env!("CARGO_MANIFEST_DIR"));
    let png_path = format!("{}/tests/picture.png", env!("CARGO_MANIFEST_DIR"));

    let xlsx_path = format!("{}/tests/picture.xlsx", env!("CARGO_MANIFEST_DIR"));
    let xlsb_path = format!("{}/tests/picture.xlsb", env!("CARGO_MANIFEST_DIR"));
    let xls_path = format!("{}/tests/picture.xls", env!("CARGO_MANIFEST_DIR"));
    let ods_path = format!("{}/tests/picture.ods", env!("CARGO_MANIFEST_DIR"));

    let jpg_hash = sha256::digest(&*std::fs::read(&jpg_path)?);
    let png_hash = sha256::digest(&*std::fs::read(&png_path)?);

    let xlsx: Xlsx<_> = open_workbook(xlsx_path)?;
    let xlsb: Xlsb<_> = open_workbook(xlsb_path)?;
    let xls: Xls<_> = open_workbook(xls_path)?;
    let ods: Ods<_> = open_workbook(ods_path)?;

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
        let pic_hash = sha256::digest(&*data);
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
    setup();

    let path = format!("{}/tests/merged_cells.ods", env!("CARGO_MANIFEST_DIR"));
    let mut ods: Ods<_> = open_workbook(&path).unwrap();
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
    setup();

    let path = format!(
        "{}/tests/number_rows_repeated.ods",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut ods: Ods<_> = open_workbook(&path).unwrap();
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
    setup();
    let path = format!("{}/tests/xls_formula.xls", env!("CARGO_MANIFEST_DIR"));
    let mut wb: Xls<_> = open_workbook(&path).unwrap();
    let formula = wb.worksheet_formula("Sheet1").unwrap();
    let mut rows = formula.rows();
    assert_eq!(rows.next(), Some(&["A1*2".to_owned()][..]));
    assert_eq!(rows.next(), Some(&["2*Sheet2!A1".to_owned()][..]));
    assert_eq!(rows.next(), Some(&["A1+Sheet2!A1".to_owned()][..]));
    assert_eq!(rows.next(), None);
}

#[test]
fn issue304_xls_values() {
    setup();
    let path = format!("{}/tests/xls_formula.xls", env!("CARGO_MANIFEST_DIR"));
    let mut wb: Xls<_> = open_workbook(&path).unwrap();
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
    setup();
    let path = format!("{}/tests/xls_ref_String.xls", env!("CARGO_MANIFEST_DIR"));
    let mut wb: Xls<_> = open_workbook(&path).unwrap();
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
    setup();

    let path = format!("{}/tests/issue281.xlsm", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

    let mut vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.to_mut().get_module("testVBA").unwrap(),
        "Attribute VB_Name = \"testVBA\"\r\nPublic Sub test()\r\n    MsgBox \"Hello from \
         vba!\"\r\nEnd Sub\r\n"
    );
}

#[test]
fn issue343() {
    setup();

    let path = format!("{}/tests/issue343.xls", env!("CARGO_MANIFEST_DIR"));

    // should not panic
    let _: Xls<_> = open_workbook(&path).unwrap();
}

#[test]
fn any_sheets_xlsx() {
    setup();

    let path = format!("{}/tests/any_sheets.xlsx", env!("CARGO_MANIFEST_DIR"));
    let workbook: Xlsx<_> = open_workbook(path).unwrap();

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
    setup();

    let path = format!("{}/tests/any_sheets.xlsb", env!("CARGO_MANIFEST_DIR"));
    let workbook: Xlsb<_> = open_workbook(path).unwrap();

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
    setup();

    let path = format!("{}/tests/any_sheets.xls", env!("CARGO_MANIFEST_DIR"));
    let workbook: Xls<_> = open_workbook(path).unwrap();

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
fn any_sheets_ods() {
    setup();

    let path = format!("{}/tests/any_sheets.ods", env!("CARGO_MANIFEST_DIR"));
    let workbook: Ods<_> = open_workbook(path).unwrap();

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
    setup();

    let path = format!("{}/tests/pass_protected.xlsx", env!("CARGO_MANIFEST_DIR"));

    assert!(
        matches!(
            open_workbook::<Xlsx<_>, std::string::String>(path),
            Err(calamine::XlsxError::Password)
        ),
        "Is expeced to return XlsxError::Password error"
    );
}

#[test]
fn issue_374() {
    let path = format!("{}/tests/biff5_write.xls", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xls<_> = open_workbook(path).unwrap();

    let first_sheet_name = workbook.sheet_names().first().unwrap().to_owned();

    assert_eq!("SheetJS", first_sheet_name);

    let range = workbook.worksheet_range(&first_sheet_name).unwrap();
    let second_row = range.rows().nth(1).unwrap();
    let cell_text = second_row.get(3).unwrap().to_string();

    assert_eq!("sheetjs", cell_text);
}

#[test]
fn issue_385() {
    let path = format!("{}/tests/issue_385.xls", env!("CARGO_MANIFEST_DIR"));

    assert!(
        matches!(
            open_workbook::<Xls<_>, std::string::String>(path),
            Err(calamine::XlsError::Password)
        ),
        "Is expeced to return XlsError::Password error"
    );
}

#[test]
fn pass_protected_xlsb() {
    let path = format!("{}/tests/pass_protected.xlsb", env!("CARGO_MANIFEST_DIR"));

    assert!(
        matches!(
            open_workbook::<Xlsb<_>, std::string::String>(path),
            Err(calamine::XlsbError::Password)
        ),
        "Is expeced to return XlsbError::Password error"
    );
}

#[test]
fn pass_protected_ods() {
    let path = format!("{}/tests/pass_protected.ods", env!("CARGO_MANIFEST_DIR"));

    assert!(
        matches!(
            open_workbook::<Ods<_>, std::string::String>(path),
            Err(calamine::OdsError::Password)
        ),
        "Is expeced to return OdsError::Password error"
    );
}

#[test]
fn issue_384_multiple_formula() {
    let path = format!("{}/tests/formula.issue.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();

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
    setup();

    let path = format!("{}/tests/date.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    excel.load_tables().unwrap();
    let tables = excel.table_names();
    assert!(tables.is_empty());
}

#[test]
fn issue_391_shared_formula() {
    setup();

    let path = format!("{}/tests/issue_391.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();
    let mut expect = Range::<std::string::String>::new((1, 0), (6, 0));
    for (i, cell) in vec!["A1+1", "A2+1", "A3+1", "A4+1", "A5+1", "A6+1"]
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
fn issue_420_empty_s_attribute() {
    setup();

    let path = format!(
        "{}/tests/empty_s_attribute.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut excel: Xlsx<_> = open_workbook(&path).unwrap();

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
