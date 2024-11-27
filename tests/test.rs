#![allow(clippy::zero_prefixed_literal)]

use calamine::Data::{Bool, DateTime, DateTimeIso, DurationIso, Empty, Error, Float, Int, String};
use calamine::{
    open_workbook, open_workbook_auto, Color, DataType, Dimensions, ExcelDateTime,
    ExcelDateTimeType, FontFormat, Ods, Range, Reader, RichText, RichTextPart, Sheet, SheetType,
    SheetVisible, Sheets, Xls, Xlsb, Xlsx,
};
use calamine::{CellErrorType::*, Data};
use calamine::{DataRef, HeaderRow, ReaderRef};
use rstest::rstest;
use std::borrow::Cow;
use std::collections::BTreeSet;
use std::fs::File;
use std::io::{BufReader, Cursor};
use std::sync::Once;

static INIT: Once = Once::new();

/// Setup function that is only run once, even if called multiple times.
fn wb<R: Reader<BufReader<File>>>(name: &str) -> R {
    INIT.call_once(|| {
        env_logger::init();
    });
    let path = format!("{}/tests/{name}", env!("CARGO_MANIFEST_DIR"));
    open_workbook(&path).expect(&path)
}

/// Setup function that is only run once, even if called multiple times.
fn wb_auto(name: &str) -> Sheets<BufReader<File>> {
    INIT.call_once(|| {
        env_logger::init();
    });
    let path = format!("{}/tests/{name}", env!("CARGO_MANIFEST_DIR"));
    open_workbook_auto(&path).expect(&path)
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
            [Data::RichText(cell2.clone())],
            [Data::RichText(cell3.clone())],
            [String("test4".to_string())]
        ]
    );
}

#[test]
fn vba() {
    let mut excel: Xlsx<_> = wb("vba.xlsm");
    let mut vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.to_mut().get_module("testVBA").unwrap(),
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
    let format = FontFormat {
        color: Color::Theme(1, 0.0),
        name: Some("Aptos Narrow".to_owned()),
        ..Default::default()
    };

    // TODO: Add XLSB, XLS, ODS once supported.
    #[allow(clippy::single_element_loop)]
    for file in ["rich_text_support.xlsx"] {
        let mut excel = wb_auto(file);
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
    let excel: Xlsx<_> = wb("issues.xlsx");
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
    let excel: Xlsb<_> = wb("issues.xlsb");
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
    let excel: Xls<_> = wb("issues.xls");
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
    let excel: Ods<_> = wb("issues.ods");
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
    assert_eq!(names, vec!["stdole", "Office"]);
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
    let root = env!("CARGO_MANIFEST_DIR");
    let ordered_names: Vec<std::string::String> = [
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
        &DataRef::SharedString(&RichText::plain("celsius".to_owned()))
    );
    assert_eq!(
        data.get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString(&RichText::plain("fahrenheit".to_owned()))
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
        &DataRef::SharedString(&RichText::plain("something".to_owned()))
    );
    assert_eq!(
        data.get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString(&RichText::plain("else".to_owned()))
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
        &DataRef::SharedString(&RichText::plain("something".to_owned()))
    );
    assert_eq!(
        owned_data
            .get((1, 0))
            .expect("Could not get data from table ref."),
        &DataRef::SharedString(&RichText::plain("else".to_owned()))
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
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), *o3))
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
            .map(|&(o1, o2, o3)| (o1.to_string(), o2.to_string(), *o3))
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
    let mut excel: Xlsx<_> = wb("merge_cells.xlsx");
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
    let excel: Xls<_> = wb("merge_cells.xls");
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
    let path = |name: &str| format!("{}/tests/{name}", env!("CARGO_MANIFEST_DIR"));
    let jpg_path = path("picture.jpg");
    let png_path = path("picture.png");

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

    let mut vba = excel.vba_project().unwrap().unwrap();
    assert_eq!(
        vba.to_mut().get_module("testVBA").unwrap(),
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
    let workbook: Xls<_> = wb("any_sheets.xls");

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
fn isssue_444_memory_allocation() {
    let mut excel: Xls<_> = wb("issue444.xls"); // should not fail
    let range = excel
        .worksheet_range("Sheet1")
        .expect("could not open worksheet range");
    assert_eq!(range.get_size(), (10, 8));
}

#[test]
fn isssue_446_formulas() {
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
