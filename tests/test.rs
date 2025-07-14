use calamine::Data::{Bool, DateTime, DateTimeIso, DurationIso, Empty, Error, Float, Int, String};
use calamine::{
    open_workbook, open_workbook_auto, CellFormat, Color, DataRef, DataWithFormatting, Dimensions, ExcelDateTime, ExcelDateTimeType, HeaderRow, Ods, PatternType, Range, Reader, ReaderRef, Sheet, SheetType, SheetVisible, Xls, Xlsb, Xlsx
};
use calamine::{CellErrorType::*, Data};
use rstest::rstest;
use std::collections::BTreeSet;
use std::fs::File;
use std::io::{BufReader, Cursor};
use std::sync::Arc;
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
fn test_worksheet_range_with_formatting() {
    let mut excel: Xlsx<_> = wb("format.xlsx");
    
    // Get the worksheet range which now returns Range<DataWithFormatting>
    let range = excel.worksheet_range("Sheet1").unwrap();
    
    // Test that we got a valid range
    assert!(range.start().is_some());
    assert!(range.end().is_some());
    
    // Test cell A1 - should have white font on black background formatting
    let cell_a1 = range.get_value((0, 0)).unwrap(); // A1
    let data_a1 = cell_a1.get_data();
    let formatting_a1 = cell_a1.get_formatting();
    
    // Verify the cell data
    assert_eq!(data_a1.to_string(), "White header on black text");
    
    // Verify the formatting is present
    assert!(formatting_a1.is_some(), "A1 should have formatting");
    
    let fmt_a1 = formatting_a1.as_ref().unwrap();
    
    // Test font formatting (white font)
    let font_a1 = fmt_a1.font.as_ref().expect("A1 should have font formatting");
    assert_eq!(
        font_a1.color,
        Some(Color::Argb {
            a: 255,
            r: 255,
            g: 255,
            b: 255
        })
    ); // White font
    
    // Test fill formatting (black background)
    let fill_a1 = fmt_a1.fill.as_ref().expect("A1 should have fill formatting");
    assert_eq!(fill_a1.pattern_type, PatternType::Solid);
    assert_eq!(
        fill_a1.foreground_color,
        Some(Color::Argb {
            a: 255,
            r: 0,
            g: 0,
            b: 0
        })
    ); // Black background
    
    // Test cell A2 - should have right alignment formatting
    let cell_a2 = range.get_value((1, 0)).unwrap(); // A2
    let data_a2 = cell_a2.get_data();
    let formatting_a2 = cell_a2.get_formatting();
    
    // Verify the cell data
    assert_eq!(data_a2.to_string(), "Right aligned");
    
    // Verify the formatting is present
    assert!(formatting_a2.is_some(), "A2 should have formatting");
    
    let fmt_a2 = formatting_a2.as_ref().unwrap();
    
    // Test alignment formatting (right aligned)
    let alignment_a2 = fmt_a2.alignment.as_ref().expect("A2 should have alignment formatting");
    assert_eq!(alignment_a2.horizontal, Some(Arc::from("right")));
    
    // Test iterating through the range and checking formatting
    let mut cells_with_formatting = 0;
    let mut cells_without_formatting = 0;
    
    for (row, col, cell) in range.used_cells() {
        let data = cell.get_data();
        let formatting = cell.get_formatting();
        
        println!("Cell ({}, {}): {} - has formatting: {}", row, col, data, formatting.is_some());
        
        if formatting.is_some() {
            cells_with_formatting += 1;
        } else {
            cells_without_formatting += 1;
        }
    }
    
    // Verify we found cells with formatting
    assert!(cells_with_formatting > 0, "Should have found cells with formatting");
    
    // Test that cells without explicit formatting still have default formatting
    for row in range.rows() {
        for cell in row.iter() {
            let data = cell.get_data();
            let formatting = cell.get_formatting();
            
            // Even if formatting is None, the data should still be accessible
            assert!(!data.to_string().is_empty() || matches!(data, Data::Empty));
            
            // If formatting is present, verify it has the expected structure
            if let Some(fmt) = formatting {
                // All formats should have a number format (check for valid variants)
                assert!(matches!(fmt.number_format, CellFormat::Other | CellFormat::DateTime));
            }
        }
    }
    
    println!("Total cells with formatting: {}", cells_with_formatting);
    println!("Total cells without formatting: {}", cells_without_formatting);
}

#[test]
fn test_comprehensive_formatting_format_xlsx() {
    let mut excel: Xlsx<_> = wb("format.xlsx");

    // === Part 1: Test cell-level formatting access ===

    let sheet_names = excel.sheet_names();
    let sheet_name = &sheet_names[0];
    assert_eq!(sheet_name, "Sheet1");

    let mut cell_reader = excel.worksheet_cells_reader(sheet_name).unwrap();

    // Read first cell - should be A1 with "White header on black text" and style 1
    let (cell_a1, formatting_a1) = cell_reader
        .next_cell_with_formatting()
        .expect("Should read first cell")
        .expect("First cell should exist");

    // Verify cell position and content
    assert_eq!(cell_a1.get_position(), (0, 0)); // A1
    if let DataRef::SharedString(text) = cell_a1.get_value() {
        assert_eq!(*text, "White header on black text");
    } else {
        panic!("A1 should contain a shared string");
    }

    // A1 should have formatting (style 1: white font on black background)
    let fmt_a1 = formatting_a1.expect("A1 should have formatting");
    let font_a1 = fmt_a1
        .font
        .as_ref()
        .expect("A1 should have font formatting");
    assert_eq!(
        font_a1.color,
        Some(Color::Argb {
            a: 255,
            r: 255,
            g: 255,
            b: 255
        })
    ); // White font

    let fill_a1 = fmt_a1
        .fill
        .as_ref()
        .expect("A1 should have fill formatting");
    assert_eq!(fill_a1.pattern_type, PatternType::Solid);
    assert_eq!(
        fill_a1.foreground_color,
        Some(Color::Argb {
            a: 255,
            r: 0,
            g: 0,
            b: 0
        })
    ); // Black background

    // Skip the remaining cells in row 1 and read A2 - should be "Right aligned" with style 4
    let mut found_a2 = false;
    let mut cell_a2_data = None;

    while let Ok(Some((cell, formatting))) = cell_reader.next_cell_with_formatting() {
        if cell.get_position() == (1, 0) {
            // A2
            cell_a2_data = Some((cell, formatting));
            found_a2 = true;
            break;
        }
    }

    assert!(found_a2, "Should find cell A2");
    let (cell_a2, formatting_a2) = cell_a2_data.unwrap();
    assert_eq!(cell_a2.get_position(), (1, 0)); // A2
    if let DataRef::SharedString(text) = cell_a2.get_value() {
        assert_eq!(*text, "Right aligned");
    } else {
        panic!("A2 should contain 'Right aligned'");
    }

    // A2 should have right alignment (style 4)
    let fmt_a2 = formatting_a2.expect("A2 should have formatting");
    let alignment_a2 = fmt_a2
        .alignment
        .as_ref()
        .expect("A2 should have alignment formatting");
    assert_eq!(alignment_a2.horizontal, Some(Arc::from("right")));

    // Test accessing formatting by index
    let format_0 = cell_reader
        .get_formatting_by_index(0)
        .expect("Should get format 0");
    assert_eq!(format_0.font.as_ref().unwrap().size, Some(10.0));

    let format_8 = cell_reader
        .get_formatting_by_index(8)
        .expect("Should get format 8");
    // Format 8 uses Comic Sans MS font
    assert_eq!(
        format_8.font.as_ref().unwrap().name,
        Some(Arc::from("Comic Sans MS"))
    );

    // === Part 2: Test all cell formats ===

    // Create a fresh instance to avoid borrow checker issues
    let excel_for_formats: Xlsx<_> = wb("format.xlsx");
    let formats = excel_for_formats.get_all_cell_formats();

    // Verify we have the expected number of cell formats (10 total: indices 0-9)
    assert_eq!(formats.len(), 10, "Should have exactly 10 cell formats");

    // Test Format 0: Default formatting with Arial 10pt, black color, bottom alignment
    let format_0 = &formats[0];
    assert_eq!(format_0.number_format, CellFormat::Other);

    let font_0 = format_0
        .font
        .as_ref()
        .expect("Format 0 should have font information");
    assert_eq!(font_0.name, Some(Arc::from("Arial")));
    assert_eq!(font_0.size, Some(10.0));
    assert_eq!(font_0.bold, None);
    assert_eq!(font_0.italic, None);
    assert_eq!(
        font_0.color,
        Some(Color::Argb {
            a: 255,
            r: 0,
            g: 0,
            b: 0
        })
    ); // Black

    let alignment_0 = format_0
        .alignment
        .as_ref()
        .expect("Format 0 should have alignment");
    assert_eq!(alignment_0.vertical, Some(Arc::from("bottom")));
    assert_eq!(alignment_0.wrap_text, Some(false));

    // Test Format 1: White font on black background
    let format_1 = &formats[1];
    let font_1 = format_1
        .font
        .as_ref()
        .expect("Format 1 should have font information");
    assert_eq!(font_1.name, Some(Arc::from("Arial")));
    assert_eq!(
        font_1.color,
        Some(Color::Argb {
            a: 255,
            r: 255,
            g: 255,
            b: 255
        })
    ); // White

    let fill_1 = format_1
        .fill
        .as_ref()
        .expect("Format 1 should have fill information");
    assert_eq!(fill_1.pattern_type, PatternType::Solid);
    assert_eq!(
        fill_1.foreground_color,
        Some(Color::Argb {
            a: 255,
            r: 0,
            g: 0,
            b: 0
        })
    ); // Black background

    // Test Format 3: White font on black background with wrap text
    let format_3 = &formats[3];
    let font_3 = format_3
        .font
        .as_ref()
        .expect("Format 3 should have font information");
    assert_eq!(font_3.name, Some(Arc::from("Arial")));
    assert_eq!(
        font_3.color,
        Some(Color::Argb {
            a: 255,
            r: 255,
            g: 255,
            b: 255
        })
    ); // White

    let alignment_3 = format_3
        .alignment
        .as_ref()
        .expect("Format 3 should have alignment");
    assert_eq!(alignment_3.wrap_text, Some(true));

    // Test Format 4: Right aligned
    let format_4 = &formats[4];
    let alignment_4 = format_4
        .alignment
        .as_ref()
        .expect("Format 4 should have alignment");
    assert_eq!(alignment_4.horizontal, Some(Arc::from("right")));

    // Test Format 5: Center aligned (both horizontal and vertical)
    let format_5 = &formats[5];
    let alignment_5 = format_5
        .alignment
        .as_ref()
        .expect("Format 5 should have alignment");
    assert_eq!(alignment_5.horizontal, Some(Arc::from("center")));
    assert_eq!(alignment_5.vertical, Some(Arc::from("center")));

    // Test Format 6: Custom currency format (numFmtId=164, detected as Other with custom format string)
    let format_6 = &formats[6];
    assert_eq!(format_6.number_format, CellFormat::Other);
    assert_eq!(
        format_6.format_string.as_ref().map(|s| s.as_ref()),
        Some("&quot;$&quot;#,##0.00"),
        "Format 6 should have format string"
    );

    // Test Format 7: Percentage format (built-in format 10)
    let format_7 = &formats[7];
    assert_eq!(format_7.number_format, CellFormat::Other);

    // Test Format 8: Comic Sans MS font
    let format_8 = &formats[8];
    let font_8 = format_8
        .font
        .as_ref()
        .expect("Format 8 should have font information");
    assert_eq!(font_8.name, Some(Arc::from("Comic Sans MS")));
    assert_eq!(
        font_8.color,
        Some(Color::Theme {
            theme: 1,
            tint: None
        })
    );

    // === Part 3: Test specific color patterns ===

    // Test specific color combinations that should exist in format.xlsx
    let black_fill_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.fill.as_ref().is_some_and(|fill| {
                fill.foreground_color
                    == Some(Color::Argb {
                        a: 255,
                        r: 0,
                        g: 0,
                        b: 0,
                    })
            })
        })
        .collect();
    assert!(
        !black_fill_formats.is_empty(),
        "Should have black filled formats"
    );

    // Test font color variations - white fonts
    let white_font_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().is_some_and(|font| {
                font.color
                    == Some(Color::Argb {
                        a: 255,
                        r: 255,
                        g: 255,
                        b: 255,
                    })
            })
        })
        .collect();
    assert!(
        !white_font_formats.is_empty(),
        "Should have white font formats"
    );

    // Test black fonts
    let black_font_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().is_some_and(|font| {
                font.color
                    == Some(Color::Argb {
                        a: 255,
                        r: 0,
                        g: 0,
                        b: 0,
                    })
            })
        })
        .collect();
    assert!(
        !black_font_formats.is_empty(),
        "Should have black font formats"
    );

    // Test theme color usage
    let theme_color_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().is_some_and(|font| {
                matches!(
                    font.color,
                    Some(Color::Theme {
                        theme: 1,
                        tint: None
                    })
                )
            })
        })
        .collect();
    assert!(
        !theme_color_formats.is_empty(),
        "Should have theme color formats"
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
            ("MyBrokenRange".to_string(), "'Sheet1'!#REF!".to_string()),
            (
                "MyDataTypes".to_string(),
                "'datatypes'!$A$1:$A$6".to_string()
            ),
            ("OneRange".to_string(), "'Sheet1'!$A$1".to_string()),
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
    assert_eq!(b.unwrap().get_data(), &Float(1.));
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
        let p = format!("{root}/tests/issue127.{ext}");
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
    assert_eq!(range.get_value((6, 2)).unwrap().get_data(), &Float(9.));
}

#[test]
fn skip_phonetic_text() {
    let mut xls: Xlsx<_> = wb("rph.xlsx");
    let range = xls.worksheet_range("Sheet1").unwrap();
    assert_eq!(
        range.get_value((0, 0)).unwrap().get_data(),
        &String("課きく　毛こ".to_string())
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
    assert_eq!(data.get((0, 0)).unwrap().get_data(), &String("celsius".to_owned()));
    assert_eq!(data.get((1, 0)).unwrap().get_data(), &String("fahrenheit".to_owned()));
    assert_eq!(data.get((0, 1)).unwrap().get_data(), &Float(22.2222));
    assert_eq!(data.get((1, 1)).unwrap().get_data(), &Float(72.0));
    // Check the second table
    let table = xls
        .table_by_name("OtherTable")
        .expect("Parsing table's sheet should not error");
    assert_eq!(table.name(), "OtherTable");
    assert_eq!(table.columns()[0], "label2");
    assert_eq!(table.columns()[1], "value2");
    let data = table.data();
    assert_eq!(data.get((0, 0)).unwrap().get_data(), &String("something".to_owned()));
    assert_eq!(data.get((1, 0)).unwrap().get_data(), &String("else".to_owned()));
    assert_eq!(data.get((0, 1)).unwrap().get_data(), &Float(12.5));
    assert_eq!(data.get((1, 1)).unwrap().get_data(), &Float(64.0));
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.632060185185185,
            ExcelDateTimeType::TimeDelta,
            false
        ))
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.632060185185185,
            ExcelDateTimeType::TimeDelta,
            true
        ))
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            false
        ))
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            true
        ))
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTimeIso("2021-01-01".to_string())
    );
    assert_eq!(
        range.get_value((1, 0)).unwrap().get_data(),
        &DateTimeIso("2021-01-01T10:10:10".to_string())
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTimeIso("10:10:10".to_string())
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTimeIso("2021-01-01".to_string())
    );
    assert_eq!(
        range.get_value((1, 0)).unwrap().get_data(),
        &DateTimeIso("2021-01-01T10:10:10".to_string())
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DurationIso("PT10H10M10S".to_string())
    );
    assert_eq!(
        range.get_value((3, 0)).unwrap().get_data(),
        &DurationIso("PT10H10M10.123456S".to_string())
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            44197.0,
            ExcelDateTimeType::DateTime,
            false
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            false
        ))
    );

    #[cfg(feature = "dates")]
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
        range.get_value((0, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            42735.0,
            ExcelDateTimeType::DateTime,
            true
        ))
    );
    assert_eq!(
        range.get_value((2, 0)).unwrap().get_data(),
        &DateTime(ExcelDateTime::new(
            10.6320601851852,
            ExcelDateTimeType::TimeDelta,
            true
        ))
    );

    #[cfg(feature = "dates")]
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
    assert_eq!(rows.next().unwrap()[0].get_data().to_string(), "A1*2");
    assert_eq!(rows.next().unwrap()[0].get_data().to_string(), "2*Sheet2!A1");
    assert_eq!(rows.next().unwrap()[0].get_data().to_string(), "A1+Sheet2!A1");
    assert_eq!(rows.next(), None);
}

#[test]
fn issue304_xls_values() {
    let mut wb: Xls<_> = wb("xls_formula.xls");
    let rge = wb.worksheet_range("Sheet1").unwrap();
    let mut rows = rge.rows();
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::Float(10.));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::Float(20.));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::Float(110.));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::Float(65.));
    assert_eq!(rows.next(), None);
}

#[test]
fn issue334_xls_values_string() {
    let mut wb: Xls<_> = wb("xls_ref_String.xls");
    let rge = wb.worksheet_range("Sheet1").unwrap();
    let mut rows = rge.rows();
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::String("aa".into()));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::String("bb".into()));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::String("aa".into()));
    assert_eq!(rows.next().unwrap()[0].get_data(), &Data::String("bb".into()));
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
        (0, 0, DataWithFormatting::new(Data::Float(23.), None)),
        (0, 2, DataWithFormatting::new(Data::Float(23.), None)),
        (12, 6, DataWithFormatting::new(Data::Float(2.), None)),
        (13, 9, DataWithFormatting::new(Data::String("US".into()), None)),
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
    let mut expect = Range::<DataWithFormatting>::new((1, 0), (6, 0));
    for (i, cell) in ["A1+1", "A2+1", "A3+1", "A4+1", "A5+1", "A6+1"]
        .iter()
        .enumerate()
    {
        expect.set_value((1 + i as u32, 0), DataWithFormatting::new(Data::String(cell.to_string()), None));
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

#[test]
fn test_malformed_format() {
    let _xls: Xls<_> = wb("malformed_format.xls");
}

#[test]
fn test_sheet_ref_quote_xlsb() {
    let mut excel: Xlsb<_> = wb("sheet_ref_quote.xlsb");

    let formula_range = excel.worksheet_formula("Sheet1").unwrap();

    range_eq!(formula_range, [["'1'!A1"], ["'G&A'!A1"], ["'a''b'!A1"]]);
}

#[test]
fn test_sheet_ref_error_xlsb() {
    let mut excel: Xlsb<_> = wb("sheet_ref_error.xlsb");

    let formula_range = excel.worksheet_formula("Sheet1").unwrap();

    range_eq!(formula_range, [["'#InvalidWorkSheet'!A1".to_string()]]);
}

#[test]
fn test_choose_xlsb() {
    let mut excel: Xlsb<_> = wb("choose.xlsb");

    let formulas = excel.worksheet_formula("Sheet1").unwrap();

    range_eq!(
        formulas,
        [["CHOOSE(A1,IF(ISERROR(1/0),B1+C1,B2+C2),IF(RAND()>=0.5,B3+C3,B4+C4),B5+C5,B6+C6)"],]
    );
}

#[test]
fn test_escape_quote_xlsb() {
    let mut excel: Xlsb<_> = wb("escape_quote.xlsb");

    let formulas = excel.worksheet_formula("Sheet1").unwrap();

    range_eq!(formulas, [["\"ab\"\"cd\""],]);
}

#[test]
fn test_advanced_formatting_features_format_xlsx() {
    let excel: Xlsx<_> = wb("format.xlsx");
    let formats = excel.get_all_cell_formats();

    // Test number format variations
    let percentage_formats: Vec<_> = formats
        .iter()
        .filter(|f| f.number_format == CellFormat::Other)
        .collect();
    assert!(
        !percentage_formats.is_empty(),
        "Should have Percentage formats"
    );

    let other_formats: Vec<_> = formats
        .iter()
        .filter(|f| matches!(f.number_format, CellFormat::Other))
        .collect();
    assert!(
        !other_formats.is_empty(),
        "Should have Other/general formats"
    );

    // Test specific font sizes that should exist (10pt)
    let font_10pt_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font
                .as_ref()
                .map_or(false, |font| font.size == Some(10.0))
        })
        .collect();
    assert!(!font_10pt_formats.is_empty(), "Should have 10pt fonts");

    // Test Arial font (should be the primary font)
    let arial_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font
                .as_ref()
                .map_or(false, |font| font.name == Some(Arc::from("Arial")))
        })
        .collect();
    assert!(!arial_formats.is_empty(), "Should have Arial fonts");

    // Test Comic Sans MS font (format 8)
    let comic_sans_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font
                .as_ref()
                .map_or(false, |font| font.name == Some(Arc::from("Comic Sans MS")))
        })
        .collect();
    assert!(
        !comic_sans_formats.is_empty(),
        "Should have Comic Sans MS fonts"
    );

    // Test border functionality (even if borders are mostly empty in this file)
    let formats_with_borders: Vec<_> = formats.iter().filter(|f| f.border.is_some()).collect();
    assert!(
        !formats_with_borders.is_empty(),
        "Should have border structures"
    );

    // Test specific color values that exist in format.xlsx
    let white_font_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().map_or(false, |font| {
                font.color
                    == Some(Color::Argb {
                        a: 255,
                        r: 255,
                        g: 255,
                        b: 255,
                    })
            })
        })
        .collect();
    assert!(
        !white_font_formats.is_empty(),
        "Should have white font formats"
    );

    // Test black font colors
    let black_font_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().map_or(false, |font| {
                font.color
                    == Some(Color::Argb {
                        a: 255,
                        r: 0,
                        g: 0,
                        b: 0,
                    })
            })
        })
        .collect();
    assert!(
        !black_font_formats.is_empty(),
        "Should have black font formats"
    );

    // Test theme color usage
    let theme_color_formats: Vec<_> = formats
        .iter()
        .filter(|f| {
            f.font.as_ref().map_or(false, |font| {
                matches!(
                    font.color,
                    Some(Color::Theme {
                        theme: 1,
                        tint: None
                    })
                )
            })
        })
        .collect();
    assert!(
        !theme_color_formats.is_empty(),
        "Should have theme color formats"
    );
}
