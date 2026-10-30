// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use crate::datatype::{Data, DataRef, ExcelDateTime, ExcelDateTimeType};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CellFormat {
    Other,
    DateTime,
    TimeDelta,
}

/// Check excel number format is datetime
pub fn detect_custom_number_format(format: &str) -> CellFormat {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;
    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            // `\` escapes, `_` skips width, `*` fills; the next char is always a literal, not a format token.
            ('_' | '\\' | '*', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return CellFormat::Other, // first format only
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return CellFormat::TimeDelta, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return CellFormat::DateTime,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return CellFormat::DateTime
            }
            _ => {
                if hms && s.eq_ignore_ascii_case(&prev) {
                    // ok ...
                } else {
                    hms = prev == '[' && matches!(s, 'm' | 'h' | 's' | 'M' | 'H' | 'S');
                }
            }
        }
        prev = s;
    }
    CellFormat::Other
}

pub fn builtin_format_by_id(id: &str) -> CellFormat {
    match id {
        // mm-dd-yy
        "14" |
        // d-mmm-yy
        "15" |
        // d-mmm
        "16" |
        // mmm-yy
        "17" |
        // h:mm AM/PM
        "18" |
        // h:mm:ss AM/PM
        "19" |
        // h:mm
        "20" |
        // h:mm:ss
        "21" |
        // m/d/yy h:mm
        "22" |
        // mm:ss
        "45" |
        // mmss.0
        "47" => CellFormat::DateTime,
        // East Asian built-in formats vary by locale and include dates and
        // clock times. Both use DateTime, rather than the duration category.
        "27" | "28" | "29" | "30" | "31" | "32" | "33" | "34" | "35" | "36" | "50" | "51"
        | "52" | "53" | "54" | "55" | "56" | "57" | "58" => CellFormat::DateTime,
        // [h]:mm:ss
        "46" => CellFormat::TimeDelta,
        _ => CellFormat::Other
    }
}

// Get the format code string for a builtin number format id.
//
// The ids and format codes are defined in ECMA-376 section 18.8.30. Returns
// `None` for ids that are not built in, such as custom formats and the reserved and
// locale-specific formats.
pub(crate) fn builtin_format_code_by_id(id: u32) -> Option<&'static str> {
    match id {
        0 => Some("General"),
        1 => Some("0"),
        2 => Some("0.00"),
        3 => Some("#,##0"),
        4 => Some("#,##0.00"),
        9 => Some("0%"),
        10 => Some("0.00%"),
        11 => Some("0.00E+00"),
        12 => Some("# ?/?"),
        13 => Some("# ??/??"),
        14 => Some("mm-dd-yy"),
        15 => Some("d-mmm-yy"),
        16 => Some("d-mmm"),
        17 => Some("mmm-yy"),
        18 => Some("h:mm AM/PM"),
        19 => Some("h:mm:ss AM/PM"),
        20 => Some("h:mm"),
        21 => Some("h:mm:ss"),
        22 => Some("m/d/yy h:mm"),
        37 => Some("#,##0 ;(#,##0)"),
        38 => Some("#,##0 ;[Red](#,##0)"),
        39 => Some("#,##0.00;(#,##0.00)"),
        40 => Some("#,##0.00;[Red](#,##0.00)"),
        41 => Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
        42 => Some("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"),
        43 => Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
        44 => Some("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
        45 => Some("mm:ss"),
        46 => Some("[h]:mm:ss"),
        47 => Some("mmss.0"),
        48 => Some("##0.0E+0"),
        49 => Some("@"),
        _ => None,
    }
}

/// Check if code corresponds to builtin date format
///
/// See `builtin_format_by_id`
pub fn builtin_format_by_code(code: u16) -> CellFormat {
    match code {
        14..=22 | 27..=36 | 45 | 47 | 50..=58 => CellFormat::DateTime,
        46 => CellFormat::TimeDelta,
        _ => CellFormat::Other,
    }
}

// Convert an i64 to a date if the cell format is a date or time format.
#[inline]
pub fn format_excel_i64_ref(
    value: i64,
    format: Option<&CellFormat>,
    is_1904: bool,
) -> DataRef<'static> {
    match format {
        Some(CellFormat::DateTime | CellFormat::TimeDelta) => {
            format_excel_f64_ref(value as f64, format, is_1904)
        }
        _ => DataRef::Int(value),
    }
}

// Convert an i64 to a date if the cell format is a date or time format.
pub fn format_excel_i64(value: i64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    format_excel_i64_ref(value, format, is_1904).into()
}

// Convert an f64 to a date if the cell format is a date or time format.
#[inline]
pub fn format_excel_f64_ref(
    value: f64,
    format: Option<&CellFormat>,
    is_1904: bool,
) -> DataRef<'static> {
    match format {
        Some(CellFormat::DateTime) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::DateTime,
            is_1904,
        )),
        Some(CellFormat::TimeDelta) => DataRef::DateTime(ExcelDateTime::new(
            value,
            ExcelDateTimeType::TimeDelta,
            is_1904,
        )),
        _ => DataRef::Float(value),
    }
}

// Convert an f64 to a date if the cell format is a date or time format.
pub fn format_excel_f64(value: f64, format: Option<&CellFormat>, is_1904: bool) -> Data {
    format_excel_f64_ref(value, format, is_1904).into()
}

/// Ported from openpyxl, MIT License
/// https://foss.heptapod.net/openpyxl/openpyxl/-/blob/a5e197c530aaa49814fd1d993dd776edcec35105/openpyxl/styles/tests/test_number_style.py
#[test]
fn test_is_date_format() {
    assert_eq!(
        detect_custom_number_format("DD/MM/YY"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("H:MM:SS;@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$\\u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("m\"M\"d\"D\";@"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm:ss"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("#,##0\\ [$''u20bd-46D]"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("0_ ;[Red]\\-0\\ "),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("\\Y000000"), CellFormat::Other);
    assert_eq!(
        detect_custom_number_format("#,##0.0####\" YMD\""),
        CellFormat::Other
    );
    assert_eq!(detect_custom_number_format("[h]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[ss]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[s].000"),
        CellFormat::TimeDelta
    );
    assert_eq!(detect_custom_number_format("[m]"), CellFormat::TimeDelta);
    assert_eq!(detect_custom_number_format("[mm]"), CellFormat::TimeDelta);
    assert_eq!(
        detect_custom_number_format("[Blue]\\+[h]:mm;[Red]\\-[h]:mm;[Green][h]:mm"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta][s].00"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[h]:mm;[=0]\\-"),
        CellFormat::TimeDelta
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta].00"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("[>=100][Magenta]General"),
        CellFormat::Other
    );
    assert_eq!(
        detect_custom_number_format("ha/p\\\\m"),
        CellFormat::DateTime
    );
    assert_eq!(
        detect_custom_number_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
        CellFormat::Other
    );
    // The `*` fill operator repeats the next character to fill the cell, so
    // that character is a literal and must not be read as a date token even
    // when it happens to be one (here `y`, `d`, `m`).
    assert_eq!(detect_custom_number_format("#,##0*y"), CellFormat::Other);
    assert_eq!(detect_custom_number_format("0\"x\"*d"), CellFormat::Other);
    assert_eq!(detect_custom_number_format("*-#,##0"), CellFormat::Other);
    // A real date token after the single fill char is still detected.
    assert_eq!(
        detect_custom_number_format("*-yyyy-mm-dd"),
        CellFormat::DateTime
    );
}
