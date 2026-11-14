// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Internal module providing handy functions

use std::borrow::Cow;
use std::collections::HashMap;
use std::io::{Read, Seek};

use quick_xml::{escape::resolve_xml_entity, events::BytesRef};
use zip::read::ZipArchive;
use zip::result::ZipResult;

const UNICODE_ESCAPE_LENGTH: usize = 7; // Length of _xHHHH_.

macro_rules! from_err {
    ($from:ty, $to:tt, $var:tt) => {
        impl From<$from> for $to {
            fn from(e: $from) -> $to {
                $to::$var(e.into())
            }
        }
    };
}

#[inline]
pub fn read_u32(s: &[u8]) -> u32 {
    u32::from_le_bytes(s[..4].try_into().unwrap())
}

#[inline]
pub fn read_i32(s: &[u8]) -> i32 {
    i32::from_le_bytes(s[..4].try_into().unwrap())
}

#[inline]
pub fn read_u16(s: &[u8]) -> u16 {
    u16::from_le_bytes(s[..2].try_into().unwrap())
}

#[inline]
pub fn read_i16(s: &[u8]) -> i16 {
    i16::from_le_bytes(s[..2].try_into().unwrap())
}

#[inline]
pub fn read_usize(s: &[u8]) -> usize {
    read_u32(s).try_into().unwrap()
}

#[inline]
pub fn read_f64(s: &[u8]) -> f64 {
    f64::from_le_bytes(s[..8].try_into().unwrap())
}

/// Convert a 0-based column index to an Excel column name (0 -> "A", 26 -> "AA").
pub fn push_column(col: u32, buf: &mut String) {
    let mut digits = [0u8; 6];
    let len = column_name_digits(col, &mut digits);
    for &d in &digits[..len] {
        buf.push(d as char);
    }
}

/// Write the Excel letters for a 0-based column index into `out` (most
/// significant first) and return how many were written.
pub(crate) fn column_name_digits(col: u32, out: &mut [u8]) -> usize {
    let mut num = col + 1;
    let mut n = 0;
    while num > 0 {
        out[n] = b'A' + ((num - 1) % 26) as u8;
        n += 1;
        num = (num - 1) / 26;
    }
    out[..n].reverse();
    n
}

/// Quote a sheet name for a formula, such as `'My Sheet'`, if Excel would.
///
/// An apostrophe in a quoted name is doubled. The rules are the same as
/// `quote_sheet_name()` in `rust_xlsxwriter`, plus the quoting of `TRUE` and
/// `FALSE`.
pub(crate) fn quote_sheet_name(name: &str) -> Cow<'_, str> {
    if sheet_name_needs_quotes(name) {
        Cow::Owned(format!("'{}'", name.replace('\'', "''")))
    } else {
        Cow::Borrowed(name)
    }
}

/// Check if Excel quotes a sheet name in a formula.
fn sheet_name_needs_quotes(name: &str) -> bool {
    const ROW_MAX: u64 = 1_048_576;
    const COL_MAX: u64 = 16_384;

    // Rule 1. Names with characters other than letters, digits, `_`, `.` and
    // emoji are quoted.
    if !name
        .chars()
        .all(|c| c.is_alphanumeric() || c == '_' || c == '.' || is_emoji(c))
    {
        return true;
    }

    // Rule 2. Names that start with a digit, `.` or an emoji are quoted.
    if name.is_empty() || name.starts_with(|c: char| c.is_ascii_digit() || c == '.' || is_emoji(c))
    {
        return true;
    }

    // The boolean values are quoted.
    if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
        return true;
    }

    // Split names like `A1` and `R1C1` into the letters before the first digit
    // and the rest. Names without a digit are only quoted if they are an R1C1
    // reference to the current row or column.
    let Some(position) = name.find(|c: char| c.is_ascii_digit()) else {
        return ["R", "C", "RC"]
            .iter()
            .any(|r| name.eq_ignore_ascii_case(r));
    };
    let (letters, number) = name.split_at(position);
    if !letters.chars().all(|c| c.is_ascii_alphabetic()) {
        return false;
    }
    let letters = letters.to_ascii_uppercase();
    let digits = number.bytes().take_while(u8::is_ascii_digit).count();
    let value = number[..digits].parse::<u64>().unwrap_or_default();

    // Rule 3. Names that are an A1 reference within Excel's limits are quoted.
    if letters.len() <= 3 && digits == number.len() {
        let col = letters
            .bytes()
            .fold(0, |col, b| col * 26 + u64::from(b - b'A' + 1));
        if (1..=ROW_MAX).contains(&value) && col <= COL_MAX {
            return true;
        }
    }

    // Rule 4. Names that start with an R1C1 reference within Excel's limits are
    // quoted. Characters after the reference are ignored.
    match letters.as_str() {
        "R" => (1..=ROW_MAX).contains(&value),
        "C" | "RC" => (1..=COL_MAX).contains(&value),
        _ => false,
    }
}

// Match the emoji characters that Excel allows in an unquoted sheet name.
fn is_emoji(c: char) -> bool {
    if c < '\u{203C}' {
        // Shortcut for most chars in the lower range. We ignore '#', '*',
        // '0-9', '©️' and '®️' which are in this range and which are, strictly
        // speaking, emoji symbols but they are not treated so by Excel in the
        // context of this check.
        return false;
    }

    if c < '\u{01F004}' {
        return matches!(c,
            '\u{203C}' | '\u{2049}' | '\u{2122}' | '\u{2139}' | '\u{2194}'..='\u{2199}' |
            '\u{21A9}' | '\u{21AA}' | '\u{231A}' | '\u{231B}' | '\u{2328}' | '\u{23CF}' |
            '\u{23E9}'..='\u{23F3}' | '\u{23F8}'..='\u{23FA}' | '\u{24C2}' | '\u{25AA}' |
            '\u{25AB}' | '\u{25B6}' | '\u{25C0}' | '\u{25FB}'..='\u{25FE}' |
            '\u{2600}'..='\u{2604}' | '\u{260E}' | '\u{2611}' | '\u{2614}' | '\u{2615}' |
            '\u{2618}' | '\u{261D}' | '\u{2620}' | '\u{2622}' | '\u{2623}' | '\u{2626}' |
            '\u{262A}' | '\u{262E}' | '\u{262F}' | '\u{2638}'..='\u{263A}' | '\u{2640}' |
            '\u{2642}' | '\u{2648}'..='\u{2653}' | '\u{265F}' | '\u{2660}' | '\u{2663}' |
            '\u{2665}' | '\u{2666}' | '\u{2668}' | '\u{267B}' | '\u{267E}' | '\u{267F}' |
            '\u{2692}'..='\u{2697}' | '\u{2699}' | '\u{269B}' | '\u{269C}' | '\u{26A0}' |
            '\u{26A1}' | '\u{26A7}' | '\u{26AA}' | '\u{26AB}' | '\u{26B0}' | '\u{26B1}' |
            '\u{26BD}' | '\u{26BE}' | '\u{26C4}' | '\u{26C5}' | '\u{26C8}' | '\u{26CE}' |
            '\u{26CF}' | '\u{26D1}' | '\u{26D3}' | '\u{26D4}' | '\u{26E9}' | '\u{26EA}' |
            '\u{26F0}'..='\u{26F5}' | '\u{26F7}'..='\u{26FA}' | '\u{26FD}' | '\u{2702}' |
            '\u{2705}' | '\u{2708}'..='\u{270D}' | '\u{270F}' | '\u{2712}' | '\u{2714}' |
            '\u{2716}' | '\u{271D}' | '\u{2721}' | '\u{2728}' | '\u{2733}' | '\u{2734}' |
            '\u{2744}' | '\u{2747}' | '\u{274C}' | '\u{274E}' | '\u{2753}'..='\u{2755}' |
            '\u{2757}' | '\u{2763}' | '\u{2764}' | '\u{2795}'..='\u{2797}' | '\u{27A1}' |
            '\u{27B0}' | '\u{27BF}' | '\u{2934}' | '\u{2935}' | '\u{2B05}'..='\u{2B07}' |
            '\u{2B1B}' | '\u{2B1C}' | '\u{2B50}' | '\u{2B55}' | '\u{3030}' | '\u{303D}' |
            '\u{3297}' | '\u{3299}'
        );
    }

    matches!(c,
        '\u{01F004}' | '\u{01F0CF}' | '\u{01F170}' | '\u{01F171}' | '\u{01F17E}' | '\u{01F17F}' |
        '\u{01F18E}' | '\u{01F191}'..='\u{01F19A}' | '\u{01F1E6}'..='\u{01F1FF}' | '\u{01F201}' |
        '\u{01F202}' | '\u{01F21A}' | '\u{01F22F}' | '\u{01F232}'..='\u{01F23A}' | '\u{01F250}' |
        '\u{01F251}' | '\u{01F300}'..='\u{01F321}' | '\u{01F324}'..='\u{01F393}' | '\u{01F396}' |
        '\u{01F397}' | '\u{01F399}'..='\u{01F39B}' | '\u{01F39E}'..='\u{01F3F0}' |
        '\u{01F3F3}'..='\u{01F3F5}' | '\u{01F3F7}'..='\u{01F4FD}' | '\u{01F4FF}'..='\u{01F53D}' |
        '\u{01F549}'..='\u{01F54E}' | '\u{01F550}'..='\u{01F567}' | '\u{01F56F}' | '\u{01F570}' |
        '\u{01F573}'..='\u{01F57A}' | '\u{01F587}' | '\u{01F58A}'..='\u{01F58D}' | '\u{01F590}' |
        '\u{01F595}' | '\u{01F596}' | '\u{01F5A4}' | '\u{01F5A5}' | '\u{01F5A8}' | '\u{01F5B1}' |
        '\u{01F5B2}' | '\u{01F5BC}' | '\u{01F5C2}'..='\u{01F5C4}' | '\u{01F5D1}'..='\u{01F5D3}' |
        '\u{01F5DC}'..='\u{01F5DE}' | '\u{01F5E1}' | '\u{01F5E3}' | '\u{01F5E8}' | '\u{01F5EF}' |
        '\u{01F5F3}' | '\u{01F5FA}'..='\u{01F64F}' | '\u{01F680}'..='\u{01F6C5}' |
        '\u{01F6CB}'..='\u{01F6D2}' | '\u{01F6D5}'..='\u{01F6D7}' | '\u{01F6DC}'..='\u{01F6E5}' |
        '\u{01F6E9}' | '\u{01F6EB}' | '\u{01F6EC}' | '\u{01F6F0}' | '\u{01F6F3}'..='\u{01F6FC}' |
        '\u{01F7E0}'..='\u{01F7EB}' | '\u{01F7F0}' | '\u{01F90C}'..='\u{01F93A}' |
        '\u{01F93C}'..='\u{01F945}' | '\u{01F947}'..='\u{01F9FF}' | '\u{01FA70}'..='\u{01FA7C}' |
        '\u{01FA80}'..='\u{01FA88}' | '\u{01FA90}'..='\u{01FABD}' | '\u{01FABF}'..='\u{01FAC5}' |
        '\u{01FACE}'..='\u{01FADB}' | '\u{01FAE0}'..='\u{01FAE8}' | '\u{01FAF0}'..='\u{01FAF8}'
    )
}

// Utility function to unescape standard XML entities or character references
// generated by a `quick_xml::Event::GeneralRef()`. Appends the result to the
// provided buffer or returns a `quick_xml::Error`.
#[inline]
pub(crate) fn unescape_entity_to_buffer(
    entity: &BytesRef,
    buffer: &mut String,
) -> Result<(), quick_xml::Error> {
    let decoded = entity.as_ref();

    // Handle standard XML entities directly.
    if let Some(unescaped_xml_entity) = resolve_xml_entity(decoded) {
        buffer.push_str(unescaped_xml_entity);
        return Ok(());
    }

    // Handle character references like `#xA0`.
    if let Some(unescaped_char) = entity.resolve_char_ref()? {
        buffer.push(unescaped_char);
        return Ok(());
    }

    // Raise a quick_xml::Error if we failed to unescape the entity. The
    // location information isn't useful to the end user in this context so we
    // return `0..0`.
    Err(quick_xml::Error::Escape(
        quick_xml::escape::EscapeError::UnrecognizedEntity(0..0, format!("&{decoded};")),
    ))
}

// Function to unescape Excel XML escapes in a string. Excel encodes a character
// like "\r" as "_x000D_", where the 4 hex digits are a UTF-16 code unit, so a
// character outside the BMP is encoded as a surrogate pair like
// "_xD83D__xDCF7_". In turn it escapes the literal string "_x000D_" as
// "_x005F_x000D_".
//
// Most strings don't contain escapes so the function returns the original
// string without allocating unless an escape is found.
pub(crate) fn unescape_xml(original: &str) -> Cow<'_, str> {
    let Some((mut index, mut unit)) = find_xml_escape(original, 0) else {
        return Cow::Borrowed(original);
    };

    let mut unescaped = String::with_capacity(original.len());
    let mut copied = 0;

    loop {
        // Copy the literal text before the escape.
        unescaped.push_str(&original[copied..index]);
        index += UNICODE_ESCAPE_LENGTH;

        match unit {
            0xD800..=0xDBFF => {
                // A high surrogate must be followed by a low surrogate escape
                // to form a character. Otherwise it is replaced by U+FFFD.
                match parse_xml_escape(original, index) {
                    Some(low @ 0xDC00..=0xDFFF) => {
                        let code_point = 0x10000
                            + ((u32::from(unit) - 0xD800) << 10)
                            + (u32::from(low) - 0xDC00);
                        unescaped.push(
                            char::from_u32(code_point).unwrap_or(char::REPLACEMENT_CHARACTER),
                        );
                        index += UNICODE_ESCAPE_LENGTH;
                    }
                    _ => unescaped.push(char::REPLACEMENT_CHARACTER),
                }
            }
            0xDC00..=0xDFFF => unescaped.push(char::REPLACEMENT_CHARACTER),
            _ => unescaped
                .push(char::from_u32(u32::from(unit)).unwrap_or(char::REPLACEMENT_CHARACTER)),
        }

        copied = index;

        match find_xml_escape(original, index) {
            Some(next) => (index, unit) = next,
            None => break,
        }
    }

    unescaped.push_str(&original[copied..]);
    Cow::Owned(unescaped)
}

// Find the first "_xHHHH_" escape at or after the byte index `start`. Returns
// the index of the escape and its UTF-16 code unit.
fn find_xml_escape(original: &str, mut start: usize) -> Option<(usize, u16)> {
    while let Some(offset) = original[start..].find('_') {
        let index = start + offset;
        if let Some(unit) = parse_xml_escape(original, index) {
            return Some((index, unit));
        }
        start = index + 1;
    }

    None
}

// Parse an "_xHHHH_" escape at the byte index `index` and return its UTF-16
// code unit. Returns None if the bytes at the index aren't a valid escape.
fn parse_xml_escape(original: &str, index: usize) -> Option<u16> {
    let bytes = original
        .as_bytes()
        .get(index..index + UNICODE_ESCAPE_LENGTH)?;

    if bytes[0] != b'_' || bytes[1] != b'x' || bytes[6] != b'_' {
        return None;
    }

    let hex = &bytes[2..6];
    if !hex.iter().all(u8::is_ascii_hexdigit) {
        return None;
    }

    // The digits are ASCII so the slice is valid UTF-8 and fits in a u16.
    u16::from_str_radix(std::str::from_utf8(hex).ok()?, 16).ok()
}

/// Build a lookup cache mapping lowercased normalized paths to original ZIP entry names.
pub fn build_zip_path_cache<RS: Read + Seek>(
    zip: &ZipArchive<RS>,
) -> ZipResult<HashMap<String, String>> {
    let mut cache = HashMap::with_capacity(zip.len());
    for zip_path in zip.file_names() {
        let zip_path = zip_path?;
        let normalized = zip_path.replace('\\', "/").to_ascii_lowercase();
        cache.insert(normalized, zip_path.to_string());
    }
    Ok(cache)
}

/// Look up a path in the cache, falling back to the path itself.
pub fn cached_zip_path<'a>(cache: &'a HashMap<String, String>, path: &'a str) -> &'a str {
    let key = path.to_ascii_lowercase();
    cache.get(&key).map(|s| s.as_str()).unwrap_or(path)
}

pub const FTAB_LEN: usize = 485;

/* [MS-XLS] 2.5.198.17 */
/* [MS-XLSB] 2.5.97.10 */
pub const FTAB: [&str; FTAB_LEN] = [
    "COUNT",
    "IF",
    "ISNA",
    "ISERROR",
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "ROW",
    "COLUMN",
    "NA",
    "NPV",
    "STDEV",
    "DOLLAR",
    "FIXED",
    "SIN",
    "COS",
    "TAN",
    "ATAN",
    "PI",
    "SQRT",
    "EXP",
    "LN",
    "LOG10",
    "ABS",
    "INT",
    "SIGN",
    "ROUND",
    "LOOKUP",
    "INDEX",
    "REPT",
    "MID",
    "LEN",
    "VALUE",
    "TRUE",
    "FALSE",
    "AND",
    "OR",
    "NOT",
    "MOD",
    "DCOUNT",
    "DSUM",
    "DAVERAGE",
    "DMIN",
    "DMAX",
    "DSTDEV",
    "VAR",
    "DVAR",
    "TEXT",
    "LINEST",
    "TREND",
    "LOGEST",
    "GROWTH",
    "GOTO",
    "HALT",
    "RETURN",
    "PV",
    "FV",
    "NPER",
    "PMT",
    "RATE",
    "MIRR",
    "IRR",
    "RAND",
    "MATCH",
    "DATE",
    "TIME",
    "DAY",
    "MONTH",
    "YEAR",
    "WEEKDAY",
    "HOUR",
    "MINUTE",
    "SECOND",
    "NOW",
    "AREAS",
    "ROWS",
    "COLUMNS",
    "OFFSET",
    "ABSREF",
    "RELREF",
    "ARGUMENT",
    "SEARCH",
    "TRANSPOSE",
    "ERROR",
    "STEP",
    "TYPE",
    "ECHO",
    "SET.NAME",
    "CALLER",
    "DEREF",
    "WINDOWS",
    "SERIES",
    "DOCUMENTS",
    "ACTIVE.CELL",
    "SELECTION",
    "RESULT",
    "ATAN2",
    "ASIN",
    "ACOS",
    "CHOOSE",
    "HLOOKUP",
    "VLOOKUP",
    "LINKS",
    "INPUT",
    "ISREF",
    "GET.FORMULA",
    "GET.NAME",
    "SET.VALUE",
    "LOG",
    "EXEC",
    "CHAR",
    "LOWER",
    "UPPER",
    "PROPER",
    "LEFT",
    "RIGHT",
    "EXACT",
    "TRIM",
    "REPLACE",
    "SUBSTITUTE",
    "CODE",
    "NAMES",
    "DIRECTORY",
    "FIND",
    "CELL",
    "ISERR",
    "ISTEXT",
    "ISNUMBER",
    "ISBLANK",
    "T",
    "N",
    "FOPEN",
    "FCLOSE",
    "FSIZE",
    "FREADLN",
    "FREAD",
    "FWRITELN",
    "FWRITE",
    "FPOS",
    "DATEVALUE",
    "TIMEVALUE",
    "SLN",
    "SYD",
    "DDB",
    "GET.DEF",
    "REFTEXT",
    "TEXTREF",
    "INDIRECT",
    "REGISTER",
    "CALL",
    "ADD.BAR",
    "ADD.MENU",
    "ADD.COMMAND",
    "ENABLE.COMMAND",
    "CHECK.COMMAND",
    "RENAME.COMMAND",
    "SHOW.BAR",
    "DELETE.MENU",
    "DELETE.COMMAND",
    "GET.CHART.ITEM",
    "DIALOG.BOX",
    "CLEAN",
    "MDETERM",
    "MINVERSE",
    "MMULT",
    "FILES",
    "IPMT",
    "PPMT",
    "COUNTA",
    "CANCEL.KEY",
    "FOR",
    "WHILE",
    "BREAK",
    "NEXT",
    "INITIATE",
    "REQUEST",
    "POKE",
    "EXECUTE",
    "TERMINATE",
    "RESTART",
    "HELP",
    "GET.BAR",
    "PRODUCT",
    "FACT",
    "GET.CELL",
    "GET.WORKSPACE",
    "GET.WINDOW",
    "GET.DOCUMENT",
    "DPRODUCT",
    "ISNONTEXT",
    "GET.NOTE",
    "NOTE",
    "STDEVP",
    "VARP",
    "DSTDEVP",
    "DVARP",
    "TRUNC",
    "ISLOGICAL",
    "DCOUNTA",
    "DELETE.BAR",
    "UNREGISTER",
    "",
    "",
    "USDOLLAR",
    "FINDB",
    "SEARCHB",
    "REPLACEB",
    "LEFTB",
    "RIGHTB",
    "MIDB",
    "LENB",
    "ROUNDUP",
    "ROUNDDOWN",
    "ASC",
    "DBCS",
    "RANK",
    "",
    "",
    "ADDRESS",
    "DAYS360",
    "TODAY",
    "VDB",
    "ELSE",
    "ELSE.IF",
    "END.IF",
    "FOR.CELL",
    "MEDIAN",
    "SUMPRODUCT",
    "SINH",
    "COSH",
    "TANH",
    "ASINH",
    "ACOSH",
    "ATANH",
    "DGET",
    "CREATE.OBJECT",
    "VOLATILE",
    "LAST.ERROR",
    "CUSTOM.UNDO",
    "CUSTOM.REPEAT",
    "FORMULA.CONVERT",
    "GET.LINK.INFO",
    "TEXT.BOX",
    "INFO",
    "GROUP",
    "GET.OBJECT",
    "DB",
    "PAUSE",
    "",
    "",
    "RESUME",
    "FREQUENCY",
    "ADD.TOOLBAR",
    "DELETE.TOOLBAR",
    "User",
    "RESET.TOOLBAR",
    "EVALUATE",
    "GET.TOOLBAR",
    "GET.TOOL",
    "SPELLING.CHECK",
    "ERROR.TYPE",
    "APP.TITLE",
    "WINDOW.TITLE",
    "SAVE.TOOLBAR",
    "ENABLE.TOOL",
    "PRESS.TOOL",
    "REGISTER.ID",
    "GET.WORKBOOK",
    "AVEDEV",
    "BETADIST",
    "GAMMALN",
    "BETAINV",
    "BINOMDIST",
    "CHIDIST",
    "CHIINV",
    "COMBIN",
    "CONFIDENCE",
    "CRITBINOM",
    "EVEN",
    "EXPONDIST",
    "FDIST",
    "FINV",
    "FISHER",
    "FISHERINV",
    "FLOOR",
    "GAMMADIST",
    "GAMMAINV",
    "CEILING",
    "HYPGEOMDIST",
    "LOGNORMDIST",
    "LOGINV",
    "NEGBINOMDIST",
    "NORMDIST",
    "NORMSDIST",
    "NORMINV",
    "NORMSINV",
    "STANDARDIZE",
    "ODD",
    "PERMUT",
    "POISSON",
    "TDIST",
    "WEIBULL",
    "SUMXMY2",
    "SUMX2MY2",
    "SUMX2PY2",
    "CHITEST",
    "CORREL",
    "COVAR",
    "FORECAST",
    "FTEST",
    "INTERCEPT",
    "PEARSON",
    "RSQ",
    "STEYX",
    "SLOPE",
    "TTEST",
    "PROB",
    "DEVSQ",
    "GEOMEAN",
    "HARMEAN",
    "SUMSQ",
    "KURT",
    "SKEW",
    "ZTEST",
    "LARGE",
    "SMALL",
    "QUARTILE",
    "PERCENTILE",
    "PERCENTRANK",
    "MODE",
    "TRIMMEAN",
    "TINV",
    "",
    "MOVIE.COMMAND",
    "GET.MOVIE",
    "CONCATENATE",
    "POWER",
    "PIVOT.ADD.DATA",
    "GET.PIVOT.TABLE",
    "GET.PIVOT.FIELD",
    "GET.PIVOT.ITEM",
    "RADIANS",
    "DEGREES",
    "SUBTOTAL",
    "SUMIF",
    "COUNTIF",
    "COUNTBLANK",
    "SCENARIO.GET",
    "OPTIONS.LISTS.GET",
    "ISPMT",
    "DATEDIF",
    "DATESTRING",
    "NUMBERSTRING",
    "ROMAN",
    "OPEN.DIALOG",
    "SAVE.DIALOG",
    "VIEW.GET",
    "GETPIVOTDATA",
    "HYPERLINK",
    "PHONETIC",
    "AVERAGEA",
    "MAXA",
    "MINA",
    "STDEVPA",
    "VARPA",
    "STDEVA",
    "VARA",
    "BAHTTEXT",
    "THAIDAYOFWEEK",
    "THAIDIGIT",
    "THAIMONTHOFYEAR",
    "THAINUMSOUND",
    "THAINUMSTRING",
    "THAISTRINGLENGTH",
    "ISTHAIDIGIT",
    "ROUNDBAHTDOWN",
    "ROUNDBAHTUP",
    "THAIYEAR",
    "RTD",
    "CUBEVALUE",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "OCT2BIN",
    "OCT2HEX",
    "OCT2DEC",
    "BIN2DEC",
    "BIN2OCT",
    "BIN2HEX",
    "IMSUB",
    "IMDIV",
    "IMPOWER",
    "IMABS",
    "IMSQRT",
    "IMLN",
    "IMLOG2",
    "IMLOG10",
    "IMSIN",
    "IMCOS",
    "IMEXP",
    "IMARGUMENT",
    "IMCONJUGATE",
    "IMAGINARY",
    "IMREAL",
    "COMPLEX",
    "IMSUM",
    "IMPRODUCT",
    "SERIESSUM",
    "FACTDOUBLE",
    "SQRTPI",
    "QUOTIENT",
    "DELTA",
    "GESTEP",
    "ISEVEN",
    "ISODD",
    "MROUND",
    "ERF",
    "ERFC",
    "BESSELJ",
    "BESSELK",
    "BESSELY",
    "BESSELI",
    "XIRR",
    "XNPV",
    "PRICEMAT",
    "YIELDMAT",
    "INTRATE",
    "RECEIVED",
    "DISC",
    "PRICEDISC",
    "YIELDDISC",
    "TBILLEQ",
    "TBILLPRICE",
    "TBILLYIELD",
    "PRICE",
    "YIELD",
    "DOLLARDE",
    "DOLLARFR",
    "NOMINAL",
    "EFFECT",
    "CUMPRINC",
    "CUMIPMT",
    "EDATE",
    "EOMONTH",
    "YEARFRAC",
    "COUPDAYBS",
    "COUPDAYS",
    "COUPDAYSNC",
    "COUPNCD",
    "COUPNUM",
    "COUPPCD",
    "DURATION",
    "MDURATION",
    "ODDLPRICE",
    "ODDLYIELD",
    "ODDFPRICE",
    "ODDFYIELD",
    "RANDBETWEEN",
    "WEEKNUM",
    "AMORDEGRC",
    "AMORLINC",
    "CONVERT",
    //     "SHEETJS",
    "ACCRINT",
    "ACCRINTM",
    "WORKDAY",
    "NETWORKDAYS",
    "GCD",
    "MULTINOMIAL",
    "LCM",
    "FVSCHEDULE",
    "CUBEKPIMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "IFERROR",
    "COUNTIFS",
    "SUMIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
];

pub const FTAB_ARGC: [u8; FTAB_LEN] = [
    255, // "COUNT",
    3,   // "IF",
    1,   // "ISNA",
    1,   // "ISERROR",
    255, // "SUM",
    255, // "AVERAGE",
    255, // "MIN",
    255, // "MAX",
    1,   // "ROW",
    1,   // "COLUMN",
    0,   // "NA",
    254, // "NPV",
    255, // "STDEV",
    2,   // "DOLLAR",
    3,   // "FIXED",
    1,   // "SIN",
    1,   // "COS",
    1,   // "TAN",
    1,   // "ATAN",
    0,   // "PI",
    1,   // "SQRT",
    1,   // "EXP",
    1,   // "LN",
    1,   // "LOG10",
    1,   // "ABS",
    1,   // "INT",
    1,   // "SIGN",
    2,   // "ROUND",
    3,   // "LOOKUP",
    4,   // "INDEX",
    2,   // "REPT",
    3,   // "MID",
    1,   // "LEN",
    1,   // "VALUE",
    0,   // "TRUE",
    0,   // "FALSE",
    255, // "AND",
    255, // "OR",
    1,   // "NOT",
    2,   // "MOD",
    3,   // "DCOUNT",
    3,   // "DSUM",
    3,   // "DAVERAGE",
    3,   // "DMIN",
    3,   // "DMAX",
    3,   // "DSTDEV",
    255, // "VAR",
    3,   // "DVAR",
    2,   // "TEXT",
    4,   // "LINEST",
    4,   // "TREND",
    4,   // "LOGEST",
    4,   // "GROWTH",
    1,   // "GOTO",
    1,   // "HALT",
    1,   // "RETURN",
    5,   // "PV",
    5,   // "FV",
    5,   // "NPER",
    5,   // "PMT",
    6,   // "RATE",
    3,   // "MIRR",
    2,   // "IRR",
    0,   // "RAND",
    3,   // "MATCH",
    3,   // "DATE",
    3,   // "TIME",
    1,   // "DAY",
    1,   // "MONTH",
    1,   // "YEAR",
    2,   // "WEEKDAY",
    1,   // "HOUR",
    1,   // "MINUTE",
    1,   // "SECOND",
    0,   // "NOW",
    1,   // "AREAS",
    1,   // "ROWS",
    1,   // "COLUMNS",
    5,   // "OFFSET",
    2,   // "ABSREF",
    2,   // "RELREF",
    3,   // "ARGUMENT",
    3,   // "SEARCH",
    1,   // "TRANSPOSE",
    2,   // "ERROR",
    0,   // "STEP",
    1,   // "TYPE",
    1,   // "ECHO",
    2,   // "SET.NAME",
    0,   // "CALLER",
    1,   // "DEREF",
    2,   // "WINDOWS",
    2,   // "SERIES",
    2,   // "DOCUMENTS",
    0,   // "ACTIVE.CELL",
    0,   // "SELECTION",
    1,   // "RESULT",
    2,   // "ATAN2",
    1,   // "ASIN",
    1,   // "ACOS",
    255, // "CHOOSE",
    4,   // "HLOOKUP",
    4,   // "VLOOKUP",
    2,   // "LINKS",
    7,   // "INPUT",
    1,   // "ISREF",
    1,   // "GET.FORMULA",
    2,   // "GET.NAME",
    2,   // "SET.VALUE",
    2,   // "LOG",
    4,   // "EXEC",
    1,   // "CHAR",
    1,   // "LOWER",
    1,   // "UPPER",
    1,   // "PROPER",
    2,   // "LEFT",
    2,   // "RIGHT",
    2,   // "EXACT",
    1,   // "TRIM",
    4,   // "REPLACE",
    4,   // "SUBSTITUTE",
    1,   // "CODE",
    3,   // "NAMES",
    1,   // "DIRECTORY",
    3,   // "FIND",
    2,   // "CELL",
    1,   // "ISERR",
    1,   // "ISTEXT",
    1,   // "ISNUMBER",
    1,   // "ISBLANK",
    1,   // "T",
    1,   // "N",
    2,   // "FOPEN",
    1,   // "FCLOSE",
    1,   // "FSIZE",
    1,   // "FREADLN",
    2,   // "FREAD",
    2,   // "FWRITELN",
    2,   // "FWRITE",
    2,   // "FPOS",
    1,   // "DATEVALUE",
    1,   // "TIMEVALUE",
    3,   // "SLN",
    4,   // "SYD",
    5,   // "DDB",
    3,   // "GET.DEF",
    2,   // "REFTEXT",
    2,   // "TEXTREF",
    2,   // "INDIRECT",
    255, // "REGISTER",
    255, // "CALL",
    1,   // "ADD.BAR",
    4,   // "ADD.MENU",
    5,   // "ADD.COMMAND",
    5,   // "ENABLE.COMMAND",
    5,   // "CHECK.COMMAND",
    5,   // "RENAME.COMMAND",
    1,   // "SHOW.BAR",
    3,   // "DELETE.MENU",
    4,   // "DELETE.COMMAND",
    3,   // "GET.CHART.ITEM",
    1,   // "DIALOG.BOX",
    1,   // "CLEAN",
    1,   // "MDETERM",
    1,   // "MINVERSE",
    1,   // "MMULT",
    2,   // "FILES",
    6,   // "IPMT",
    6,   // "PPMT",
    255, // "COUNTA",
    2,   // "CANCEL.KEY",
    4,   // "FOR",
    1,   // "WHILE",
    0,   // "BREAK",
    0,   // "NEXT",
    2,   // "INITIATE",
    2,   // "REQUEST",
    3,   // "POKE",
    2,   // "EXECUTE",
    1,   // "TERMINATE",
    1,   // "RESTART",
    1,   // "HELP",
    4,   // "GET.BAR",
    255, // "PRODUCT",
    1,   // "FACT",
    2,   // "GET.CELL",
    1,   // "GET.WORKSPACE",
    2,   // "GET.WINDOW",
    2,   // "GET.DOCUMENT",
    3,   // "DPRODUCT",
    1,   // "ISNONTEXT",
    3,   // "GET.NOTE",
    4,   // "NOTE",
    255, // "STDEVP",
    255, // "VARP",
    3,   // "DSTDEVP",
    3,   // "DVARP",
    2,   // "TRUNC",
    1,   // "ISLOGICAL",
    3,   // "DCOUNTA",
    1,   // "DELETE.BAR",
    1,   // "UNREGISTER",
    0,   // "",
    0,   // "",
    2,   // "USDOLLAR",
    3,   // "FINDB",
    3,   // "SEARCHB",
    4,   // "REPLACEB",
    2,   // "LEFTB",
    2,   // "RIGHTB",
    3,   // "MIDB",
    3,   // "LENB",
    2,   // "ROUNDUP",
    2,   // "ROUNDDOWN",
    1,   // "ASC",
    1,   // "DBCS",
    3,   // "RANK",
    0,   // "",
    0,   // "",
    5,   // "ADDRESS",
    3,   // "DAYS360",
    0,   // "TODAY",
    7,   // "VDB",
    0,   // "ELSE",
    1,   // "ELSE.IF",
    0,   // "END.IF",
    3,   // "FOR.CELL",
    255, // "MEDIAN",
    255, // "SUMPRODUCT",
    1,   // "SINH",
    1,   // "COSH",
    1,   // "TANH",
    1,   // "ASINH",
    1,   // "ACOSH",
    1,   // "ATANH",
    3,   // "DGET",
    11,  // "CREATE.OBJECT",
    1,   // "VOLATILE",
    0,   // "LAST.ERROR",
    2,   // "CUSTOM.UNDO",
    3,   // "CUSTOM.REPEAT",
    5,   // "FORMULA.CONVERT",
    4,   // "GET.LINK.INFO",
    4,   // "TEXT.BOX",
    1,   // "INFO",
    0,   // "GROUP",
    5,   // "GET.OBJECT",
    5,   // "DB",
    1,   // "PAUSE",
    0,   // "",
    0,   // "",
    1,   // "RESUME",
    2,   // "FREQUENCY",
    2,   // "ADD.TOOLBAR",
    1,   // "DELETE.TOOLBAR",
    255, // "User",
    1,   // "RESET.TOOLBAR",
    1,   // "EVALUATE",
    2,   // "GET.TOOLBAR",
    3,   // "GET.TOOL",
    3,   // "SPELLING.CHECK",
    1,   // "ERROR.TYPE",
    1,   // "APP.TITLE",
    1,   // "WINDOW.TITLE",
    2,   // "SAVE.TOOLBAR",
    3,   // "ENABLE.TOOL",
    3,   // "PRESS.TOOL",
    3,   // "REGISTER.ID",
    2,   // "GET.WORKBOOK",
    255, // "AVEDEV",
    5,   // "BETADIST",
    1,   // "GAMMALN",
    5,   // "BETAINV",
    4,   // "BINOMDIST",
    2,   // "CHIDIST",
    2,   // "CHIINV",
    2,   // "COMBIN",
    3,   // "CONFIDENCE",
    3,   // "CRITBINOM",
    1,   // "EVEN",
    3,   // "EXPONDIST",
    3,   // "FDIST",
    3,   // "FINV",
    1,   // "FISHER",
    1,   // "FISHERINV",
    2,   // "FLOOR",
    4,   // "GAMMADIST",
    3,   // "GAMMAINV",
    2,   // "CEILING",
    4,   // "HYPGEOMDIST",
    3,   // "LOGNORMDIST",
    3,   // "LOGINV",
    3,   // "NEGBINOMDIST",
    4,   // "NORMDIST",
    1,   // "NORMSDIST",
    3,   // "NORMINV",
    1,   // "NORMSINV",
    3,   // "STANDARDIZE",
    1,   // "ODD",
    2,   // "PERMUT",
    3,   // "POISSON",
    3,   // "TDIST",
    4,   // "WEIBULL",
    2,   // "SUMXMY2",
    2,   // "SUMX2MY2",
    2,   // "SUMX2PY2",
    2,   // "CHITEST",
    2,   // "CORREL",
    2,   // "COVAR",
    3,   // "FORECAST",
    2,   // "FTEST",
    2,   // "INTERCEPT",
    2,   // "PEARSON",
    2,   // "RSQ",
    2,   // "STEYX",
    2,   // "SLOPE",
    4,   // "TTEST",
    4,   // "PROB",
    255, // "DEVSQ",
    255, // "GEOMEAN",
    255, // "HARMEAN",
    255, // "SUMSQ",
    255, // "KURT",
    255, // "SKEW",
    3,   // "ZTEST",
    2,   // "LARGE",
    2,   // "SMALL",
    2,   // "QUARTILE",
    2,   // "PERCENTILE",
    3,   // "PERCENTRANK",
    255, // "MODE",
    2,   // "TRIMMEAN",
    2,   // "TINV",
    4,   // "",
    4,   // "MOVIE.COMMAND",
    3,   // "GET.MOVIE",
    255, // "CONCATENATE",
    2,   // "POWER",
    9,   // "PIVOT.ADD.DATA",
    2,   // "GET.PIVOT.TABLE",
    3,   // "GET.PIVOT.FIELD",
    4,   // "GET.PIVOT.ITEM",
    1,   // "RADIANS",
    1,   // "DEGREES",
    255, // "SUBTOTAL",
    3,   // "SUMIF",
    2,   // "COUNTIF",
    1,   // "COUNTBLANK",
    2,   // "SCENARIO.GET",
    1,   // "OPTIONS.LISTS.GET",
    4,   // "ISPMT",
    3,   // "DATEDIF",
    1,   // "DATESTRING",
    2,   // "NUMBERSTRING",
    2,   // "ROMAN",
    4,   // "OPEN.DIALOG",
    5,   // "SAVE.DIALOG",
    2,   // "VIEW.GET",
    128, // "GETPIVOTDATA",
    2,   // "HYPERLINK",
    1,   // "PHONETIC",
    255, // "AVERAGEA",
    255, // "MAXA",
    255, // "MINA",
    255, // "STDEVPA",
    255, // "VARPA",
    255, // "STDEVA",
    255, // "VARA",
    1,   // "BAHTTEXT",
    1,   // "THAIDAYOFWEEK",
    1,   // "THAIDIGIT",
    1,   // "THAIMONTHOFYEAR",
    1,   // "THAINUMSOUND",
    1,   // "THAINUMSTRING",
    1,   // "THAISTRINGLENGTH",
    1,   // "ISTHAIDIGIT",
    1,   // "ROUNDBAHTDOWN",
    1,   // "ROUNDBAHTUP",
    1,   // "THAIYEAR",
    255, // "RTD",
    255, // "CUBEVALUE",
    3,   // "CUBEMEMBER",
    3,   // "CUBEMEMBERPROPERTY",
    4,   // "CUBERANKEDMEMBER",
    2,   // "HEX2BIN",
    1,   // "HEX2DEC",
    2,   // "HEX2OCT",
    2,   // "DEC2BIN",
    2,   // "DEC2HEX",
    2,   // "DEC2OCT",
    2,   // "OCT2BIN",
    2,   // "OCT2HEX",
    1,   // "OCT2DEC",
    1,   // "BIN2DEC",
    2,   // "BIN2OCT",
    2,   // "BIN2HEX",
    2,   // "IMSUB",
    2,   // "IMDIV",
    2,   // "IMPOWER",
    1,   // "IMABS",
    1,   // "IMSQRT",
    1,   // "IMLN",
    1,   // "IMLOG2",
    1,   // "IMLOG10",
    1,   // "IMSIN",
    1,   // "IMCOS",
    1,   // "IMEXP",
    1,   // "IMARGUMENT",
    1,   // "IMCONJUGATE",
    1,   // "IMAGINARY",
    1,   // "IMREAL",
    3,   // "COMPLEX",
    255, // "IMSUM",
    255, // "IMPRODUCT",
    4,   // "SERIESSUM",
    1,   // "FACTDOUBLE",
    1,   // "SQRTPI",
    2,   // "QUOTIENT",
    2,   // "DELTA",
    2,   // "GESTEP",
    1,   // "ISEVEN",
    1,   // "ISODD",
    2,   // "MROUND",
    2,   // "ERF",
    1,   // "ERFC",
    2,   // "BESSELJ",
    2,   // "BESSELK",
    2,   // "BESSELY",
    2,   // "BESSELI",
    3,   // "XIRR",
    3,   // "XNPV",
    6,   // "PRICEMAT",
    6,   // "YIELDMAT",
    5,   // "INTRATE",
    5,   // "RECEIVED",
    5,   // "DISC",
    5,   // "PRICEDISC",
    5,   // "YIELDDISC",
    3,   // "TBILLEQ",
    3,   // "TBILLPRICE",
    3,   // "TBILLYIELD",
    7,   // "PRICE",
    7,   // "YIELD",
    2,   // "DOLLARDE",
    2,   // "DOLLARFR",
    2,   // "NOMINAL",
    2,   // "EFFECT",
    6,   // "CUMPRINC",
    6,   // "CUMIPMT",
    2,   // "EDATE",
    2,   // "EOMONTH",
    3,   // "YEARFRAC",
    4,   // "COUPDAYBS",
    4,   // "COUPDAYS",
    4,   // "COUPDAYSNC",
    4,   // "COUPNCD",
    4,   // "COUPNUM",
    4,   // "COUPPCD",
    6,   // "DURATION",
    6,   // "MDURATION",
    8,   // "ODDLPRICE",
    8,   // "ODDLYIELD",
    8,   // "ODDFPRICE",
    8,   // "ODDFYIELD",
    2,   // "RANDBETWEEN",
    2,   // "WEEKNUM",
    7,   // "AMORDEGRC",
    7,   // "AMORLINC",
    8,   // "CONVERT",
    //  1, //    "SHEETJS",
    8,   // "ACCRINT",
    5,   // "ACCRINTM",
    3,   // "WORKDAY",
    3,   // "NETWORKDAYS",
    255, // "GCD",
    255, // "MULTINOMIAL",
    255, // "LCM",
    2,   // "FVSCHEDULE",
    4,   // "CUBEKPIMEMBER",
    5,   // "CUBESET",
    1,   // "CUBESETCOUNT",
    2,   // "IFERROR",
    128, // "COUNTIFS",
    129, // "SUMIFS",
    3,   // "AVERAGEIF",
    129, // "AVERAGEIFS"
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_push_column() {
        let check = |col: u32, expected: &str| {
            let mut got = String::new();
            push_column(col, &mut got);
            assert_eq!(got, expected, "push_column({col})");
        };
        check(0, "A");
        check(25, "Z");
        check(26, "AA");
        check(27, "AB");
        check(51, "AZ");
        check(53, "BB");
        check(701, "ZZ");
        check(702, "AAA");
        check(16383, "XFD");
    }

    #[test]
    fn test_quote_sheet_name() {
        // Unquoted and quoted sheet names from Excel files, from the tests of
        // `quote_sheet_name()` in `rust_xlsxwriter`.
        let tests = [
            // ----------------------------------------------------------------
            // Rule 1.
            // ----------------------------------------------------------------
            // Some simple variants on standard sheet names.
            ("Sheet1", "Sheet1"),
            ("Sheet.1", "Sheet.1"),
            ("Sheet_1", "Sheet_1"),
            ("Sheet-1", "'Sheet-1'"),
            ("Sheet 1", "'Sheet 1'"),
            ("Sheet#1", "'Sheet#1'"),
            // Sheetnames with single quotes.
            ("Sheet'1", "'Sheet''1'"),
            ("Sheet''1", "'Sheet''''1'"),
            // Single special chars that are unquoted in sheetnames. These are
            // variants of the first char rule.
            ("_", "_"),
            (".", "'.'"),
            // White space only.
            (" ", "' '"),
            ("    ", "'    '"),
            // Sheetnames with unicode or emojis.
            ("été", "été"),
            ("mangé", "mangé"),
            ("Sheet©", "'Sheet©'"),
            ("Sheet😀", "Sheet😀"),
            ("Sheet🤌1", "Sheet🤌1"),
            ("Sheet⟦1", "'Sheet⟦1'"), // Unicode punctuation.
            ("Sheet᠅1", "'Sheet᠅1'"), // Unicode punctuation.
            // ----------------------------------------------------------------
            // Rule 2.
            // ----------------------------------------------------------------
            // Sheetnames starting with non-word characters.
            ("_Sheet1", "_Sheet1"),
            (".Sheet1", "'.Sheet1'"),
            ("1Sheet1", "'1Sheet1'"),
            ("-Sheet1", "'-Sheet1'"),
            ("#Sheet1", "'#Sheet1'"),
            ("©Sheet", "'©Sheet'"),
            ("😀Sheet", "'😀Sheet'"),
            ("🤌Sheet", "'🤌Sheet'"),
            // Sheetnames that are digits only also start with a non word char.
            ("1", "'1'"),
            ("2", "'2'"),
            ("1234", "'1234'"),
            ("12345678", "'12345678'"),
            // ----------------------------------------------------------------
            // Rule 3.
            // ----------------------------------------------------------------
            // Worksheet names that look like A1 style references (with the
            // row/column number in the Excel allowable range). These are case
            // insensitive.
            ("A0", "A0"),
            ("A1", "'A1'"),
            ("a1", "'a1'"),
            ("XFD", "XFD"),
            ("xfd", "xfd"),
            ("XFE1", "XFE1"),
            ("ZZZ1", "ZZZ1"),
            ("XFD1", "'XFD1'"),
            ("xfd1", "'xfd1'"),
            ("B1048577", "B1048577"),
            ("A1048577", "A1048577"),
            ("A1048576", "'A1048576'"),
            ("B1048576", "'B1048576'"),
            ("B1048576a", "B1048576a"),
            ("XFD048576", "'XFD048576'"),
            ("XFD1048576", "'XFD1048576'"),
            ("XFD01048577", "XFD01048577"),
            ("XFD01048576", "'XFD01048576'"),
            ("A123456789012345678901", "A123456789012345678901"), // Exceeds u64.
            // Sheet names where the characters before the digits aren't ASCII
            // column letters aren't treated as cell references.
            ("Q.1", "Q.1"),
            ("_1", "_1"),
            ("é1", "é1"),
            ("一1", "一1"),
            ("éa1", "éa1"),
            ("ß1", "ß1"), // Uppercases to "SS" but isn't treated as "SS1".
            ("ﬁ1", "ﬁ1"), // Uppercases to "FI" but isn't treated as "FI1".
            // ----------------------------------------------------------------
            // Rule 4.
            // ----------------------------------------------------------------
            // Sheet names that *start* with RC style references (with the
            // row/column number in the Excel allowable range). These are case
            // insensitive.
            ("A", "A"),
            ("B", "B"),
            ("D", "D"),
            ("Q", "Q"),
            ("S", "S"),
            ("c", "'c'"),
            ("C", "'C'"),
            ("CR", "CR"),
            ("CZ", "CZ"),
            ("r", "'r'"),
            ("R", "'R'"),
            ("C8", "'C8'"),
            ("rc", "'rc'"),
            ("RC", "'RC'"),
            ("RCZ", "RCZ"),
            ("RRC", "RRC"),
            ("R0C0", "R0C0"),
            ("R4C", "'R4C'"),
            ("R5C", "'R5C'"),
            ("rc2", "'rc2'"),
            ("RC2", "'RC2'"),
            ("RC8", "'RC8'"),
            ("bR1C1", "bR1C1"),
            ("R1C1", "'R1C1'"),
            ("r1c2", "'r1c2'"),
            ("rc2z", "'rc2z'"),
            ("bR1C1b", "bR1C1b"),
            ("R1C1b", "'R1C1b'"),
            ("R1C1R", "'R1C1R'"),
            ("C16384", "'C16384'"),
            ("C16385", "'C16385'"),
            ("C16385Z", "C16385Z"),
            ("C16386", "'C16386'"),
            ("C16384Z", "'C16384Z'"),
            ("PC16384Z", "PC16384Z"),
            ("RC16383", "'RC16383'"),
            ("RC16385Z", "RC16385Z"),
            ("R1048576", "'R1048576'"),
            ("R1048577C", "R1048577C"),
            ("R1C16384", "'R1C16384'"),
            ("R1C16385", "'R1C16385'"),
            ("RC16384Z", "'RC16384Z'"),
            ("R1048576C", "'R1048576C'"),
            ("R1048577C1", "R1048577C1"),
            ("R1C16384Z", "'R1C16384Z'"),
            ("R1048575C1", "'R1048575C1'"),
            ("R1048576C1", "'R1048576C1'"),
            ("R1048577C16384", "R1048577C16384"),
            ("R1048576C16384", "'R1048576C16384'"),
            ("R1048576C16385", "'R1048576C16385'"),
            ("ZR1048576C16384", "ZR1048576C16384"),
            ("C123456789012345678901Z", "C123456789012345678901Z"), // Exceeds u64.
            ("R123456789012345678901Z", "R123456789012345678901Z"), // Exceeds u64.
            // ----------------------------------------------------------------
            // Boolean values.
            // ----------------------------------------------------------------
            ("TRUE", "'TRUE'"),
            ("false", "'false'"),
        ];

        for (name, expected) in tests {
            assert_eq!(quote_sheet_name(name), expected, "for name {name:?}");
        }
    }

    #[test]
    fn unescape_entity() {
        // Tests for standard named XML entities.
        let mut buffer = String::new();
        let entity = BytesRef::new("amp");
        unescape_entity_to_buffer(&entity, &mut buffer).unwrap();
        assert_eq!(buffer, "&");

        buffer.clear();
        let entity = BytesRef::new("lt");
        unescape_entity_to_buffer(&entity, &mut buffer).unwrap();
        assert_eq!(buffer, "<");

        // Tests for numeric character references.
        buffer.clear();
        let entity = BytesRef::new("#xA");
        let result = unescape_entity_to_buffer(&entity, &mut buffer);
        assert_eq!(buffer, "\n");
        assert!(result.is_ok());

        buffer.clear();
        let entity = BytesRef::new("#x41");
        let result = unescape_entity_to_buffer(&entity, &mut buffer);
        assert!(result.is_ok());
        assert_eq!(buffer, "A");

        buffer.clear();
        let entity = BytesRef::new("#65"); // Decimal.
        let result = unescape_entity_to_buffer(&entity, &mut buffer);
        assert!(result.is_ok());
        assert_eq!(buffer, "A");

        // Test for failure on empty entity.
        buffer.clear();
        let entity = BytesRef::new("");
        let result = unescape_entity_to_buffer(&entity, &mut buffer);
        assert!(result.is_err());

        // Test the error message.
        let entity = BytesRef::new("not_a_real_entity");
        let result = unescape_entity_to_buffer(&entity, &mut buffer);
        assert!(result.is_err());
        if let Err(quick_xml::Error::Escape(quick_xml::escape::EscapeError::UnrecognizedEntity(
            _,
            msg,
        ))) = result
        {
            assert!(msg.contains("not_a_real_entity"));
        }
    }

    #[test]
    fn xml_with_escapes() {
        let test_cases = vec![
            ("_", "_"),
            ("_x", "_x"),
            ("_x0", "_x0"),
            ("_x00", "_x00"),
            ("_x005F_", "_"),
            ("_x000D_", "\r"),
            ("_x000", "_x000"),
            ("_x001F_", "\x1F"),
            ("_x000D", "_x000D"),
            ("_x00ZZ_", "_x00ZZ_"),
            ("_x_x_x", "_x_x_x"),
            ("_x00½_", "_x00½_"),
            ("_x000G_", "_x000G_"),
            ("_x597G_", "_x597G_"),
            ("😀_x000D_😀", "😀\r😀"),
            ("_x005F_x0000_", "_x0000_"),
            ("_x005F_x000a_", "_x000a_"),
            ("_x005F_x000A_", "_x000A_"),
            ("_x005F_x005F_", "_x005F_"),
            ("_x005F_x597D_", "_x597D_"),
            ("_x005F_x597d_", "_x597d_"),
            ("__x005F_x0000__", "__x0000__"),
            ("Hello_x000D_World", "Hello\rWorld"),
            ("Control_x001F_Char", "Control\x1FChar"),
            ("Hello_x000D_World_x000D_", "Hello\rWorld\r"),
            ("Just_a_normal_string", "Just_a_normal_string"),
            ("Hello_x000D_World_x000D__x000D_", "Hello\rWorld\r\r"),
            ("Multiple_x000D__x000D__x000D_Chars", "Multiple\r\r\rChars"),
            // Escapes of code units above U+00FF, including surrogate pairs.
            ("_x2764_", "\u{2764}"),
            ("_x2764__xFE0F_", "\u{2764}\u{FE0F}"),
            ("_x597d_", "\u{597D}"),
            ("_x00E9_", "\u{E9}"),
            ("_xFFFE_", "\u{FFFE}"),
            ("_xFFFF_", "\u{FFFF}"),
            ("_xD83D__xDCF7_", "\u{1F4F7}"),
            ("_xd83d__xdcf7_", "\u{1F4F7}"),
            ("a_xD83D__xDE00_b", "a\u{1F600}b"),
            ("_xD83D_", "\u{FFFD}"),
            ("_xDCF7_", "\u{FFFD}"),
            ("_xDCF7__xD83D_", "\u{FFFD}\u{FFFD}"),
            ("_xD83D__xD83D__xDCF7_", "\u{FFFD}\u{1F4F7}"),
            ("_xD83D_x", "\u{FFFD}x"),
            ("_xD83D__x0041_", "\u{FFFD}A"),
            ("_x005F_x2764_", "_x2764_"),
            ("_x005F_xD83D__xDCF7_", "_xD83D_\u{FFFD}"),
            ("_x+123_", "_x+123_"),
            ("_x00+1_", "_x00+1_"),
            // Round trips of the rust_xlsxwriter escape_xml_escapes() tests.
            ("hello world", "hello world"),
            ("/path/to/kernel32_x64d8.pyd", "/path/to/kernel32_x64d8.pyd"),
            ("test_x", "test_x"),
            ("test_x00", "test_x00"),
            ("test_x005F_x0000_end", "test_x0000_end"),
            ("test_xGGGG_end", "test_xGGGG_end"),
            ("test_x00GG_end", "test_x00GG_end"),
            ("test_x005F_xABCD_end", "test_xABCD_end"),
            ("test_x005F_xabcd_end", "test_xabcd_end"),
            ("test_x005F_xAbCd_end", "test_xAbCd_end"),
            ("test_xABCD_end", "test\u{ABCD}end"),
            ("test_xabcd_end", "test\u{ABCD}end"),
            ("_x005F_x0000_middle_x005F_x0001_", "_x0000_middle_x0001_"),
            (
                "/path/file_x64d8f123xc8c311aa.pyd",
                "/path/file_x64d8f123xc8c311aa.pyd",
            ),
            (
                "/path/file_x005F_x64d8_name.pyd",
                "/path/file_x64d8_name.pyd",
            ),
            ("/path/file_x64d8_name.pyd", "/path/file\u{64D8}name.pyd"),
            ("_x005F_x0000_rest", "_x0000_rest"),
            ("start_x005F_x0000_", "start_x0000_"),
            ("_x005F_x0000_", "_x0000_"),
            ("日本語_x005F_x0000_テスト", "日本語_x0000_テスト"),
            ("test_xäöüü_end", "test_xäöüü_end"),
            ("_x005F_x0000__x005F_x0001_", "_x0000__x0001_"),
            ("_x005F_x005F_x005F_x000D_", "_x005F_x000D_"),
        ];

        for (input, expected) in test_cases {
            assert_eq!(unescape_xml(input), expected);
        }
    }

    #[test]
    fn test_unescape_xml_borrows_without_escapes() {
        // Strings without a valid escape are returned without allocating.
        for input in [
            "",
            "_",
            "pos_x",
            "_x",
            "_x00ZZ_",
            "snake_case_x_y",
            "a_x005F",
        ] {
            assert!(matches!(unescape_xml(input), Cow::Borrowed(_)), "{input:?}");
        }

        for input in ["_x000D_", "a_x2764_", "_xD83D_"] {
            assert!(matches!(unescape_xml(input), Cow::Owned(_)), "{input:?}");
        }
    }
}
