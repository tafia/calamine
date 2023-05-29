/// Check excel number format is datetime
pub fn is_custom_date_format(format: &str) -> bool {
    let mut escaped = false;
    let mut is_quote = false;
    let mut brackets = 0u8;
    let mut prev = ' ';
    let mut hms = false;
    let mut ap = false;
    for s in format.chars() {
        match (s, escaped, is_quote, ap, brackets) {
            (_, true, ..) => escaped = false, // if escaped, ignore
            ('_' | '\\', ..) => escaped = true,
            ('"', _, true, _, _) => is_quote = false,
            (_, _, true, _, _) => (),
            ('"', _, _, _, _) => is_quote = true,
            (';', ..) => return false, // first format only
            ('[', ..) => brackets += 1,
            (']', .., 1) if hms => return true, // if closing
            (']', ..) => brackets = brackets.saturating_sub(1),
            ('a' | 'A', _, _, false, 0) => ap = true,
            ('p' | 'm' | '/' | 'P' | 'M', _, _, true, 0) => return true,
            ('d' | 'm' | 'h' | 'y' | 's' | 'D' | 'M' | 'H' | 'Y' | 'S', _, _, false, 0) => {
                return true
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
    false
}

pub fn is_builtin_date_format_id(id: &[u8]) -> bool {
    matches!(
        id,
        // mm-dd-yy
        b"14" |
        // d-mmm-yy
        b"15" |
        // d-mmm
        b"16" |
        // mmm-yy
        b"17" |
        // h:mm AM/PM
        b"18" |
        // h:mm:ss AM/PM
        b"19" |
        // h:mm
        b"20" |
        // h:mm:ss
        b"21" |
        // m/d/yy h:mm
        b"22" |
        // mm:ss
        b"45" |
        // [h]:mm:ss
        b"46" |
        // mmss.0
        b"47"
    )
}

/// Check if code corresponds to builtin date format
///
/// See `is_builtin_date_format_id`
pub fn is_builtin_date_format_code(code: u16) -> bool {
    matches!(code, 14..=22 | 45..=47)
}

/// Ported from openpyxl, MIT License
/// https://foss.heptapod.net/openpyxl/openpyxl/-/blob/a5e197c530aaa49814fd1d993dd776edcec35105/openpyxl/styles/tests/test_number_style.py
#[test]
fn test_is_date_format() {
    assert_eq!(is_custom_date_format("DD/MM/YY"), true);
    assert_eq!(is_custom_date_format("H:MM:SS;@"), true);
    assert_eq!(is_custom_date_format("#,##0\\ [$\\u20bd-46D]"), false);
    assert_eq!(is_custom_date_format("m\"M\"d\"D\";@"), true);
    assert_eq!(is_custom_date_format("[h]:mm:ss"), true);
    assert_eq!(
        is_custom_date_format("\"Y: \"0.00\"m\";\"Y: \"-0.00\"m\";\"Y: <num>m\";@"),
        false
    );
    assert_eq!(is_custom_date_format("#,##0\\ [$''u20bd-46D]"), false);
    assert_eq!(
        is_custom_date_format("\"$\"#,##0_);[Red](\"$\"#,##0)"),
        false
    );
    assert_eq!(
        is_custom_date_format("[$-404]e\"\\xfc\"m\"\\xfc\"d\"\\xfc\""),
        true
    );
    assert_eq!(is_custom_date_format("0_ ;[Red]\\-0\\ "), false);
    assert_eq!(is_custom_date_format("\\Y000000"), false);
    assert_eq!(is_custom_date_format("#,##0.0####\" YMD\""), false);
    assert_eq!(is_custom_date_format("[h]"), true);
    assert_eq!(is_custom_date_format("[ss]"), true);
    assert_eq!(is_custom_date_format("[s].000"), true);
    assert_eq!(is_custom_date_format("[m]"), true);
    assert_eq!(is_custom_date_format("[mm]"), true);
    assert_eq!(
        is_custom_date_format("[Blue]\\+[h]:mm;[Red]\\-[h]:mm;[Green][h]:mm"),
        true
    );
    assert_eq!(is_custom_date_format("[>=100][Magenta][s].00"), true);
    assert_eq!(is_custom_date_format("[h]:mm;[=0]\\-"), true);
    assert_eq!(is_custom_date_format("[>=100][Magenta].00"), false);
    assert_eq!(is_custom_date_format("[>=100][Magenta]General"), false);
    assert_eq!(is_custom_date_format("ha/p\\\\m"), true);
    assert_eq!(
        is_custom_date_format("#,##0.00\\ _M\"H\"_);[Red]#,##0.00\\ _M\"S\"_)"),
        false
    );
}
