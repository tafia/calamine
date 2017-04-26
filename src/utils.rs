//! Internal module providing handy function

use encoding_rs::{self, Encoding};

/// Converts a &[u8] into a &[u32]
pub fn to_u32(s: &[u8]) -> &[u32] {
    assert!(s.len() % 4 == 0);
    unsafe { ::std::slice::from_raw_parts(s as *const [u8] as *const u32, s.len() / 4) }
}
pub fn read_slice<T>(s: &[u8]) -> T {
    unsafe { ::std::ptr::read(&s[..::std::mem::size_of::<T>()] as *const [u8] as *const T) }
}

pub fn read_u32(s: &[u8]) -> u32 {
    read_slice(s)
}

pub fn read_u16(s: &[u8]) -> u16 {
    read_slice(s)
}

pub fn read_usize(s: &[u8]) -> usize {
    read_u32(s) as usize
}

/// Returns an encoding from Windows code page number.
/// http://msdn.microsoft.com/en-us/library/windows/desktop/dd317756%28v=vs.85%29.aspx
/// Sometimes it can return a *superset* of the requested encoding, e.g. for several CJK encodings.
///
/// The code is copied from [rust-encoding](https://github.com/lifthrasiir/rust-encoding)
pub fn encoding_from_windows_code_page(cp: usize) -> Option<&'static Encoding> {
    match cp {
        65001 => Some(encoding_rs::UTF_8),
        866 => Some(encoding_rs::IBM866),
        28592 => Some(encoding_rs::ISO_8859_2),
        28593 => Some(encoding_rs::ISO_8859_3),
        28594 => Some(encoding_rs::ISO_8859_4),
        28595 => Some(encoding_rs::ISO_8859_5),
        28596 => Some(encoding_rs::ISO_8859_6),
        28597 => Some(encoding_rs::ISO_8859_7),
        28598 => Some(encoding_rs::ISO_8859_8),
        28603 => Some(encoding_rs::ISO_8859_13),
        28605 => Some(encoding_rs::ISO_8859_15),
        20866 => Some(encoding_rs::KOI8_R),
        21866 => Some(encoding_rs::KOI8_U),
        10000 => Some(encoding_rs::MACINTOSH),
        874 => Some(encoding_rs::WINDOWS_874),
        1250 => Some(encoding_rs::WINDOWS_1250),
        1251 => Some(encoding_rs::WINDOWS_1251),
        1252 => Some(encoding_rs::WINDOWS_1252),
        1253 => Some(encoding_rs::WINDOWS_1253),
        1254 => Some(encoding_rs::WINDOWS_1254),
        1255 => Some(encoding_rs::WINDOWS_1255),
        1256 => Some(encoding_rs::WINDOWS_1256),
        1257 => Some(encoding_rs::WINDOWS_1257),
        1258 => Some(encoding_rs::WINDOWS_1258),
        1259 | 10007 => Some(encoding_rs::X_MAC_CYRILLIC),
        936 | 54936 => Some(encoding_rs::GB18030), // XXX technicencoding_rsy wrong
        950 => Some(encoding_rs::BIG5),
        20932 => Some(encoding_rs::EUC_JP),
        50220 => Some(encoding_rs::ISO_2022_JP),
        932 => Some(encoding_rs::SHIFT_JIS),
        1201 => Some(encoding_rs::UTF_16BE),
        1200 => Some(encoding_rs::UTF_16LE),
        949 => Some(encoding_rs::EUC_KR),
        // This is actually not a valid codepage but it happens with some unknown writer
        // Looking on internet, it looks like excel treats this as codepage 1200
        21010 => Some(encoding_rs::UTF_16LE),
        // Not available because not in the Encoding Standard
        //28591 => Some(encoding_rs::ISO_8859_1),
        //38598 => Some(encoding_rs::whatwg::ISO_8859_8_I),
        //52936 => Some(encoding_rs::HZ),
        _ => None,
    }
}

/// Push literal column into a String buffer
pub fn push_column(mut col: u32, buf: &mut String) {
    if col < 26 {
        buf.push((b'A' + col as u8) as char);
    } else {
        let mut rev = String::new();
        while col >= 26 {
            let c = col % 26;
            rev.push((b'A' + c as u8) as char);
            col -= c;
            col /= 26;
        }
        buf.extend(rev.chars().rev());
    }
}
