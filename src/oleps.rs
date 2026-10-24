// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Read document properties from OLE property set streams.
//!
//! XLS files store their document properties in the `\005SummaryInformation`
//! and `\005DocumentSummaryInformation` streams of the compound file.
//!
//! # Reference
//!
//! [MS-OLEPS]: Object Linking and Embedding (OLE) Property Set Data Structures.
//!
//! [MS-OLEPS]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps

use std::collections::HashMap;

use crate::cfb::XlsEncoding;
use crate::WorkbookProperties;

/// FMTID of the SummaryInformation property set.
const FMTID_SUMMARY_INFORMATION: [u8; 16] = guid(
    0xF29F_85E0,
    0x4FF9,
    0x1068,
    [0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3, 0xD9],
);

/// FMTID of the DocumentSummaryInformation property set.
const FMTID_DOC_SUMMARY_INFORMATION: [u8; 16] = guid(
    0xD5CD_D502,
    0x2E9C,
    0x101B,
    [0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE],
);

/// FMTID of the user defined (custom) property set.
const FMTID_USER_DEFINED_PROPERTIES: [u8; 16] = guid(
    0xD5CD_D505,
    0x2E9C,
    0x101B,
    [0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9, 0xAE],
);

// Property identifiers.
const PID_DICTIONARY: u32 = 0x00;
const PID_CODEPAGE: u32 = 0x01;

// Property types.
const VT_I2: u16 = 0x0002;
const VT_I4: u16 = 0x0003;
const VT_R8: u16 = 0x0005;
const VT_BOOL: u16 = 0x000B;
const VT_UI4: u16 = 0x0013;
const VT_LPSTR: u16 = 0x001E;
const VT_LPWSTR: u16 = 0x001F;
const VT_FILETIME: u16 = 0x0040;

/// Code page identifier for UTF-16LE strings.
const CP_WINUNICODE: u16 = 1200;

/// Read the properties from a `\005SummaryInformation` stream.
pub(crate) fn read_summary_information(stream: &[u8], props: &mut WorkbookProperties) {
    let Some(section) = find_section(stream, &FMTID_SUMMARY_INFORMATION) else {
        return;
    };

    for (id, value) in section.properties() {
        let field = match id {
            0x02 => &mut props.title,
            0x03 => &mut props.subject,
            0x04 => &mut props.creator,
            0x05 => &mut props.keywords,
            0x06 => &mut props.description,
            0x07 => &mut props.template,
            0x08 => &mut props.last_modified_by,
            0x09 => &mut props.revision,
            0x0C => &mut props.created,
            0x0D => &mut props.modified,
            0x12 => &mut props.application,
            _ => continue,
        };
        *field = value.filter(|v| !v.is_empty());
    }
}

/// Read the properties, including the custom properties, from a
/// `\005DocumentSummaryInformation` stream.
pub(crate) fn read_doc_summary_information(stream: &[u8], props: &mut WorkbookProperties) {
    if let Some(section) = find_section(stream, &FMTID_DOC_SUMMARY_INFORMATION) {
        for (id, value) in section.properties() {
            let field = match id {
                0x02 => &mut props.category,
                0x0E => &mut props.manager,
                0x0F => &mut props.company,
                0x1B => &mut props.content_status,
                0x1D => &mut props.version,
                // The application version, as major.minor in the high and
                // low words, for example "16.0000".
                0x17 => {
                    props.app_version = section
                        .u32_value(id)
                        .map(|version| format!("{}.{:04}", version >> 16, version & 0xFFFF));
                    continue;
                }
                _ => continue,
            };
            *field = value.filter(|v| !v.is_empty());
        }
    }

    if let Some(section) = find_section(stream, &FMTID_USER_DEFINED_PROPERTIES) {
        let names = section.dictionary();
        for (id, value) in section.properties() {
            if let (Some(name), Some(value)) = (names.get(&id), value) {
                props.custom_properties.insert(name.clone(), value);
            }
        }
    }
}

/// A property set section within a property set stream.
struct Section<'a> {
    /// The section data, starting at the section header.
    data: &'a [u8],
    /// The code page used for 8-bit strings in the section.
    codepage: u16,
}

impl<'a> Section<'a> {
    fn new(data: &'a [u8]) -> Self {
        let mut section = Section {
            data,
            codepage: CP_WINUNICODE,
        };

        if let Some(codepage) = section
            .value_offset(PID_CODEPAGE)
            .filter(|&offset| read_u16(data, offset) == Some(VT_I2))
            .and_then(|offset| read_u16(data, offset.checked_add(4)?))
        {
            section.codepage = codepage;
        }

        section
    }

    /// Iterate over the property identifiers and their offsets.
    fn entries(&self) -> impl Iterator<Item = (u32, usize)> + '_ {
        let count = read_u32(self.data, 4).unwrap_or(0) as usize;

        (0..count).map_while(move |i| {
            let pos = i.checked_mul(8)?.checked_add(8)?;
            let id = read_u32(self.data, pos)?;
            let offset = read_u32(self.data, pos.checked_add(4)?)? as usize;
            Some((id, offset))
        })
    }

    /// Get the offset of a property's typed value.
    fn value_offset(&self, id: u32) -> Option<usize> {
        self.entries()
            .find(|&(pid, _)| pid == id)
            .map(|(_, offset)| offset)
    }

    /// Iterate over the property values, converted to strings.
    fn properties(&self) -> impl Iterator<Item = (u32, Option<String>)> + '_ {
        self.entries()
            .filter(|&(id, _)| id != PID_DICTIONARY && id != PID_CODEPAGE)
            .map(|(id, offset)| (id, self.string_value(offset)))
    }

    /// Get an unsigned integer property value.
    fn u32_value(&self, id: u32) -> Option<u32> {
        let offset = self.value_offset(id)?;
        match read_u16(self.data, offset)? {
            VT_I4 | VT_UI4 => read_u32(self.data, offset.checked_add(4)?),
            _ => None,
        }
    }

    /// Convert the typed property value at `offset` to a string.
    fn string_value(&self, offset: usize) -> Option<String> {
        let value = offset.checked_add(4)?;
        let value = match read_u16(self.data, offset)? {
            VT_I2 => (read_u16(self.data, value)? as i16).to_string(),
            VT_I4 => (read_u32(self.data, value)? as i32).to_string(),
            VT_UI4 => read_u32(self.data, value)?.to_string(),
            VT_R8 => f64::from_bits(read_u64(self.data, value)?).to_string(),
            VT_BOOL => (read_u16(self.data, value)? != 0).to_string(),
            VT_LPSTR => {
                let len = read_u32(self.data, value)? as usize;
                self.decode_string(read_bytes(self.data, value.checked_add(4)?, len)?)
            }
            VT_LPWSTR => {
                let len = read_u32(self.data, value)? as usize;
                decode_utf16(read_bytes(
                    self.data,
                    value.checked_add(4)?,
                    len.checked_mul(2)?,
                )?)
            }
            VT_FILETIME => filetime_to_iso8601(read_u64(self.data, value)?)?,
            _ => return None,
        };
        Some(value)
    }

    /// Read the dictionary of custom property names.
    fn dictionary(&self) -> HashMap<u32, String> {
        let mut names = HashMap::new();
        let Some(offset) = self.value_offset(PID_DICTIONARY) else {
            return names;
        };

        let count = read_u32(self.data, offset).unwrap_or(0);
        let mut pos = offset.saturating_add(4);
        for _ in 0..count {
            let (Some(id), Some(len)) = (
                read_u32(self.data, pos),
                read_u32(self.data, pos.saturating_add(4)).map(|len| len as usize),
            ) else {
                break;
            };
            pos = pos.saturating_add(8);

            // Unicode names are UTF-16 and padded to a multiple of 4 bytes.
            // Other names are 8-bit strings in the section code page.
            let (name, size) = if self.codepage == CP_WINUNICODE {
                let size = len.saturating_mul(2);
                let Some(bytes) = read_bytes(self.data, pos, size) else {
                    break;
                };
                (decode_utf16(bytes), size.saturating_add(3) & !3)
            } else {
                let Some(bytes) = read_bytes(self.data, pos, len) else {
                    break;
                };
                (self.decode_string(bytes), len)
            };

            names.insert(id, name);
            pos = pos.saturating_add(size);
        }

        names
    }

    /// Decode an 8-bit string in the section code page.
    fn decode_string(&self, bytes: &[u8]) -> String {
        let s = match XlsEncoding::from_codepage(self.codepage) {
            Ok(encoding) => encoding.decode_all(bytes),
            Err(_) => String::from_utf8_lossy(bytes).into_owned(),
        };
        s.trim_end_matches('\0').to_string()
    }
}

/// Find the property set section with the given FMTID.
fn find_section<'a>(stream: &'a [u8], fmtid: &[u8; 16]) -> Option<Section<'a>> {
    // The stream starts with a byte order mark of 0xFFFE.
    if read_u16(stream, 0)? != 0xFFFE {
        return None;
    }

    // The header is followed by up to 2 (FMTID, offset) pairs.
    let count = read_u32(stream, 24)?.min(2) as usize;
    (0..count).find_map(|i| {
        let pos = 28 + i * 20;
        if read_bytes(stream, pos, 16)? != fmtid {
            return None;
        }
        let offset = read_u32(stream, pos + 16)? as usize;
        let size = read_u32(stream, offset)? as usize;
        let end = offset.saturating_add(size).min(stream.len());
        Some(Section::new(stream.get(offset..end)?))
    })
}

/// Build a GUID in its little endian byte layout.
const fn guid(data1: u32, data2: u16, data3: u16, data4: [u8; 8]) -> [u8; 16] {
    let d1 = data1.to_le_bytes();
    let d2 = data2.to_le_bytes();
    let d3 = data3.to_le_bytes();
    [
        d1[0], d1[1], d1[2], d1[3], d2[0], d2[1], d3[0], d3[1], data4[0], data4[1], data4[2],
        data4[3], data4[4], data4[5], data4[6], data4[7],
    ]
}

fn read_bytes(data: &[u8], pos: usize, len: usize) -> Option<&[u8]> {
    data.get(pos..pos.checked_add(len)?)
}

fn read_u16(data: &[u8], pos: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        read_bytes(data, pos, 2)?.try_into().ok()?,
    ))
}

fn read_u32(data: &[u8], pos: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        read_bytes(data, pos, 4)?.try_into().ok()?,
    ))
}

fn read_u64(data: &[u8], pos: usize) -> Option<u64> {
    Some(u64::from_le_bytes(
        read_bytes(data, pos, 8)?.try_into().ok()?,
    ))
}

/// Decode a UTF-16LE string, removing any trailing null characters.
fn decode_utf16(bytes: &[u8]) -> String {
    let chars = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|c| u16::from_le_bytes(*c));
    let s = char::decode_utf16(chars)
        .map(|c| c.unwrap_or(char::REPLACEMENT_CHARACTER))
        .collect::<String>();
    s.trim_end_matches('\0').to_string()
}

/// Convert a Windows FILETIME, the number of 100ns intervals since
/// 1601-01-01, to an ISO 8601 UTC date time such as "2024-01-15T08:30:00Z".
///
/// A zero FILETIME means the time isn't set.
fn filetime_to_iso8601(filetime: u64) -> Option<String> {
    if filetime == 0 {
        return None;
    }

    let secs = filetime / 10_000_000;
    let days = (secs / 86_400) as i64;
    let secs = secs % 86_400;

    // Convert the days since 1601-01-01 to a civil date, using the algorithm
    // from https://howardhinnant.github.io/date_algorithms.html#civil_from_days
    // with the days shifted to be relative to 0000-03-01.
    let days = days + 584_694;
    let era = days / 146_097;
    let doe = days - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + i64::from(month <= 2);

    Some(format!(
        "{year:04}-{month:02}-{day:02}T{:02}:{:02}:{:02}Z",
        secs / 3_600,
        secs % 3_600 / 60,
        secs % 60
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_filetime_to_iso8601() {
        assert_eq!(filetime_to_iso8601(0), None);
        assert_eq!(
            filetime_to_iso8601(1).as_deref(),
            Some("1601-01-01T00:00:00Z")
        );
        // The Unix epoch.
        assert_eq!(
            filetime_to_iso8601(116_444_736_000_000_000).as_deref(),
            Some("1970-01-01T00:00:00Z")
        );
        // A leap day.
        assert_eq!(
            filetime_to_iso8601(132_274_080_000_000_000).as_deref(),
            Some("2020-02-29T00:00:00Z")
        );
        assert_eq!(
            filetime_to_iso8601(133_497_810_000_000_000).as_deref(),
            Some("2024-01-15T08:30:00Z")
        );
    }

    // Build a property set stream with a single section. Each property is an
    // id and its value bytes, including the type for typed values.
    fn property_set_stream(fmtid: &[u8; 16], props: &[(u32, Vec<u8>)]) -> Vec<u8> {
        let header_size = 8 + props.len() * 8;
        let mut entries = Vec::new();
        let mut values = Vec::new();
        for (id, value) in props {
            entries.extend(id.to_le_bytes());
            entries.extend(((header_size + values.len()) as u32).to_le_bytes());
            values.extend(value);
            values.resize(values.len().next_multiple_of(4), 0);
        }

        let mut stream = vec![0xFE, 0xFF, 0, 0, 0, 0, 0, 0];
        stream.extend([0; 16]);
        stream.extend(1u32.to_le_bytes());
        stream.extend(fmtid);
        stream.extend(48u32.to_le_bytes());
        stream.extend(((header_size + values.len()) as u32).to_le_bytes());
        stream.extend((props.len() as u32).to_le_bytes());
        stream.extend(entries);
        stream.extend(values);
        stream
    }

    fn typed(vt: u16, value: &[u8]) -> Vec<u8> {
        let mut bytes = u32::from(vt).to_le_bytes().to_vec();
        bytes.extend(value);
        bytes
    }

    // A null terminated UTF-16 string, as a character count and the characters.
    fn utf16(s: &str) -> Vec<u8> {
        let chars: Vec<u16> = s.encode_utf16().chain([0]).collect();
        let mut bytes = (chars.len() as u32).to_le_bytes().to_vec();
        bytes.extend(chars.iter().flat_map(|c| c.to_le_bytes()));
        bytes
    }

    #[test]
    fn test_custom_properties() {
        let names = ["Int", "Float", "Bool", "Date", "Text"];
        let mut dictionary = (names.len() as u32).to_le_bytes().to_vec();
        for (id, name) in (2u32..).zip(names) {
            dictionary.extend(id.to_le_bytes());
            dictionary.extend(utf16(name));
            dictionary.resize(dictionary.len().next_multiple_of(4), 0);
        }

        let stream = property_set_stream(
            &FMTID_USER_DEFINED_PROPERTIES,
            &[
                (PID_DICTIONARY, dictionary),
                (PID_CODEPAGE, typed(VT_I2, &CP_WINUNICODE.to_le_bytes())),
                (2, typed(VT_I4, &(-4i32).to_le_bytes())),
                (3, typed(VT_R8, &2.5f64.to_le_bytes())),
                (4, typed(VT_BOOL, &0xFFFFu16.to_le_bytes())),
                (
                    5,
                    typed(VT_FILETIME, &133_497_810_000_000_000u64.to_le_bytes()),
                ),
                (6, typed(VT_LPWSTR, &utf16("Café"))),
            ],
        );

        let mut props = WorkbookProperties::default();
        read_doc_summary_information(&stream, &mut props);

        let custom = &props.custom_properties;
        assert_eq!(custom.len(), 5);
        assert_eq!(custom["Int"], "-4");
        assert_eq!(custom["Float"], "2.5");
        assert_eq!(custom["Bool"], "true");
        assert_eq!(custom["Date"], "2024-01-15T08:30:00Z");
        assert_eq!(custom["Text"], "Café");
    }

    #[test]
    fn test_code_page_strings() {
        // "Ålesund" in Windows-1252, with a null terminator.
        let mut title = 8u32.to_le_bytes().to_vec();
        title.extend(b"\xC5lesund\0");

        let stream = property_set_stream(
            &FMTID_SUMMARY_INFORMATION,
            &[
                (PID_CODEPAGE, typed(VT_I2, &1252u16.to_le_bytes())),
                (0x02, typed(VT_LPSTR, &title)),
                (0x04, typed(VT_LPSTR, &1u32.to_le_bytes())),
            ],
        );

        let mut props = WorkbookProperties::default();
        read_summary_information(&stream, &mut props);

        assert_eq!(props.title.as_deref(), Some("Ålesund"));
        // An empty string is returned as `None`, as for the other formats.
        assert_eq!(props.creator, None);
    }

    #[test]
    fn test_malformed_streams_are_ignored() {
        let mut props = WorkbookProperties::default();
        read_summary_information(&[], &mut props);
        read_summary_information(&[0xFE, 0xFF, 0, 0], &mut props);
        read_doc_summary_information(&[0xFF; 64], &mut props);
        assert_eq!(props, WorkbookProperties::default());
    }
}
