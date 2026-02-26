// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Zero-allocation XML attribute extraction utilities.
//!
//! These replace quick_xml's own `Attributes` iterator, avoiding its per-item
//! `Cow`/`QName` newtypes, quote-type tracking, namespace bookkeeping, entity
//! decoding, and UTF8 validation. We can do this as we have a constrained
//! set of keys that we access internally (all of them are simple ASCII).
//!
//! Like quick_xml, malformed attributes are reported as [`AttrError`] items rather
//! than being skipped; a key with no `=` yields [`AttrError::ExpectedEq`] and an
//! unquoted values yield [`AttrError::UnquotedValue`]. (Unclosed quotes never
//! reach here as quick_xml's tokenizer rejects them at `read_event` time).

use quick_xml::escape::unescape;
use quick_xml::events::attributes::AttrError;
use quick_xml::events::BytesStart;
use quick_xml::Decoder;

/// Zero-allocation iterator over raw XML attribute bytes, yielding
/// `(key, value)` byte-slice pairs or an [`AttrError`] for malformed input.
pub(crate) struct RawAttrIter<'a> {
    raw: &'a [u8],
    pos: usize,
}

impl<'a> RawAttrIter<'a> {
    #[inline]
    fn new(raw: &'a [u8]) -> Self {
        Self { raw, pos: 0 }
    }
}

impl<'a> Iterator for RawAttrIter<'a> {
    type Item = Result<(&'a [u8], &'a [u8]), AttrError>;

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let raw = self.raw;
        let len = raw.len();

        // note: make "pos" local to ensure the compiler keeps it in
        // a register across the scan loops (struct field otherwise)
        let mut pos = self.pos;

        // skip whitespace before the key
        while pos < len && raw[pos].is_ascii_whitespace() {
            pos += 1;
        }
        if pos >= len {
            self.pos = pos;
            return None;
        }

        // key: scan to '=' with a single comparison
        // per byte (handling whitespace around '=')
        let key_start = pos;
        while pos < len && raw[pos] != b'=' {
            pos += 1;
        }
        if pos >= len {
            // a key with no '=' is malformed; park the cursor at the end so the
            // next call returns None rather than re-reporting the same error
            self.pos = len;
            return Some(Err(AttrError::ExpectedEq(key_start)));
        }
        let key = raw[key_start..pos].trim_ascii_end();
        pos += 1; // skip '='

        // skip whitespace after '=' (normally none: loop exits immediately)
        while pos < len && raw[pos].is_ascii_whitespace() {
            pos += 1;
        }

        // quoted value (XML mandates quotes; anything else, including a missing
        // value at end-of-input, is malformed)
        let quote = raw.get(pos).copied();
        if quote != Some(b'"') && quote != Some(b'\'') {
            self.pos = len;
            return Some(Err(AttrError::UnquotedValue(pos)));
        }
        let quote = quote.unwrap();
        pos += 1; // skip opening quote
        let val_start = pos;
        while pos < len && raw[pos] != quote {
            pos += 1;
        }
        let val = &raw[val_start..pos];
        if pos < len {
            pos += 1; // skip closing quote
        }
        self.pos = pos;
        Some(Ok((key, val)))
    }
}

/// Extension trait for fast/raw attribute access on XML elements.
pub(crate) trait RawAttributes {
    /// Iterate over attributes, yielding `(key, value)` byte-slice
    /// pairs (or an [`AttrError`] for malformed input).
    fn iter_raw_attrs(&self) -> RawAttrIter<'_>;

    /// Get a single attribute by name, returning `Ok(None)` if absent and
    /// an `Err` if a malformed attribute is encountered while scanning.
    #[inline]
    fn raw_attr(&self, name: &[u8]) -> Result<Option<&[u8]>, AttrError> {
        for item in self.iter_raw_attrs() {
            let (k, v) = item?;
            if k == name {
                return Ok(Some(v));
            }
        }
        Ok(None)
    }
}

impl RawAttributes for BytesStart<'_> {
    #[inline]
    fn iter_raw_attrs(&self) -> RawAttrIter<'_> {
        RawAttrIter::new(self.attributes_raw())
    }
}

/// Get a set of named attributes from an element in a single
/// pass, with early exit as soon as all items are found.
macro_rules! get_attrs {
    ($e:expr, $($key:expr => $var:ident),+ $(,)?) => {{
        $(let mut $var = None;)+
        let mut found = 0u8;
        let total = get_attrs!(@count $($key),+);
        let mut result = Ok(());
        for item in $e.iter_raw_attrs() {
            match item {
                Ok((k, v)) => {
                    match k {
                        $($key => { $var = Some(v); found += 1; })+
                        _ => {}
                    }
                    if found == total {
                        break;
                    }
                }
                Err(e) => {
                    result = Err(e);
                    break;
                }
            }
        }
        result.map(|()| ($($var),+))
    }};
    (@count $first:expr $(, $rest:expr)*) => {
        1u8 $(+ get_attrs!(@count_one $rest))*
    };
    (@count_one $e:expr) => { 1u8 };
}

/// Decode raw attribute bytes into a `String`, with XML entity unescaping.
/// Only needed for values that can contain entities (eg: sheet names, table names, etc).
pub(crate) fn decode_attr(decoder: &Decoder, val: &[u8]) -> Result<String, quick_xml::Error> {
    let decoded = decoder.decode(val)?;
    let unescaped = unescape(&decoded).map_err(quick_xml::Error::from)?;
    Ok(unescaped.into_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Collect well-formed attributes, panicking on any malformed item.
    fn collect(raw: &[u8]) -> Vec<(&[u8], &[u8])> {
        RawAttrIter::new(raw)
            .map(|r| r.expect("unexpected malformed attribute"))
            .collect()
    }

    #[test]
    fn test_basic_attrs() {
        let bytes = b"key1=\"val1\" key2='val2'";
        let mut iter = RawAttrIter::new(bytes);
        assert_eq!(iter.next(), Some(Ok((&b"key1"[..], &b"val1"[..]))));
        assert_eq!(iter.next(), Some(Ok((&b"key2"[..], &b"val2"[..]))));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_whitespace_around_equals() {
        let bytes = b"key = \"value\"";
        let mut iter = RawAttrIter::new(bytes);
        assert_eq!(iter.next(), Some(Ok((&b"key"[..], &b"value"[..]))));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_empty_value() {
        let bytes = b"key=\"\"";
        let mut iter = RawAttrIter::new(bytes);
        assert_eq!(iter.next(), Some(Ok((&b"key"[..], &b""[..]))));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_no_trailing_space() {
        let bytes = b"key=\"value\"";
        let mut iter = RawAttrIter::new(bytes);
        assert_eq!(iter.next(), Some(Ok((&b"key"[..], &b"value"[..]))));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_no_attributes() {
        assert_eq!(collect(b""), vec![]);
        assert_eq!(collect(b"   "), vec![]);
    }

    #[test]
    fn test_leading_and_trailing_whitespace() {
        // `attributes_raw()` yields a leading space; a self-closing tag can
        // leave a trailing one (e.g. `<c r="A1" />` -> ` r="A1" `).
        assert_eq!(collect(b" r=\"A1\" "), vec![(&b"r"[..], &b"A1"[..])]);
    }

    #[test]
    fn test_mixed_whitespace_and_quotes() {
        // Whitespace around `=` combined with both quote styles, as a real
        // (if unusual) tag like `<row r = "50" spans = '1:4'>` would produce.
        let bytes = b" r = \"50\" spans = '1:4'";
        assert_eq!(
            collect(bytes),
            vec![(&b"r"[..], &b"50"[..]), (&b"spans"[..], &b"1:4"[..])]
        );
    }

    #[test]
    fn test_value_containing_other_quote() {
        // A single-quoted value may contain double quotes and vice versa.
        assert_eq!(collect(b"k='a\"b'"), vec![(&b"k"[..], &b"a\"b"[..])]);
        assert_eq!(collect(b"k=\"a'b\""), vec![(&b"k"[..], &b"a'b"[..])]);
    }

    #[test]
    fn test_namespaced_attributes() {
        // Namespace declarations and prefixed names are yielded verbatim with
        // their prefix; call sites match the full raw key (see issue #632).
        let bytes = b" xmlns:foo=\"bar\" foo:id=\"x\" r:id=\"rId1\"";
        assert_eq!(
            collect(bytes),
            vec![
                (&b"xmlns:foo"[..], &b"bar"[..]),
                (&b"foo:id"[..], &b"x"[..]),
                (&b"r:id"[..], &b"rId1"[..]),
            ]
        );
    }

    #[test]
    fn test_malformed_unquoted_value_reports_error() {
        // Unquoted values are malformed: yield the valid attr,
        // then the UnquotedValue error, then stop.
        let mut iter = RawAttrIter::new(b"good=\"1\" bad=2");
        assert_eq!(iter.next(), Some(Ok((&b"good"[..], &b"1"[..]))));
        assert!(matches!(
            iter.next(),
            Some(Err(AttrError::UnquotedValue(_)))
        ));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_no_equals_reports_error() {
        // A bare name with no '=' is reported as ExpectedEq, then iteration ends.
        let mut iter = RawAttrIter::new(b"disabled");
        assert!(matches!(iter.next(), Some(Err(AttrError::ExpectedEq(_)))));
        assert_eq!(iter.next(), None);

        // Leading whitespace then a bare name behaves the same.
        let mut iter = RawAttrIter::new(b"  spans  ");
        assert!(matches!(iter.next(), Some(Err(AttrError::ExpectedEq(_)))));
        assert_eq!(iter.next(), None);
    }

    #[test]
    fn test_tabs_and_newlines_as_separators() {
        // Attributes may be split by tabs/newlines, not just spaces.
        assert_eq!(
            collect(b"\tr=\"A1\"\n\ts=\"3\""),
            vec![(&b"r"[..], &b"A1"[..]), (&b"s"[..], &b"3"[..])]
        );
    }

    #[test]
    fn test_raw_attr_lookup() {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        let mut reader = Reader::from_str(r#"<c r="A1" s="3" t="s"/>"#);
        let Event::Empty(e) = reader.read_event().unwrap() else {
            panic!("expected empty element");
        };
        assert_eq!(e.raw_attr(b"r"), Ok(Some(&b"A1"[..])));
        assert_eq!(e.raw_attr(b"s"), Ok(Some(&b"3"[..])));
        assert_eq!(e.raw_attr(b"t"), Ok(Some(&b"s"[..])));
        assert_eq!(e.raw_attr(b"missing"), Ok(None));
    }

    #[test]
    fn test_get_attrs_macro() {
        use quick_xml::events::Event;
        use quick_xml::Reader;

        // Test with/without whitespace around `=`
        let mut reader = Reader::from_str(r#"<c r = "A1" t="s"/>"#);
        let Event::Empty(e) = reader.read_event().unwrap() else {
            panic!("expected empty element");
        };
        let (r, s, t) = get_attrs!(e, b"r" => r, b"s" => s, b"t" => t).unwrap();
        assert_eq!(r, Some(&b"A1"[..]));
        assert_eq!(s, None);
        assert_eq!(t, Some(&b"s"[..]));
    }
}
