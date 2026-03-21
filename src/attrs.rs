// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Zero-allocation XML attribute extraction utilities.
//!
//! These replace quick_xml's own `Attributes` iterator,
//! avoiding per-item overhead from `Result` wrapping,
//! `Cow`/`QName` newtypes, quote-type tracking, etc.

use quick_xml::escape::unescape;
use quick_xml::events::BytesStart;
use quick_xml::Decoder;

/// Zero-allocation iterator over raw XML attribute
/// bytes, yielding `(key, value)` byte-slice pairs.
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
    type Item = (&'a [u8], &'a [u8]);

    #[inline]
    fn next(&mut self) -> Option<Self::Item> {
        let raw = self.raw;
        let len = raw.len();

        // skip whitespace
        while self.pos < len && raw[self.pos].is_ascii_whitespace() {
            self.pos += 1;
        }
        if self.pos >= len {
            return None;
        }

        // key
        let key_start = self.pos;
        while self.pos < len && raw[self.pos] != b'=' {
            self.pos += 1;
        }
        if self.pos >= len {
            return None;
        }
        let key = &raw[key_start..self.pos];
        self.pos += 1; // skip '='
        if self.pos >= len {
            return None;
        }

        // quoted value
        let quote = raw[self.pos];
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        self.pos += 1; // skip opening quote
        let val_start = self.pos;
        while self.pos < len && raw[self.pos] != quote {
            self.pos += 1;
        }
        let val = &raw[val_start..self.pos];
        if self.pos < len {
            self.pos += 1; // skip closing quote
        }
        Some((key, val))
    }
}

/// Extension trait for fast/raw attribute access on XML elements.
pub(crate) trait RawAttributes {
    /// Iterate over all attributes as `(key, value)` byte-slice pairs.
    fn iter_raw_attrs(&self) -> RawAttrIter<'_>;

    /// Get a single attribute by name.
    #[inline]
    fn raw_attr(&self, name: &[u8]) -> Option<&[u8]> {
        self.iter_raw_attrs()
            .find_map(|(k, v)| (k == name).then_some(v))
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
        for (k, v) in $e.iter_raw_attrs() {
            match k {
                $($key => { $var = Some(v); found += 1; })+
                _ => {}
            }
            if found == total {
                break;
            }
        }
        ($($var),+)
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
