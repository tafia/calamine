// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Some utility functions to help with dealing with Compound File Binary (CFB)
//! format MS-CFB.

use std::borrow::Cow;
use std::cmp::min;

use encoding_rs::{Encoding, UTF_8};

use crate::utils::*;

/// A Cfb specific error enum
#[derive(Debug)]
pub enum CfbError {
    Io(std::io::Error),
    Invalid {
        name: &'static str,
        expected: &'static str,
        found: u16,
    },
    CodePageNotFound(u16),
}

impl std::fmt::Display for CfbError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            CfbError::Io(e) => write!(f, "I/O error: {e}"),
            CfbError::Invalid {
                name,
                expected,
                found,
            } => write!(f, "Invalid {name}, expecting {expected} found {found:X}"),
            CfbError::CodePageNotFound(e) => write!(f, "Codepage {e:X} not found"),
        }
    }
}

impl std::error::Error for CfbError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            CfbError::Io(e) => Some(e),
            _ => None,
        }
    }
}

/// Decompresses stream
pub fn decompress_stream(s: &[u8]) -> Result<Vec<u8>, CfbError> {
    const POWER_2: [usize; 16] = [
        1,
        1 << 1,
        1 << 2,
        1 << 3,
        1 << 4,
        1 << 5,
        1 << 6,
        1 << 7,
        1 << 8,
        1 << 9,
        1 << 10,
        1 << 11,
        1 << 12,
        1 << 13,
        1 << 14,
        1 << 15,
    ];

    let mut res = Vec::new();

    if s[0] != 0x01 {
        return Err(CfbError::Invalid {
            name: "signature",
            expected: "0x01",
            found: s[0] as u16,
        });
    }

    let mut i = 1;
    while i < s.len() {
        let chunk_header = read_u16(&s[i..]);
        i += 2;

        // each 'chunk' is 4096 wide, let's reserve that space
        let start = res.len();
        res.reserve(4096);

        let chunk_size = chunk_header & 0x0FFF;
        let chunk_signature = (chunk_header & 0x7000) >> 12;
        let chunk_flag = (chunk_header & 0x8000) >> 15;

        assert_eq!(chunk_signature, 0b011, "i={}, len={}", i, s.len());

        if chunk_flag == 0 {
            // uncompressed
            res.extend_from_slice(&s[i..i + 4096]);
            i += 4096;
        } else {
            let mut chunk_len = 0;
            let mut buf = [0u8; 4096];
            'chunk: loop {
                if i >= s.len() {
                    break;
                }

                let bit_flags = s[i];
                i += 1;
                chunk_len += 1;

                for bit_index in 0..8 {
                    if chunk_len > chunk_size {
                        break 'chunk;
                    }

                    if (bit_flags & (1 << bit_index)) == 0 {
                        // literal token
                        res.push(s[i]);
                        i += 1;
                        chunk_len += 1;
                    } else {
                        // copy token
                        let token = read_u16(&s[i..]);
                        i += 2;
                        chunk_len += 2;

                        let decomp_len = res.len() - start;
                        let bit_count = (4..16).find(|i| POWER_2[*i] >= decomp_len).unwrap();
                        let len_mask = 0xFFFF >> bit_count;
                        let mut len = (token & len_mask) as usize + 3;
                        let offset = ((token & !len_mask) >> (16 - bit_count)) as usize + 1;

                        while len > offset {
                            buf[..offset].copy_from_slice(&res[res.len() - offset..]);
                            res.extend_from_slice(&buf[..offset]);
                            len -= offset;
                        }
                        buf[..len]
                            .copy_from_slice(&res[res.len() - offset..res.len() - offset + len]);
                        res.extend_from_slice(&buf[..len]);
                    }
                }
            }
        }
    }
    Ok(res)
}

#[derive(Clone)]
pub struct XlsEncoding {
    encoding: &'static Encoding,
}

impl XlsEncoding {
    pub fn from_codepage(codepage: u16) -> Result<XlsEncoding, CfbError> {
        let e = codepage::to_encoding(codepage).ok_or(CfbError::CodePageNotFound(codepage))?;
        Ok(XlsEncoding { encoding: e })
    }

    fn high_byte(&self, high_byte: Option<bool>) -> Option<bool> {
        high_byte.or_else(|| {
            if self.encoding == UTF_8 || self.encoding.is_single_byte() {
                None
            } else {
                Some(false)
            }
        })
    }

    pub fn decode_to(
        &self,
        stream: &[u8],
        len: usize,
        s: &mut String,
        high_byte: Option<bool>,
    ) -> (usize, usize) {
        let (l, ub, bytes) = match self.high_byte(high_byte) {
            None => {
                let l = min(stream.len(), len);
                (l, l, Cow::Borrowed(&stream[..l]))
            }
            Some(false) => {
                let l = min(stream.len(), len);

                // add 0x00 high bytes to unicodes
                let mut bytes = vec![0; l * 2];
                for (i, sce) in stream.iter().take(l).enumerate() {
                    bytes[2 * i] = *sce;
                }
                (l, l, Cow::Owned(bytes))
            }
            Some(true) => {
                let l = min(stream.len() / 2, len);
                (l, 2 * l, Cow::Borrowed(&stream[..2 * l]))
            }
        };

        s.push_str(&self.encoding.decode(&bytes).0);
        (l, ub)
    }

    pub fn decode_all(&self, stream: &[u8]) -> String {
        self.encoding.decode(stream).0.into_owned()
    }
}
