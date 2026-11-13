// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use std::io::{Read, Seek};

use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, format_excel_i64_ref, CellFormat},
    utils::{read_f64, read_i32, read_u32, read_usize},
    Cell, CellErrorType, Dimensions, XlsbError,
};

use super::{cell_format, check_record_len, parse_formula, wide_str, RecordIter};

/// A cells reader for xlsb files
pub struct XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    iter: RecordIter<'a, RS>,
    formats: &'a [CellFormat],
    strings: &'a [String],
    extern_sheets: &'a [String],
    metadata_names: &'a [(String, String)],
    typ: u16,
    row: u32,
    is_1904: bool,
    dimensions: Dimensions,
    buf: Vec<u8>,
}

impl<'a, RS> XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub(crate) fn new(
        mut iter: RecordIter<'a, RS>,
        formats: &'a [CellFormat],
        strings: &'a [String],
        extern_sheets: &'a [String],
        metadata_names: &'a [(String, String)],
        is_1904: bool,
    ) -> Result<Self, XlsbError> {
        let mut buf = Vec::with_capacity(1024);
        // BrtWsDim
        let len = iter.next_skip_blocks(
            0x0094,
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
        // The buffer is reused for each record and is never shrunk, so it can be
        // longer than this record. Check the record length instead.
        check_record_len("BrtWsDim record length", len, 16)?;
        let dimensions = parse_dimensions(&buf[..16]);

        // BrtBeginSheetData
        let _ = iter.next_skip_blocks(
            0x0091,
            &[
                (0x0085, Some(0x0086)), // Views
                (0x0025, Some(0x0026)), // AC blocks
                (0x01E5, None),         // BrtWsFmtInfo
                (0x0186, Some(0x0187)), // Col Infos
            ],
            &mut buf,
        )?;

        Ok(XlsbCellsReader {
            iter,
            formats,
            is_1904,
            strings,
            extern_sheets,
            metadata_names,
            dimensions,
            typ: 0,
            row: 0,
            buf,
        })
    }

    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsbError> {
        // loop until end of sheet
        let value = loop {
            self.typ = self.iter.read_type()?;
            let len = self.iter.fill_buffer(&mut self.buf)?;
            // The fixed-size part of the records that are read below.
            let min_len = match self.typ {
                0x0000 => 4,                   // BrtRowHdr
                0x0002 | 0x0007 => 12,         // BrtCellRk, BrtCellIsst
                0x0003 | 0x0004 | 0x000A => 9, // BrtCellError, BrtCellBool, BrtFmlaBool
                0x0005 | 0x0009 => 16,         // BrtCellReal, BrtFmlaNum
                0x0006 | 0x0008 => 8,          // BrtCellSt, BrtFmlaString
                _ => 0,
            };
            check_record_len("cell record length", len, min_len)?;
            let value = match self.typ {
                // 0x0001 => continue, // Data::Empty, // BrtCellBlank
                0x0002 => {
                    // BrtCellRk MS-XLSB 2.5.122
                    let d100 = (self.buf[8] & 1) != 0;
                    let is_int = (self.buf[8] & 2) != 0;
                    self.buf[8] &= 0xFC;

                    if is_int {
                        let v = (read_i32(&self.buf[8..12]) >> 2) as i64;
                        if d100 {
                            let v = (v as f64) / 100.0;
                            format_excel_f64_ref(
                                v,
                                cell_format(self.formats, &self.buf),
                                self.is_1904,
                            )
                        } else {
                            // A whole-number date (e.g. 2021-01-01, 44197) is
                            // often stored as an integer RK, so it needs the
                            // cell's format too.
                            format_excel_i64_ref(
                                v,
                                cell_format(self.formats, &self.buf),
                                self.is_1904,
                            )
                        }
                    } else {
                        let mut v = [0u8; 8];
                        v[4..].copy_from_slice(&self.buf[8..12]);
                        let v = read_f64(&v);
                        let v = if d100 { v / 100.0 } else { v };
                        format_excel_f64_ref(v, cell_format(self.formats, &self.buf), self.is_1904)
                    }
                }
                0x0003 => {
                    let error = match self.buf[8] {
                        0x00 => CellErrorType::Null,
                        0x07 => CellErrorType::Div0,
                        0x0F => CellErrorType::Value,
                        0x17 => CellErrorType::Ref,
                        0x1D => CellErrorType::Name,
                        0x24 => CellErrorType::Num,
                        0x2A => CellErrorType::NA,
                        0x2B => CellErrorType::GettingData,
                        c => return Err(XlsbError::CellError(c)),
                    };
                    // BrtCellError
                    DataRef::Error(error)
                }
                0x0004 | 0x000A => DataRef::Bool(self.buf[8] != 0), // BrtCellBool or BrtFmlaBool
                0x0005 | 0x0009 => {
                    let v = read_f64(&self.buf[8..16]);
                    format_excel_f64_ref(v, cell_format(self.formats, &self.buf), self.is_1904)
                } // BrtCellReal or BrtFmlaNum
                0x0006 | 0x0008 => {
                    DataRef::String(wide_str(&self.buf[8..len], &mut 0)?.into_owned())
                } // BrtCellSt or BrtFmlaString
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&self.buf[8..12]);
                    let Some(string) = self.strings.get(isst) else {
                        return Err(XlsbError::Unrecognized {
                            typ: "shared string index",
                            val: isst.to_string(),
                        });
                    };
                    DataRef::SharedString(string)
                }
                0x0000 => {
                    // BrtRowHdr
                    self.row = read_u32(&self.buf);
                    if self.row > 0x0010_0000 {
                        return Ok(None); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(None), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };
            break value;
        };
        let col = read_u32(&self.buf);
        Ok(Some(Cell::new((self.row, col), value)))
    }

    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsbError> {
        let value = loop {
            self.typ = self.iter.read_type()?;
            let len = self.iter.fill_buffer(&mut self.buf)?;
            let record = &self.buf[..len];

            let value = match self.typ {
                // 0x0001 => continue, // Data::Empty, // BrtCellBlank
                0x0008 => {
                    // BrtFmlaString
                    check_record_len("formula record length", len, 12)?;
                    let cch = read_u32(&record[8..]) as usize;
                    let rgce = formula_rgce(record, cch.saturating_mul(2).saturating_add(14))?;
                    parse_formula(rgce, self.extern_sheets, self.metadata_names)?
                }
                0x0009 => {
                    // BrtFmlaNum
                    let rgce = formula_rgce(record, 18)?;
                    parse_formula(rgce, self.extern_sheets, self.metadata_names)?
                }
                0x000A | 0x000B => {
                    // BrtFmlaBool | BrtFmlaError
                    let rgce = formula_rgce(record, 11)?;
                    parse_formula(rgce, self.extern_sheets, self.metadata_names)?
                }
                0x0000 => {
                    // BrtRowHdr
                    check_record_len("cell record length", len, 4)?;
                    self.row = read_u32(&self.buf);
                    if self.row > 0x0010_0000 {
                        return Ok(None); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(None), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };
            break value;
        };
        let col = read_u32(&self.buf);
        Ok(Some(Cell::new((self.row, col), value)))
    }
}

/// Returns the rgce bytes of the formula that starts at `offset` in `record`.
fn formula_rgce(record: &[u8], offset: usize) -> Result<&[u8], XlsbError> {
    let too_short = || XlsbError::Unrecognized {
        typ: "formula record length",
        val: record.len().to_string(),
    };
    let formula = record.get(offset..).ok_or_else(too_short)?;
    if formula.len() < 4 {
        return Err(too_short());
    }
    let cce = read_u32(formula) as usize;
    formula[4..].get(..cce).ok_or_else(too_short)
}

fn parse_dimensions(buf: &[u8]) -> Dimensions {
    Dimensions {
        start: (read_u32(&buf[0..4]), read_u32(&buf[8..12])),
        end: (read_u32(&buf[4..8]), read_u32(&buf[12..16])),
    }
}
