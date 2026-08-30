// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::{read_f64, read_i32, read_u32, read_usize},
    Cell, CellErrorType, Dimensions, XlsbError,
};

use super::{cell_format, parse_formula, shared_formula_anchor_row, wide_str, RecordIter};

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
        let _ = iter.next_skip_blocks(
            0x0094,
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
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
            let _ = self.iter.fill_buffer(&mut self.buf)?;
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
                            DataRef::Int(v)
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
                0x0006 | 0x0008 => DataRef::String(wide_str(&self.buf[8..], &mut 0)?.into_owned()), // BrtCellSt or BrtFmlaString
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&self.buf[8..12]);
                    DataRef::SharedString(&self.strings[isst])
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

    /// Read the next formula cell, skipping shared and array formula members.
    ///
    /// Members of a shared or array formula group decode to an empty string,
    /// because their token stream is a single `PtgExp` pointing at a definition
    /// stored elsewhere in the sheet. Resolving them needs a second pass over
    /// the sheet, so use [`XlsbCellsReader::formulas`] to get them.
    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsbError> {
        loop {
            match self.next_formula_record()? {
                None => return Ok(None),
                Some(FormulaRecord::Cell(cell)) => return Ok(Some(cell)),
                // Preserve the historical shape of this API: a member yields an
                // empty formula rather than disappearing from the iteration.
                Some(FormulaRecord::Member { pos }) => {
                    return Ok(Some(Cell::new(pos, String::new())))
                }
                Some(FormulaRecord::Definition { .. }) => continue,
            }
        }
    }

    /// Read every formula on the sheet, resolving shared and array formulas.
    ///
    /// Excel stores a run of repeated formulas once, as a `BrtShrFmla` or
    /// `BrtArrFmla` definition covering a range, and gives each cell in that
    /// range a lone `PtgExp` token pointing back at it. A single pass cannot
    /// resolve those, because the definition may appear after the members that
    /// refer to it, so members are collected and resolved at the end.
    ///
    /// A definition's token stream uses the relative `PtgRefN` and `PtgAreaN`
    /// forms, so it is decoded once per member, anchored at that member's own
    /// position. That is what makes each row of a filled-down column come out
    /// with its own correct references.
    pub fn formulas(&mut self) -> Result<Vec<Cell<String>>, XlsbError> {
        let mut cells = Vec::new();
        let mut members: Vec<(u32, u32)> = Vec::new();
        let mut definitions: Vec<(Dimensions, Vec<u8>)> = Vec::new();

        while let Some(record) = self.next_formula_record()? {
            match record {
                FormulaRecord::Cell(cell) => {
                    if !cell.get_value().is_empty() {
                        cells.push(cell);
                    }
                }
                FormulaRecord::Member { pos } => members.push(pos),
                FormulaRecord::Definition { range, rgce } => definitions.push((range, rgce)),
            }
        }

        // Index definitions by column. A sheet can hold thousands of them —
        // Excel splits a filled-down column into groups of about 64 rows — and
        // scanning them all for each of millions of members is quadratic enough
        // to dominate the read. Groups are nearly always one column wide;
        // anything wider stays in a small list scanned linearly, which keeps the
        // index compact without losing matches.
        const WIDE_GROUP_COLS: u32 = 64;
        let mut by_column: HashMap<u32, Vec<(u32, u32, usize)>> = HashMap::new();
        let mut wide: Vec<usize> = Vec::new();
        for (i, (range, _)) in definitions.iter().enumerate() {
            if range.end.1.saturating_sub(range.start.1) >= WIDE_GROUP_COLS {
                wide.push(i);
                continue;
            }
            for col in range.start.1..=range.end.1 {
                by_column
                    .entry(col)
                    .or_default()
                    .push((range.start.0, range.end.0, i));
            }
        }
        for spans in by_column.values_mut() {
            spans.sort_unstable();
        }

        for pos in members {
            let found = by_column.get(&pos.1).and_then(|spans| {
                // Groups in a column do not overlap, so the only candidate is
                // the last one starting at or before this row.
                let i = spans.partition_point(|&(start, _, _)| start <= pos.0);
                let &(_, end, index) = spans.get(i.checked_sub(1)?)?;
                (pos.0 <= end).then_some(index)
            });
            let Some(index) = found.or_else(|| {
                wide.iter().rev().copied().find(|&i| {
                    let range = &definitions[i].0;
                    pos.0 >= range.start.0
                        && pos.0 <= range.end.0
                        && pos.1 >= range.start.1
                        && pos.1 <= range.end.1
                })
            }) else {
                continue;
            };
            let rgce = &definitions[index].1;
            let formula = parse_formula(rgce, self.extern_sheets, self.metadata_names, pos)?;
            if !formula.is_empty() {
                cells.push(Cell::new(pos, formula));
            }
        }

        // `Range::from_sparse` expects row-major order, which the second pass
        // breaks by appending resolved members after the cells read inline.
        cells.sort_unstable_by_key(|c| c.get_position());
        Ok(cells)
    }

    /// Read the next formula-related record from the sheet stream.
    fn next_formula_record(&mut self) -> Result<Option<FormulaRecord>, XlsbError> {
        loop {
            self.typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;

            let rgce = match self.typ {
                0x0008 => {
                    // BrtFmlaString
                    let cch = read_u32(&self.buf[8..]) as usize;
                    let formula = &self.buf[14 + cch * 2..];
                    let cce = read_u32(formula) as usize;
                    &formula[4..4 + cce]
                }
                0x0009 => {
                    // BrtFmlaNum
                    let formula = &self.buf[18..];
                    let cce = read_u32(formula) as usize;
                    &formula[4..4 + cce]
                }
                0x000A | 0x000B => {
                    // BrtFmlaBool | BrtFmlaError
                    let formula = &self.buf[11..];
                    let cce = read_u32(formula) as usize;
                    &formula[4..4 + cce]
                }
                0x0000 => {
                    // BrtRowHdr
                    self.row = read_u32(&self.buf);
                    if self.row > 0x0010_0000 {
                        return Ok(None); // invalid row
                    }
                    continue;
                }
                // BrtArrFmla carries a one-byte flags field between the range
                // and the token stream; BrtShrFmla does not.
                0x01AA | 0x01AB => {
                    let offset = if self.typ == 0x01AA { 17 } else { 16 };
                    match parse_formula_definition(&self.buf, offset) {
                        Some(def) => return Ok(Some(def)),
                        None => continue,
                    }
                }
                0x0092 => return Ok(None), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };

            let pos = (self.row, read_u32(&self.buf));
            if shared_formula_anchor_row(rgce).is_some() {
                return Ok(Some(FormulaRecord::Member { pos }));
            }
            let formula = parse_formula(rgce, self.extern_sheets, self.metadata_names, pos)?;
            return Ok(Some(FormulaRecord::Cell(Cell::new(pos, formula))));
        }
    }
}

/// A record encountered while scanning a sheet for formulas.
enum FormulaRecord {
    /// A cell whose formula stands on its own.
    Cell(Cell<String>),
    /// A cell belonging to a shared or array formula group.
    Member { pos: (u32, u32) },
    /// The definition a group's members point at.
    Definition { range: Dimensions, rgce: Vec<u8> },
}

/// Decode a `BrtShrFmla` or `BrtArrFmla` payload into a range and token stream.
///
/// Returns `None` rather than a malformed definition if the declared token
/// length does not fit the record, so a misread never turns into a wrong
/// formula on every cell of a range.
fn parse_formula_definition(buf: &[u8], offset: usize) -> Option<FormulaRecord> {
    if buf.len() < offset + 4 {
        return None;
    }
    let range = parse_dimensions(buf.get(..16)?);
    let cce = read_u32(&buf[offset..]) as usize;
    if cce == 0 || buf.len() < offset + 4 + cce {
        return None;
    }
    Some(FormulaRecord::Definition {
        range,
        rgce: buf[offset + 4..offset + 4 + cce].to_vec(),
    })
}

fn parse_dimensions(buf: &[u8]) -> Dimensions {
    Dimensions {
        start: (read_u32(&buf[0..4]), read_u32(&buf[8..12])),
        end: (read_u32(&buf[4..8]), read_u32(&buf[12..16])),
    }
}

#[cfg(test)]
mod definition_tests {
    use super::{parse_formula_definition, FormulaRecord};

    /// Build a `BrtShrFmla`-shaped payload: `rfx(16) cce(4) rgce`.
    fn shr_fmla(
        rw_first: u32,
        rw_last: u32,
        col_first: u32,
        col_last: u32,
        rgce: &[u8],
    ) -> Vec<u8> {
        let mut v = Vec::new();
        for f in [rw_first, rw_last, col_first, col_last] {
            v.extend_from_slice(&f.to_le_bytes());
        }
        v.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
        v.extend_from_slice(rgce);
        v
    }

    #[test]
    fn shared_definition_layout() {
        // Excel chunks a filled-down column into groups of about 64 rows, so a
        // single-column range like this is the common shape.
        let payload = shr_fmla(2, 65, 3, 3, &[0x4C, 0, 0, 0, 0, 0xFF, 0xC0]);
        let Some(FormulaRecord::Definition { range, rgce }) =
            parse_formula_definition(&payload, 16)
        else {
            panic!("expected a definition");
        };
        assert_eq!((range.start, range.end), ((2, 3), (65, 3)));
        assert_eq!(rgce.len(), 7);
    }

    #[test]
    fn array_definition_skips_its_flags_byte() {
        // BrtArrFmla carries one flags byte between the range and the tokens.
        let mut payload = shr_fmla(0, 2, 2, 2, &[0x1C, 0x2A]);
        payload.insert(16, 0x02);
        let Some(FormulaRecord::Definition { range, rgce }) =
            parse_formula_definition(&payload, 17)
        else {
            panic!("expected a definition");
        };
        assert_eq!((range.start, range.end), ((0, 2), (2, 2)));
        assert_eq!(rgce, vec![0x1C, 0x2A]);
    }

    #[test]
    fn malformed_definitions_are_rejected_rather_than_guessed() {
        // A wrong offset would otherwise turn into a plausible but wrong
        // formula on every cell of a range, which is worse than none at all.
        assert!(parse_formula_definition(&[0u8; 8], 16).is_none());
        assert!(parse_formula_definition(&shr_fmla(0, 1, 0, 0, &[]), 16).is_none());

        // Token length overruns the record.
        let mut truncated = shr_fmla(0, 1, 0, 0, &[0x4C, 0, 0]);
        let len = truncated.len();
        truncated[16..20].copy_from_slice(&99u32.to_le_bytes());
        assert_eq!(truncated.len(), len);
        assert!(parse_formula_definition(&truncated, 16).is_none());
    }
}
