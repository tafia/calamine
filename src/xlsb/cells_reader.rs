// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

use std::io::{Read, Seek};

use log::debug;

use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, format_excel_i64_ref, CellFormat},
    utils::{read_f64, read_i32, read_u32, read_usize, SharedFormulaIndex},
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

    /// Read the next formula cell.
    ///
    /// Each call returns the next cell in the worksheet that contains a
    /// formula, along with the formula text.
    ///
    /// Formulas that were filled or copied across a range, and array
    /// formulas, are returned as empty strings. Use
    /// [`XlsbCellsReader::formulas`] to read all formulas with these cells
    /// filled in.
    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsbError> {
        loop {
            match self.next_formula_record()? {
                None => return Ok(None),
                Some(FormulaRecord::Cell(cell)) => return Ok(Some(cell)),
                Some(FormulaRecord::SharedFormulaCell { pos }) => {
                    return Ok(Some(Cell::new(pos, String::new())))
                }
                Some(FormulaRecord::SharedFormula { .. }) => continue,
            }
        }
    }

    /// Read all the formula cells in the worksheet.
    ///
    /// Returns every cell in the worksheet that contains a formula, along with
    /// the formula text, in row and column order.
    ///
    /// Unlike [`XlsbCellsReader::next_formula`], formulas that were filled or
    /// copied across a range, and array formulas, are returned for every cell
    /// in the range.
    pub fn formulas(&mut self) -> Result<Vec<Cell<String>>, XlsbError> {
        // A shared or array formula is stored once, as a BrtShrFmla or
        // BrtArrFmla record over a range, and each cell in the range holds a
        // lone PtgExp pointing back to it. The shared formula follows the first
        // cell, so the shared formula cells are collected in one pass and
        // resolved at the end. Its tokens use the relative PtgRefN and PtgAreaN
        // forms, so it is decoded once per cell, anchored at that cell's
        // position.
        let mut cells = Vec::new();
        let mut shared_formula_cells: Vec<(u32, u32)> = Vec::new();
        let mut shared_formulas: Vec<(Dimensions, Vec<u8>)> = Vec::new();

        while let Some(record) = self.next_formula_record()? {
            match record {
                FormulaRecord::Cell(cell) => {
                    if !cell.get_value().is_empty() {
                        cells.push(cell);
                    }
                }
                FormulaRecord::SharedFormulaCell { pos } => shared_formula_cells.push(pos),
                FormulaRecord::SharedFormula { range, rgce } => shared_formulas.push((range, rgce)),
            }
        }

        let shared_formula_index =
            SharedFormulaIndex::new(shared_formulas.iter().map(|(range, _)| range));
        for pos in shared_formula_cells {
            let Some(index) = shared_formula_index.find(pos) else {
                continue;
            };

            // Decode the shared formula anchored at this cell's position, and
            // add the result if it isn't empty. If the shared formula can't be
            // decoded, log the error and skip the cell, so that the other
            // formulas on the sheet are still returned.
            let rgce = &shared_formulas[index].1;
            match parse_formula(rgce, self.extern_sheets, self.metadata_names, pos) {
                Ok(formula) if !formula.is_empty() => cells.push(Cell::new(pos, formula)),
                Ok(_) => {}
                Err(e) => debug!("{e}"),
            }
        }

        // Sort the cells into row and column order, as documented.
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
                    // BrtFmlaString.
                    let cch = read_u32(&self.buf[8..]) as usize;
                    let formula = &self.buf[14 + cch * 2..];
                    let cce = read_u32(formula) as usize;
                    &formula[4..4 + cce]
                }
                0x0009 => {
                    // BrtFmlaNum.
                    let formula = &self.buf[18..];
                    let cce = read_u32(formula) as usize;
                    &formula[4..4 + cce]
                }
                0x000A | 0x000B => {
                    // BrtFmlaBool | BrtFmlaError.
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
                0x01AA | 0x01AB => {
                    // BrtArrFmla | BrtShrFmla.
                    // BrtArrFmla has a one-byte flags field between the range
                    // and the token stream. BrtShrFmla doesn't.
                    let offset = if self.typ == 0x01AA { 17 } else { 16 };
                    match parse_shared_formula(&self.buf, offset) {
                        Some(shared_formula) => return Ok(Some(shared_formula)),
                        None => continue,
                    }
                }
                // BrtEndSheetData marks the end of the sheet's cell data.
                0x0092 => return Ok(None),
                // Ignore any other record and read the next one.
                _ => continue,
            };

            let pos = (self.row, read_u32(&self.buf));
            if shared_formula_anchor_row(rgce).is_some() {
                return Ok(Some(FormulaRecord::SharedFormulaCell { pos }));
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

    /// A cell whose formula is a pointer to a shared formula.
    SharedFormulaCell { pos: (u32, u32) },

    /// A shared formula, with the range of cells that use it.
    SharedFormula { range: Dimensions, rgce: Vec<u8> },
}

/// Split a `BrtShrFmla` or `BrtArrFmla` payload into the range it covers and
/// its token stream, which starts at `offset`.
///
/// Returns `None` if the record is too short to hold the token length, if the
/// token stream is empty, or if the token length runs past the end of the
/// record.
fn parse_shared_formula(buf: &[u8], offset: usize) -> Option<FormulaRecord> {
    if buf.len() < offset + 4 {
        return None;
    }

    let range = parse_dimensions(buf.get(..16)?);
    let cce = read_u32(&buf[offset..]) as usize;
    if cce == 0 || buf.len() < offset + 4 + cce {
        return None;
    }

    Some(FormulaRecord::SharedFormula {
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
mod tests {
    use super::{parse_shared_formula, FormulaRecord};

    /// Build a `BrtShrFmla` payload: the range as four `u32` values (first
    /// row, last row, first column, last column), then the token stream
    /// length and the tokens.
    fn shared_formula_payload(
        rw_first: u32,
        rw_last: u32,
        col_first: u32,
        col_last: u32,
        rgce: &[u8],
    ) -> Vec<u8> {
        let mut payload = Vec::new();
        for value in [rw_first, rw_last, col_first, col_last] {
            payload.extend_from_slice(&value.to_le_bytes());
        }
        payload.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
        payload.extend_from_slice(rgce);
        payload
    }

    #[test]
    fn test_parse_shared_formula() {
        // Test that a BrtShrFmla payload is split into its range and token
        // stream.
        let tokens = [0x4C, 0, 0, 0, 0, 0xFF, 0xC0];
        let payload = shared_formula_payload(2, 65, 3, 3, &tokens);

        let Some(FormulaRecord::SharedFormula { range, rgce }) = parse_shared_formula(&payload, 16)
        else {
            panic!("expected a shared formula");
        };

        assert_eq!((range.start, range.end), ((2, 3), (65, 3)));
        assert_eq!(rgce, tokens);
    }

    #[test]
    fn test_parse_shared_formula_array() {
        // Test that a BrtArrFmla payload is split into its range and token
        // stream. It has a one-byte flags field between the range and the
        // token stream, so the token stream starts at offset 17.
        let tokens = [0x1C, 0x2A];
        let mut payload = shared_formula_payload(0, 2, 2, 2, &tokens);
        let flags = 0x02;
        payload.insert(16, flags);

        let Some(FormulaRecord::SharedFormula { range, rgce }) = parse_shared_formula(&payload, 17)
        else {
            panic!("expected a shared formula");
        };

        assert_eq!((range.start, range.end), ((0, 2), (2, 2)));
        assert_eq!(rgce, tokens);
    }

    #[test]
    fn test_parse_shared_formula_malformed() {
        // Test that a malformed shared formula record is rejected.

        // The record is too short to hold the token stream length.
        assert!(parse_shared_formula(&[0u8; 8], 16).is_none());

        // The token stream is empty.
        let empty = shared_formula_payload(0, 1, 0, 0, &[]);
        assert!(parse_shared_formula(&empty, 16).is_none());

        // The token stream length runs past the end of the record.
        let mut overrun = shared_formula_payload(0, 1, 0, 0, &[0x4C, 0, 0]);
        overrun[16..20].copy_from_slice(&99u32.to_le_bytes());
        assert!(parse_shared_formula(&overrun, 16).is_none());
    }
}
