// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Demonstrates streaming XLSX cell values and formulas in one worksheet pass.
//!
//! Run the example like this:
//!
//! ```text
//! $ cargo run -q --example xlsx_formula_stream
//! ```

use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("tests/formula.issue.xlsx")?;
    let mut reader = workbook.worksheet_cells_reader("Sheet1")?;

    while let Some(cell) = reader.next_cell_with_formula()? {
        if let Some(formula) = cell.formula {
            println!(
                "row={}, col={}, value={:?}, formula={}",
                cell.pos.0 + 1,
                cell.pos.1 + 1,
                cell.value,
                formula
            );
        }
    }

    Ok(())
}
