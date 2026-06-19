// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Demonstrates reading worksheet hyperlinks from an XLSX file.
//!
//! Run the example like this:
//!
//! ```text
//! $ cargo run -q --example read_hyperlinks
//! ```

use calamine::{open_workbook, Hyperlink, Xlsx};

fn print_hyperlink(hyperlink: &Hyperlink) {
    let (start, end) = (hyperlink.range.start, hyperlink.range.end);
    println!(
        "  cell=({},{})..=({},{}) target={:?} location={:?} text={:?} tooltip={:?}",
        start.0 + 1,
        start.1 + 1,
        end.0 + 1,
        end.1 + 1,
        hyperlink.target.as_deref(),
        hyperlink.location.as_deref(),
        hyperlink.displayed_text.as_deref(),
        hyperlink.tooltip.as_deref(),
    );
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("tests/hyperlinks.xlsx")?;

    // Hyperlinks can be read by sheet name ...
    println!("Sheet \"Links\":");
    for hyperlink in workbook.hyperlinks_by_sheet_name("Links")? {
        print_hyperlink(&hyperlink);
    }

    // ... or by zero-based sheet index.
    println!("Sheet 0:");
    for hyperlink in workbook.hyperlinks_by_sheet_id(0)? {
        print_hyperlink(&hyperlink);
    }

    Ok(())
}
