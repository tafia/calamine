// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Demonstrates reading workbook properties from an XLSX or XLSB file.
//!
//! This example reads the document properties (such as
//! creator, application, and company) from a workbook.
//!
//! Run the example like this:
//!
//! ```text
//! $ cargo run -q --example read_properties -- tests/workbook_properties.xlsx
//!
//! Workbook properties:
//!   creator: Some("Test Creator")
//!   last_modified_by: Some("Last Modifier")
//!   created: Some("2024-01-15T08:30:00Z")
//!   modified: Some("2024-06-20T14:22:00Z")
//!   title: Some("Workbook Title")
//!   application: Some("Microsoft Excel")
//!   app_version: Some("16.0300")
//!   company: Some("Contoso & Co")
//!   template: Some("Book.xltx")
//!   manager: Some("Jane Doe")
//! ```

use calamine::open_workbook_auto;
use std::env;
use std::process::exit;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        eprintln!("Usage: {} <xlsx/xlsb path>", args[0]);
        exit(1);
    }

    let path = &args[1];
    let mut excel = match open_workbook_auto(path) {
        Ok(excel) => excel,
        Err(e) => {
            eprintln!("Cannot open {path}: {e}");
            exit(1);
        }
    };

    let props = match excel.workbook_properties() {
        Ok(props) => props,
        Err(e) => {
            eprintln!("Cannot read workbook properties from {path}: {e}");
            exit(1);
        }
    };

    println!("Workbook properties:");
    println!("  creator: {:?}", props.creator);
    println!("  last_modified_by: {:?}", props.last_modified_by);
    println!("  created: {:?}", props.created);
    println!("  modified: {:?}", props.modified);
    println!("  title: {:?}", props.title);
    println!("  application: {:?}", props.application);
    println!("  app_version: {:?}", props.app_version);
    println!("  company: {:?}", props.company);
    println!("  template: {:?}", props.template);
    println!("  manager: {:?}", props.manager);
}
