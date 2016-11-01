# office

A Excel file reader, in Rust.

[![Build Status](https://travis-ci.org/tafia/office.svg?branch=master)](https://travis-ci.org/tafia/office)
[![Build status](https://ci.appveyor.com/api/projects/status/nqagdg5o9evq31qu/branch/master?svg=true)](https://ci.appveyor.com/project/tafia/office/branch/master)

## Description

**office** is a pure Rust library to process excel files. 

As long as your files are *simple enough*, this library is ready for use.

## Usage

```rust
use office::{Excel, Range, DataType};

// opens a new workbook
let path = "/path/to/my/excel/file.xlsm";
let mut workbook = Excel::open(path).unwrap();

// Check if the workbook has a vba project
if workbook.has_vba() {
    let mut vba = workbook.vba_project().unwrap();
    if let Ok((references, modules)) = vba.read_vba() {
        for m in modules {
            if &m.name == "module1" {
                println!("Module 1 code:");
                println!("{}", vba.read_module(&m).unwrap());
            }
        }
        for r in references {
            if r.is_missing() {
                println!("Reference {} is broken or not accessible", r.name);
            }
        }
    }
}

// Read whole worksheet data and provide some statistics
if let Ok(range) = workbook.worksheet_range("Sheet1") {
    let total_cells = range.get_size().0 * range.get_size().1;
    let non_empty_cells: usize = range.rows().map(|r| {
        r.iter().filter(|cell| cell != &&DataType::Empty).count()
    }).sum();
    println!("Found {} cells in 'Sheet1', including {} non empty cells",
             total_cells, non_empty_cells);
}
```

## Examples

Please look at [examples](https://github.com/tafia/office/tree/master/examples) folder.

## Performance

While is no official benchmark yet, my first tests show a major boost compared to official C# libraries:
- Reading cell values: at least 3 times faster
- Reading vba code: you do not need to read the cells before processing the vba part, which saves a crazy amount of time

## Warning

This library is very young and is not a transcription of an existing library.
As a result there is a large room for improvement: only items related to either cell values or vba is implemented.

## Unsupported

Many (most) part of the specifications are not implemented, the attention has been put on reading cell values and vba code.

The main unsupported items are:
- no support for reading `.xls` files (vba is ok though)
- no support for decoding MBSC vba code, office tries to decode as normal utf8, which is ok most of the time but not accurate
- no support for writing excel files, this is a read-only libray

## License

MIT
