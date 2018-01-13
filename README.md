# calamine

An Excel/OpenDocument Spreadsheets file reader/deserializer, in pure Rust.

[![Build Status](https://travis-ci.org/tafia/calamine.svg?branch=master)](https://travis-ci.org/tafia/calamine)
[![Build status](https://ci.appveyor.com/api/projects/status/njpnhq54h5hxsgel/branch/master?svg=true)](https://ci.appveyor.com/project/tafia/calamine/branch/master)

[Documentation](https://docs.rs/calamine/)

## Description

**calamine** is a pure Rust library to read and deserialize any spreadsheet file:
- excel like (`xls`, `xlsx`, `xlsm`, `xlsb`, `xla`, `xlam`)
- opendocument spreadsheets (`ods`)

As long as your files are *simple enough*, this library should just work.
For anything else, please file an issue with a failing test or send a pull request!

## Examples

### Serde deserialization

It is as simple as:

```rust
use calamine::{Result, Sheets, RangeDeserializerBuilder};

fn example() -> Result<()> {
    let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook = Sheets::open(path)?;
    let range = workbook.worksheet_range("Sheet1")?;

    let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;

    if let Some(result) = iter.next() {
        let (label, value): (String, f64) = result?;
        assert_eq!(label, "celcius");
        assert_eq!(value, 22.2222);
        Ok(())
    } else {
        Err(From::from("expected at least one record but got none"))
    }
}
```

### Reader: Simple
```rust
let mut excel = Sheets::open("file.xlsx").unwrap();
let r = excel.worksheet_range("Sheet1").unwrap();
for row in r.rows() {
    println!("row={:?}, row[0]={:?}", row, row[0]);
}
```

### Reader: More complex

```rust
use calamine::{Sheets, Range, DataType};

// opens a new workbook
let path = "/path/to/my/excel/file.xlsm";
let mut workbook = Sheets::open(path).unwrap();

// Read whole worksheet data and provide some statistics
if let Ok(range) = workbook.worksheet_range("Sheet1") {
    let total_cells = range.get_size().0 * range.get_size().1;
    let non_empty_cells: usize = range.rows().map(|r| {
        r.iter().filter(|cell| cell != &&DataType::Empty).count()
    }).sum();
    println!("Found {} cells in 'Sheet1', including {} non empty cells",
             total_cells, non_empty_cells);
}

// Check if the workbook has a vba project
if workbook.has_vba() {
    let mut vba = workbook.vba_project().expect("Cannot find VbaProject");
    let vba = vba.to_mut();
    let module1 = vba.get_module("Module 1").unwrap();
    println!("Module 1 code:");
    println!("{}", module1);
    for r in vba.get_references() {
        if r.is_missing() {
            println!("Reference {} is broken or not accessible", r.name);
        }
    }
}

// You can also get defined names definition (string representation only)
for &(ref name, ref formula) in excel.defined_names().expect("Cannot get defined names!") {
    println!("name: {}, formula: {}", name, formula);
}

// Now get all formula!
let sheets = workbook.sheet_names().expect("Cannot get sheet names");
for s in sheets {
    println!("found {} formula in '{}'",
             workbook
                .worksheet_formula(&s)
                .expect("error while getting formula")
                .rows().flat_map(|r| r.iter().filter(|f| !f.is_empty()))
                .count(),
             s);
}
```

### Others

Browse the [examples](https://github.com/tafia/calamine/tree/master/examples) directory.

## Performance

While there is no official benchmark yet, my first tests show a significant boost compared to official C# libraries:
- Reading cell values: at least 3 times faster
- Reading vba code: calamine does not read all sheets when opening your workbook, this is not fair

## Unsupported

Many (most) part of the specifications are not implemented, the focus has been put on reading cell **values** and **vba** code.

The main unsupported items are:
- no support for writing excel files, this is a read-only library
- no support for reading extra contents, such as formatting, excel parameter, encrypted components etc ...
- no support for reading VB for opendocuments

## Credits

Thanks to [xlsx-js](https://github.com/SheetJS/js-xlsx) developers!
This library is by far the simplest open source implementation I could find and helps making sense out of official documentation.

## License

MIT
