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
use calamine::{open_workbook, Error, Xlsx, Reader, RangeDeserializerBuilder};

fn example() -> Result<(), Error> {
    let path = format!("{}/tests/tempurature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")
        .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;

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
use calamine::{Reader, Xlsx, open_workbook};

let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {i
    for row in r.rows() {
        println!("row={:?}, row[0]={:?}", row, row[0]);
    }
}
```

### Reader: More complex

Let's assume
- the file type (xls, xlsx ...) cannot be known at static time
- we need to get all data from the workbook
- we need to parse the vba
- we need to see the defined names
- and the formula!

```rust
use calamine::{Reader, open_workbook_auto, Xlsx, DataType};

// opens a new workbook
let path = ...; // we do not know the file type
let mut workbook = open_workbook_auto(path).expect("Cannot open file");

// Read whole worksheet data and provide some statistics
if let Some(Ok(range)) = workbook.worksheet_range("Sheet1") {
    let total_cells = range.get_size().0 * range.get_size().1;
    let non_empty_cells: usize = range.used_cells().count();
    println!("Found {} cells in 'Sheet1', including {} non empty cells",
             total_cells, non_empty_cells);
    // alternatively, we can manually filter rows
    assert_eq!(non_empty_cells, range.rows()
        .flat_map(|r| r.iter().filter(|&c| c != &DataType::Empty)).count());
}

// Check if the workbook has a vba project
if let Some(Ok(mut vba)) = workbook.vba_project() {
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
for name in workbook.defined_names() {
    println!("name: {}, formula: {}", name.0, name.1);
}

// Now get all formula!
let sheets = workbook.sheet_names().to_owned();
for s in sheets {
    println!("found {} formula in '{}'",
             workbook
                .worksheet_formula(&s)
                .expect("sheet not found")
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

Thanks also to all the contributors!

## License

MIT
