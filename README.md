# calamine

An Excel/OpenDocument Spreadsheets file reader/deserializer, in pure Rust.

[![GitHub CI Rust tests](https://github.com/tafia/calamine/workflows/Rust/badge.svg)](https://github.com/tafia/calamine/actions)
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
    let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")
        .ok_or(Error::Msg("Cannot find 'Sheet1'"))??;

    let mut iter = RangeDeserializerBuilder::new().from_range(&range)?;

    if let Some(result) = iter.next() {
        let (label, value): (String, f64) = result?;
        assert_eq!(label, "celsius");
        assert_eq!(value, 22.2222);
        Ok(())
    } else {
        Err(From::from("expected at least one record but got none"))
    }
}
```

Note if you want to deserialize a column that may have invalid types (i.e. a float where some values may be strings), you can use Serde's `deserialize_with` field attribute:

```rust
use serde::Deserialize;
use calamine::{RangeDeserializerBuilder, Reader, Xlsx};


#[derive(Deserialize)]
struct ExcelRow {
    metric: String,
    #[serde(deserialize_with = "de_opt_f64")]
    value: Option<f64>,
}


// Convert value cell to Some(f64) if float or int, else None
fn de_opt_f64<'de, D>(deserializer: D) -> Result<Option<f64>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    let data_type = calamine::DataType::deserialize(deserializer)?;
    if let Some(float) = data_type.as_f64() {
        Ok(Some(float))
    } else {
        Ok(None)
    }
}

fn main() ->  Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}/tests/excel.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut excel: Xlsx<_> = open_workbook(path)?;

    let range = excel
      .worksheet_range("Sheet1")
      .ok_or(calamine::Error::Msg("Cannot find Sheet1"))??;

    let iter_result =
        RangeDeserializerBuilder::with_headers(&COLUMNS).from_range::<_, ExcelRow>(&range)?;
  }
```


### Reader: Simple

```rust
use calamine::{Reader, Xlsx, open_workbook};

let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
if let Some(Ok(r)) = excel.worksheet_range("Sheet1") {
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

## Features

- `dates`: Add date related fn to `DataType`. 
- `picture`: Extract picture data.

### Others

Browse the [examples](https://github.com/tafia/calamine/tree/master/examples) directory.

## Performance

As `calamine` is readonly, the comparisons will only involve reading an excel `xlsx` file and then iterating over the rows. Along with `calamine`, two other libraries were chosen, from two different languages, [`excelize`](https://github.com/qax-os/excelize) written in `go` and [`ClosedXML`](https://github.com/ClosedXML/ClosedXML) written in `C#`. The benchmarks were done using this [dataset](https://raw.githubusercontent.com/wiki/jqnatividad/qsv/files/NYC_311_SR_2010-2020-sample-1M.7z), a `186MB` `xlsx` file. The plotting data was gotten from the [`sysinfo`](https://github.com/GuillaumeGomez/sysinfo) crate, at a sample interval of `200ms`. The program samples the reported values for the running process and records it.

The programs are all structured to follow the same constructs:

`calamine`:
```rust
use calamine::{open_workbook, Reader, Xlsx};

fn main() {
    // Open workbook 
    let mut excel: Xlsx<_> =
        open_workbook("NYC_311_SR_2010-2020-sample-1M.xlsx").expect("failed to find file");

    // Get worksheet
    let sheet = excel
        .worksheet_range("NYC_311_SR_2010-2020-sample-1M")
        .unwrap()
        .unwrap();

    // iterate over rows
    for _row in sheet.rows() {}
}
```
`excelize`:
```go
package main

import (
        "fmt"
        "github.com/xuri/excelize/v2"
)

func main() {
        // Open workbook
        file, err := excelize.OpenFile(`NYC_311_SR_2010-2020-sample-1M.xlsx`)

        if err != nil {
                fmt.Println(err)
                return
        }

        defer func() {
                // Close the spreadsheet.
                if err := file.Close(); err != nil {
                        fmt.Println(err)
                }
        }()

        // Get worksheet
        rows, err := file.GetRows("NYC_311_SR_2010-2020-sample-1M")
        if err != nil {
                fmt.Println(err)
                return
        }

        // Iterate over rows
        for _, row := range rows {
                _ = row
        }
}
```
`ClosedXML`:
```csharp
using ClosedXML.Excel;

internal class Program
{
        private static void Main(string[] args)
        {
                // Open workbook
                using var workbook = new XLWorkbook("NYC_311_SR_2010-2020-sample-1M.xlsx");

                // Get Worksheet
                // "NYC_311_SR_2010-2020-sample-1M"
                var worksheet = workbook.Worksheet(1);

                // Iterate over rows
                foreach (var row in worksheet.Rows())
                {

                }
        }
}
```

### Benchmarks

The benchmarking was done using [`hyperfine`](https://github.com/sharkdp/hyperfine) with `--warmup 3` on an `AMD RYZEN 9 5900X @ 4.0GHz` running `Windows 11`. Both `calamine` and `ClosedXML` were built in release mode.

```bash
0.22.1 calamine.exe
  Time (mean ± σ):     25.278 s ±  0.424 s    [User: 24.852 s, System: 0.470 s]
  Range (min … max):   24.980 s … 26.369 s    10 runs

v2.8.0 excelize.exe
  Time (mean ± σ):     199.709 s ± 11.671 s    [User: 158.678 s, System: 69.350 s]
  Range (min … max):   193.934 s … 232.725 s    10 runs

0.102.1 closedxml.exe
  Time (mean ± σ):     178.343 s ±  3.673 s    [User: 177.442 s, System: 2.612 s]
  Range (min … max):   173.232 s … 185.086 s    10 runs
```

The spreadsheet has a range of 1,000,001 rows and 41 columns, for a total of 41,000,041 cells in the range. Of those, 28,056,975 cells had values. Going off of that number, `calamine` had a rate of 1,122,279 cells per second. `excelize` had 140,989 cells per second. `ClosedXML` had 157,320 cells per second.

`calamine` comes is 7.9x faster than `excelize`, and 7x faster than `ClosedXML`.

### Plots

#### Disk Read
![bytes_from_disk](https://github.com/RoloEdits/calamine/assets/12489689/b662c037-f34b-437e-b6b8-331f07e2a15b)

The filesize on disk is `186MB`. `calamine` read just that and nothing more. `ClosedXML` read a little more, `208MB`. `Excelize` on the other hand managed to read `2.12GB` from disk. `calamine` did a 1:1 here, with `ClosedXML` just a little over that. And `excelize` read the file 11x over.

#### Disk Write
![bytes_to_disk](https://github.com/RoloEdits/calamine/assets/12489689/1bde4d93-eb50-472d-bd4a-40471f39b2da)

As the programs have no writing involved, I thought this would be a pointless exercise and an empty section, but Excelize is doing something very strange here. I have no clue what its writing. The other two had expected behavior. The programs had no logic for writing and therefore nothing was written.

#### Memory
![mem_usage](https://github.com/RoloEdits/calamine/assets/12489689/fdc1129e-bc0f-4133-9a2c-9eb09d7c0390)

![virt_mem_usage](https://github.com/RoloEdits/calamine/assets/12489689/59009854-3863-4fbc-8c3a-5ba422429f13)
> [!NOTE]
> `ClosedXML` was reporting a constant `2.5TB` of virtual memory usage, so it was excluded from the chart.

The stepping and falling for `calamine` is from the grows of `Vec`s and the freeing of of memory right after, with the memory usage dropping down again. The sudden jump at the end is when the sheet is being read into memory. The other two, being garbage collected, have a more linear climb all the way through.

#### CPU
![cpu_usage](https://github.com/RoloEdits/calamine/assets/12489689/3cf046ba-7661-4bc5-8507-24a8960dfdc3)

Very noisy chart, but `calamine` and `ClosedXML` are very close on CPU usage. `excelize`'s spikes must be from the GC?

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
