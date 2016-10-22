# office

Excel files reader in rust

[![Build Status](https://travis-ci.org/tafia/office.svg?branch=master)](https://travis-ci.org/tafia/office)
[![Build status](https://ci.appveyor.com/api/projects/status/nqagdg5o9evq31qu/branch/master?svg=true)](https://ci.appveyor.com/project/tafia/office/branch/master)

## Status

This is a work in progress, it evolves mailnly following my current needs.

So far it can read:
- worksheet and ranges for *.xlsx, *.xlsb and *.xlsm
- vba for all excel files, including old *.xls

## Performance

There is no benchmark but the first tests shows REALLY good performance, in particular for reading Compound File Binary [MS-CFB] data (vba).
I mainly compared to c# Office Interop, which obviously should do much more than what is done in this library, so it is not exactly fair.

## Usage

```rust
extern crate office;

use office::{Excel, VbaProject};

// ...

let excel = Excel::open(path).unwrap();
let vba = excel.vba().unwrap();

// ...
```
