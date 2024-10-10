> Legend:
  - feat: A new feature
  - fix: A bug fix
  - docs: Documentation only changes
  - style: White-space, formatting, missing semi-colons, etc
  - refactor: A code change that neither fixes a bug nor adds a feature
  - perf: A code change that improves performance
  - test: Adding missing tests
  - chore: Changes to the build process or auxiliary tools/libraries/documentation

## Unreleased

## 0.26.1 (2024-10-10)

- fix: sparse celle expect 0 index rows, even when using `header_row`

## 0.26.0 (2024-10-08)

- feat: add ability to merge cells from xls and xlsx
- feat: add options to keep first empty rows for xlsx
- feat: support consecutive repeated empty cells for ods
- feat: add new `header_row` config
- test: fix broken fuzzer
- docs: fix broken links in README
- docs: enable dates and pictures features in docs.rs build
- chore: bump MRSV to 1.73


## 0.25.0 (2024-05-25)

- feat: added `is_error` and `get_error` methods to the `DataType` trait
- feat: add deserializer helper functions
- feat: support get merged region
- feat: add `Range::headers` method
- feat: expose some `Dimensions` methods
- fix: use case insensitive comparison when searching for file in xlsx
- fix: do not panic when reading cell format with invalid index
- chore: use `OnceLock` instead of once_cell crate (MSRV: 1.71)

## 0.24.0

- refactor (breaking): rename `DataType` enum to `Data` and `DataTypeRef` to `DataRef`
- feat: introduce a `DataType` trait implemented by both `Data` and `DataRef`.
- feat: `Data` and `DataType` now return `Some(0{.0})` and `Some(1{.0})` rather than `None` when `.as_i64` or `.as_f64`
  is used on a Bool value
- fix: getting tables names on xlsx workbook without _rels files
- refactor: DateTime(f64) to DateTime(ExcelDateTime)
- feat: detect xlsb/ods password protected files
- feat: introduce is_x methods for date and time variants

## 0.23.1

- fix: `worksheet_formula` not returning all formula

## 0.23.0

- feat: add new `DataTypeRef` available from `worksheet_range_ref` to reduce memory usage
- docs: add benchmark plot
- fix: truncated text in xls
- feat: detect if workbook is password protected

## 0.22.1

- fix: regression on `Range::get`
- chore: update GitHub actions
- fix: spelling of formula error type
- feat: support label cells for xls
- refactor: clippy
- perf: preallocate several buffers

## 0.22.0

- docs: improve `de_opt_f64` example.
- docs: remove datetime notice from README
- feat: add support of sheet type and visibility
- style: clippy
- chore: bump MSRV to 1.63 (breaking)
- chore: set edition to 2021
- feat: implement blank string handling

## 0.21.2:

- fix: formula with string not displaying properly

## 0.21.1

- chore: bump MSRv to 1.60.0 due to log dependencies
- fix: xls: formula values ignored
- fix: xls: formula (string) not displayed properly

## 0.21.0

- fix: (xlsx) support `r` attribute
- fix: support `PROJECTCOMPATVERSION` in vba
- fix: incorrect date parsing due to excel bug
- feat: support for duration
- test: add MSRV

## 0.20.0

- refactor: bump dependencies
- feat: __(all) parse format/style information to infer cell as datetime__
- feat: (ods) support number-columns-repeated attribute
- refactor: multiple clippy refactorings

## 0.19.2

- feat: Extract picture data by turning `picture` feature on

## 0.19.1

- fix: wrong range len calculation
- fix: date precision

## 0.19.0

- fix: Several decoding issues in xls and xlsb
- fix: Wrong decimal parsing
- refactor: bump dependencies (quick-xml in particular and chrono)
- feat: always return sheet names in lexicographic order (BTreeMap)
- chore: remove travis

## 0.18.0

- fix: Allow empty value cells in xlsx
- fix: obscure xls parsing errors (#195)
- feat: Improve conversions from raw data to primitives
- docs: fix two typos in readme
- feat: replace macro matches! by match expression to reduce MSRV

## 0.17.0

- feat: use `chunks_exact` instead of chunks where possible
- fix: make `to_u32`, `read_slice` safe and sound
- fix: security issue #199
- test: add regression tests that fail with miri
- feat: detect date/time formatted cells in XLSX
- feat: brute force file detection if extension is not known
- feat: support xlsx sheet sizes beyond u32::MAX
- test: ensure doctest functions actually run
- test: run cargo fmt to fix travis tests
- fix: fix Float reading for XLSB

## 0.16.2
- docs: add `deserialize_with` example in readme
- fix: Skip phonetic run
- fix: Fix XLS float parsing error
- docs: Correct MBSC to MBCS in vba.rs (misspelled before)
- style: use 2018 edition paths
- fix: Add the ability to read formula values from XLSB
- fix: support integral date types

## 0.16.1
- feat: Make `Metadata.sheets` (and `Reader.sheet_names`) always return names in workbook order
- style: fix warnings in tests

## 0.16.0
- feat: deprecate failure and impl `std::error::Error` for all errors.
- feat: add `dates` feature to enrich `DataType` with date conversions fns.

## 0.15.6
- feat: update dependencies

## 0.15.5
- fix: wrong bound comparisons

## 0.15.4
- feat: improve deserializer
- feat: bump dependencies

## 0.15.3
- feat: add several new convenient fn to `DataType`
- feat: add a `Range::range` fn to get subranges
- feat: add a new `Range::cells` iterator
- feat: impl DoubleEndedIterator when possible
- perf: add some missing `size_hint` impl in iterators
- feat: add a `Range::get` fn (similar to slice's)
- perf: add some `ExactSizeIterator`

## 0.15.2
- feat: consider empty cell as empty str if deserializing to str or String

## 0.15.1
- fix: xls - allow sectors ending after eof (truncate them!)

## 0.15.0
- feat: codepage/encoding_rs for codpage mapping

## 0.14.10
- fix: serde map do not stop at first empty value

## 0.14.9
- fix: do not return map keys for empty cells. Fixes not working `#[serde(default)]`

## 0.14.8
- feat: bump dependencies
- feat: add a `RangeDeserializerBuilder::with_headers` fn to improve serde deserializer

## 0.14.7
- feat: ods, support *text:s* and *text:p*

## 0.14.6
- fix: support MulRk for xls files

## 0.14.5
- fix: properly parse richtext ods files
- refactor: bump dependencies

## 0.14.4
- feat: ods: display sheet names in order.

## 0.14.3
- feat: handle 'covered cells' which are behind merge-cells in ODS

## 0.14.2
- fix: boolean detection and missing repeated cells in ODS
- refactor: bump dependencies

## 0.14.1
- fix: possibility of index out of bound in get_value and eventually in Index<(usize, usize)>

## 0.14.0
- feat: have Range `start`/`end` return None if the range is actually empty
- feat: Have `Range::get_value` return an Option if the index is out of range

## 0.13.1
- refactor: bump dependencies
- feat: make `Range::from_sparse` public

## 0.13.0
- feat: migrate from error-chain to failure
- refactor: simplify Reader trait (enable direct Xlsx read etc ...)
- refactor: always initialize at creation
- feat: more documentation on error
- feat: bump dependencies (calamine, encoding_rs and zip)
- feat: process any Read not only Files
- docs: fix various typos

## 0.12.1
- feat: update dependencies

## 0.12.0
- feat: add serde deserialization

## 0.11.8
- perf: update dependencies, in particular quick-xml 0.9.1

## 0.11.7
- fix: add a bound check when decoding cfb
- refactor: bump dependencies

## 0.11.6
- refactor: bump dependencies
- style: ignore .bk files

## 0.11.5
- refactor: bump dependencies

## 0.11.4
- refactor: update to quick-xml 0.7.3 and encoding_rs 0.6.6

## 0.11.3
- feat: implement Display for DataType and CellTypeError
- feat: add a CellType alias trait

## 0.11.2
- perf: update to quick-xml 0.7.1

## 0.11.1
- refactor: update encoding_rs to 0.6.2
- perf: add benches and avoid clearing a buffer supposed to be reused

## 0.11.0
- feat: add support for formula parsing/decoding
- refactor: make `Range` generic over its content
- fix: convert codepage 21010 as codepage 1200
- fix: support EUC_KR encoding

## 0.10.2
- fix: error while using a singlebyte encoding for xls files (read_dbcs)

## 0.10.1
- fix: error while using a singlebyte encoding for xls files (short_strings)

## 0.10.0
- feat: support defined names for named ranges
- refactor: better internal logics

## 0.9.0
- refactor: rename `Excel` in `Sheets` to accommodate OpenDocuments
- feat: add Index/IndexMut for Range

## 0.8.0
- feat: add basic support for opendocument spreadsheets
- style: apply rustfmt
- feat: force rustfmt on travis checks

## 0.7.0
- fix: extend appveyor paths to be able to use curl
- refactor: update deps
- fix: extract richtext reading from `read_shared_strings` to `read_string`,
and use for inlineStr instead of `read_inline_str`
- style: rustfmt
- fix: enable namespaced xmls when parsing xlsx files

## 0.6.0
- refactor: bump dependencies
- refactor: move from rust-encoding to encoding_rs (faster), loses some decoders ...

## 0.5.1
- refactor: bump to quick-xml 0.6.0 (supposedly faster)

## 0.5.0
- style: rustfmt the code
- feat: xlsx - support 'inlineStr' elements (`<is>` nodes)
- fix: xlsx - support sheetnames prefixed with 'xl/' or '/xl/'
- chore: bump deps (error-chain 0.8.1, quick-xml 0.5.0)

## 0.4.0
- refactor: replace `try!` with `?` operator
- feat: adds a new `worksheet_range_by_index` function.
- feat: adds new `ErrorKind`s
- refactor: simplify `search_error` example by using a `run()` function

## 0.3.3
- refactor: update dependencies (error-chain and byteorder)

## 0.3.2
- refactor: update dependencies

## 0.3.1
- perf: [xls] preload vba only instead of sheets only
- refactor: [vba] consume cfb in constructor and do not store cfb

## 0.3.0
- feat: [all] better `Range` initialization via `Range::from_sparse`
- feat: [all] several new fn in `Range` (`used_cells`, `start`, `end` ...)
- refactor: adds a `range_eq!` macro in tests

## 0.2.1
- fix: [xls] allow directory start to empty sector if version = 3
- fix: [vba] support all project codepage encodings
- feat: [xls] early exit if workbook is password protected
- fix: [xls] better decoding based on codepage
- fix: [xlsb] simplify setting values and early exit when stepping into an invalid BrtRowHdr
- fix: [xlsb] fix record length calculation

## 0.2.0
- fix: [all] allow range to resize when we try to set a value out of bounds
- docs: less `unwrap`s, no unused imports
- refactor: range bounds is not (`start`, `end`) instead of (`position`, `size`)
- feat: add new methods for `Range`: `width`, `height`, `is_empty`

## 0.1.3
- fix: [xls] better management of continue record for rich_extended_strings

## 0.1.2
- fix: [all] return error when trying to set out of bound values in `Range`
- fix: [xls] do a proper encoding when reading cells (force 2 bytes unicode instead of utf8)
- fix: [xls] support continue records
- fix: [all] allow empty rows iterator

## 0.1.1
- fix: remove some development `println!`

## 0.1.0
- first release!
