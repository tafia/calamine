# Changelog

This is the changelog/release notes for the `calamine` crate.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.33.0] - 2025-XX-XX (Draft for next release)

### Added

### Changed

### Fixed


## [0.32.0] - 2025-11-20

### Changed

- Refactored VBA reading functions to be on-demand for better performance.

- Simplified `vba_project()` function return type from `Option<Result<T>>` to
  `Result<Option<T>>` for more idiomatic error handling. This is a breaking
  change.

### Fixed

- Fixed out-of-memory vulnerabilities in XLS file parsing by bounding
  allocations.

- Fixed and extended support for XLSX shared formulas with handling of ranges,
  absolute references, and column/row ranges in XLSX files.

- Fixed XLSX issue with missing shared string sub-elements. Also improved error
  messages for shared string parsing issues.

- Fixed acceptance of XLS `XLUnicodeRichExtendedString` records without reserved
  tags.

- Fixed various edge cases in XLS handling that could lead to parsing errors.


## [0.31.0] - 2025-09-27

### Changed

- Upgraded `quick-xml` to v0.38. This was a significant change in `quick-xml`
  relative to v0.37 and required changes in `calamine` to entity handling. It
  also fixes EOL handling which may lead to regressions in `calamine`
  applications if they expected to see `"\r\n"` in strings instead of the
  correct (for XML and Excel) `"\n"`.

  For most users these will be inconsequential changes but please take note
  before upgrading production code.

- Renamed the `"dates"` feature flag to `"chrono"` since there is now some
  native date handling features without `"chrono"`. The `"chrono"` flag is more
  specific and accurate. The `"dates"` flag is still supported as before for
  backward compatibility.

  This change also made some datatype methods related to date/times available in
  the `"default"` feature set. They were previously hidden unnecessarily behind
  the "dates/"chrono" flag.

### Added

- Added a conversion function to `ExcelDateTime` to convert the inner serial
  Excel datetime to standard year, month, date, hour, minute, second and
  millisecond components. Works for 1900 and 1904 epochs.

### Fixed

- Fixed issue where Excel xlsx shared formula failed if it contained Unicode
  characters. [Issue #553].

  [Issue #553]: https://github.com/tafia/calamine/issues/553

- Fixed issue where Excel XML escapes in strings weren't unescaped. For example
  `"_x000D_" -> "\r"`. [Issue #469].

  [Issue #469]: https://github.com/tafia/calamine/issues/469


## [0.30.1] - 2025-09-06

### Added

- Added `Debug` and `Clone` to `Table` for easier debugging. [PR #547].

  [PR #547]: https://github.com/tafia/calamine/issues/547

### Fixed

- Fixed issue [Issue #548] for xls files where the `SST` record had an incorrect
  number of unique strings.

  [Issue #548]: https://github.com/tafia/calamine/issues/548


## [0.30.0] - 2025-08-07

### Changed

- Unpinned the `zip.rs` dependency from v4.2.0 to allow cargo to choose the
  correct version for the user's rustc version.

  The Rust MSRV was bumped to v1.75.0 (which it should have been for for
  `zip.rs` compatibility in previous releases).

  See the discussion at [Issue #527].

  [Issue #527]: https://github.com/tafia/calamine/issues/527


## [0.29.0] - 2025-07-17

### Added

- Add additional documentation and examples for the `Range`, `Cell`, `XlsxError`
  and `Table` structs, and `Xlsx` Table and Merge methods. Issue #459

### Changed

- Pin zip.rs to v4.2.0.

  The current latest release of `zip.rs`, v4.3.0, requires a MSRV of v1.85.0.
  This release pins `zip.rs` to v4.2.0 to allow users to maintain a MSRV of
  v1.73.0 for at least one more release. It is likely that `calamine` v0.30.0 or
  later will move back to the latest `zip.rs` v4.x and require rustc v1.85.0.

### Fixed

- Fixed issue where XLSX files had Windows style directory separators for
  internal paths instead of the required Unix style separators. Issue #530.
- Fixed several XLS parsing issues which could lead to out-of-memory errors. PR
  #525.
- Fixed numeric underflow in `Xlsx::from_sparse()` and also ensured that the
  associated `Range` of cells would be in row-column order. PR #524.

## [0.28.0] - 2025-06-19

### Changed

- Bump zip to 4.0.

## [0.27.0] - 2025-04-22

### Added

- (xls): add one more `Error` variant related to formatting.

### Changed

- Bump dependencies.
- (xls): Invalid formats parsing.
- Always parse string cell as string.
- Pin zip crate to 2.5.
- (xlsx): check 'closing' tag name with more prefixes.

## [0.26.1] - 2024-10-10

### Fixed

- Sparse cells expect 0 index rows, even when using `header_row`.

## [0.26.0] - 2024-10-08

### Added

- Ability to merge cells from xls and xlsx.
- Options to keep first empty rows for xlsx.
- Support consecutive repeated empty cells for ods.
- New `header_row` config.

### Changed

- Bump MSRV to 1.73.
- Fix broken links in README.
- Enable dates and pictures features in `docs.rs` build.
- Fix broken fuzzer.

## [0.25.0] - 2024-05-25

### Added

- Add `is_error` and `get_error` methods to the `DataType` trait.
- Add deserializer helper functions.
- Support getting merged region.
- `Range::headers` method.
- Expose some `Dimensions` methods.

### Changed

- Use `OnceLock` instead of `once_cell` crate (MSRV: 1.71).

### Fixed

- Use case insensitive comparison when searching for file in xlsx.
- Do not panic when reading cell format with invalid index.

## [0.24.0] - 2024-02-08

### Added

- Introduce a `DataType` trait implemented by both `Data` and `DataRef`.
- `Data` and `DataType` now return `Some(0{.0})` and `Some(1{.0})` rather than
  `None` when `.as_i64` or `.as_f64` is used on a Bool value.
- Detect xlsb/ods password protected files.
- Introduce `is_x` methods for date and time variants.

### Changed

- **BREAKING**: rename `DataType` enum to `Data` and `DataTypeRef` to `DataRef`.
- DateTime(f64) to DateTime(ExcelDateTime).

### Fixed

- Getting tables names from xlsx workbooks without `_rels` files.

## [0.23.1] - 2023-12-19

### Fixed

- `worksheet_formula` not returning all formula.

## [0.23.0] - 2023-12-12

### Added

- New `DataTypeRef` available from `worksheet_range_ref` to reduce memory usage.
- Detect if workbook is password protected.

### Changed

- Add benchmark plot.

### Fixed

- Truncated text in xls.

## [0.22.1] - 2023-10-08

### Added

- Support label cells for xls.

### Changed

- Update GitHub actions.
- Clippy.
- Preallocate several buffers.

### Fixed

- Regression on `Range::get`.
- Spelling of formula error type.

## [0.22.0] - 2023-09-04

### Added

- Add support of sheet type and visibility.
- Implement blank string handling.

### Changed

- Improve `de_opt_f64` example.
- Remove datetime notice from README.
- Clippy.
- Bump MSRV to 1.63 (breaking).
- Set edition to 2021.

## [0.21.2] - 2023-06-25

### Fixed

- Formula with string not displaying properly.

## [0.21.1] - 2023-06-17

### Changed

- Bump MSRV to 1.60.0 due to log dependencies.

### Fixed

- Xls: formula values ignored.
- Xls: formula (string) not displayed properly.

## [0.21.0] - 2023-06-13

### Added

- Support for duration.

### Changed

- Add MSRV.

### Fixed

- (xlsx) support `r` attribute.
- Support `PROJECTCOMPATVERSION` in vba.
- Incorrect date parsing due to excel bug.

## [0.20.0] - 2023-05-29

### Added

- (all) parse format/style information to infer cell as datetime.
- (ods) support number-columns-repeated attribute.

### Changed

- Bump dependencies.
- Multiple clippy refactorings.

## [0.19.2] - 2023-02-09

### Added

- Extract picture data by turning `picture` feature on.

## [0.19.1] - 2022-10-20

### Fixed

- Wrong range len calculation.
- Date precision.

## [0.19.0] - 2022-10-20

### Added

- Always return sheet names in lexicographic order (`BTreeMap`).

### Changed

- Bump dependencies (quick-xml in particular and chrono).
- Remove travis.

### Fixed

- Several decoding issues in xls and xlsb.
- Wrong decimal parsing.

## [0.18.0] - 2021-02-23

### Added

- Improve conversions from raw data to primitives.
- Replace macro matches! by match expression to reduce MSRV.

### Changed

- Fix two typos in README.

### Fixed

- Allow empty value cells in xlsx.
- Obscure xls parsing errors (#195).

## [0.17.0] - 2021-02-03

### Added

- Use `chunks_exact` instead of chunks where possible.
- Detect date/time formatted cells in XLSX.
- Brute force file detection if extension is not known.
- Support xlsx sheet sizes beyond `u32::MAX`.

### Changed

- Add regression tests that fail with miri.
- Ensure doctest functions actually run.
- Run cargo fmt to fix travis tests.

### Fixed

- Make `to_u32`, `read_slice` safe and sound.
- Security issue #199.
- Fix Float reading for XLSB.

## [0.16.2] - 2020-09-26

### Changed

- Add `deserialize_with` example in README.
- Correct MBSC to MBCS in vba.rs (misspelled before).
- Use 2018 edition paths.

### Fixed

- Skip phonetic run.
- Fix XLS float parsing error.
- Add the ability to read formula values from XLSB.
- Support integral date types.

## [0.16.1] - 2019-11-20

### Added

- Make `Metadata.sheets` (and `Reader.sheet_names`) always return names in
  workbook order.

### Changed

- Fix warnings in tests.

## [0.16.0] - 2019-10-11

### Added

- Deprecate failure and impl `std::error::Error` for all errors.
- Add `dates` feature to enrich `DataType` with date conversions functions.

## [0.15.6] - 2019-08-24

### Added

- Update dependencies.

## [0.15.5] - 2019-07-15

### Fixed

- Wrong bound comparisons.

## [0.15.4] - 2019-04-11

### Added

- Improve deserializer.
- Bump dependencies.

## [0.15.3] - 2018-12-14

### Added

- Add several new convenient fn to `DataType`.
- Add a `Range::range` fn to get sub-ranges.
- Add a new `Range::cells` iterator.
- Implement `DoubleEndedIterator` when possible.
- Add a `Range::get` fn (similar to slice's).

### Changed

- Add some missing `size_hint` impl in iterators.
- Add some `ExactSizeIterator`.

## [0.15.2] - 2018-12-14

### Added

- Consider empty cell as empty str if deserializing to str or String.

## [0.15.1] - 2018-12-13

### Fixed

- Xls - allow sectors ending after eof (truncate them!).

## [0.15.0] - 2018-12-13

### Added

- Codepage/`encoding_rs` for codepage mapping.

## [0.14.10] - 2018-11-23

### Fixed

- Serde map do not stop at first empty value.

## [0.14.9] - 2018-11-23

### Fixed

- Do not return map keys for empty cells. Fixes not working `#[serde(default)]`.

## [0.14.8] - 2018-11-23

### Added

- Bump dependencies.
- Add a `RangeDeserializerBuilder::with_headers` fn to improve serde deserializer.

## [0.14.7] - 2018-10-23

### Added

- Ods, support *text:s* and *text:p*.

## [0.14.6] - 2018-09-20

### Fixed

- Support `MulRk` for xls files.

## [0.14.5] - 2018-08-28

### Changed

- Bump dependencies.

### Fixed

- Properly parse richtext ods files.

## [0.14.4] - 2018-08-28

### Added

- Ods: display sheet names in order.

## [0.14.3] - 2018-08-09

### Added

- Handle 'covered cells' which are behind merge-cells in ODS.

## [0.14.2] - 2018-08-03

### Changed

- Bump dependencies.

### Fixed

- Boolean detection and missing repeated cells in ODS.

## [0.14.1] - 2018-05-08

### Fixed

- Possibility of index out of bound in `get_value` and eventually in Index<(usize, usize)>.

## [0.14.0] - 2018-04-27

### Added

- Have Range `start`/`end` return None if the range is actually empty.
- Have `Range::get_value` return an Option if the index is out of range.

## [0.13.1] - 2018-03-23

### Added

- Make `Range::from_sparse` public.

### Changed

- Bump dependencies.

## [0.13.0] - 2018-01-27

### Added

- Migrate from error-chain to failure.
- More documentation on error.
- Bump dependencies (calamine, `encoding_rs` and zip).
- Process any Read not only Files.

### Changed

- Simplify Reader trait (enable direct Xlsx read etc ...).
- Always initialize at creation.
- Fix various typos.

## [0.12.1] - 2017-11-27

### Added

- Update dependencies.

## [0.12.0] - 2017-10-27

### Added

- Add serde deserialization.

## [0.11.8] - 2017-08-22

### Changed

- Update dependencies, in particular quick-xml 0.9.1.

## [0.11.7] - 2017-07-08

### Changed

- Bump dependencies.

### Fixed

- Add a bound check when decoding cfb.

## [0.11.6] - 2017-07-05

### Changed

- Bump dependencies.
- Ignore .bk files.

## [0.11.5] - 2017-05-12

### Changed

- Bump dependencies.

## [0.11.4] - 2017-05-08

### Changed

- Update to quick-xml 0.7.3 and `encoding_rs` 0.6.6.

## [0.11.3] - 2017-05-05

### Added

- Implement `Display` for `DataType` and `CellTypeError`.
- Add a `CellType` alias trait.

## [0.11.2] - 2017-05-04

### Changed

- Update to quick-xml 0.7.1.

## [0.11.1] - 2017-05-03

### Changed

- Update `encoding_rs` to 0.6.2.
- Add benches and avoid clearing a buffer supposed to be reused.

## [0.11.0] - 2017-04-27

### Added

- Add support for formula parsing/decoding.

### Changed

- Make `Range` generic over its content.

### Fixed

- Convert codepage 21010 as codepage 1200.
- Support `EUC_KR` encoding.

## [0.10.2] - 2017-04-18

### Fixed

- Error while using a singlebyte encoding for xls files (`read_dbcs`).

## [0.10.1] - 2017-04-18

### Fixed

- Error while using a singlebyte encoding for xls files (`short_strings`).

## [0.10.0] - 2017-04-14

### Added

- Support defined names for named ranges.

### Changed

- Better internal logic.

## [0.9.0] - 2017-04-12

### Added

- Add Index/IndexMut for Range.

### Changed

- Rename `Excel` in `Sheets` to accommodate `OpenDocument`.

## [0.8.0] - 2017-04-12

### Added

- Add basic support for `OpenDocument` spreadsheets.
- Force rustfmt on travis checks.

### Changed

- Apply rustfmt.

## [0.7.0] - 2017-03-23

### Changed

- Update dependencies.
- Rustfmt.

### Fixed

- Extend appveyor paths to be able to use curl.
- Extract richtext reading from `read_shared_strings` to `read_string`.
- Enable namespaced xmls when parsing xlsx files.

## [0.6.0] - 2017-03-06

### Changed

- Bump dependencies.
- Move from rust-encoding to `encoding_rs` (faster), loses some decoders.

## [0.5.1] - 2017-03-06

### Changed

- Bump to quick-xml 0.6.0 (supposedly faster).

## [0.5.0] - 2017-02-07

### Added

- Xlsx - support 'inlineStr' elements (`<is>` nodes).

### Changed

- Rustfmt the code.
- Bump dependencies (error-chain 0.8.1, quick-xml 0.5.0).

### Fixed

- Xlsx - support sheetnames prefixed with 'xl/' or '/xl/'.

## [0.4.0] - 2017-01-09

### Added

- Adds a new `worksheet_range_by_index` function.
- Adds new `ErrorKind`s.

### Changed

- Replace `try!` with `?` operator.
- Simplify `search_error` example by using a `run()` function.

## [0.3.3] - 2017-01-09

### Changed

- Update dependencies (error-chain and byteorder).

## [0.3.2] - 2016-11-27

### Changed

- Update dependencies.

## [0.3.1] - 2016-11-17

### Changed

- (xls) preload vba only instead of sheets only.
- (vba) consume cfb in constructor and do not store cfb.

## [0.3.0] - 2016-11-16

### Added

- (all) better `Range` initialization via `Range::from_sparse`.
- (all) several new fn in `Range` (`used_cells`, `start`, `end` ...).

### Changed

- Adds a `range_eq!` macro in tests.

## [0.2.1] - 2016-11-15

### Added

- (xls) early exit if workbook is password protected.

### Fixed

- (xls) allow directory start to empty sector if version = 3.
- (vba) support all project codepage encodings.
- (xls) better decoding based on codepage.
- (xlsb) simplify setting values and early exit when stepping into an invalid
  `BrtRowHdr`.
- (xlsb) fix record length calculation.

## [0.2.0] - 2016-11-14

### Added

- Add new methods for `Range`: `width`, `height`, `is_empty`.

### Changed

- Less `unwrap`s, no unused imports.
- Range bounds is not (`start`, `end`) instead of (`position`, `size`).

### Fixed

- (all) allow range to resize when we try to set a value out of bounds.

## [0.1.3] - 2016-11-11

### Fixed

- (xls) better management of continue record for `rich_extended_strings`.

## [0.1.2] - 2016-11-11

### Fixed

- (all) return error when trying to set out of bound values in `Range`.
- (xls) do a proper encoding when reading cells (force 2 bytes unicode instead of utf8).
- (xls) support continue records.
- (all) allow empty rows iterator.

## [0.1.1] - 2016-11-09

### Fixed

- Remove some development `println!`.

## [0.1.0] - 2016-11-09

### Changed

- First release.
