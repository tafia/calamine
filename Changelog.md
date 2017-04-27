> Legend:
  - feat: A new feature
  - fix: A bug fix
  - docs: Documentation only changes
  - style: White-space, formatting, missing semi-colons, etc
  - refactor: A code change that neither fixes a bug nor adds a feature
  - perf: A code change that improves performance
  - test: Adding missing tests
  - chore: Changes to the build process or auxiliary tools/libraries/documentation

## 0.11.0
- feat: add support for formula parsing/decoding
- refactor: make `Range` generic over its content

## 0.10.2
- fix: error while using a singlebyte encoding for xls files (read_dbcs)

## 0.10.1
- fix: error while using a singlebyte encoding for xls files (short_strings)

## 0.10.0
- feat: support defined names for named ranges
- refactor: better internal logics

## 0.9.0
- refactor: rename `Excel` in `Sheets` to accomodate OpenDocuments
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
