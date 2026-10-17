# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

calamine is a pure-Rust **read-only** library for Excel/OpenDocument spreadsheet files (`xls`, `xlsx`, `xlsm`, `xlsb`, `xla`, `xlam`, `ods`) including VBA project extraction. It never writes spreadsheet files; formatting/styles are deliberately unsupported — the focus is cell values, formulas, and VBA.

## Commands

```bash
cargo build
cargo test --all-features          # CI runs tests with all features
cargo test --all-features --test test.rs        # only the integration test suite
cargo test --all-features --test test.rs dates   # run tests matching a name substring
cargo fmt -- --check
cargo clippy --all-targets --all-features -- -Dwarnings   # CI treats warnings as errors
cargo check --target wasm32-unknown-unknown [--all-features]  # wasm is a supported target
cargo bench                        # criterion benchmarks in benches/basic.rs
```

- MSRV is 1.88 (see `rust-version` in Cargo.toml); CI tests 1.88, stable, beta, and nightly. Don't use features newer than the MSRV.
- Tests are almost entirely in the single integration file `tests/test.rs`, driven by binary fixture files in `tests/` (e.g. `date.xlsx`, `any_sheets.ods`). When adding a regression test, add the fixture file alongside it.
- `.typos.toml` configures the typos spell-checker (CI runs it); ignore-list entries are used to keep fixture filenames/strings passing.

## Architecture

The core abstraction is the `Reader<RS>` trait ([src/lib.rs](src/lib.rs)), implemented once per file format. All formats produce the same output type `Range<Data>` — a dense 2D grid of `Data` cells (the cell value enum: `Int`, `Float`, `String`, `Bool`, `DateTime`, `DateTimeIso`, `DurationIso`, `Error`, `Empty`).

- `src/xlsx/` — zip + quick-xml parser for OOXML. Supports **lazy loading**: `worksheet_range` re-parses the sheet XML on each call via `cells_reader.rs` rather than caching. xlsb is the binary counterpart (`src/xlsb/`), same lazy design.
- `src/xls.rs` — the largest module; BIFF binary format (xls/xla/xlam), parsed from a CFB (compound file) container. `XlsOptions` can relax OOM guards (see the `OOM_*.xls` fixtures).
- `src/ods.rs` — OpenDocument.
- `src/cfb.rs` — generic compound-file binary container used by xls and vba.
- `src/vba.rs` — public module for extracting VBA modules/references from the CFB stream.
- `src/auto.rs` — `open_workbook_auto`: sniffs file magic bytes and dispatches to the right format implementation (`Sheets` enum).
- `src/datatype.rs` — `Data`/`DataType` and `ExcelDateTime`; date conversion depends on workbook epoch (1900 vs 1904) and number formats.
- `src/formats.rs` — Excel number-format string parsing to detect date/time/duration cells.
- `src/de.rs` — serde integration: `RangeDeserializerBuilder` turns a `Range` into records; helper fns like `deserialize_as_f64_or_none`.
- `src/errors.rs` — top-level `Error` enum wrapping the per-format error types (`XlsxError`, `XlsError`, `XlsbError`, `OdsError`, defined in their respective modules).

Two read APIs per format:
- `Reader::worksheet_range` — allocates an owned `Range<Data>`.
- `ReaderRef::worksheet_range_ref` — zero-copy borrowing variant returning `Range<DataRef>` (only implemented where the format allows, i.e. xlsx/xlsb).

Some format-specific functionality lives outside the trait as inherent methods on the format type (e.g. `Xlsx::load_merged_regions`, `Xlsx::table_by_name`, `Xlsx::pivot_tables`, `Xlsx::hyperlinks_by_sheet_name`) — these often need an explicit `load_*()` call before querying because of lazy parsing.

Note the duplication pattern between `worksheet_range*` and `worksheet_formula`: formulas are fetched as a separate `Range<String>`, not merged into cell values.
