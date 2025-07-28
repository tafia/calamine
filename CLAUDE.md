# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

### Build
```bash
cargo build
cargo build --release
```

### Test
```bash
# Run all tests
cargo test

# Run tests with specific name pattern
cargo test <pattern>

# Run tests for a specific package feature
cargo test --features dates
cargo test --features picture
cargo test --all-features

# Run a single test
cargo test test_name -- --exact
```

### Lint and Format
```bash
# Format code
cargo fmt

# Check formatting without making changes
cargo fmt -- --check

# Run clippy linter
cargo clippy

# Run clippy with all features
cargo clippy --all-features
```

### Documentation
```bash
# Build and open documentation
cargo doc --open

# Build docs with all features
cargo doc --all-features
```

## Architecture Overview

Calamine is a pure Rust library for reading Excel and OpenDocument spreadsheet files. It supports multiple file formats through a unified API.

### Core Components

1. **Format-Specific Readers**: Each supported format has its own module:
   - `xls.rs`: Legacy Excel format (.xls) - uses CFB (Compound File Binary) format
   - `xlsx/`: Modern Excel format (.xlsx, .xlsm) - uses ZIP with XML
   - `xlsb/`: Excel Binary format (.xlsb) - uses ZIP with binary records  
   - `ods.rs`: OpenDocument Spreadsheet (.ods) - uses ZIP with XML

2. **Unified API**: The `Reader` trait (defined in `lib.rs`) provides a consistent interface across all formats:
   - `worksheet_range()`: Read cell data into a `Range`
   - `worksheet_formula()`: Get formulas
   - `vba_project()`: Access VBA/macro code
   - `defined_names()`: Get named ranges
   - `sheet_names()`: List all sheets

3. **Data Types** (`datatype.rs`):
   - `Data`: Owned cell values (String, Float, Int, Bool, DateTime, Error, Empty)
   - `DataRef`: Borrowed cell values for better performance
   - `Range`: 2D grid of cells with efficient access patterns

4. **Deserialization** (`de.rs`): 
   - Serde integration for deserializing rows into Rust structs
   - `RangeDeserializer` with header row support

5. **Cell Formatting** (`formats.rs`):
   - Number format detection and parsing
   - Cell styles (font, fill, borders, alignment)
   - Built-in and custom format support

6. **VBA Support** (`vba.rs`):
   - Parse and extract VBA project modules
   - Handle references and dependencies

### Key Design Patterns

- **Lazy Loading**: XLSX/XLSB readers support streaming to handle large files efficiently
- **Zero-Copy Where Possible**: Uses `Cow<str>` and `DataRef` to minimize allocations
- **Format Auto-Detection**: `open_workbook_auto()` detects format from file content
- **Modular Parsers**: Each format has its own `cells_reader` for parsing cell data

## NOBIE-Specific Notes

This is a fork of the upstream calamine repository with NOBIE-specific patches:
- Commits marked `[PATCH]` should be converted to upstream PRs
- Reference this repo's `master` branch using exact SHA in dependencies
- For local development in the nobie project, use: `calamine = { path = "YOUR_PATH_TO_CALAMINE" }`