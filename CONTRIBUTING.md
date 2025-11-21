# Contributing to Calamine

Here are some guidelines for contributing to Calamine.

If you are unsure about anything please open a GitHub issue to ask questions or
start a discussion around the feature/change you intend to add.


## Code Style

Use `rustfmt` on the code:

```bash
cargo fmt
```

## Lint with Clippy

Run Clippy on the code:

```bash
cargo clippy --lib --examples --tests
```

## Testing

### Unit tests

Unit tests are generally added in the same file as the code being tested:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    // Comment explaining what is being tested.
    #[test]
    fn test_parse_number() {
        let data = &[0x01, 0x02, 0x03, 0x04];
        let result = parse_number(data);
        assert_eq!(result, 67305985);
    }
}
```

### Integration tests

Integration tests should be added in the `tests/test.rs` file. These usually
test against a sample `xlsx`, `xls`, `xlsb`, or `ods` file:

```rust
// Comment explaining what is being tested and the `#123` GitHub
// issue number if applicable.
#[test]
fn test_open_sample_file() {
    let mut workbook: Xlsx<_> = wb("issues.xlsx");
    // Test specific functionality.
}
```

### Running Tests

```bash
# Run all tests.
cargo test

# Run integration tests only.
cargo test --test '*'

# Run lib tests only.
cargo test --lib
```

### Benchmarking

The `benches` benchmark test can be used to check for any major performance
regressions:

```bash
cargo +nightly bench
```

## Documentation

### API Documentation Guidelines

- Explain the general use case first.
- Explain any edge cases after that.
- Add a `# Parameters` section for functions with parameters.
- Add a `# Errors` section for functions with `Result<T, E>`.
- Add a `# Panics` section if required (although it is better to avoid
  panics in the code).
- Add an `# Examples` section with one or more examples.
- Link to any related functions or types using the `[function_name]` syntax.
- Check spelling with the [`typos`](https://github.com/crate-ci/typos) command.

Here is an example:

```rust
/// Get a worksheet table by name.
///
/// This method retrieves a [`Table`] from the workbook by name. The
/// table will contain an owned copy of the worksheet data within the table
/// range.
///
/// # Parameters
///
/// - `table_name`: The name of the table to retrieve.
///
/// # Errors
///
/// - [`XlsxError::TableNotFound`].
/// - [`XlsxError::NotAWorksheet`].
///
/// # Panics
///
/// Panics if tables have not been loaded via [`Xlsx::load_tables()`].
///
/// # Examples
///
/// An example of getting an Excel worksheet table by its name. The file in
/// this example contains 4 tables spread across 3 worksheets. This example
/// gets an owned copy of the worksheet data in the table area.
///
/// ```
/// use calamine::{open_workbook, Data, Error, Xlsx};
///
/// fn main() -> Result<(), Error> {
///     let path = "tests/table-multiple.xlsx";
///
///     // Open the workbook.
///     let mut workbook: Xlsx<_> = open_workbook(path)?;
///
///     ...
///
///     Ok(())
/// }
/// ```
///
pub fn table_by_name(&mut self, table_name: &str) -> Result<Table<Data>, XlsxError> {
    // ...
```

