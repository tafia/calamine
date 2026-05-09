## Intro

This directory contains some example of Calamine usage.

## Examples

- `simple_read.rs`: This is a minimal "hello world" example that demonstrates
  how to open a workbook and read some information from it using the `calamine`
  crate.
- `excel_to_csv.rs`:  An example for using the `calamine` crate to convert an
  Excel file to CSV.
- `search_errors.rs`: Recursively searches a directory for Excel files and
  checks them for errors.

### Serialization examples

- `deserialize_range.rs`: This example demonstrates the simplest way to
  deserialize a spreadsheet row. An anonymous tuple is read one row at a time
  using [`Range::deserialize`].
- `deserialize_struct.rs`: Shows how to deserialize rows into named structs with
  header-based column matching.
- `deserialize_fallible.rs`: Handles cells that may be empty or contain
  unexpected types during deserialization.
- `deserialize_flatten.rs`: Uses `#[serde(flatten)]` to capture unknown columns
  in a `HashMap`.
- `deserialize_no_headers.rs`: An example of positional deserialization without
  assuming a header row.
- `deserialize_seed.rs`: Demonstrates stateful deserialization using
  [`RowDeserializer`] and [`DeserializeSeed`] for cases where column names are
  only known at runtime or deserialization depends on runtime context.

