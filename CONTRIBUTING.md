# Contributing to Calamine

Thanks for contributing. These notes describe what reviewers look for. The
pull request template has the per-PR checklist.

## Before you start

- Most pull requests should start with a linked issue, unless the change is very
  small. The issue issue describes the problem and the pull request should
  describe the fix.
- New features and public API changes should always start with an issue so
  the design can be discussed before the code is written.
- For a file parsing bug, attach a file that reproduces the problem.

## Checks

The CI runs the following on every pull request. Run them locally first:

```sh
cargo test --all-features
cargo fmt -- --check
cargo clippy --all-targets --all-features -- -Dwarnings
cargo +1.88 check --all-features
cargo check --target wasm32-unknown-unknown --all-features
typos  # cargo install typos-cli
```

## Code style

- Formatting is `rustfmt` with the default settings.
- Don't reference issue numbers in code or doc comments. Describe the
  behaviour instead. The exception is the end of a test comment, see below.
- Avoid adding panics. Malformed input should return an error.
- Comments should be formatted as sentences with proper punctuation.

## Tests

- Unit tests go in a `#[cfg(test)] mod tests` block in the source file.
- Tests that read a file go in `tests/test.rs`, with the file in `tests/`.
- The test name should describe what is is testing. The format of the test name
  is `fn test_format_what_it_tests()` where format is one of
  all/xlsx/xlsb/xls/ods/xlsm. For example `test_xlsb_truncated_name_record`. The
  test name shouldn't include the issue number.
- Start each test with a comment saying what it tests. An issue reference, if
  any, goes at the end:

  ```rust
  #[test]
  fn test_xlsb_truncated_name_record() {
      // Test that a truncated BrtName record returns an error instead of
      // panicking. See issue #760.
      ...
  }
  ```

## Commits and pull requests

- Title the pull request `area: what the change does`, for example
  `xlsb: return an error on truncated BrtName records`.
- The body should be short and descriptive and should start with either:
  - Fixed issue where ...
  - Added feature ...
- Reference the issue in the description with `Fixes #NNN` or `Closes #NNN`.
- If you make incremental changes due to review comments please rebase the
  changes back into the relevant commit.

