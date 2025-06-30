//! An example of creating a deserializer fora calamine `Range`.
//!
//! The sample Excel file `temperature.xlsx` contains a single sheet named
//! "Sheet1" with the following data:
//!
//! ```text
//!  ____________________________________________
//! |         ||                |                |
//! |         ||       A        |       B        |
//! |_________||________________|________________|
//! |    1    || label          | value          |
//! |_________||________________|________________|
//! |    2    || celsius        | 22.2222        |
//! |_________||________________|________________|
//! |    3    || fahrenheit     | 72             |
//! |_________||________________|________________|
//! |_          _________________________________|
//!   \ Sheet1 /
//!     ------
//! ```

use calamine::{open_workbook, Error, Reader, Xlsx};

fn main() -> Result<(), Error> {
    let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));

    // Open the workbook.
    let mut workbook: Xlsx<_> = open_workbook(path)?;

    // Get the data range from the first sheet.
    let sheet_range = workbook.worksheet_range("Sheet1")?;

    // Get an iterator over data in the range.
    let mut iter = sheet_range.deserialize()?;

    // Get the next record in the range. The first row is assumed to be the
    // header.
    if let Some(result) = iter.next() {
        let (label, value): (String, f64) = result?;

        assert_eq!(label, "celsius");
        assert_eq!(value, 22.2222);

        Ok(())
    } else {
        Err(From::from("Expected at least one record but got none"))
    }
}
