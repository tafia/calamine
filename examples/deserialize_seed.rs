//! Calamine example to demonstrate stateful deserialization using
//! [`RowDeserializer`] and [`serde::de::DeserializeSeed`].
//!
//! Use this approach when:
//! - Column names are only known at runtime (discovered from the header row),
//!   or,
//! - The deserialized value depends on context that cannot be expressed with
//!   `#[serde(...)]` attributes alone.
//!
//! The sample Excel file `temperature.xlsx` used in this example contains a
//! single sheet named "Sheet1" with the following layout:
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

use std::collections::HashMap;

use calamine::{open_workbook, Reader, RowDeserializer, Xlsx};
use serde::de::{DeserializeSeed, MapAccess, Visitor};

// ---------------------------------------------------------------------------
// Target type. The deserialized value we want to produce.
// ---------------------------------------------------------------------------

#[derive(Debug, PartialEq)]
struct Row {
    label: String,
    /// The raw cell value multiplied by `RowSeed::multiplier`.
    value: f64,
}

// ---------------------------------------------------------------------------
// Seed: Carries the runtime context that influences deserialization.
// ---------------------------------------------------------------------------

/// Carries state that is only known at runtime.
///
/// The `multiplier` is a trivial stand-in for any runtime value — a unit
/// conversion factor, a per-sheet configuration value read from another cell,
/// or a database look-up result, etc.
struct RowSeed {
    multiplier: f64,
}

impl<'de> DeserializeSeed<'de> for RowSeed {
    type Value = Row;

    fn deserialize<D: serde::Deserializer<'de>>(
        self,
        deserializer: D,
    ) -> Result<Self::Value, D::Error> {
        deserializer.deserialize_map(RowVisitor {
            multiplier: self.multiplier,
        })
    }
}

// ---------------------------------------------------------------------------
// Visitor: walks the map entries produced by RowDeserializer.
// ---------------------------------------------------------------------------

struct RowVisitor {
    multiplier: f64,
}

impl<'de> Visitor<'de> for RowVisitor {
    type Value = Row;

    fn expecting(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "A map of spreadsheet cells")
    }

    fn visit_map<A: MapAccess<'de>>(self, mut map: A) -> Result<Self::Value, A::Error> {
        let mut label: Option<String> = None;
        let mut value: Option<f64> = None;

        while let Some(key) = map.next_key::<String>()? {
            match key.as_str() {
                "label" => label = Some(map.next_value()?),
                "value" => value = Some(map.next_value::<f64>()? * self.multiplier),
                _ => {
                    map.next_value::<serde::de::IgnoredAny>()?;
                }
            }
        }

        Ok(Row {
            label: label.ok_or_else(|| serde::de::Error::missing_field("label"))?,
            value: value.ok_or_else(|| serde::de::Error::missing_field("value"))?,
        })
    }
}

// ---------------------------------------------------------------------------
// Main.
// ---------------------------------------------------------------------------

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook.worksheet_range("Sheet1")?;
    let mut rows = range.rows();

    // Read the header row and build a name/index map at runtime.
    let header_row = rows.next().ok_or("missing header row")?;
    let headers: HashMap<String, usize> = header_row
        .iter()
        .enumerate()
        .map(|(idx, cell)| (cell.to_string(), idx))
        .collect();

    // Build the ordered (index, name) pairs that `RowDeserializer` expects.
    let mut header_pairs: Vec<(usize, String)> = headers
        .iter()
        .map(|(name, &idx)| (idx, name.clone()))
        .collect();
    header_pairs.sort_by_key(|&(idx, _)| idx);
    let column_indexes: Vec<usize> = header_pairs.iter().map(|&(idx, _)| idx).collect();
    let header_names: Vec<String> = header_pairs.iter().map(|(_, name)| name.clone()).collect();

    // Some runtime value — here we use a fixed multiplier, but in practice this
    // could come from another cell, a config file, a database, etc.
    let multiplier = 2.0_f64;

    // Deserialize each data row using the seed.
    let mut results: Vec<Row> = Vec::new();
    for (row_idx, row) in rows.enumerate() {
        let de = RowDeserializer::new(
            &column_indexes,
            Some(&header_names),
            row,
            (row_idx as u32 + 1, 0),
        );
        results.push(RowSeed { multiplier }.deserialize(de)?);
    }

    assert_eq!(
        results[0],
        Row {
            label: "celsius".into(),
            value: 22.2222 * multiplier
        }
    );

    assert_eq!(
        results[1],
        Row {
            label: "fahrenheit".into(),
            value: 72.0 * multiplier
        }
    );

    println!(
        "Deserialized {} rows (multiplier = {multiplier}):",
        results.len()
    );

    for row in &results {
        println!("  {:?}", row);
    }

    Ok(())
}
