// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Example of reading the charts embedded in an XLSX worksheet, including
//! 3D charts, using the `calamine` crate.

use calamine::{Error, Xlsx};

fn main() -> Result<(), Error> {
    let path = "tests/charts.xlsx";

    let mut workbook: Xlsx<_> = calamine::open_workbook(path)?;

    let charts = workbook.worksheet_charts("Sheet1")?;

    for chart in &charts {
        println!(
            "{}: {:?}",
            chart.name.as_deref().unwrap_or("(unnamed)"),
            chart.chart_type()
        );

        if let Some(title) = &chart.title {
            if let Some(text) = title.text() {
                println!("  title: {text}");
            }
        }

        if let Some(view) = &chart.view_3d {
            println!(
                "  3D view: rotX={:?} rotY={:?} perspective={:?}",
                view.rot_x, view.rot_y, view.perspective
            );
        }

        for series in chart.series() {
            let name = series.name_text().unwrap_or("(unnamed)");
            let values = series.values.as_ref();
            let formula = values.and_then(|v| v.formula.as_deref()).unwrap_or("-");
            let count = values.map_or(0, |v| v.values.len());
            println!("  series '{name}': {formula} ({count} cached points)");
        }
    }

    Ok(())
}
