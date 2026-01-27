// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Benchmarks for style parsing and extraction features.
//!
//! Uses styles_1M.xlsx (1M styled cells) for realistic performance measurement.
//!
//! ## Setup
//!
//! Generate the test file first:
//! ```bash
//! cargo run --example generate_styles_1M
//! ```
//!
//! ## Run benchmarks
//!
//! ```bash
//! cargo bench --bench style
//! ```
//!
//! ## Profiling (identify bottlenecks)
//!
//! Install samply (cross-platform, works on macOS and Linux):
//! ```bash
//! cargo install samply
//! ```
//!
//! Profile a specific benchmark:
//! ```bash
//! samply record cargo bench --bench style -- "style/worksheet_style" --profile-time 5
//! ```
//!
//! This opens Firefox Profiler with an interactive flamegraph showing where time is spent.

use calamine::{open_workbook, Reader, Xlsx};
use criterion::{criterion_group, criterion_main, Criterion, SamplingMode};
use std::fs::File;
use std::hint::black_box;
use std::io::BufReader;
use std::time::Duration;

const LARGE_FILE: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/styles_1M.xlsx");

fn configure(c: &mut Criterion) -> criterion::BenchmarkGroup<'_, criterion::measurement::WallTime> {
    let mut group = c.benchmark_group("style");
    group.sample_size(10);
    group.warm_up_time(Duration::from_millis(100));
    group.measurement_time(Duration::from_secs(15)); // Accommodate slowest benchmark (~1.2s × 10)
    group.sampling_mode(SamplingMode::Flat); // 1 iteration per sample for slow benchmarks
    group
}

fn bench_style_parsing(c: &mut Criterion) {
    if !std::path::Path::new(LARGE_FILE).exists() {
        eprintln!(
            "ERROR: styles_1M.xlsx not found.\n\
             Generate with: cargo run --example generate_styles_1M"
        );
        return;
    }

    let mut group = configure(c);

    // Core style parsing
    group.bench_function("worksheet_style", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            black_box(excel.worksheet_style("Sheet 1").unwrap())
        })
    });

    // Layout parsing (column widths, row heights)
    group.bench_function("worksheet_layout", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            black_box(excel.worksheet_layout("Sheet 1").unwrap())
        })
    });

    // Range parsing (cell values only, no styles)
    group.bench_function("worksheet_range", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            black_box(excel.worksheet_range("Sheet 1").unwrap())
        })
    });

    // Combined range + style (common real-world usage)
    group.bench_function("range_and_style", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            let range = excel.worksheet_range("Sheet 1").unwrap();
            let style = excel.worksheet_style("Sheet 1").unwrap();
            black_box((range.cells().count(), style.cells().count()))
        })
    });

    // Cell-by-cell iteration via cells_reader
    group.bench_function("cells_reader", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            let mut reader = excel.worksheet_cells_reader("Sheet 1").unwrap();
            let mut count = 0usize;
            while let Ok(Some(_)) = reader.next_cell() {
                count += 1;
            }
            black_box(count)
        })
    });

    // Iterate and access ALL style properties
    group.bench_function("iterate_all_properties", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            let styles = excel.worksheet_style("Sheet 1").unwrap();
            let mut count = 0usize;
            for (_, _, style) in styles.cells() {
                if style.get_font().is_some() {
                    count += 1;
                }
                if style.get_fill().is_some() {
                    count += 1;
                }
                if style.borders.is_some() {
                    count += 1;
                }
                if style.get_alignment().is_some() {
                    count += 1;
                }
                if style.get_number_format().is_some() {
                    count += 1;
                }
            }
            black_box(count)
        })
    });

    // RLE-compressed style parsing (deduplicates styles, compresses consecutive runs)
    group.bench_function("worksheet_style_rle", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            let styles = excel.worksheet_style_rle("Sheet 1").unwrap();
            black_box((styles.unique_style_count(), styles.run_count()))
        })
    });

    // RLE iterate all cells
    group.bench_function("rle_iterate_all", |b| {
        b.iter(|| {
            let mut excel: Xlsx<BufReader<File>> =
                open_workbook(LARGE_FILE).expect("cannot open file");
            let styles = excel.worksheet_style_rle("Sheet 1").unwrap();
            let mut count = 0usize;
            for (_, _, style) in styles.cells() {
                if style.get_font().is_some() {
                    count += 1;
                }
            }
            black_box(count)
        })
    });

    group.finish();
}

criterion_group!(benches, bench_style_parsing);
criterion_main!(benches);
