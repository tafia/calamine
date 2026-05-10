// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

// Simple benchmarks to check for performance regressions.
//
// Run with: `cargo bench`
// HTML reports are written to target/criterion/.

use calamine::{open_workbook, Ods, Reader, Xls, Xlsb, Xlsx};
use criterion::{criterion_group, criterion_main, Criterion};
use std::fs::File;
use std::io::BufReader;

fn count<R: Reader<BufReader<File>>>(path: &str) -> usize {
    let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
    let mut excel: R = open_workbook(&path).expect("cannot open excel file");

    let sheets = excel.sheet_names();
    let mut count = 0;
    for s in sheets {
        count += excel
            .worksheet_range(&s)
            .unwrap()
            .rows()
            .flat_map(|r| r.iter())
            .count();
    }
    count
}

fn count_cells_reader_xlsx(path: &str) -> usize {
    let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
    let mut excel: Xlsx<_> = open_workbook(&path).expect("cannot open excel file");

    let sheets = excel.sheet_names();
    let mut count = 0;
    for s in sheets {
        let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
        while cells_reader.next_cell().unwrap().is_some() {
            count += 1;
        }
    }
    count
}

fn count_cells_reader_xlsb(path: &str) -> usize {
    let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), path);
    let mut excel: Xlsb<_> = open_workbook(&path).expect("cannot open excel file");

    let sheets = excel.sheet_names();
    let mut count = 0;
    for s in sheets {
        let mut cells_reader = excel.worksheet_cells_reader(&s).unwrap();
        while cells_reader.next_cell().unwrap().is_some() {
            count += 1;
        }
    }
    count
}

fn bench_xls(c: &mut Criterion) {
    c.bench_function("xls", |b| b.iter(|| count::<Xls<_>>("tests/issues.xls")));
}

fn bench_xlsx(c: &mut Criterion) {
    c.bench_function("xlsx", |b| b.iter(|| count::<Xlsx<_>>("tests/issues.xlsx")));
}

fn bench_xlsb(c: &mut Criterion) {
    c.bench_function("xlsb", |b| b.iter(|| count::<Xlsb<_>>("tests/issues.xlsb")));
}

fn bench_ods(c: &mut Criterion) {
    c.bench_function("ods", |b| b.iter(|| count::<Ods<_>>("tests/issues.ods")));
}

fn bench_xlsx_cells_reader(c: &mut Criterion) {
    c.bench_function("xlsx_cells_reader", |b| {
        b.iter(|| count_cells_reader_xlsx("tests/issues.xlsx"))
    });
}

fn bench_xlsb_cells_reader(c: &mut Criterion) {
    c.bench_function("xlsb_cells_reader", |b| {
        b.iter(|| count_cells_reader_xlsb("tests/issues.xlsb"))
    });
}

criterion_group!(
    benches,
    bench_xls,
    bench_xlsx,
    bench_xlsb,
    bench_ods,
    bench_xlsx_cells_reader,
    bench_xlsb_cells_reader
);
criterion_main!(benches);
