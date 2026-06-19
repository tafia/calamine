// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

// Simple benchmarks to check for performance regressions.
//
// Run with: `cargo bench`
// HTML reports are written to target/criterion/.

use calamine::{open_workbook, Ods, Reader, Xls, Xlsb, Xlsx, XlsxFormulaMetadata};
use criterion::{criterion_group, criterion_main, Criterion};
use std::fs::{read, File};
use std::hint::black_box;
use std::io::{BufReader, Cursor};

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

fn count_shared_formula_expanded(bytes: &[u8]) -> usize {
    let mut excel =
        Xlsx::new(Cursor::new(bytes.to_vec())).expect("cannot open shared formula xlsx");
    let mut cells_reader = excel.worksheet_cells_reader("Sheet1").unwrap();
    let mut count = 0usize;
    let mut formula_len = 0usize;
    while let Some(record) = cells_reader.next_cell_with_formula().unwrap() {
        count += 1;
        formula_len += record.formula.as_deref().map_or(0, str::len);
    }
    count ^ formula_len
}

fn count_shared_formula_metadata(bytes: &[u8]) -> usize {
    let mut excel =
        Xlsx::new(Cursor::new(bytes.to_vec())).expect("cannot open shared formula xlsx");
    let mut cells_reader = excel.worksheet_cells_reader("Sheet1").unwrap();
    let mut count = 0usize;
    let mut shared_tags = 0usize;
    while let Some(record) = cells_reader.next_cell_with_formula_metadata().unwrap() {
        count += 1;
        if matches!(
            record.formula,
            Some(XlsxFormulaMetadata::Shared { .. })
                | Some(XlsxFormulaMetadata::SharedDerived { .. })
        ) {
            shared_tags += 1;
        }
    }
    count ^ shared_tags
}

fn bench_xlsx_shared_formula_cells_reader(c: &mut Criterion) {
    let path = format!(
        "{}/tests/shared_formula_bench.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let bytes = read(path).expect("cannot read shared formula benchmark fixture");

    let mut group = c.benchmark_group("xlsx_shared_formula_cells_reader");
    group.bench_function("expanded/shared_formula_bench", |b| {
        b.iter(|| count_shared_formula_expanded(black_box(&bytes)))
    });
    group.bench_function("metadata/shared_formula_bench", |b| {
        b.iter(|| count_shared_formula_metadata(black_box(&bytes)))
    });
    group.finish();
}

criterion_group!(
    benches,
    bench_xls,
    bench_xlsx,
    bench_xlsb,
    bench_ods,
    bench_xlsx_cells_reader,
    bench_xlsb_cells_reader,
    bench_xlsx_shared_formula_cells_reader
);
criterion_main!(benches);
