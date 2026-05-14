// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

// Simple benchmarks to check for performance regressions.
//
// Run with: `cargo bench`
// HTML reports are written to target/criterion/.

use calamine::{open_workbook, Ods, Reader, Xls, Xlsb, Xlsx, XlsxFormulaMetadata};
use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion};
use std::fs::File;
use std::hint::black_box;
use std::io::{BufReader, Cursor, Write};
use zip::write::SimpleFileOptions;

fn shared_formula_xlsx(rows: u32) -> Vec<u8> {
    let mut sheet_data = String::with_capacity(rows as usize * 96);
    for row in 1..=rows {
        let formula = if row == 1 {
            format!(r#"<f t="shared" si="0" ref="B1:B{rows}">A1*2</f>"#)
        } else {
            r#"<f t="shared" si="0"/>"#.to_string()
        };
        sheet_data.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c><c r="B{row}">{formula}<v>{}</v></c></row>"#,
            row * 2
        ));
    }
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B{rows}"/>
  <sheetData>{sheet_data}</sheetData>
</worksheet>"#
    );

    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#).unwrap();
    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#).unwrap();
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#).unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#).unwrap();
    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#).unwrap();
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();
    zip.finish().unwrap().into_inner()
}

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
            Some(XlsxFormulaMetadata::SharedAnchor { .. })
                | Some(XlsxFormulaMetadata::SharedDerived { .. })
        ) {
            shared_tags += 1;
        }
    }
    count ^ shared_tags
}

fn bench_xlsx_shared_formula_cells_reader(c: &mut Criterion) {
    let mut group = c.benchmark_group("xlsx_shared_formula_cells_reader");
    for rows in [1_000u32, 10_000, 100_000] {
        let bytes = shared_formula_xlsx(rows);
        group.bench_with_input(BenchmarkId::new("expanded", rows), &bytes, |b, bytes| {
            b.iter(|| count_shared_formula_expanded(black_box(bytes)))
        });
        group.bench_with_input(BenchmarkId::new("metadata", rows), &bytes, |b, bytes| {
            b.iter(|| count_shared_formula_metadata(black_box(bytes)))
        });
    }
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
