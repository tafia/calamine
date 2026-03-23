// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

//! Benchmarks for conditional formatting parsing.
//!
//! Run with:
//! ```bash
//! cargo bench --bench conditional_formatting
//! ```

use calamine::{open_workbook, Xlsx};
use criterion::{criterion_group, criterion_main, Criterion};
use std::fs::File;
use std::hint::black_box;
use std::io::{BufReader, Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const SMALL_FILE: &str = concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/tests/conditional_formatting.xlsx"
);

fn build_large_cf_xlsx() -> Vec<u8> {
    let mut buf = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(&mut buf));
    let opts = SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    write!(
        zip,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    write!(
        zip,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
    )
    .unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
    write!(
        zip,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#
    )
    .unwrap();

    zip.start_file("xl/workbook.xml", opts).unwrap();
    write!(
        zip,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#
    )
    .unwrap();

    zip.start_file("xl/styles.xml", opts).unwrap();
    let mut styles = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>
  <dxfs count="100">"#,
    );
    for i in 0u8..100 {
        let r = i.wrapping_mul(37);
        let g = i.wrapping_mul(73);
        let b = i.wrapping_mul(131);
        styles.push_str(&format!(
            r#"<dxf><font><color rgb="FF{r:02X}{g:02X}{b:02X}"/><b/></font></dxf>"#
        ));
    }
    styles.push_str("</dxfs></styleSheet>");
    write!(zip, "{styles}").unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", opts).unwrap();
    let mut sheet = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="n"><v>1</v></c></row></sheetData>"#,
    );

    // 200 CF blocks, each with 5 rules -> 1000 rules total
    for block in 0u32..200 {
        let row_start = block * 10 + 1;
        let row_end = row_start + 9;
        let priority_base = block * 5 + 1;
        let dxf = (block % 100) as u8;
        sheet.push_str(&format!(
            r#"<conditionalFormatting sqref="A{row_start}:Z{row_end}">"#
        ));

        sheet.push_str(&format!(
            r#"<cfRule type="cellIs" dxfId="{dxf}" priority="{}" operator="greaterThan"><formula>50</formula></cfRule>"#,
            priority_base
        ));
        sheet.push_str(&format!(
            r#"<cfRule type="colorScale" priority="{}"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>"#,
            priority_base + 1
        ));
        sheet.push_str(&format!(
            r#"<cfRule type="dataBar" priority="{}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>"#,
            priority_base + 2
        ));
        sheet.push_str(&format!(
            r#"<cfRule type="iconSet" priority="{}"><iconSet iconSet="3TrafficLights"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>"#,
            priority_base + 3
        ));
        sheet.push_str(&format!(
            r#"<cfRule type="containsText" dxfId="{}" priority="{}" operator="containsText" text="test"><formula>NOT(ISERROR(SEARCH("test",A{row_start})))</formula></cfRule>"#,
            (dxf + 1) % 100,
            priority_base + 4
        ));

        sheet.push_str("</conditionalFormatting>");
    }

    sheet.push_str("</worksheet>");
    write!(zip, "{sheet}").unwrap();

    zip.finish().unwrap();
    drop(zip);
    buf
}

fn bench_cf_small(c: &mut Criterion) {
    c.bench_function("cf_small_25_blocks", |b| {
        b.iter(|| {
            let mut wb: Xlsx<BufReader<File>> =
                open_workbook(SMALL_FILE).expect("cannot open file");
            black_box(wb.worksheet_conditional_formatting("Sheet1").unwrap())
        })
    });
}

fn bench_cf_large(c: &mut Criterion) {
    let bytes = build_large_cf_xlsx();

    c.bench_function("cf_large_200_blocks_1000_rules", |b| {
        b.iter(|| {
            let cursor = Cursor::new(&bytes);
            let mut wb: Xlsx<Cursor<&Vec<u8>>> = Xlsx::new(cursor).expect("cannot open file");
            black_box(wb.worksheet_conditional_formatting("Sheet1").unwrap())
        })
    });
}

fn bench_cf_dxf_resolution(c: &mut Criterion) {
    let bytes = build_large_cf_xlsx();

    c.bench_function("cf_large_with_dxf_access", |b| {
        b.iter(|| {
            let cursor = Cursor::new(&bytes);
            let mut wb: Xlsx<Cursor<&Vec<u8>>> = Xlsx::new(cursor).expect("cannot open file");
            let cfs = wb.worksheet_conditional_formatting("Sheet1").unwrap();
            let mut format_count = 0usize;
            for cf in &cfs {
                for rule in &cf.rules {
                    if rule.format.is_some() {
                        format_count += 1;
                    }
                }
            }
            black_box(format_count)
        })
    });
}

criterion_group!(benches, bench_cf_small, bench_cf_large, bench_cf_dxf_resolution);
criterion_main!(benches);
