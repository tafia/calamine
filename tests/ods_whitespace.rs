// SPDX-License-Identifier: MIT

use std::io::{Cursor, Write};

use calamine::{Data, Ods, OdsError, Reader};
use zip::{write::SimpleFileOptions, ZipWriter};

const CONTENT: &str = include_str!("fixtures/ods-whitespace.xml");
const MANIFEST: &str = include_str!("fixtures/ods-whitespace-manifest.xml");
const PICTURE: &[u8] = b"synthetic-picture-bytes";

type TestResult = Result<(), Box<dyn std::error::Error>>;

fn document(separator: &str, picture: bool) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    assert_eq!(CONTENT.matches("{{ROW_SEPARATOR}}").count(), 4);
    let content = CONTENT.replace("{{ROW_SEPARATOR}}", separator);
    let manifest = if picture {
        MANIFEST.replace(
            "</manifest:manifest>",
            concat!(
                "<manifest:file-entry manifest:full-path=\"Pictures/synthetic.png\" ",
                "manifest:media-type=\"image/png\" /></manifest:manifest>"
            ),
        )
    } else {
        MANIFEST.to_owned()
    };
    let mut entries = vec![
        (
            "mimetype",
            b"application/vnd.oasis.opendocument.spreadsheet".as_slice(),
        ),
        ("META-INF/manifest.xml", manifest.as_bytes()),
        ("content.xml", content.as_bytes()),
    ];
    if picture {
        entries.push(("Pictures/synthetic.png", PICTURE));
    }
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    for (name, bytes) in entries {
        writer.start_file(
            name,
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
        )?;
        writer.write_all(bytes)?;
    }
    let bytes = writer.finish()?.into_inner();
    assert_mimetype(&bytes)?;
    Ok(bytes)
}

fn assert_mimetype(bytes: &[u8]) -> TestResult {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes))?;
    let first = archive.by_index(0)?;
    assert_eq!(first.name(), "mimetype");
    assert_eq!(first.compression(), zip::CompressionMethod::Stored);
    Ok(())
}

fn assert_cells(separator: &str) -> TestResult {
    let mut reader = Ods::new(Cursor::new(document(separator, false)?))?;
    assert_eq!(reader.sheet_names(), ["Invoice"]);
    let cells = reader.worksheet_range("Invoice")?;
    assert_eq!(cells.get_size(), (1, 3));
    assert_eq!(
        cells.get((0, 0)),
        Some(&Data::String("  INV-SEC-1  ".into()))
    );
    assert_eq!(cells.get((0, 1)), Some(&Data::Float(125.5)));
    assert_eq!(cells.get((0, 2)), Some(&Data::Bool(true)));
    Ok(())
}

#[test]
fn should_read_compact_ods_cells() -> TestResult {
    assert_cells("")
}

#[test]
fn should_read_indented_ods_cells_without_losing_values() -> TestResult {
    assert_cells("\n          ")
}

#[test]
fn should_accept_space_between_ods_cells() -> TestResult {
    assert_cells(" ")
}

#[test]
fn should_accept_tab_between_ods_cells() -> TestResult {
    assert_cells("\t")
}

#[test]
fn should_accept_carriage_return_between_ods_cells() -> TestResult {
    assert_cells("\r")
}

#[test]
fn should_accept_line_feed_between_ods_cells() -> TestResult {
    assert_cells("\n")
}

#[test]
fn should_accept_mixed_xml_whitespace_between_ods_cells() -> TestResult {
    assert_cells(" \t\r\n")
}

#[test]
fn should_preserve_leading_and_trailing_spaces_inside_ods_cell_text() -> TestResult {
    assert_cells("\n  ")
}

#[test]
fn should_reject_non_whitespace_text_between_ods_cells() -> TestResult {
    let error = Ods::new(Cursor::new(document(" x ", false)?))
        .err()
        .ok_or("unexpected success")?;
    assert!(matches!(
        error,
        OdsError::Mismatch {
            expected: "table-cell",
            ..
        }
    ));
    Ok(())
}

#[test]
fn should_reject_non_xml_unicode_whitespace_between_ods_cells() -> TestResult {
    let error = Ods::new(Cursor::new(document("\u{a0}", false)?))
        .err()
        .ok_or("unexpected success")?;
    assert!(matches!(
        error,
        OdsError::Mismatch {
            expected: "table-cell",
            ..
        }
    ));
    Ok(())
}

#[cfg(feature = "picture")]
#[test]
fn should_preserve_exact_picture_payload_in_indented_ods() -> TestResult {
    let reader = Ods::new(Cursor::new(document("\n  ", true)?))?;
    assert_eq!(
        reader.pictures(),
        Some(vec![("png".to_owned(), PICTURE.to_vec())])
    );
    Ok(())
}
