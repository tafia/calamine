use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

use calamine::{open_workbook_auto, DataType, Error, Reader};
use glob::{glob, GlobError, GlobResult};

#[derive(Debug)]
enum FileStatus {
    VbaError(Error),
    RangeError(Error),
    Glob(GlobError),
}

fn main() {
    // Search recursively for all excel files matching argument pattern
    // Output statistics: nb broken references, nb broken cells etc...
    let folder = env::args().nth(1).unwrap_or_else(|| ".".to_string());
    let pattern = format!("{}/**/*.xl*", folder);
    let mut filecount = 0;

    let mut output = pattern
        .chars()
        .take_while(|c| *c != '*')
        .filter_map(|c| match c {
            ':' => None,
            '/' | '\\' | ' ' => Some('_'),
            c => Some(c),
        })
        .collect::<String>();
    output.push_str("_errors.csv");
    let mut output = BufWriter::new(File::create(output).unwrap());

    for f in glob(&pattern).expect(
        "Failed to read excel glob,\
         the first argument must correspond to a directory",
    ) {
        filecount += 1;
        match run(f) {
            Ok((f, missing, cell_errors)) => {
                writeln!(output, "{:?}~{:?}~{}", f, missing, cell_errors)
            }
            Err(e) => writeln!(output, "{:?}", e),
        }
        .unwrap_or_else(|e| println!("{:?}", e))
    }

    println!("Found {} excel files", filecount);
}

fn run(f: GlobResult) -> Result<(PathBuf, Option<usize>, usize), FileStatus> {
    let f = f.map_err(FileStatus::Glob)?;

    println!("Analysing {:?}", f.display());
    let mut xl = open_workbook_auto(&f).unwrap();

    let mut missing = None;
    let mut cell_errors = 0;
    match xl.vba_project() {
        Some(Ok(vba)) => {
            missing = Some(
                vba.get_references()
                    .iter()
                    .filter(|r| r.is_missing())
                    .count(),
            );
        }
        Some(Err(e)) => return Err(FileStatus::VbaError(e)),
        None => (),
    }

    // get owned sheet names
    let sheets = xl.sheet_names().to_owned();

    for s in sheets {
        let range = xl
            .worksheet_range(&s)
            .unwrap()
            .map_err(FileStatus::RangeError)?;
        cell_errors += range
            .rows()
            .flat_map(|r| {
                r.iter().filter(|c| {
                    if let DataType::Error(_) = **c {
                        true
                    } else {
                        false
                    }
                })
            })
            .count();
    }

    Ok((f, missing, cell_errors))
}
