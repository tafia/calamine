extern crate office;
extern crate glob;

use std::env;
use glob::{glob, GlobError};
use office::{Excel, DataType, Error};

type MissingReference = Option<usize>;

#[derive(Debug)]
enum FileStatus {
    ExcelError(Error),
    VbaError(Error),
    RangeError(Error),
    Valid(MissingReference, usize),
    Glob(GlobError),
}

fn main() {
    
    // Search recursively for all excel files matching argument pattern
    // Output statistics: nb broken references, nb broken cells etc...
    let pattern = format!("{}/**/*.xl*", env::args().skip(1).next().unwrap_or(".".to_string()));
    let mut filecount = 0;

    for f in glob(&pattern).expect("Failed to read excel glob, the first \
                                    argument must correspond to a directory") {
        filecount += 1;
        let mut status = Vec::new();
        let f = match f {
            Ok(f) => f,
            Err(e) => {
                status.push(FileStatus::Glob(e));
                continue;
            }
        };
        println!("Analysing {:?}", f.display());

        match Excel::open(&f) {
            Ok(mut xl) => {
                let mut missing = None;
                let mut cell_errors = 0;
                if xl.has_vba() {
                    match xl.vba_project() {
                        Ok(ref mut vba) => {
                            let refs = vba.get_references();
                            missing = Some(refs.into_iter()
                                .filter(|r| r.is_missing()).count());
                        },
                        Err(e) => status.push(FileStatus::VbaError(e)),
                    }
                }

                match xl.sheet_names() {
                    Ok(sheets) => {
                        for s in sheets {
                            match xl.worksheet_range(&s) {
                                Ok(range) => {
                                    cell_errors += range.rows()
                                        .map(|r| r.iter()
                                             .map(|c| if let &DataType::Error(_) = c { 1usize } else { 0 })
                                             .sum::<usize>())
                                        .sum::<usize>();
                                },
                                Err(e) => status.push(FileStatus::RangeError(e)),
                            }
                        }
                    },
                    Err(e) => status.push(FileStatus::ExcelError(e)),
                }

                status.push(FileStatus::Valid(missing, cell_errors));
            },
            Err(e) => status.push(FileStatus::ExcelError(e)),
        }
        println!("{:#?}\r\n", status);
    }

    println!("Found {} excel files", filecount);
}
