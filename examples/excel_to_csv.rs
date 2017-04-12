extern crate calamine;

use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;
use calamine::{Sheets, Range, DataType, Result};

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .skip(1)
        .next()
        .expect("Please provide an excel file to convert");
    let sheet = env::args()
        .skip(2)
        .next()
        .expect("Expecting a sheet name as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = Sheets::open(&sce).unwrap();
    let range = xl.worksheet_range(&sheet).unwrap();

    write_range(&mut dest, range).unwrap();
}

fn write_range<W: Write>(dest: &mut W, range: Range) -> Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            let _ = match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => write!(dest, "{}", s),
                DataType::Float(ref f) => write!(dest, "{}", f),
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                let _ = write!(dest, ";")?;
            }
        }
        let _ = write!(dest, "\r\n")?;
    }
    Ok(())
}
