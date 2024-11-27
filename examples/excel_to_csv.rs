use calamine::{open_workbook_auto, Data, Range, Reader};
use std::env;
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

fn main() {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let file = env::args()
        .nth(1)
        .expect("Please provide an excel file to convert");
    let sheet = env::args()
        .nth(2)
        .expect("Expecting a sheet name as second argument");

    let sce = PathBuf::from(file);
    match sce.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = sce.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&sce).unwrap();
    let range = xl.worksheet_range(&sheet).unwrap();

    write_range(&mut dest, &range).unwrap();
}

fn write_range<W: Write>(dest: &mut W, range: &Range<Data>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                Data::Empty => Ok(()),
                Data::String(ref s) | Data::DateTimeIso(ref s) | Data::DurationIso(ref s) => {
                    write!(dest, "{}", s)
                }
                Data::RichText(ref s) => write!(dest, "{}", s.text()),
                Data::Float(ref f) => write!(dest, "{}", f),
                Data::DateTime(ref d) => write!(dest, "{}", d.as_f64()),
                Data::Int(ref i) => write!(dest, "{}", i),
                Data::Error(ref e) => write!(dest, "{:?}", e),
                Data::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ";")?;
            }
        }
        write!(dest, "\r\n")?;
    }
    Ok(())
}
