use byteorder::{LittleEndian, ReadBytesExt};
use log::LogLevel;
use errors::*;

/// Reads a variable length record
/// 
/// `mult` is a multiplier of the length (e.g 2 when parsing XLWideString)
pub fn read_variable_record<'a>(r: &mut &'a[u8], mult: usize) -> Result<&'a[u8]> {
    let len = try!(r.read_u32::<LittleEndian>()) as usize * mult;
    let (read, next) = r.split_at(len);
    *r = next;
    Ok(read)
}

/// Check that next record matches `id` and returns a variable length record
pub fn check_variable_record<'a>(id: u16, r: &mut &'a[u8]) -> Result<&'a[u8]> {
    try!(check_record(id, r));
    let record = try!(read_variable_record(r, 1));
    if log_enabled!(LogLevel::Warn) && record.len() > 100_000 {
        warn!("record id {} as a suspicious huge length of {} (hex: {:x})", 
              id, record.len(), record.len() as u32);
    }
    Ok(record)
}

/// Check that next record matches `id`
pub fn check_record(id: u16, r: &mut &[u8]) -> Result<()> {
    debug!("check record {:x}", id);
    let record_id = try!(r.read_u16::<LittleEndian>());
    if record_id != id {
        Err(format!("invalid record id, found {:x}, expecting {:x}", record_id, id).into())
    } else {
        Ok(())
    }
}

/// An iterator over a byte slice that iterates u32 elements
pub struct U32Iter<'a> {
    chunks: ::std::slice::Chunks<'a, u8>,
}

/// Converts a slice into a u32 iterator
pub fn to_u32(s: &[u8]) -> U32Iter {
    U32Iter { chunks: s.chunks(4), }
}

impl<'a> Iterator for U32Iter<'a> {
    type Item=u32;
    fn next(&mut self) -> Option<u32> {
        self.chunks.next().map(|c| unsafe { ::std::ptr::read(c as *const [u8] as *const u32) }) 
    }
}

/// Converts the first 4 bytes into u32
pub fn start_u32(s: &[u8]) -> u32 {
    unsafe { ::std::ptr::read(&s[..4] as *const [u8] as *const u32) }
}

/// converts a text representation (e.g. "A6:G67") of a dimension into integers
/// - top left (row, column), 
/// - size (width, height)
pub fn get_dimension(dimension: &[u8]) -> Result<((u32, u32), (u32, u32))> {
    let parts: Vec<_> = try!(dimension.split(|c| *c == b':')
        .map(|s| get_row_column(s))
        .collect::<Result<Vec<_>>>());

    match parts.len() {
        0 => Err("dimension cannot be empty".into()),
        1 => Ok((parts[0], (1, 1))),
        2 => Ok((parts[0], (parts[1].0 - parts[0].0 + 1, parts[1].1 - parts[0].1 + 1))),
        len => Err(format!("range dimension has 0 or 1 ':', got {}", len).into()),
    }
}

/// converts a text range name into its position (row, column)
pub fn get_row_column(range: &[u8]) -> Result<(u32, u32)> {
    let (mut row, mut col) = (0, 0);
    let mut pow = 1;
    let mut readrow = true;
    for c in range.iter().rev() {
        match *c {
            c @ b'0'...b'9' => {
                if readrow {
                    row += ((c - b'0') as u32) * pow;
                    pow *= 10;
                } else {
                    return Err(format!("Numeric character are only allowed \
                        at the end of the range: {:x}", c).into());
                }
            }
            c @ b'A'...b'Z' => {
                if readrow { 
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'A') as u32 + 1) * pow;
                pow *= 26;
            },
            c @ b'a'...b'z' => {
                if readrow { 
                    pow = 1;
                    readrow = false;
                }
                col += ((c - b'a') as u32 + 1) * pow;
                pow *= 26;
            },
            _ => return Err(format!("Expecting alphanumeric character, got {:x}", c).into()),
        }
    }
    Ok((row, col))
}

