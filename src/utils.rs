use errors::*;

/// Converts a &[u8] into a &[u32]
pub fn to_u32(s: &[u8]) -> &[u32] {
    assert!(s.len() % 4 == 0);
    unsafe { ::std::slice::from_raw_parts(s as *const [u8] as *const u32, s.len() / 4) }
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

