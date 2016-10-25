use utils;
use std::io::Read;
use std::string::String;
use errors::*;

struct BinarySheet {
    name: String,
}

impl BinarySheet {

    /// Creates a new `BinarySheet`
    pub fn new<R: Read>(mut f: R, len: usize) -> Result<BinarySheet> {
        debug!("new binary sheet");

        let mut data = Vec::with_capacity(len);
        try!(f.read_to_end(&mut data));
        let data = &mut &*data;

        try!(utils::check_record(0x0081, data)); // BrtBeginSheet
        let name = try!(read_ws_prop(data));

        Ok(BinarySheet {
            name: name,
        })
    }

}

/// Parses BrtWsProp record and returns sheet name
/// MS-XLSB 2.4.820
fn read_ws_prop(data: &mut &[u8]) -> Result<String> {
    *data = &data[19 * 8..]; // discard first 19 bytes
    let name = try!(utils::read_variable_record(data, 2));
    let name = try!(String::from_utf8(name.to_vec()));
    Ok(name)
}
