//! An example of creating a new empty calamine `Range`.

use calamine::{Data, Range};

fn main() {
    let range: Range<Data> = Range::empty();

    assert!(range.is_empty());
}
