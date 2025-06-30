//! An example of getting the (height, width) size of a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    let range: Range<Data> = Range::new((2, 3), (9, 3));

    assert_eq!(range.get_size(), (8, 1));
}
