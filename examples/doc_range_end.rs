//! An example of getting the end position of a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    let range: Range<Data> = Range::new((2, 3), (9, 3));

    assert_eq!(range.end(), Some((9, 3)));
}
