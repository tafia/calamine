//! An example of creating a new calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a 8x1 Range.
    let range: Range<Data> = Range::new((2, 2), (9, 2));

    assert_eq!(range.width(), 1);
    assert_eq!(range.height(), 8);
    assert_eq!(range.cells().count(), 8);
    assert_eq!(range.used_cells().count(), 0);
}
