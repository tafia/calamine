//! An example of cell indexing for a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a range with a value.
    let mut range = Range::new((1, 1), (3, 3));
    range.set_value((2, 2), Data::Int(123));

    // Get the value via cell indexing.
    assert_eq!(range[(1, 1)], Data::Int(123));
}
