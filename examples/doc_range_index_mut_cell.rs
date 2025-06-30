//! An example of mutable cell indexing for a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a new empty range.
    let mut range = Range::new((1, 1), (3, 3));

    // Set a value in the range using cell indexing.
    range[(1, 1)] = Data::Int(123);

    // Test the value was set correctly.
    assert_eq!(range.get((1, 1)), Some(&Data::Int(123)));
}
