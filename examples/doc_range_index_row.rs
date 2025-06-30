//! An example of row indexing for a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a range with a value.
    let mut range = Range::new((1, 1), (3, 3));
    range.set_value((2, 2), Data::Int(123));

    // Get the second row via indexing.
    assert_eq!(range[1], [Data::Empty, Data::Int(123), Data::Empty]);
}
