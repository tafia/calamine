//! An example of setting a value in a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    let mut range = Range::new((0, 0), (5, 2));

    // The initial range is empty.
    assert_eq!(range.get_value((2, 1)), Some(&Data::Empty));

    // Set a value at a specific position.
    range.set_value((2, 1), Data::Float(1.0));

    // The value at the specified position should now be set.
    assert_eq!(range.get_value((2, 1)), Some(&Data::Float(1.0)));
}
