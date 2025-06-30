//! An example of getting a value in a calamine `Range`, using relative
//! positioning.

use calamine::{Data, Range};

fn main() {
    let mut range = Range::new((1, 1), (5, 5));

    // Set a cell value using the cell absolute position.
    range.set_value((2, 3), Data::Int(123));

    // Get the value using the range relative position.
    assert_eq!(range.get((1, 2)), Some(&Data::Int(123)));
}
