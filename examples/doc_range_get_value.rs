//! An example of getting a value in a calamine `Range`, using absolute
//! positioning.

use calamine::{Data, Range};

fn main() {
    let range = Range::new((1, 1), (5, 5));

    // Get the value for a cell in the range.
    assert_eq!(range.get_value((2, 2)), Some(&Data::Empty));

    // Get the value for a cell outside the range.
    assert_eq!(range.get_value((0, 0)), None);
}
