//! An example of getting a sub range of a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a range with some values.
    let mut a = Range::new((1, 1), (3, 3));
    a.set_value((1, 1), Data::Bool(true));
    a.set_value((2, 2), Data::Bool(true));
    a.set_value((3, 3), Data::Bool(true));

    // Get a sub range of the main range.
    let b = a.range((1, 1), (2, 2));
    assert_eq!(b.get_value((1, 1)), Some(&Data::Bool(true)));
    assert_eq!(b.get_value((2, 2)), Some(&Data::Bool(true)));

    // Get a larger range with default values.
    let c = a.range((0, 0), (5, 5));
    assert_eq!(c.get_value((0, 0)), Some(&Data::Empty));
    assert_eq!(c.get_value((3, 3)), Some(&Data::Bool(true)));
    assert_eq!(c.get_value((5, 5)), Some(&Data::Empty));
}
