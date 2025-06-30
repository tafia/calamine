//! An example of getting the header row of a calamine `Range`.

use calamine::{Data, Range};

fn main() {
    // Create a range with some values.
    let mut range = Range::new((0, 0), (5, 2));
    range.set_value((0, 0), Data::String(String::from("a")));
    range.set_value((0, 1), Data::Int(1));
    range.set_value((0, 2), Data::Bool(true));

    // Get the headers of the range.
    let headers = range.headers();

    assert_eq!(
        headers,
        Some(vec![
            String::from("a"),
            String::from("1"),
            String::from("true")
        ])
    );
}
