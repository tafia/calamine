//! An example of iterating over the used cells in a calamine `Range`.

use calamine::{Cell, Data, Range};

fn main() {
    let cells = vec![
        Cell::new((1, 1), Data::Int(1)),
        Cell::new((1, 2), Data::Int(2)),
        Cell::new((3, 1), Data::Int(3)),
    ];

    // Create a Range from the cells.
    let range = Range::from_sparse(cells);

    // Iterate over the used cells in the range.
    for (row, col, data) in range.used_cells() {
        println!("({row}, {col}): {data}");
    }
}

// Output:
//
// (0, 0): 1
// (0, 1): 2
// (2, 0): 3
