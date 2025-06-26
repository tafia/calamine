//! An example of creating a new calamine `Range` for a sparse vector of Cells.

use calamine::{Cell, Data, Range};

fn main() {
    let cells = vec![
        Cell::new((2, 2), Data::Int(1)),
        Cell::new((5, 2), Data::Int(1)),
        Cell::new((9, 2), Data::Int(1)),
    ];

    let range = Range::from_sparse(cells);

    assert_eq!(range.width(), 1);
    assert_eq!(range.height(), 8);
    assert_eq!(range.cells().count(), 8);
    assert_eq!(range.used_cells().count(), 3);
}
