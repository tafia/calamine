//! An example of using a `Row` iterator with a calamine `Range`.

use calamine::{Cell, Data, Range};

fn main() {
    let cells = vec![
        Cell::new((1, 1), Data::Int(1)),
        Cell::new((1, 2), Data::Int(2)),
        Cell::new((3, 1), Data::Int(3)),
    ];

    // Create a Range from the cells.
    let range = Range::from_sparse(cells);

    // Iterate over the rows of the range.
    for (row_num, row) in range.rows().enumerate() {
        for (col_num, data) in row.iter().enumerate() {
            // Print the data in each cell of the row.
            println!("({row_num}, {col_num}): {data}");
        }
    }
}

// Output in relative coordinates:
//
// (0, 0): 1
// (0, 1): 2
// (1, 0):
// (1, 1):
// (2, 0): 3
// (2, 1):
