//! A normalized set of 0-based indices used to project worksheet reads onto a
//! subset of columns and/or rows.

use std::ops::{Range, RangeFrom, RangeFull, RangeInclusive, RangeTo, RangeToInclusive};

/// A normalized set of 0-based indices for projecting columns or rows on a
/// worksheet read.
///
/// Construct via `From`/`Into` from a range, a single index, a list, or a list
/// of ranges. An empty set (e.g. from `..` or `IndexSet::default()`) selects
/// **everything** (no projection).
///
/// ```
/// use calamine::IndexSet;
///
/// let _: IndexSet = (0..5).into();        // a contiguous range
/// let _: IndexSet = (5..).into();         // open-ended: 5 to the last index
/// let _: IndexSet = [1, 3, 5].into();     // a discrete list
/// let _: IndexSet = [0..3, 8..10].into(); // disjoint ranges
/// let _: IndexSet = (..).into();          // all (no projection)
/// ```
///
/// Overlapping or duplicate inputs merge silently; the order of inputs does not
/// matter.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct IndexSet {
    /// Sorted, merged, non-overlapping half-open `[start, end)` intervals.
    /// Empty == "all" (no projection). An open-ended upper bound is `u32::MAX`.
    ///
    /// Indices are assumed to be well below `u32::MAX`; `u32::MAX` is reserved
    /// as the exclusive open-ended upper-bound sentinel, so it is not itself a
    /// representable/selectable index.
    spans: Vec<(u32, u32)>,
}

impl IndexSet {
    /// Build a normalized `IndexSet` from raw half-open `[start, end)` spans.
    /// Spans with `start >= end` contribute nothing. Overlapping and adjacent
    /// spans are merged.
    fn from_spans(mut raw: Vec<(u32, u32)>) -> Self {
        raw.retain(|&(s, e)| s < e);
        raw.sort_unstable_by_key(|&(s, _)| s);
        let mut spans: Vec<(u32, u32)> = Vec::with_capacity(raw.len());
        for (s, e) in raw {
            match spans.last_mut() {
                // Overlapping or adjacent with the previous span: extend it.
                Some(last) if s <= last.1 => last.1 = last.1.max(e),
                _ => spans.push((s, e)),
            }
        }
        IndexSet { spans }
    }

    /// True if this set selects everything (no projection).
    pub fn is_all(&self) -> bool {
        self.spans.is_empty()
    }

    /// True if index `i` is selected. An empty set selects everything.
    pub(crate) fn keep(&self, i: u32) -> bool {
        if self.spans.is_empty() {
            return true;
        }
        // Find the last span whose start is <= i, then bounds-check its end.
        match self.spans.binary_search_by(|&(s, _)| s.cmp(&i)) {
            Ok(_) => true,   // i is exactly a span start
            Err(0) => false, // before the first span
            Err(idx) => i < self.spans[idx - 1].1,
        }
    }

    /// Number of selected indices that fall within `0..bound`. Used only for a
    /// capacity estimate, so saturating arithmetic is fine. Returns `bound` for
    /// the "all" set.
    pub(crate) fn selected_count(&self, bound: u32) -> u64 {
        if self.spans.is_empty() {
            return bound as u64;
        }
        // Relies on `spans` being normalized (non-overlapping), so summing
        // clamped span widths counts each selected index exactly once.
        self.spans
            .iter()
            .map(|&(s, e)| {
                let s = s.min(bound);
                let e = e.min(bound);
                (e - s) as u64
            })
            .sum()
    }
}

impl From<&[Range<u32>]> for IndexSet {
    fn from(ranges: &[Range<u32>]) -> Self {
        IndexSet::from_spans(ranges.iter().map(|r| (r.start, r.end)).collect())
    }
}

impl From<&[u32]> for IndexSet {
    fn from(list: &[u32]) -> Self {
        IndexSet::from_spans(list.iter().map(|&i| (i, i.saturating_add(1))).collect())
    }
}

impl<const N: usize> From<[Range<u32>; N]> for IndexSet {
    fn from(ranges: [Range<u32>; N]) -> Self {
        IndexSet::from(&ranges[..])
    }
}

impl<const N: usize> From<[u32; N]> for IndexSet {
    fn from(list: [u32; N]) -> Self {
        IndexSet::from(&list[..])
    }
}

impl From<Range<u32>> for IndexSet {
    fn from(r: Range<u32>) -> Self {
        IndexSet::from_spans(vec![(r.start, r.end)])
    }
}

impl From<RangeFrom<u32>> for IndexSet {
    fn from(r: RangeFrom<u32>) -> Self {
        IndexSet::from_spans(vec![(r.start, u32::MAX)])
    }
}

impl From<RangeFull> for IndexSet {
    fn from(_: RangeFull) -> Self {
        IndexSet::default()
    }
}

impl From<RangeInclusive<u32>> for IndexSet {
    fn from(r: RangeInclusive<u32>) -> Self {
        let (start, end) = (*r.start(), *r.end());
        IndexSet::from_spans(vec![(start, end.saturating_add(1))])
    }
}

impl From<RangeTo<u32>> for IndexSet {
    fn from(r: RangeTo<u32>) -> Self {
        IndexSet::from_spans(vec![(0, r.end)])
    }
}

impl From<RangeToInclusive<u32>> for IndexSet {
    fn from(r: RangeToInclusive<u32>) -> Self {
        IndexSet::from_spans(vec![(0, r.end.saturating_add(1))])
    }
}

impl From<u32> for IndexSet {
    fn from(i: u32) -> Self {
        IndexSet::from_spans(vec![(i, i.saturating_add(1))])
    }
}

impl From<Vec<Range<u32>>> for IndexSet {
    fn from(ranges: Vec<Range<u32>>) -> Self {
        IndexSet::from(&ranges[..])
    }
}

impl From<Vec<u32>> for IndexSet {
    fn from(list: Vec<u32>) -> Self {
        IndexSet::from(&list[..])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Each input set reduces to the expected normalized spans.
    #[test]
    fn normalization() {
        let cases: &[(IndexSet, &[(u32, u32)])] = &[
            ((..).into(), &[]),
            ((5..5).into(), &[]), // degenerate -> nothing -> "all"
            ((0..5).into(), &[(0, 5)]),
            ((0..=4).into(), &[(0, 5)]),
            ((5..).into(), &[(5, u32::MAX)]),
            ([3u32, 1, 2, 1].into(), &[(1, 4)]), // unsorted+dup, adjacent merge
            ([0..3, 8..10].into(), &[(0, 3), (8, 10)]), // disjoint
            ([0..5, 3..10].into(), &[(0, 10)]),  // overlapping
            ([0..10, 2..5].into(), &[(0, 10)]),  // contained
        ];
        for (set, expected) in cases {
            assert_eq!(set.spans, expected.to_vec(), "spans for {set:?}");
            assert_eq!(set.is_all(), expected.is_empty(), "is_all for {set:?}");
        }
        // The empty set is the canonical "all", so `..` and `default()` should agree.
        assert_eq!(IndexSet::from(..), IndexSet::default());
    }

    /// `keep(i)` for representative in/out indices, including half-open boundaries, gaps
    /// between disjoint ranges, and open-ended ranges probed past any real sheet bound.
    #[test]
    fn membership() {
        let cases: &[(IndexSet, &[(u32, bool)])] = &[
            (IndexSet::default(), &[(0, true), (999, true)]), // all
            ((0..5).into(), &[(0, true), (4, true), (5, false)]), // half-open
            ((0..=4).into(), &[(4, true), (5, false)]),
            ((5..).into(), &[(4, false), (5, true), (1_000_000, true)]),
            (
                [3u32, 1, 2, 1].into(),
                &[(0, false), (1, true), (3, true), (4, false)],
            ),
            (
                [0..3, 8..10].into(),
                &[
                    (2, true),
                    (3, false),
                    (7, false),
                    (8, true),
                    (9, true),
                    (10, false),
                ],
            ),
        ];
        for (set, probes) in cases {
            for &(i, kept) in *probes {
                assert_eq!(set.keep(i), kept, "keep({i}) for {set:?}");
            }
        }
    }

    /// `selected_count(bound)` capacity estimate, clamped to the bound.
    #[test]
    fn selected_count_clamps_to_bound() {
        let cases: &[(IndexSet, u32, u64)] = &[
            ([0..3, 8..10].into(), 100, 5), // 3 + 2
            ([0..3, 8..10].into(), 9, 4),   // 3 + (9 - 8)
            (IndexSet::default(), 7, 7),    // "all" -> bound
            ((8..10).into(), 5, 0),         // span starts beyond bound
        ];
        for (set, bound, expected) in cases {
            assert_eq!(
                set.selected_count(*bound),
                *expected,
                "selected_count({bound}) for {set:?}"
            );
        }
    }
}
