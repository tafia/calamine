use std::ops::{Range, RangeFrom, RangeFull, RangeInclusive, RangeTo, RangeToInclusive};

/// A normalized set of 0-based indices for projecting columns or rows.
///
/// Constructed via `From`/`Into` from a range, single index, list, or list of
/// ranges. An empty set (eg: `..` or `IndexSet::default()`) selects *everything*.
///
/// ```
/// use calamine::IndexSet;
///
/// let _: IndexSet = (0..5).into();        // contiguous range
/// let _: IndexSet = (5..).into();         // open-ended range
/// let _: IndexSet = [1, 3, 5].into();     // list of indexes
/// let _: IndexSet = [0..3, 8..10].into(); // disjoint ranges
/// let _: IndexSet = (..).into();          // everything
/// ```
///
/// Note: overlapping or duplicate inputs are merged.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct IndexSet {
    /// Sorted, merged, non-overlapping half-open `[start, end)` intervals.
    /// Empty == "everything", [`IndexSet::UNBOUNDED`] is the upper-bound sentinel.
    spans: Vec<(u32, u32)>,
}

impl IndexSet {
    /// Upper-bound sentinel; real worksheet ranges end far below
    /// this, so it can never collide with an addressable index.
    pub(crate) const UNBOUNDED: u32 = u32::MAX;

    /// Build normalized `IndexSet` from raw half-open `[start, end)` spans. Spans
    /// with `start >= end` contribute nothing, overlapping/adjacent spans merge.
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

    /// The maximum index this set can select, or [`IndexSet::UNBOUNDED`] if
    /// unbounded above (either "all", or an open-ended span such as `5..`).
    pub(crate) fn max_index(&self) -> u32 {
        match self.spans.last() {
            Some(&(_, end)) if end == Self::UNBOUNDED => Self::UNBOUNDED, // open-ended
            Some(&(_, end)) => end - 1, // half-open -> inclusive last
            None => Self::UNBOUNDED,    // "all"
        }
    }

    /// Number of selected indices that fall within `0..bound`.
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
        IndexSet::from_spans(vec![(r.start, IndexSet::UNBOUNDED)])
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
    use rstest::rstest;

    /// Each input set reduces to the expected normalized spans.
    #[rstest]
    #[case::all((..).into(), &[])]
    #[case::degenerate((5..5).into(), &[])] // start >= end -> nothing -> "all"
    #[case::half_open((0..5).into(), &[(0, 5)])]
    #[case::inclusive((0..=4).into(), &[(0, 5)])]
    #[case::open_ended((5..).into(), &[(5, IndexSet::UNBOUNDED)])]
    #[case::unsorted_dup([3u32, 1, 2, 1].into(), &[(1, 4)])] // adjacent merge
    #[case::disjoint([0..3, 8..10].into(), &[(0, 3), (8, 10)])]
    #[case::overlapping([0..5, 3..10].into(), &[(0, 10)])]
    #[case::contained([0..10, 2..5].into(), &[(0, 10)])]
    fn normalization(#[case] set: IndexSet, #[case] expected: &[(u32, u32)]) {
        assert_eq!(set.spans, expected.to_vec(), "spans for {set:?}");
        assert_eq!(set.is_all(), expected.is_empty(), "is_all for {set:?}");
    }

    /// The empty set is the canonical "all", so `..` and `default()` should agree.
    #[test]
    fn full_range_equals_default() {
        assert_eq!(IndexSet::from(..), IndexSet::default());
    }

    /// `keep(i)` for representative in/out indices, including half-open boundaries, gaps
    /// between disjoint ranges, and open-ended ranges probed past any real sheet bound.
    #[rstest]
    #[case::all(IndexSet::default(), &[(0, true), (999, true)])]
    #[case::disjoint(
        [0..3, 8..10].into(),
        &[(2, true), (3, false), (7, false), (8, true), (9, true), (10, false)]
    )]
    #[case::half_open((0..5).into(), &[(0, true), (4, true), (5, false)])]
    #[case::inclusive((0..=4).into(), &[(4, true), (5, false)])]
    #[case::list([3u32, 1, 2, 1].into(), &[(0, false), (1, true), (3, true), (4, false)])]
    #[case::open_ended((5..).into(), &[(4, false), (5, true), (1_000_000, true)])]
    fn membership(#[case] set: IndexSet, #[case] probes: &[(u32, bool)]) {
        for &(i, kept) in probes {
            assert_eq!(set.keep(i), kept, "keep({i}) for {set:?}");
        }
    }

    /// `max_index()` returns the inclusive upper bound, or `UNBOUNDED` when unbounded.
    #[rstest]
    #[case::all(IndexSet::default(), IndexSet::UNBOUNDED)]
    #[case::disjoint([0..3, 8..10].into(), 9)] // last span's end - 1
    #[case::half_open((0..5).into(), 4)] // half-open -> inclusive last
    #[case::inclusive((0..=4).into(), 4)]
    #[case::open_ended((5..).into(), IndexSet::UNBOUNDED)]
    #[case::single(3u32.into(), 3)]
    fn max_index_is_inclusive_ceiling(#[case] set: IndexSet, #[case] expected: u32) {
        assert_eq!(set.max_index(), expected, "max_index for {set:?}");
    }

    /// `selected_count(bound)` capacity estimate, clamped to the bound.
    #[rstest]
    #[case::all(IndexSet::default(), 7, 7)] // "all" -> bound
    #[case::beyond_bound((8..10).into(), 5, 0)] // span starts beyond bound
    #[case::clamped([0..3, 8..10].into(), 9, 4)] // 3 + (9 - 8)
    #[case::within_bound([0..3, 8..10].into(), 100, 5)] // 3 + 2
    fn selected_count_clamps_to_bound(
        #[case] set: IndexSet,
        #[case] bound: u32,
        #[case] expected: u64,
    ) {
        assert_eq!(
            set.selected_count(bound),
            expected,
            "selected_count({bound}) for {set:?}"
        );
    }
}
