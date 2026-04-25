//! Merged cell ranges from xlsx imports.  See
//! `docs/XLSX_IMPORT_PLAN.md` §2.4.
//!
//! 1-2-3 R3.4a had no first-class "merge cells" command; this type
//! exists to round-trip Excel merges and to let the grid renderer
//! paint the anchor's content across the span.
//!
//! ## Anchor + end convention
//!
//! `anchor` is the top-left cell of the merged region (Excel calls
//! this the "active cell" of the merge — the only one whose content
//! is preserved). `end` is the bottom-right.  Single-cell "merges"
//! (`anchor == end`) are technically valid but pointless; the
//! engine adapter skips emitting them.
//!
//! Both fields share the same `sheet`; merges never span sheets.

use crate::Address;

#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub struct Merge {
    pub anchor: Address,
    pub end: Address,
}

impl Merge {
    /// Construct a merge from two corner cells.  The constructor
    /// normalizes so `anchor` is top-left and `end` is bottom-right
    /// regardless of which order the caller passed them in.  Returns
    /// `None` when the two cells are on different sheets.
    pub fn from_corners(a: Address, b: Address) -> Option<Self> {
        if a.sheet != b.sheet {
            return None;
        }
        Some(Self {
            anchor: Address::new(a.sheet, a.col.min(b.col), a.row.min(b.row)),
            end: Address::new(a.sheet, a.col.max(b.col), a.row.max(b.row)),
        })
    }

    /// True when `addr` falls inside the merged rectangle (inclusive
    /// of both corners). False when `addr` is on a different sheet.
    pub fn contains(&self, addr: Address) -> bool {
        addr.sheet == self.anchor.sheet
            && addr.col >= self.anchor.col
            && addr.col <= self.end.col
            && addr.row >= self.anchor.row
            && addr.row <= self.end.row
    }

    /// True when `addr` is the merge's anchor (top-left) cell.
    pub fn is_anchor(&self, addr: Address) -> bool {
        addr == self.anchor
    }

    /// Number of columns the merge spans (always >= 1).
    pub fn col_span(&self) -> u16 {
        self.end.col - self.anchor.col + 1
    }

    /// Number of rows the merge spans (always >= 1).
    pub fn row_span(&self) -> u32 {
        self.end.row - self.anchor.row + 1
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SheetId;

    fn a(col: u16, row: u32) -> Address {
        Address::new(SheetId::A, col, row)
    }

    #[test]
    fn from_corners_normalizes_corners() {
        let m = Merge::from_corners(a(2, 5), a(0, 3)).unwrap();
        assert_eq!(m.anchor, a(0, 3));
        assert_eq!(m.end, a(2, 5));
    }

    #[test]
    fn from_corners_rejects_cross_sheet() {
        let s2 = Address::new(SheetId(1), 0, 0);
        assert!(Merge::from_corners(a(0, 0), s2).is_none());
    }

    #[test]
    fn contains_is_inclusive() {
        let m = Merge::from_corners(a(1, 1), a(3, 3)).unwrap();
        assert!(m.contains(a(1, 1)));
        assert!(m.contains(a(3, 3)));
        assert!(m.contains(a(2, 2)));
        assert!(!m.contains(a(0, 1)));
        assert!(!m.contains(a(4, 1)));
        assert!(!m.contains(a(1, 0)));
        assert!(!m.contains(a(1, 4)));
    }

    #[test]
    fn contains_is_false_for_other_sheet() {
        let m = Merge::from_corners(a(1, 1), a(3, 3)).unwrap();
        let other = Address::new(SheetId(1), 2, 2);
        assert!(!m.contains(other));
    }

    #[test]
    fn is_anchor_only_for_top_left() {
        let m = Merge::from_corners(a(1, 1), a(3, 3)).unwrap();
        assert!(m.is_anchor(a(1, 1)));
        assert!(!m.is_anchor(a(2, 2)));
        assert!(!m.is_anchor(a(3, 3)));
    }

    #[test]
    fn spans_are_one_when_single_cell() {
        let m = Merge::from_corners(a(5, 5), a(5, 5)).unwrap();
        assert_eq!(m.col_span(), 1);
        assert_eq!(m.row_span(), 1);
    }

    #[test]
    fn spans_count_inclusive_endpoints() {
        let m = Merge::from_corners(a(2, 3), a(5, 7)).unwrap();
        assert_eq!(m.col_span(), 4); // cols 2,3,4,5
        assert_eq!(m.row_span(), 5); // rows 3,4,5,6,7
    }
}
