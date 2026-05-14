//! `/Range Compare` — pure diff between two same-shape ranges (v0.4).
//!
//! The UI gathers two ranges via three POINTs (left, right, output anchor),
//! collects values for both ranges in row-major order, and passes them
//! through [`diff_ranges`]. The result is one [`DiffRow`] per differing
//! cell; equal cells produce no row. Shape mismatch (different widths,
//! heights, or multi-sheet inputs) is reported via [`DiffError`].

use l123_core::{Address, Range, Value};
use thiserror::Error;

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum DiffKind {
    /// Left cell has a value, right is `Empty`.
    OnlyLeft,
    /// Right cell has a value, left is `Empty`.
    OnlyRight,
    /// Both cells have values of the same type, but they aren't equal.
    BothDifferent,
    /// Both cells have values but their types disagree (e.g. number vs label).
    TypeMismatch,
}

impl DiffKind {
    /// Kebab-case label written into the result range's diff_kind column.
    pub fn label(self) -> &'static str {
        match self {
            DiffKind::OnlyLeft => "only-left",
            DiffKind::OnlyRight => "only-right",
            DiffKind::BothDifferent => "both-different",
            DiffKind::TypeMismatch => "type-mismatch",
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct DiffRow {
    /// Address inside the left range — anchors the diff back to the user's
    /// original data so they can navigate to the cell after running compare.
    pub addr: Address,
    pub left: Value,
    pub right: Value,
    pub kind: DiffKind,
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum DiffError {
    #[error("ranges have different shape")]
    ShapeMismatch,
    #[error("compare across sheets is not supported")]
    MultiSheet,
}

/// Diff `left` against `right`. Both value slices are row-major over their
/// respective ranges. Returns one [`DiffRow`] per differing cell, in
/// row-major order over the left range; equal cells produce no row.
pub fn diff_ranges(
    left: Range,
    right: Range,
    left_values: &[Value],
    right_values: &[Value],
) -> Result<Vec<DiffRow>, DiffError> {
    let ln = left.normalized();
    let rn = right.normalized();

    if !ln.is_single_sheet() || !rn.is_single_sheet() {
        return Err(DiffError::MultiSheet);
    }

    let (lw, lh) = dims(ln);
    let (rw, rh) = dims(rn);
    if lw != rw || lh != rh {
        return Err(DiffError::ShapeMismatch);
    }

    let expected = (lw as usize) * (lh as usize);
    if left_values.len() != expected || right_values.len() != expected {
        return Err(DiffError::ShapeMismatch);
    }

    let mut out = Vec::new();
    for dy in 0..lh {
        for dx in 0..lw {
            let idx = (dy * lw + dx) as usize;
            let lv = &left_values[idx];
            let rv = &right_values[idx];
            if let Some(kind) = classify(lv, rv) {
                let addr =
                    Address::new(ln.start.sheet, ln.start.col + dx as u16, ln.start.row + dy);
                out.push(DiffRow {
                    addr,
                    left: lv.clone(),
                    right: rv.clone(),
                    kind,
                });
            }
        }
    }
    Ok(out)
}

fn dims(r: Range) -> (u32, u32) {
    let w = (r.end.col - r.start.col + 1) as u32;
    let h = r.end.row - r.start.row + 1;
    (w, h)
}

fn classify(l: &Value, r: &Value) -> Option<DiffKind> {
    match (l, r) {
        (Value::Empty, Value::Empty) => None,
        (Value::Empty, _) => Some(DiffKind::OnlyRight),
        (_, Value::Empty) => Some(DiffKind::OnlyLeft),
        (Value::Number(a), Value::Number(b)) => {
            if a == b {
                None
            } else {
                Some(DiffKind::BothDifferent)
            }
        }
        (Value::Text(a), Value::Text(b)) => {
            if a == b {
                None
            } else {
                Some(DiffKind::BothDifferent)
            }
        }
        (Value::Bool(a), Value::Bool(b)) => {
            if a == b {
                None
            } else {
                Some(DiffKind::BothDifferent)
            }
        }
        (Value::Error(a), Value::Error(b)) => {
            if a == b {
                None
            } else {
                Some(DiffKind::BothDifferent)
            }
        }
        _ => Some(DiffKind::TypeMismatch),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use l123_core::SheetId;

    fn a(col: u16, row: u32) -> Address {
        Address::new(SheetId::A, col, row)
    }

    fn r(start: Address, end: Address) -> Range {
        Range { start, end }
    }

    #[test]
    fn equal_ranges_produce_no_rows() {
        let left = r(a(0, 0), a(0, 2));
        let right = r(a(2, 0), a(2, 2));
        let vals = vec![Value::Number(1.0), Value::Number(2.0), Value::Number(3.0)];
        let out = diff_ranges(left, right, &vals, &vals).unwrap();
        assert!(out.is_empty());
    }

    #[test]
    fn numeric_difference_is_both_different() {
        let left = r(a(0, 0), a(0, 1));
        let right = r(a(2, 0), a(2, 1));
        let lv = vec![Value::Number(10.0), Value::Number(20.0)];
        let rv = vec![Value::Number(10.0), Value::Number(99.0)];
        let out = diff_ranges(left, right, &lv, &rv).unwrap();
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].addr, a(0, 1));
        assert_eq!(out[0].kind, DiffKind::BothDifferent);
        assert_eq!(out[0].left, Value::Number(20.0));
        assert_eq!(out[0].right, Value::Number(99.0));
    }

    #[test]
    fn number_vs_label_is_type_mismatch() {
        let left = r(a(0, 0), a(0, 0));
        let right = r(a(2, 0), a(2, 0));
        let lv = vec![Value::Number(10.0)];
        let rv = vec![Value::Text("ten".into())];
        let out = diff_ranges(left, right, &lv, &rv).unwrap();
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].kind, DiffKind::TypeMismatch);
    }

    #[test]
    fn left_value_vs_empty_is_only_left() {
        let left = r(a(0, 0), a(0, 0));
        let right = r(a(2, 0), a(2, 0));
        let lv = vec![Value::Number(40.0)];
        let rv = vec![Value::Empty];
        let out = diff_ranges(left, right, &lv, &rv).unwrap();
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].kind, DiffKind::OnlyLeft);
    }

    #[test]
    fn empty_vs_right_value_is_only_right() {
        let left = r(a(0, 0), a(0, 0));
        let right = r(a(2, 0), a(2, 0));
        let lv = vec![Value::Empty];
        let rv = vec![Value::Number(7.0)];
        let out = diff_ranges(left, right, &lv, &rv).unwrap();
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].kind, DiffKind::OnlyRight);
    }

    #[test]
    fn shape_mismatch_returns_error() {
        let left = r(a(0, 0), a(0, 2)); // 1 col × 3 rows
        let right = r(a(2, 0), a(2, 1)); // 1 col × 2 rows
        let lv = vec![Value::Number(1.0); 3];
        let rv = vec![Value::Number(1.0); 2];
        assert_eq!(
            diff_ranges(left, right, &lv, &rv),
            Err(DiffError::ShapeMismatch)
        );
    }

    #[test]
    fn row_major_order_preserved() {
        // 2 cols × 2 rows, three diffs in row-major order
        let left = r(a(0, 0), a(1, 1));
        let right = r(a(2, 0), a(3, 1));
        let lv = vec![
            Value::Number(1.0),
            Value::Number(2.0),
            Value::Number(3.0),
            Value::Number(4.0),
        ];
        let rv = vec![
            Value::Number(9.0), // diff at (0,0)
            Value::Number(2.0),
            Value::Number(9.0), // diff at (0,1)
            Value::Number(9.0), // diff at (1,1)
        ];
        let out = diff_ranges(left, right, &lv, &rv).unwrap();
        assert_eq!(out.len(), 3);
        assert_eq!(out[0].addr, a(0, 0));
        assert_eq!(out[1].addr, a(0, 1));
        assert_eq!(out[2].addr, a(1, 1));
    }

    #[test]
    fn diff_kind_labels() {
        assert_eq!(DiffKind::OnlyLeft.label(), "only-left");
        assert_eq!(DiffKind::OnlyRight.label(), "only-right");
        assert_eq!(DiffKind::BothDifferent.label(), "both-different");
        assert_eq!(DiffKind::TypeMismatch.label(), "type-mismatch");
    }
}
