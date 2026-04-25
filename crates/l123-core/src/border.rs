//! Per-cell borders carried from xlsx imports.
//!
//! 1-2-3 R3.4a had no first-class "border" command; this type exists
//! only to preserve Excel borders on round-trip and to let the grid
//! renderer paint visible column / row separators.  See
//! `docs/XLSX_IMPORT_PLAN.md` §2.2.
//!
//! ## Fidelity
//!
//! Excel/IronCalc model ~9 border styles (several dash-dot variants).
//! [`BorderStyle`] collapses those to 6 broad buckets — thin, medium,
//! thick, double, dashed, dotted — since a terminal cell only has a
//! handful of visually distinguishable glyphs.  Round-trip therefore
//! drops sub-variants: e.g. `MediumDashDotDot` maps to `Dashed` on
//! load and writes back as `MediumDashed` on save.  Documented loss;
//! not a bug.
//!
//! Diagonals are *not* represented here; they're skipped for v1 per
//! the plan.  A cell that has only diagonal borders in xlsx will
//! round-trip with the diagonals dropped.

use crate::RgbColor;

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq, PartialOrd, Ord)]
pub enum BorderStyle {
    /// Default "fine line" — maps to IronCalc `Thin`.
    #[default]
    Thin,
    Medium,
    Thick,
    /// Two parallel thin lines — rendered with the `║` box-drawing
    /// glyph.  Heavier than `Medium` for the "heavier wins" edge
    /// merge, ranked one rung below `Thick`.
    Double,
    Dashed,
    Dotted,
}

impl BorderStyle {
    /// Ordering used by [`BorderEdge::merge_heavier`] to pick which
    /// style wins when two adjacent cells both set a border on the
    /// seam between them.  Roughly: thicker > double > medium >
    /// dashed/dotted variants > thin.  Picked to match Excel's own
    /// "later write wins" tie-break as closely as we can from
    /// import-time data alone (we can't see write order, so we lean
    /// on visual weight instead).
    pub fn weight(self) -> u8 {
        match self {
            BorderStyle::Thin => 1,
            BorderStyle::Dotted => 2,
            BorderStyle::Dashed => 3,
            BorderStyle::Medium => 4,
            BorderStyle::Double => 5,
            BorderStyle::Thick => 6,
        }
    }

    /// Unicode box-drawing glyph for a vertical border.  Used by the
    /// grid renderer to paint the seam between two cells.  Distinct
    /// glyphs are limited by what terminals render consistently —
    /// Thin and Medium share `│` since most fixed-width fonts can't
    /// express a "slightly heavier" variant.
    pub fn vertical_glyph(self) -> char {
        match self {
            BorderStyle::Thin | BorderStyle::Medium => '│',
            BorderStyle::Thick => '┃',
            BorderStyle::Double => '║',
            BorderStyle::Dashed => '╎',
            BorderStyle::Dotted => '┊',
        }
    }
}

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct BorderEdge {
    pub style: BorderStyle,
    pub color: Option<RgbColor>,
}

impl BorderEdge {
    /// When cells on both sides of a seam set a border on the shared
    /// edge, pick the heavier one to render.  Style weight breaks
    /// ties first; when styles tie, `self.color` wins (the cell on
    /// the *left* of a vertical seam owns the color for tie-breaks —
    /// an arbitrary-but-stable rule).
    pub fn merge_heavier(self, other: BorderEdge) -> BorderEdge {
        if other.style.weight() > self.style.weight() {
            other
        } else {
            self
        }
    }
}

/// All four edges of a cell.  Each is `Option<BorderEdge>` so "no
/// border" is distinct from "default-weight border".  A cell with no
/// entry in the UI's `cell_borders` map has `Border::default()` —
/// every edge `None`.
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Border {
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
}

impl Border {
    pub const NONE: Border = Border {
        left: None,
        right: None,
        top: None,
        bottom: None,
    };

    pub fn is_default(self) -> bool {
        self == Border::NONE
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_all_none() {
        assert_eq!(Border::default(), Border::NONE);
        assert!(Border::default().is_default());
    }

    #[test]
    fn border_style_weight_is_monotonic_for_common_order() {
        // Thin < Dotted < Dashed < Medium < Double < Thick.
        assert!(BorderStyle::Thin.weight() < BorderStyle::Dotted.weight());
        assert!(BorderStyle::Dotted.weight() < BorderStyle::Dashed.weight());
        assert!(BorderStyle::Dashed.weight() < BorderStyle::Medium.weight());
        assert!(BorderStyle::Medium.weight() < BorderStyle::Double.weight());
        assert!(BorderStyle::Double.weight() < BorderStyle::Thick.weight());
    }

    #[test]
    fn merge_heavier_picks_thicker() {
        let thin = BorderEdge {
            style: BorderStyle::Thin,
            color: None,
        };
        let thick = BorderEdge {
            style: BorderStyle::Thick,
            color: Some(RgbColor::BLACK),
        };
        assert_eq!(thin.merge_heavier(thick), thick);
        // Commutative for the "heavier" choice.
        assert_eq!(thick.merge_heavier(thin), thick);
    }

    #[test]
    fn merge_heavier_tie_keeps_self() {
        // Tie on weight → `self` wins (the left-of-seam cell).
        let left = BorderEdge {
            style: BorderStyle::Medium,
            color: Some(RgbColor { r: 255, g: 0, b: 0 }),
        };
        let right = BorderEdge {
            style: BorderStyle::Medium,
            color: Some(RgbColor { r: 0, g: 255, b: 0 }),
        };
        assert_eq!(left.merge_heavier(right), left);
    }

    #[test]
    fn vertical_glyph_distinguishes_styles() {
        assert_eq!(BorderStyle::Thin.vertical_glyph(), '│');
        assert_eq!(BorderStyle::Medium.vertical_glyph(), '│');
        assert_eq!(BorderStyle::Thick.vertical_glyph(), '┃');
        assert_eq!(BorderStyle::Double.vertical_glyph(), '║');
        assert_eq!(BorderStyle::Dashed.vertical_glyph(), '╎');
        assert_eq!(BorderStyle::Dotted.vertical_glyph(), '┊');
    }

    #[test]
    fn non_default_is_not_is_default() {
        let b = Border {
            right: Some(BorderEdge::default()),
            ..Default::default()
        };
        assert!(!b.is_default());
    }
}
