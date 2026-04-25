//! Cell fill: background color + pattern.  Read out of xlsx imports
//! (§2.3 of `docs/XLSX_IMPORT_PLAN.md`).
//!
//! Excel has ~18 pattern types (solid, gray125, dark-horizontal, …).
//! Terminal cells can only render a solid background; v1 collapses
//! every patterned fill to [`FillPattern::Solid`] so users still see
//! *some* color where Excel shows one.  The distinction is documented
//! as round-trip-lossy (the hatch pattern is dropped; only the
//! foreground color survives).

use crate::RgbColor;

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum FillPattern {
    /// No fill — the terminal default background shows through.
    #[default]
    None,
    /// Solid color fill.  This covers both the xlsx `"solid"` pattern
    /// and every patterned fill we collapse down.
    Solid,
}

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Fill {
    pub pattern: FillPattern,
    /// The color we actually paint.  For `Solid` fills this is the
    /// Excel "foreground color" (counter-intuitively: Excel stores the
    /// single-color fill under `fg_color`, not `bg_color`).  `None`
    /// when the fill has no color to paint (e.g. `FillPattern::None`
    /// or a corrupt xlsx).
    pub bg: Option<RgbColor>,
}

impl Fill {
    /// The no-override sentinel.  UI code treats this as "inherit
    /// the terminal default" and does not store it in the per-cell map.
    pub const DEFAULT: Fill = Fill {
        pattern: FillPattern::None,
        bg: None,
    };

    pub fn is_default(self) -> bool {
        self == Fill::DEFAULT
    }

    /// Convenience constructor for the overwhelmingly common case
    /// (`solid` fill with one color).
    pub fn solid(color: RgbColor) -> Self {
        Self {
            pattern: FillPattern::Solid,
            bg: Some(color),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_no_fill() {
        let f = Fill::default();
        assert_eq!(f.pattern, FillPattern::None);
        assert!(f.bg.is_none());
        assert!(f.is_default());
        assert_eq!(f, Fill::DEFAULT);
    }

    #[test]
    fn solid_sets_pattern_and_color() {
        let red = RgbColor { r: 255, g: 0, b: 0 };
        let f = Fill::solid(red);
        assert_eq!(f.pattern, FillPattern::Solid);
        assert_eq!(f.bg, Some(red));
        assert!(!f.is_default());
    }

    #[test]
    fn solid_with_black_is_not_default() {
        // Important: `is_default` is "no fill", not "bg == black".
        let f = Fill::solid(RgbColor::BLACK);
        assert!(!f.is_default());
    }
}
