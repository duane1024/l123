//! Extended per-cell font attributes carried from xlsx.  Sits
//! *alongside* [`crate::TextStyle`] rather than replacing it:
//!
//! * `TextStyle` is the native 1-2-3 WYSIWYG triple (bold / italic /
//!   underline) set by `:Format`.  Its marker contract
//!   (`{Bold Italic Underline}` on control-panel line 1) is held
//!   constant here.
//! * `FontStyle` is everything xlsx adds that L123 has no
//!   first-class command to set — color, point size, strikethrough.
//!   Read-and-preserve for `size`; rendered for `color` and `strike`.
//!
//! See `docs/XLSX_IMPORT_PLAN.md` §3.1 for the wiring contract.

use crate::RgbColor;

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct FontStyle {
    /// Foreground text color.  `None` means "inherit the terminal
    /// default" (no explicit fg applied).
    pub color: Option<RgbColor>,
    /// Font size in points.  `None` means default (Excel's default
    /// is 11pt; IronCalc's fresh-Model default happens to be 13).
    /// Terminal cells are uniform-size so this field is preserve-only:
    /// we read it out of xlsx, carry it through `set_cell_font_style`
    /// on save, but never resize the rendered glyph.
    pub size: Option<u8>,
    /// Strikethrough on the glyph.  Rendered via
    /// `Modifier::CROSSED_OUT` at the UI layer.
    pub strike: bool,
}

impl FontStyle {
    /// The no-override sentinel.  UI code treats this as "inherit the
    /// terminal default" and does not store it in the per-cell map.
    pub const DEFAULT: FontStyle = FontStyle {
        color: None,
        size: None,
        strike: false,
    };

    pub fn is_default(self) -> bool {
        self == FontStyle::DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_all_none_no_strike() {
        let f = FontStyle::default();
        assert!(f.color.is_none());
        assert!(f.size.is_none());
        assert!(!f.strike);
        assert!(f.is_default());
        assert_eq!(f, FontStyle::DEFAULT);
    }

    #[test]
    fn any_attribute_breaks_default() {
        let red = RgbColor { r: 255, g: 0, b: 0 };
        assert!(!FontStyle {
            color: Some(red),
            ..Default::default()
        }
        .is_default());
        assert!(!FontStyle {
            size: Some(12),
            ..Default::default()
        }
        .is_default());
        assert!(!FontStyle {
            strike: true,
            ..Default::default()
        }
        .is_default());
    }
}
