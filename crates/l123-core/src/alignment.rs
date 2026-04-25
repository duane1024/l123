//! Per-cell alignment (horizontal, vertical, wrap-text) carried through
//! from xlsx imports.  1-2-3 R3.4a had no first-class alignment command
//! other than label prefixes (`'`, `"`, `^`, `\`); this type exists to
//! preserve Excel alignment on round-trip and to override the default
//! label=left / number=right rendering when a workbook explicitly sets
//! a horizontal alignment.  See `docs/XLSX_IMPORT_PLAN.md` §2.1.

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum HAlign {
    /// Excel's default: labels flush left, numbers flush right.
    #[default]
    General,
    Left,
    Center,
    Right,
    /// Repeats the cell's content to fill the column width.
    Fill,
    Justify,
    /// Excel's `centerContinuous` — centered across a horizontal run of
    /// otherwise-empty cells.  Parsed for round-trip; render behavior is
    /// a later pass.
    CenterAcross,
}

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum VAlign {
    Top,
    Center,
    #[default]
    Bottom,
}

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct Alignment {
    pub horizontal: HAlign,
    pub vertical: VAlign,
    pub wrap_text: bool,
}

impl Alignment {
    /// The no-override sentinel.  UI code treats this as "inherit
    /// 1-2-3 defaults" and does not store it in the per-cell map.
    pub const DEFAULT: Alignment = Alignment {
        horizontal: HAlign::General,
        vertical: VAlign::Bottom,
        wrap_text: false,
    };

    pub fn is_default(self) -> bool {
        self == Alignment::DEFAULT
    }
}

impl HAlign {
    /// Parse the `horizontal` attribute string found on an xlsx
    /// `<alignment>` element.  Unknown tokens fall back to `General`
    /// so we never panic on weird inputs; callers may log.
    pub fn from_xlsx_str(s: &str) -> Self {
        match s {
            "general" => HAlign::General,
            "left" => HAlign::Left,
            "center" => HAlign::Center,
            "right" => HAlign::Right,
            "fill" => HAlign::Fill,
            "justify" => HAlign::Justify,
            "centerContinuous" => HAlign::CenterAcross,
            // `distributed` is Asian-text-justify; map to Justify for v1.
            "distributed" => HAlign::Justify,
            _ => HAlign::General,
        }
    }

    pub fn as_xlsx_str(self) -> &'static str {
        match self {
            HAlign::General => "general",
            HAlign::Left => "left",
            HAlign::Center => "center",
            HAlign::Right => "right",
            HAlign::Fill => "fill",
            HAlign::Justify => "justify",
            HAlign::CenterAcross => "centerContinuous",
        }
    }
}

impl VAlign {
    pub fn from_xlsx_str(s: &str) -> Self {
        match s {
            "top" => VAlign::Top,
            "center" => VAlign::Center,
            "bottom" => VAlign::Bottom,
            // `justify` and `distributed` collapse to `Bottom` since the
            // terminal grid is one line per cell.
            _ => VAlign::Bottom,
        }
    }

    pub fn as_xlsx_str(self) -> &'static str {
        match self {
            VAlign::Top => "top",
            VAlign::Center => "center",
            VAlign::Bottom => "bottom",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_general_bottom_no_wrap() {
        let a = Alignment::default();
        assert_eq!(a.horizontal, HAlign::General);
        assert_eq!(a.vertical, VAlign::Bottom);
        assert!(!a.wrap_text);
        assert!(a.is_default());
        assert_eq!(a, Alignment::DEFAULT);
    }

    #[test]
    fn halign_known_strings_round_trip() {
        for h in [
            HAlign::General,
            HAlign::Left,
            HAlign::Center,
            HAlign::Right,
            HAlign::Fill,
            HAlign::Justify,
            HAlign::CenterAcross,
        ] {
            assert_eq!(HAlign::from_xlsx_str(h.as_xlsx_str()), h);
        }
    }

    #[test]
    fn halign_unknown_falls_back_to_general() {
        assert_eq!(HAlign::from_xlsx_str(""), HAlign::General);
        assert_eq!(HAlign::from_xlsx_str("bogus"), HAlign::General);
        // Asian-language distributed collapses to Justify intentionally.
        assert_eq!(HAlign::from_xlsx_str("distributed"), HAlign::Justify);
    }

    #[test]
    fn valign_known_strings_round_trip() {
        for v in [VAlign::Top, VAlign::Center, VAlign::Bottom] {
            assert_eq!(VAlign::from_xlsx_str(v.as_xlsx_str()), v);
        }
    }

    #[test]
    fn valign_unknown_falls_back_to_bottom() {
        assert_eq!(VAlign::from_xlsx_str(""), VAlign::Bottom);
        assert_eq!(VAlign::from_xlsx_str("justify"), VAlign::Bottom);
        assert_eq!(VAlign::from_xlsx_str("distributed"), VAlign::Bottom);
    }

    #[test]
    fn non_default_is_not_is_default() {
        let a = Alignment {
            horizontal: HAlign::Center,
            ..Default::default()
        };
        assert!(!a.is_default());
    }

    #[test]
    fn wrap_text_alone_breaks_default() {
        let a = Alignment {
            wrap_text: true,
            ..Default::default()
        };
        assert!(!a.is_default());
    }
}
