//! Chrome theme — controls how the row/column header strip is painted.
//!
//! Today the theme only colors the column-letter row across the top of
//! the grid and the row-number gutter on the left. Cell content,
//! `:Format Color`, xlsx fills, and the rest of the chrome are
//! unaffected. The intent is "give the DOS look of black-on-cyan
//! headers without rebuilding the whole palette."

use ratatui::style::{Color, Modifier, Style};

/// Recognized themes. New variants land here; CLI/env/file parsers
/// dispatch through [`Theme::parse`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Theme {
    /// Today's look — headers paint with `Modifier::REVERSED` so the
    /// terminal's own foreground/background invert. Untouched on
    /// purpose: existing acceptance transcripts snapshot this.
    #[default]
    Default,
    /// Lotus 1-2-3 R3.4a DOS chrome: black row/column labels on a
    /// cyan strip. Mirrors the DOSBox-X reference screenshot.
    Dos,
}

impl Theme {
    /// Parse a `--theme` value or `theme=` config key. Case-insensitive;
    /// `default` and `dos` are the canonical names.
    pub fn parse(s: &str) -> Option<Self> {
        match s.trim().to_ascii_lowercase().as_str() {
            "default" => Some(Theme::Default),
            "dos" => Some(Theme::Dos),
            _ => None,
        }
    }

    /// Canonical lowercase name — what `l123 config` prints and what
    /// the sample `L123.CNF` documents.
    pub fn name(self) -> &'static str {
        match self {
            Theme::Default => "default",
            Theme::Dos => "dos",
        }
    }

    /// Style used to paint the column-header row and the row-number
    /// gutter. The grid renderer is the only caller.
    pub fn header_style(self) -> Style {
        match self {
            Theme::Default => Style::default().add_modifier(Modifier::REVERSED),
            Theme::Dos => Style::default()
                .bg(Color::Rgb(0, 170, 170))
                .fg(Color::Rgb(0, 0, 0)),
        }
    }

    /// Style used to highlight the column letter / row number that
    /// contains the active pointer, and to paint the upper-left
    /// (sheet-identity) corner. Default keeps the same look as
    /// [`Self::header_style`] (today's behavior — no special active
    /// emphasis); DOS uses CGA blue (`#0000aa`) so the active
    /// row/column reads at a glance against the resting cyan strip.
    pub fn header_active_style(self) -> Style {
        match self {
            Theme::Default => self.header_style(),
            Theme::Dos => Style::default()
                .bg(Color::Rgb(0, 0, 170))
                .fg(Color::Rgb(255, 255, 255)),
        }
    }

    /// Style used to paint the highlighted cell (the pointer in
    /// READY, the whole range in POINT). Default uses `REVERSED` so
    /// the terminal's own pair inverts; DOS paints the same teal
    /// (`#00aaaa`) used by the resting row/column labels — the
    /// selection blends visually with the chrome strip and is set
    /// apart by the active row/column emphasis instead.
    pub fn cell_highlight_style(self) -> Style {
        match self {
            Theme::Default => Style::default().add_modifier(Modifier::REVERSED),
            Theme::Dos => Style::default()
                .bg(Color::Rgb(0, 170, 170))
                .fg(Color::Rgb(0, 0, 0)),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_canonical_names() {
        assert_eq!(Theme::parse("default"), Some(Theme::Default));
        assert_eq!(Theme::parse("dos"), Some(Theme::Dos));
    }

    #[test]
    fn parse_is_case_insensitive_and_trims() {
        assert_eq!(Theme::parse("  DOS  "), Some(Theme::Dos));
        assert_eq!(Theme::parse("Default"), Some(Theme::Default));
    }

    #[test]
    fn parse_rejects_unknown() {
        assert_eq!(Theme::parse(""), None);
        assert_eq!(Theme::parse("amber"), None);
        assert_eq!(Theme::parse("classic"), None);
    }

    #[test]
    fn name_round_trips_through_parse() {
        for t in [Theme::Default, Theme::Dos] {
            assert_eq!(Theme::parse(t.name()), Some(t));
        }
    }

    #[test]
    fn default_theme_is_default_variant() {
        assert_eq!(Theme::default(), Theme::Default);
    }

    #[test]
    fn default_header_style_uses_reversed_only() {
        let s = Theme::Default.header_style();
        assert!(s.add_modifier.contains(Modifier::REVERSED));
        assert_eq!(s.fg, None);
        assert_eq!(s.bg, None);
    }

    #[test]
    fn dos_header_style_is_black_on_cyan_no_reverse() {
        let s = Theme::Dos.header_style();
        assert_eq!(s.bg, Some(Color::Rgb(0, 170, 170)));
        assert_eq!(s.fg, Some(Color::Rgb(0, 0, 0)));
        assert!(!s.add_modifier.contains(Modifier::REVERSED));
    }

    #[test]
    fn default_active_header_matches_header_no_special_emphasis() {
        assert_eq!(
            Theme::Default.header_active_style(),
            Theme::Default.header_style(),
            "default theme should not differentiate the active header",
        );
    }

    #[test]
    fn dos_active_header_is_white_on_deep_blue() {
        let s = Theme::Dos.header_active_style();
        assert_eq!(s.bg, Some(Color::Rgb(0, 0, 170)));
        assert_eq!(s.fg, Some(Color::Rgb(255, 255, 255)));
        // And it's distinct from the resting header style — that's
        // the whole point of the `_active_` variant.
        assert_ne!(s.bg, Theme::Dos.header_style().bg);
    }

    #[test]
    fn default_cell_highlight_uses_reversed_only() {
        let s = Theme::Default.cell_highlight_style();
        assert!(s.add_modifier.contains(Modifier::REVERSED));
        assert_eq!(s.fg, None);
        assert_eq!(s.bg, None);
    }

    #[test]
    fn dos_cell_highlight_matches_resting_label_teal() {
        let s = Theme::Dos.cell_highlight_style();
        // Same teal as the resting row/column labels — the selection
        // blends with the chrome and is set apart by the active
        // row/column emphasis instead.
        assert_eq!(s.bg, Some(Color::Rgb(0, 170, 170)));
        assert_eq!(s.bg, Theme::Dos.header_style().bg);
        assert_eq!(s.fg, Some(Color::Rgb(0, 0, 0)));
        assert!(!s.add_modifier.contains(Modifier::REVERSED));
    }
}
