//! Chrome color palettes selectable via `--theme` / `theme = …`.
//!
//! A [`Theme`] is a small bundle of RGB triples used to paint
//! non-document chrome: status bar, menu/help selection highlights,
//! splash background, monochrome status overlay. Most themes also
//! carry optional [`Rgb`] fields that let them paint the column /
//! row-number headers and the default cell field — these are `None`
//! on the [`ThemeName::Default`] palette so the out-of-the-box look
//! falls back to the terminal's own colors.
//!
//! Cell colors that come from the document itself (`:Format Color`,
//! xlsx fills/fonts, sheet tab tints) flow through unchanged — those
//! are data, not chrome.
//!
//! The palettes mirror the reference looks: the bare-terminal default,
//! classic 1-2-3 R3.4a DOS (cyan headers on a black field), R3.4a
//! WYSIWYG (paper, magenta gridlines, blue selection), and two CRT
//! phosphor schemes (amber and green) for users who want the
//! terminal-monitor aesthetic.

use std::fmt;

use ratatui::style::Color;

/// One of the named themes selectable from the CLI / config / env.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeName {
    /// Out-of-the-box look — leaves cell and header colors at the
    /// terminal default. Used when no `--theme` / `L123_THEME` /
    /// `theme =` is set.
    #[default]
    Default,
    /// Classic 1-2-3 R3.4a DOS — black field, cyan headers, white text.
    Dos,
    /// 1-2-3 R3.4a WYSIWYG — light paper field, teal headers, magenta gridlines.
    Wysiwyg,
    /// CRT amber phosphor.
    Amber,
    /// CRT green phosphor.
    Green,
}

impl ThemeName {
    /// Canonical (lowercase) name as accepted by `--theme` / config.
    pub fn as_str(self) -> &'static str {
        match self {
            ThemeName::Default => "default",
            ThemeName::Dos => "dos",
            ThemeName::Wysiwyg => "wysiwyg",
            ThemeName::Amber => "amber",
            ThemeName::Green => "green",
        }
    }

    /// Every theme name in CLI/help/error order.
    pub const ALL: &'static [ThemeName] = &[
        ThemeName::Default,
        ThemeName::Dos,
        ThemeName::Wysiwyg,
        ThemeName::Amber,
        ThemeName::Green,
    ];

    /// Comma-separated list for error messages, e.g.
    /// `"default, dos, wysiwyg, amber, green"`.
    pub fn list_for_error() -> String {
        Self::ALL
            .iter()
            .map(|t| t.as_str())
            .collect::<Vec<_>>()
            .join(", ")
    }

    /// Resolve a theme by name. Accepts the canonical name plus a
    /// couple of obvious aliases (`classic` → dos, `paper` → wysiwyg).
    /// Case-insensitive. Returns `None` for unknown names.
    pub fn parse(name: &str) -> Option<Self> {
        match name.trim().to_ascii_lowercase().as_str() {
            "default" | "plain" | "none" => Some(ThemeName::Default),
            "dos" | "classic" => Some(ThemeName::Dos),
            "wysiwyg" | "paper" => Some(ThemeName::Wysiwyg),
            "amber" => Some(ThemeName::Amber),
            "green" => Some(ThemeName::Green),
            _ => None,
        }
    }

    /// Materialize the palette for this name.
    pub fn palette(self) -> Theme {
        match self {
            ThemeName::Default => Theme::DEFAULT,
            ThemeName::Dos => Theme::DOS,
            ThemeName::Wysiwyg => Theme::WYSIWYG,
            ThemeName::Amber => Theme::AMBER,
            ThemeName::Green => Theme::GREEN,
        }
    }
}

impl fmt::Display for ThemeName {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_str())
    }
}

/// 24-bit RGB triple. We ship our own newtype rather than depending on
/// ratatui or `l123-core::RgbColor` so this module stays free of
/// upward edges and can be unit-tested without any rendering crate.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Rgb {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Rgb {
    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Lift this triple into a ratatui [`Color::Rgb`]. Lets the
    /// renderer write `theme.status_fg.color()` instead of repeating
    /// the field destructure at every call site.
    pub fn color(self) -> Color {
        Color::Rgb(self.r, self.g, self.b)
    }
}

/// Resolved chrome palette. Every named theme materializes one of
/// these; the renderer reads `app.theme` and never branches on the
/// theme name itself.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Theme {
    pub name: ThemeName,
    /// Background of the status line (also the full-screen "field"
    /// behind splash text). DOS = teal, WYSIWYG = paper, CRT = black.
    pub status_bg: Rgb,
    /// Foreground of status text rendered on `status_bg`.
    pub status_fg: Rgb,
    /// Highlight background for the active menu item, focused F1-help
    /// link, header/footer of the help overlay, and READY-mode cell
    /// pointer when no document fill overrides it.
    pub selection_bg: Rgb,
    /// Foreground of selection text on `selection_bg`.
    pub selection_fg: Rgb,
    /// Foreground for unfocused F1-help hyperlinks.
    pub help_link_fg: Rgb,
    /// Background fill for the splash screen "wallpaper" area.
    pub splash_bg: Rgb,
    /// Foreground for splash banner / centered logo.
    pub splash_fg: Rgb,
    /// Accent foreground used for splash-screen labels and the mode
    /// indicator on the right of the control panel.
    pub accent_fg: Rgb,
    /// Background + foreground of the monochrome `/Worksheet Status`
    /// and `/Worksheet Global Default Status` overlays. The default
    /// theme keeps the green-on-black phosphor look, but CRT themes
    /// pull these in line with their own primary color so the overlay
    /// doesn't fight the rest of the chrome.
    pub mono_overlay_bg: Rgb,
    pub mono_overlay_fg: Rgb,
    /// Foreground for the comment-marker and other one-off error/alert
    /// glyphs. Stays red on every theme except amber/green where the
    /// pure red would clash with the phosphor monochrome.
    pub alert_fg: Rgb,
    /// Foreground for the dim gridline glyph painted at each cell's
    /// rightmost column when `:Display Options Grid Yes`.
    pub gridline_fg: Rgb,
    /// Background of the column-letter row and row-number gutter.
    /// `None` on the bare-terminal default theme — falls back to the
    /// `Modifier::REVERSED` look that shipped before themes existed.
    pub header_bg: Option<Rgb>,
    /// Foreground of column-letter / row-number text. `None` follows
    /// the same rule as `header_bg`.
    pub header_fg: Option<Rgb>,
    /// Default cell background painted behind cells with no
    /// xlsx-imported fill and no `:Display Mode` override. `None`
    /// leaves the cell at the terminal background.
    pub cell_bg: Option<Rgb>,
    /// Default cell foreground paired with `cell_bg`. `None` follows
    /// the same rule.
    pub cell_fg: Option<Rgb>,
}

impl Theme {
    /// Bare-terminal default — matches the look L123 shipped before
    /// themes existed. Cell field and headers fall through to the
    /// terminal's own colors; only the chrome bits (status bar,
    /// splash, menu highlights) carry concrete RGB values.
    pub const DEFAULT: Theme = Theme {
        name: ThemeName::Default,
        status_bg: Rgb::new(0x00, 0xAA, 0xAA),
        status_fg: Rgb::new(0x80, 0x80, 0x80),
        selection_bg: Rgb::new(0x00, 0xAA, 0xAA),
        selection_fg: Rgb::new(0x00, 0x00, 0x00),
        help_link_fg: Rgb::new(0x00, 0xAA, 0x00),
        splash_bg: Rgb::new(0x00, 0xAA, 0xAA),
        splash_fg: Rgb::new(0xFF, 0xFF, 0xFF),
        accent_fg: Rgb::new(0xFF, 0xFF, 0x55),
        mono_overlay_bg: Rgb::new(0x00, 0x00, 0x00),
        mono_overlay_fg: Rgb::new(0x00, 0xAA, 0x55),
        alert_fg: Rgb::new(0xFF, 0x00, 0x00),
        gridline_fg: Rgb::new(0x55, 0x55, 0x55),
        header_bg: None,
        header_fg: None,
        cell_bg: None,
        cell_fg: None,
    };

    /// Classic 1-2-3 R3.4a DOS — black field, white text, cyan
    /// column-letter and row-number gutter. Mirrors the screenshot
    /// of Lotus 1-2-3 R3.0 for MS-DOS.
    pub const DOS: Theme = Theme {
        name: ThemeName::Dos,
        // DOS VGA cyan (#00AAAA) is the same teal the splash uses.
        status_bg: Rgb::new(0x00, 0xAA, 0xAA),
        status_fg: Rgb::new(0xFF, 0xFF, 0xFF),
        selection_bg: Rgb::new(0x00, 0xAA, 0xAA),
        selection_fg: Rgb::new(0x00, 0x00, 0x00),
        help_link_fg: Rgb::new(0x00, 0xAA, 0x00),
        splash_bg: Rgb::new(0x00, 0xAA, 0xAA),
        splash_fg: Rgb::new(0xFF, 0xFF, 0xFF),
        accent_fg: Rgb::new(0xFF, 0xFF, 0x55),
        mono_overlay_bg: Rgb::new(0x00, 0x00, 0x00),
        mono_overlay_fg: Rgb::new(0x00, 0xAA, 0x55),
        alert_fg: Rgb::new(0xFF, 0x00, 0x00),
        gridline_fg: Rgb::new(0x55, 0x55, 0x55),
        header_bg: Some(Rgb::new(0x00, 0xAA, 0xAA)),
        header_fg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_bg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_fg: Some(Rgb::new(0xFF, 0xFF, 0xFF)),
    };

    /// 1-2-3 R3.4a WYSIWYG paper-look: light grey field, teal headers,
    /// magenta gridlines, blue selection.
    pub const WYSIWYG: Theme = Theme {
        name: ThemeName::Wysiwyg,
        status_bg: Rgb::new(0xC0, 0xC0, 0xC0),
        status_fg: Rgb::new(0x40, 0x40, 0x40),
        selection_bg: Rgb::new(0x00, 0x00, 0xAA),
        selection_fg: Rgb::new(0xFF, 0xFF, 0xFF),
        help_link_fg: Rgb::new(0x00, 0x00, 0xAA),
        splash_bg: Rgb::new(0xC0, 0xC0, 0xC0),
        splash_fg: Rgb::new(0x00, 0x00, 0x00),
        accent_fg: Rgb::new(0x00, 0x00, 0xAA),
        mono_overlay_bg: Rgb::new(0x00, 0x00, 0x00),
        mono_overlay_fg: Rgb::new(0x00, 0xAA, 0x55),
        alert_fg: Rgb::new(0xAA, 0x00, 0x00),
        gridline_fg: Rgb::new(0xAA, 0x00, 0xAA),
        header_bg: Some(Rgb::new(0x55, 0xAA, 0xAA)),
        header_fg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_bg: Some(Rgb::new(0xC0, 0xC0, 0xC0)),
        cell_fg: Some(Rgb::new(0x00, 0x00, 0x00)),
    };

    /// CRT amber phosphor — black field, two amber tones.
    pub const AMBER: Theme = Theme {
        name: ThemeName::Amber,
        status_bg: Rgb::new(0x00, 0x00, 0x00),
        status_fg: Rgb::new(0xFF, 0xB0, 0x00),
        selection_bg: Rgb::new(0xFF, 0xB0, 0x00),
        selection_fg: Rgb::new(0x00, 0x00, 0x00),
        help_link_fg: Rgb::new(0xFF, 0xD7, 0x6E),
        splash_bg: Rgb::new(0x00, 0x00, 0x00),
        splash_fg: Rgb::new(0xFF, 0xB0, 0x00),
        accent_fg: Rgb::new(0xFF, 0xD7, 0x6E),
        mono_overlay_bg: Rgb::new(0x00, 0x00, 0x00),
        mono_overlay_fg: Rgb::new(0xFF, 0xB0, 0x00),
        alert_fg: Rgb::new(0xFF, 0xD7, 0x6E),
        gridline_fg: Rgb::new(0x80, 0x58, 0x00),
        header_bg: Some(Rgb::new(0x80, 0x58, 0x00)),
        header_fg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_bg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_fg: Some(Rgb::new(0xFF, 0xB0, 0x00)),
    };

    /// CRT green phosphor — black field, two green tones.
    pub const GREEN: Theme = Theme {
        name: ThemeName::Green,
        status_bg: Rgb::new(0x00, 0x00, 0x00),
        status_fg: Rgb::new(0x00, 0xCC, 0x33),
        selection_bg: Rgb::new(0x00, 0xCC, 0x33),
        selection_fg: Rgb::new(0x00, 0x00, 0x00),
        help_link_fg: Rgb::new(0x99, 0xFF, 0x99),
        splash_bg: Rgb::new(0x00, 0x00, 0x00),
        splash_fg: Rgb::new(0x00, 0xCC, 0x33),
        accent_fg: Rgb::new(0x99, 0xFF, 0x99),
        mono_overlay_bg: Rgb::new(0x00, 0x00, 0x00),
        mono_overlay_fg: Rgb::new(0x00, 0xCC, 0x33),
        alert_fg: Rgb::new(0x99, 0xFF, 0x99),
        gridline_fg: Rgb::new(0x00, 0x55, 0x11),
        header_bg: Some(Rgb::new(0x00, 0x55, 0x11)),
        header_fg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_bg: Some(Rgb::new(0x00, 0x00, 0x00)),
        cell_fg: Some(Rgb::new(0x00, 0xCC, 0x33)),
    };
}

impl Default for Theme {
    fn default() -> Self {
        Theme::DEFAULT
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_canonical_names() {
        assert_eq!(ThemeName::parse("default"), Some(ThemeName::Default));
        assert_eq!(ThemeName::parse("dos"), Some(ThemeName::Dos));
        assert_eq!(ThemeName::parse("wysiwyg"), Some(ThemeName::Wysiwyg));
        assert_eq!(ThemeName::parse("amber"), Some(ThemeName::Amber));
        assert_eq!(ThemeName::parse("green"), Some(ThemeName::Green));
    }

    #[test]
    fn parse_is_case_insensitive_and_trims() {
        assert_eq!(ThemeName::parse("  DOS  "), Some(ThemeName::Dos));
        assert_eq!(ThemeName::parse("Wysiwyg"), Some(ThemeName::Wysiwyg));
        assert_eq!(ThemeName::parse("AMBER"), Some(ThemeName::Amber));
        assert_eq!(ThemeName::parse("DEFAULT"), Some(ThemeName::Default));
    }

    #[test]
    fn parse_accepts_aliases() {
        assert_eq!(ThemeName::parse("plain"), Some(ThemeName::Default));
        assert_eq!(ThemeName::parse("none"), Some(ThemeName::Default));
        assert_eq!(ThemeName::parse("classic"), Some(ThemeName::Dos));
        assert_eq!(ThemeName::parse("paper"), Some(ThemeName::Wysiwyg));
    }

    #[test]
    fn parse_rejects_unknown() {
        assert_eq!(ThemeName::parse(""), None);
        assert_eq!(ThemeName::parse("solarized"), None);
        assert_eq!(ThemeName::parse("d0s"), None);
    }

    #[test]
    fn default_theme_is_bare_terminal() {
        assert_eq!(ThemeName::default(), ThemeName::Default);
        assert_eq!(Theme::default().name, ThemeName::Default);
    }

    #[test]
    fn as_str_round_trips_through_parse() {
        for &t in ThemeName::ALL {
            assert_eq!(ThemeName::parse(t.as_str()), Some(t));
        }
    }

    #[test]
    fn list_for_error_contains_every_name() {
        let list = ThemeName::list_for_error();
        for &t in ThemeName::ALL {
            assert!(list.contains(t.as_str()), "{list} missing {}", t.as_str());
        }
    }

    #[test]
    fn each_theme_has_distinct_palette() {
        // Sanity guard so we don't accidentally ship two clones of the
        // same palette under different names — the chrome would be
        // indistinguishable to a user comparing flags.
        for (i, &a) in ThemeName::ALL.iter().enumerate() {
            for &b in &ThemeName::ALL[i + 1..] {
                assert_ne!(
                    a.palette(),
                    b.palette(),
                    "themes {a} and {b} share a palette"
                );
            }
        }
    }

    #[test]
    fn palette_name_field_matches_lookup() {
        for &t in ThemeName::ALL {
            assert_eq!(t.palette().name, t);
        }
    }

    #[test]
    fn display_writes_canonical_name() {
        assert_eq!(format!("{}", ThemeName::Wysiwyg), "wysiwyg");
        assert_eq!(format!("{}", ThemeName::Default), "default");
    }

    #[test]
    fn default_theme_leaves_field_colors_unset() {
        // The bare-terminal default theme must not paint over the
        // user's terminal background — that's the whole point of it
        // being distinct from `dos`. If any of these turn `Some(_)`,
        // the look diverges from master and the user has to opt out
        // by passing `--theme default` to get their own colors back.
        let t = Theme::DEFAULT;
        assert!(t.header_bg.is_none(), "default header_bg should be None");
        assert!(t.header_fg.is_none(), "default header_fg should be None");
        assert!(t.cell_bg.is_none(), "default cell_bg should be None");
        assert!(t.cell_fg.is_none(), "default cell_fg should be None");
    }

    #[test]
    fn dos_theme_paints_authentic_field_colors() {
        // The DOS reference screenshot shows: cyan column/row gutter
        // with black labels, black cell field, white cell text. If
        // these drift, the `--theme dos` look stops matching the
        // Lotus-123-3.0-MSDOS reference image.
        let t = Theme::DOS;
        assert_eq!(t.header_bg, Some(Rgb::new(0x00, 0xAA, 0xAA)));
        assert_eq!(t.header_fg, Some(Rgb::new(0x00, 0x00, 0x00)));
        assert_eq!(t.cell_bg, Some(Rgb::new(0x00, 0x00, 0x00)));
        assert_eq!(t.cell_fg, Some(Rgb::new(0xFF, 0xFF, 0xFF)));
    }

    #[test]
    fn wysiwyg_theme_paints_paper_field_with_teal_headers() {
        // The WYSIWYG reference screenshot shows a light-gray paper
        // field with black text, teal column/row gutter, and magenta
        // gridlines (the gridline color stays in `gridline_fg` and
        // is asserted by render-time tests).
        let t = Theme::WYSIWYG;
        assert_eq!(t.header_bg, Some(Rgb::new(0x55, 0xAA, 0xAA)));
        assert_eq!(t.header_fg, Some(Rgb::new(0x00, 0x00, 0x00)));
        assert_eq!(t.cell_bg, Some(Rgb::new(0xC0, 0xC0, 0xC0)));
        assert_eq!(t.cell_fg, Some(Rgb::new(0x00, 0x00, 0x00)));
    }
}
