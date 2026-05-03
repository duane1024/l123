//! Chrome color palettes selectable via `--theme` / `theme = …`.
//!
//! A [`Theme`] is a small bundle of RGB triples used to paint
//! non-document chrome: status bar, menu/help selection highlights,
//! splash background, monochrome status overlay. Colors that come from
//! the document itself (`:Format Color`, xlsx fills/fonts, sheet tab
//! tints) flow through unchanged — those are data, not chrome.
//!
//! The palettes mirror three reference looks: classic 1-2-3 R3.4a DOS
//! (cyan/black/teal — the default), R3.4a WYSIWYG (paper, magenta
//! gridlines, blue selection), and two CRT phosphor schemes (amber and
//! green) for users who want the terminal-monitor aesthetic.

use std::fmt;

use ratatui::style::Color;

/// One of the named themes selectable from the CLI / config / env.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ThemeName {
    /// Classic 1-2-3 R3.4a DOS (cyan headers, blue menu bar). Default.
    #[default]
    Dos,
    /// 1-2-3 R3.4a WYSIWYG (light paper, magenta gridlines).
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
            ThemeName::Dos => "dos",
            ThemeName::Wysiwyg => "wysiwyg",
            ThemeName::Amber => "amber",
            ThemeName::Green => "green",
        }
    }

    /// Every theme name in CLI/help/error order.
    pub const ALL: &'static [ThemeName] = &[
        ThemeName::Dos,
        ThemeName::Wysiwyg,
        ThemeName::Amber,
        ThemeName::Green,
    ];

    /// Comma-separated list for error messages, e.g. `"dos, wysiwyg, amber, green"`.
    pub fn list_for_error() -> String {
        Self::ALL
            .iter()
            .map(|t| t.as_str())
            .collect::<Vec<_>>()
            .join(", ")
    }

    /// Resolve a theme by name. Accepts the canonical name plus a
    /// couple of obvious aliases (`default` → dos, `paper` → wysiwyg).
    /// Case-insensitive. Returns `None` for unknown names.
    pub fn parse(name: &str) -> Option<Self> {
        match name.trim().to_ascii_lowercase().as_str() {
            "dos" | "default" | "classic" => Some(ThemeName::Dos),
            "wysiwyg" | "paper" => Some(ThemeName::Wysiwyg),
            "amber" => Some(ThemeName::Amber),
            "green" => Some(ThemeName::Green),
            _ => None,
        }
    }

    /// Materialize the palette for this name.
    pub fn palette(self) -> Theme {
        match self {
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
}

impl Theme {
    /// Classic 1-2-3 R3.4a DOS — preserves the look the project
    /// shipped before themes existed.
    ///
    /// `status_fg` is a dim gray rather than DOS-cyan because the
    /// status bar paints text only (no background fill), so it has to
    /// read against whatever bg the user's terminal happens to be.
    pub const DOS: Theme = Theme {
        name: ThemeName::Dos,
        // DOS VGA cyan (#00AAAA) is the same teal the splash uses.
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
    };

    /// 1-2-3 R3.4a WYSIWYG paper-look: light grey field, magenta
    /// gridlines, blue selection.
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
    };
}

impl Default for Theme {
    fn default() -> Self {
        Theme::DOS
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_canonical_names() {
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
    }

    #[test]
    fn parse_accepts_aliases() {
        assert_eq!(ThemeName::parse("default"), Some(ThemeName::Dos));
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
    fn default_theme_is_dos() {
        assert_eq!(ThemeName::default(), ThemeName::Dos);
        assert_eq!(Theme::default().name, ThemeName::Dos);
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
    }
}
