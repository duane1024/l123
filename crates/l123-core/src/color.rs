//! 24-bit RGB color.  Used by xlsx fill, font, and sheet-tab color.
//!
//! xlsx stores colors as ARGB hex strings (`"FFCCEEFF"`); L123 drops
//! the alpha channel since a terminal cell cannot render transparency.
//! The RGB fields stay pure data — no rendering logic lives here.

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq, Hash)]
pub struct RgbColor {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl RgbColor {
    pub const BLACK: RgbColor = RgbColor { r: 0, g: 0, b: 0 };
    pub const WHITE: RgbColor = RgbColor {
        r: 255,
        g: 255,
        b: 255,
    };

    /// Parse a hex string written in xlsx style.  Accepts `"RRGGBB"`
    /// (6 chars) and `"AARRGGBB"` (8 chars); a leading `#` is tolerated.
    /// Alpha is dropped.  Returns `None` for any other length or for
    /// non-hex characters — callers may log and fall back to a neutral
    /// default.
    pub fn from_hex(s: &str) -> Option<Self> {
        let s = s.trim();
        let s = s.strip_prefix('#').unwrap_or(s);
        let bytes = s.as_bytes();
        let (r, g, b) = match bytes.len() {
            6 => (
                u8::from_str_radix(std::str::from_utf8(&bytes[0..2]).ok()?, 16).ok()?,
                u8::from_str_radix(std::str::from_utf8(&bytes[2..4]).ok()?, 16).ok()?,
                u8::from_str_radix(std::str::from_utf8(&bytes[4..6]).ok()?, 16).ok()?,
            ),
            // ARGB — skip alpha (bytes 0..2), take R/G/B.
            8 => (
                u8::from_str_radix(std::str::from_utf8(&bytes[2..4]).ok()?, 16).ok()?,
                u8::from_str_radix(std::str::from_utf8(&bytes[4..6]).ok()?, 16).ok()?,
                u8::from_str_radix(std::str::from_utf8(&bytes[6..8]).ok()?, 16).ok()?,
            ),
            _ => return None,
        };
        Some(RgbColor { r, g, b })
    }

    /// Render back to an 8-char ARGB hex string with fully opaque alpha
    /// (`FF` prefix).  Used when a caller needs Excel's on-disk shape.
    pub fn to_argb_hex(self) -> String {
        format!("FF{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// Render to a 6-char RGB hex string, no alpha.  Used when talking
    /// to code that will prepend its own alpha byte (e.g. IronCalc's
    /// xlsx exporter, which blindly writes `FF` + whatever the caller
    /// supplied, so passing an already-ARGB value doubles the alpha).
    pub fn to_rgb_hex(self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// WCAG 2.x relative luminance in [0.0, 1.0]. Used by callers that
    /// need to pick a contrasting foreground for an arbitrary fill —
    /// see [`Self::contrasting_text`].
    pub fn relative_luminance(self) -> f64 {
        fn channel(c: u8) -> f64 {
            let v = c as f64 / 255.0;
            if v <= 0.03928 {
                v / 12.92
            } else {
                ((v + 0.055) / 1.055).powf(2.4)
            }
        }
        0.2126 * channel(self.r) + 0.7152 * channel(self.g) + 0.0722 * channel(self.b)
    }

    /// WCAG contrast ratio against `other`, in the range [1.0, 21.0].
    pub fn contrast_ratio(self, other: RgbColor) -> f64 {
        let a = self.relative_luminance();
        let b = other.relative_luminance();
        let (lighter, darker) = if a >= b { (a, b) } else { (b, a) };
        (lighter + 0.05) / (darker + 0.05)
    }

    /// Pick `BLACK` or `WHITE` — whichever has higher WCAG contrast
    /// against `self`.  Used when a cell carries an explicit fill but
    /// no font color: Excel's "automatic" font color flips with the
    /// fill's luminance, and a TUI clone needs the same behavior so
    /// imported workbooks stay legible on a dark terminal.
    pub fn contrasting_text(self) -> RgbColor {
        if self.contrast_ratio(RgbColor::BLACK) >= self.contrast_ratio(RgbColor::WHITE) {
            RgbColor::BLACK
        } else {
            RgbColor::WHITE
        }
    }

    /// Pick a contrasting text color *only when* the typical dark-terminal
    /// foreground (a light, near-white color) would wash out against
    /// `self`.  Returns `Some(BLACK)` for light fills that clash with
    /// light terminal text; returns `None` for dark fills where the
    /// terminal default already reads fine — letting the caller leave
    /// fg unset so the WYSIWYG view stays consistent with non-filled
    /// cells.
    ///
    /// Threshold: WCAG AA contrast ratio for large text / UI components
    /// (3.0).  Above that, defer to the terminal.
    pub fn auto_contrast_for_dark_terminal(self) -> Option<RgbColor> {
        const WCAG_AA_LARGE: f64 = 3.0;
        if self.contrast_ratio(RgbColor::WHITE) < WCAG_AA_LARGE {
            Some(self.contrasting_text())
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn from_hex_rgb_6_chars() {
        let c = RgbColor::from_hex("FFCCEE").unwrap();
        assert_eq!(
            c,
            RgbColor {
                r: 0xFF,
                g: 0xCC,
                b: 0xEE
            }
        );
    }

    #[test]
    fn from_hex_argb_8_chars_drops_alpha() {
        let c = RgbColor::from_hex("FFCCEEFF").unwrap();
        // AA = FF, RR = CC, GG = EE, BB = FF → alpha dropped, RGB kept.
        assert_eq!(
            c,
            RgbColor {
                r: 0xCC,
                g: 0xEE,
                b: 0xFF
            }
        );
    }

    #[test]
    fn from_hex_tolerates_hash_prefix() {
        assert_eq!(
            RgbColor::from_hex("#00ff00"),
            Some(RgbColor { r: 0, g: 255, b: 0 })
        );
    }

    #[test]
    fn from_hex_case_insensitive() {
        assert_eq!(RgbColor::from_hex("aabbcc"), RgbColor::from_hex("AABBCC"));
    }

    #[test]
    fn from_hex_rejects_invalid() {
        assert_eq!(RgbColor::from_hex(""), None);
        assert_eq!(RgbColor::from_hex("xyz"), None);
        assert_eq!(RgbColor::from_hex("12345"), None); // wrong length
        assert_eq!(RgbColor::from_hex("GGGGGG"), None); // non-hex
    }

    #[test]
    fn to_argb_hex_emits_fully_opaque() {
        assert_eq!(
            RgbColor {
                r: 0x12,
                g: 0x34,
                b: 0x56
            }
            .to_argb_hex(),
            "FF123456"
        );
    }

    #[test]
    fn to_rgb_hex_omits_alpha() {
        assert_eq!(
            RgbColor {
                r: 0x12,
                g: 0x34,
                b: 0x56
            }
            .to_rgb_hex(),
            "123456"
        );
    }

    #[test]
    fn hex_round_trip_via_argb() {
        let c = RgbColor {
            r: 0xDE,
            g: 0xAD,
            b: 0xBE,
        };
        assert_eq!(RgbColor::from_hex(&c.to_argb_hex()), Some(c));
    }

    #[test]
    fn contrasting_text_picks_black_on_light_backgrounds() {
        // Pure white → obvious black.
        assert_eq!(RgbColor::WHITE.contrasting_text(), RgbColor::BLACK);
        // Excel-yellow (#FFFF00) — the canonical case from the user's
        // screenshot: light bg, must yield black text.
        assert_eq!(
            RgbColor::from_hex("FFFF00").unwrap().contrasting_text(),
            RgbColor::BLACK
        );
        // A pale "highlight green" in the same family.
        assert_eq!(
            RgbColor::from_hex("CCFFCC").unwrap().contrasting_text(),
            RgbColor::BLACK
        );
    }

    #[test]
    fn contrasting_text_picks_white_on_dark_backgrounds() {
        // Pure black → obvious white.
        assert_eq!(RgbColor::BLACK.contrasting_text(), RgbColor::WHITE);
        // Dark red used in the user's "Op Exp" row — white text reads.
        assert_eq!(
            RgbColor::from_hex("C00000").unwrap().contrasting_text(),
            RgbColor::WHITE
        );
        // Navy blue — clearly dark.
        assert_eq!(
            RgbColor::from_hex("000080").unwrap().contrasting_text(),
            RgbColor::WHITE
        );
    }

    #[test]
    fn auto_contrast_for_dark_terminal_overrides_only_on_light_fills() {
        // Yellow / pale green / light blue clash with the dark
        // terminal's light fg; auto-contrast must kick in.
        assert_eq!(
            RgbColor::from_hex("FFFF00")
                .unwrap()
                .auto_contrast_for_dark_terminal(),
            Some(RgbColor::BLACK)
        );
        assert_eq!(
            RgbColor::from_hex("CCFFCC")
                .unwrap()
                .auto_contrast_for_dark_terminal(),
            Some(RgbColor::BLACK)
        );
        assert_eq!(
            RgbColor::from_hex("CCDDFF")
                .unwrap()
                .auto_contrast_for_dark_terminal(),
            Some(RgbColor::BLACK)
        );

        // Dark fills read fine with the terminal's light default fg
        // — return `None` so the renderer leaves fg unset and the
        // surrounding non-filled cells stay visually consistent.
        assert_eq!(
            RgbColor::BLACK.auto_contrast_for_dark_terminal(),
            None,
            "pure black bg keeps terminal default"
        );
        assert_eq!(
            RgbColor::from_hex("C00000")
                .unwrap()
                .auto_contrast_for_dark_terminal(),
            None,
            "dark red keeps terminal default"
        );
        assert_eq!(
            RgbColor::from_hex("000080")
                .unwrap()
                .auto_contrast_for_dark_terminal(),
            None,
            "navy blue keeps terminal default"
        );
    }

    #[test]
    fn contrasting_text_handles_midtone_consistently() {
        // Mid-grey #808080 sits near the perceptual coin-flip; we only
        // care that the function returns a definite color and that the
        // chosen color has *strictly higher* WCAG contrast than the
        // alternative — never picks the worse option.
        let mid = RgbColor::from_hex("808080").unwrap();
        let picked = mid.contrasting_text();
        assert!(picked == RgbColor::BLACK || picked == RgbColor::WHITE);
        let other = if picked == RgbColor::BLACK {
            RgbColor::WHITE
        } else {
            RgbColor::BLACK
        };
        assert!(
            mid.contrast_ratio(picked) >= mid.contrast_ratio(other),
            "contrasting_text should pick the color with the higher contrast \
             ratio (picked={picked:?}, other={other:?}, mid={mid:?})"
        );
    }
}
