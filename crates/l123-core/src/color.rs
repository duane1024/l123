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
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn from_hex_rgb_6_chars() {
        let c = RgbColor::from_hex("FFCCEE").unwrap();
        assert_eq!(c, RgbColor { r: 0xFF, g: 0xCC, b: 0xEE });
    }

    #[test]
    fn from_hex_argb_8_chars_drops_alpha() {
        let c = RgbColor::from_hex("FFCCEEFF").unwrap();
        // AA = FF, RR = CC, GG = EE, BB = FF → alpha dropped, RGB kept.
        assert_eq!(c, RgbColor { r: 0xCC, g: 0xEE, b: 0xFF });
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
            RgbColor { r: 0x12, g: 0x34, b: 0x56 }.to_argb_hex(),
            "FF123456"
        );
    }

    #[test]
    fn to_rgb_hex_omits_alpha() {
        assert_eq!(
            RgbColor { r: 0x12, g: 0x34, b: 0x56 }.to_rgb_hex(),
            "123456"
        );
    }

    #[test]
    fn hex_round_trip_via_argb() {
        let c = RgbColor { r: 0xDE, g: 0xAD, b: 0xBE };
        assert_eq!(RgbColor::from_hex(&c.to_argb_hex()), Some(c));
    }
}
