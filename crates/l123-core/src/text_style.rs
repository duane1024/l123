//! Per-cell text attributes (bold, italic, underline) as shown via the
//! R3.4a WYSIWYG `:Format` commands and rendered as `{Bold}` /
//! `{Bold Italic}` on control-panel line 1.  See SPEC §20 item 11.

use std::fmt;

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub struct TextStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
}

impl TextStyle {
    pub const PLAIN: TextStyle = TextStyle {
        bold: false,
        italic: false,
        underline: false,
    };
    pub const BOLD: TextStyle = TextStyle {
        bold: true,
        italic: false,
        underline: false,
    };
    pub const ITALIC: TextStyle = TextStyle {
        bold: false,
        italic: true,
        underline: false,
    };
    pub const UNDERLINE: TextStyle = TextStyle {
        bold: false,
        italic: false,
        underline: true,
    };

    pub fn is_empty(self) -> bool {
        !self.bold && !self.italic && !self.underline
    }

    /// Bitwise union: bits set in either are set in the result.  Used when
    /// `:Format Bold Set` is applied to a cell that already has italic.
    pub fn merge(self, other: TextStyle) -> TextStyle {
        TextStyle {
            bold: self.bold || other.bold,
            italic: self.italic || other.italic,
            underline: self.underline || other.underline,
        }
    }

    /// Bitwise difference: bits set in `other` are cleared in `self`.
    /// Used for `:Format Bold Clear`.
    pub fn without(self, other: TextStyle) -> TextStyle {
        TextStyle {
            bold: self.bold && !other.bold,
            italic: self.italic && !other.italic,
            underline: self.underline && !other.underline,
        }
    }

    /// Marker string for control-panel line 1, without the surrounding
    /// braces.  Returns `None` for the empty style (no marker shown).
    pub fn marker(self) -> Option<String> {
        if self.is_empty() {
            return None;
        }
        let mut parts: Vec<&str> = Vec::with_capacity(3);
        if self.bold {
            parts.push("Bold");
        }
        if self.italic {
            parts.push("Italic");
        }
        if self.underline {
            parts.push("Underline");
        }
        Some(parts.join(" "))
    }
}

impl fmt::Display for TextStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.marker() {
            Some(s) => write!(f, "{{{s}}}"),
            None => Ok(()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_plain_and_empty() {
        let s = TextStyle::default();
        assert_eq!(s, TextStyle::PLAIN);
        assert!(s.is_empty());
    }

    #[test]
    fn constants_set_one_bit() {
        assert_eq!(
            TextStyle::BOLD,
            TextStyle {
                bold: true,
                italic: false,
                underline: false
            }
        );
        assert_eq!(
            TextStyle::ITALIC,
            TextStyle {
                bold: false,
                italic: true,
                underline: false
            }
        );
        assert_eq!(
            TextStyle::UNDERLINE,
            TextStyle {
                bold: false,
                italic: false,
                underline: true
            }
        );
        assert!(!TextStyle::BOLD.is_empty());
    }

    #[test]
    fn merge_is_bitwise_or() {
        let s = TextStyle::BOLD.merge(TextStyle::ITALIC);
        assert!(s.bold);
        assert!(s.italic);
        assert!(!s.underline);

        // Idempotent.
        let s2 = s.merge(TextStyle::BOLD);
        assert_eq!(s, s2);

        // Identity with PLAIN.
        assert_eq!(TextStyle::BOLD.merge(TextStyle::PLAIN), TextStyle::BOLD);
    }

    #[test]
    fn without_clears_only_named_bits() {
        let all = TextStyle {
            bold: true,
            italic: true,
            underline: true,
        };
        let cleared = all.without(TextStyle::BOLD);
        assert!(!cleared.bold);
        assert!(cleared.italic);
        assert!(cleared.underline);

        // Clearing a bit that isn't set is a no-op.
        assert_eq!(
            TextStyle::ITALIC.without(TextStyle::BOLD),
            TextStyle::ITALIC
        );
    }

    #[test]
    fn marker_none_for_empty() {
        assert_eq!(TextStyle::default().marker(), None);
        assert_eq!(TextStyle::PLAIN.marker(), None);
    }

    #[test]
    fn marker_single_bit() {
        assert_eq!(TextStyle::BOLD.marker().as_deref(), Some("Bold"));
        assert_eq!(TextStyle::ITALIC.marker().as_deref(), Some("Italic"));
        assert_eq!(TextStyle::UNDERLINE.marker().as_deref(), Some("Underline"));
    }

    #[test]
    fn marker_combined_is_space_joined_in_canonical_order() {
        let s = TextStyle::BOLD.merge(TextStyle::ITALIC);
        assert_eq!(s.marker().as_deref(), Some("Bold Italic"));

        let all = TextStyle {
            bold: true,
            italic: true,
            underline: true,
        };
        assert_eq!(all.marker().as_deref(), Some("Bold Italic Underline"));

        // Italic alone is still "Italic", not "Bold Italic".
        assert_eq!(TextStyle::ITALIC.marker().as_deref(), Some("Italic"));
    }

    #[test]
    fn display_wraps_marker_in_braces() {
        assert_eq!(format!("{}", TextStyle::BOLD), "{Bold}");
        assert_eq!(
            format!("{}", TextStyle::BOLD.merge(TextStyle::ITALIC)),
            "{Bold Italic}"
        );
        // Empty style renders as empty string so it can be concatenated
        // into the control-panel line unconditionally.
        assert_eq!(format!("{}", TextStyle::default()), "");
    }
}
