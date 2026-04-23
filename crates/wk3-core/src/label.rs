//! Label prefixes.

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum LabelPrefix {
    /// `'` — left-align (default).
    #[default]
    Apostrophe,
    /// `"` — right-align.
    Quote,
    /// `^` — center.
    Caret,
    /// `\` — repeat label across cell width (`\-` → dashes).
    Backslash,
    /// `|` — non-print row (first column only).
    Pipe,
}

impl LabelPrefix {
    pub fn char(self) -> char {
        match self {
            LabelPrefix::Apostrophe => '\'',
            LabelPrefix::Quote => '"',
            LabelPrefix::Caret => '^',
            LabelPrefix::Backslash => '\\',
            LabelPrefix::Pipe => '|',
        }
    }

    pub fn from_char(c: char) -> Option<Self> {
        Some(match c {
            '\'' => LabelPrefix::Apostrophe,
            '"' => LabelPrefix::Quote,
            '^' => LabelPrefix::Caret,
            '\\' => LabelPrefix::Backslash,
            '|' => LabelPrefix::Pipe,
            _ => return None,
        })
    }
}

/// Returns true if the first typed character puts the entry into VALUE mode.
/// Otherwise the entry is a LABEL (with an auto-inserted default prefix).
///
/// Value starters: `0-9 + - . ( @ # $`.
pub fn is_value_starter(c: char) -> bool {
    matches!(c, '0'..='9' | '+' | '-' | '.' | '(' | '@' | '#' | '$')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn value_starters() {
        for c in "0123456789+-.(@#$".chars() {
            assert!(is_value_starter(c), "{c} should be a value starter");
        }
        for c in "abcABC'\"^\\| ".chars() {
            assert!(!is_value_starter(c), "{c} should not be a value starter");
        }
    }

    #[test]
    fn label_prefix_roundtrip() {
        for p in [
            LabelPrefix::Apostrophe,
            LabelPrefix::Quote,
            LabelPrefix::Caret,
            LabelPrefix::Backslash,
            LabelPrefix::Pipe,
        ] {
            assert_eq!(LabelPrefix::from_char(p.char()), Some(p));
        }
    }
}
