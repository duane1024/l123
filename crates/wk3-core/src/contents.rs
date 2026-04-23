//! Cell contents.
//!
//! SPEC §16. `Label` preserves its prefix (for round-trip and display);
//! `Constant` carries a pure value; `Formula` is introduced in M2.

use crate::{label::is_value_starter, LabelPrefix, Value};

#[derive(Clone, Debug, Default, PartialEq)]
pub enum CellContents {
    #[default]
    Empty,
    Label { prefix: LabelPrefix, text: String },
    Constant(Value),
}

impl CellContents {
    pub fn is_empty(&self) -> bool {
        matches!(self, CellContents::Empty)
    }

    /// The "source form" of a cell — the string the user would see in the
    /// control panel and would type to recreate the cell.  For M1 this
    /// doubles as the control-panel line-1 readout; once formats (e.g. `(C2)`)
    /// land, the readout gains a tag prefix while `source_form` stays clean.
    pub fn source_form(&self) -> String {
        match self {
            CellContents::Empty => String::new(),
            CellContents::Label { prefix, text } => format!("{}{}", prefix.char(), text),
            CellContents::Constant(v) => match v {
                Value::Number(n) => format_number_general(*n),
                Value::Text(s) => s.clone(),
                Value::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
                Value::Error(e) => e.lotus_tag().into(),
                Value::Empty => String::new(),
            },
        }
    }

    /// Alias retained for callers that want the control-panel-line-1 readout.
    pub fn control_panel_readout(&self) -> String {
        self.source_form()
    }

    /// Parse a source-form string into `CellContents`, applying the
    /// first-character rule (SPEC §8):
    /// - empty → Empty
    /// - leading `'`, `"`, `^`, `\` → Label with that prefix, rest as text
    /// - leading value-starter (`0-9 + - . ( @ # $`) that parses as a
    ///   number → Constant(Number)
    /// - anything else → Label with the supplied default prefix
    ///
    /// Formulas (leading `+`/`@` that don't parse as numbers) land in M2;
    /// for now they become labels, which is safe and reversible.
    pub fn from_source(s: &str, default_prefix: LabelPrefix) -> CellContents {
        let mut chars = s.chars();
        match chars.next() {
            None => CellContents::Empty,
            Some(c) if matches!(c, '\'' | '"' | '^' | '\\') => {
                let prefix = LabelPrefix::from_char(c).expect("matched above");
                CellContents::Label { prefix, text: chars.collect() }
            }
            Some(c) if is_value_starter(c) => match s.parse::<f64>() {
                Ok(n) => CellContents::Constant(Value::Number(n)),
                Err(_) => CellContents::Label {
                    prefix: default_prefix,
                    text: s.to_string(),
                },
            },
            Some(_) => CellContents::Label {
                prefix: default_prefix,
                text: s.to_string(),
            },
        }
    }
}

/// General-format number rendering. Trims trailing zeros after the point;
/// drops the point entirely for integers.
///
/// Not a full Lotus General format yet — sci fallback and `********`
/// overflow come with M2 when formats are fully wired. Enough for M1.
pub fn format_number_general(n: f64) -> String {
    if n == n.trunc() && n.is_finite() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        let mut s = format!("{n}");
        if s.contains('.') {
            while s.ends_with('0') {
                s.pop();
            }
            if s.ends_with('.') {
                s.pop();
            }
        }
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_readout_is_empty() {
        assert_eq!(CellContents::Empty.control_panel_readout(), "");
        assert!(CellContents::Empty.is_empty());
    }

    #[test]
    fn label_readout_includes_prefix() {
        let c = CellContents::Label {
            prefix: LabelPrefix::Apostrophe,
            text: "hello".into(),
        };
        assert_eq!(c.control_panel_readout(), "'hello");
        assert!(!c.is_empty());
    }

    #[test]
    fn label_readout_right_align_prefix() {
        let c = CellContents::Label {
            prefix: LabelPrefix::Quote,
            text: "right".into(),
        };
        assert_eq!(c.control_panel_readout(), "\"right");
    }

    #[test]
    fn label_readout_caret_prefix() {
        let c = CellContents::Label {
            prefix: LabelPrefix::Caret,
            text: "center".into(),
        };
        assert_eq!(c.control_panel_readout(), "^center");
    }

    #[test]
    fn label_readout_backslash_prefix() {
        let c = CellContents::Label {
            prefix: LabelPrefix::Backslash,
            text: "-".into(),
        };
        assert_eq!(c.control_panel_readout(), "\\-");
    }

    #[test]
    fn number_readout_integer() {
        let c = CellContents::Constant(Value::Number(123.0));
        assert_eq!(c.control_panel_readout(), "123");
    }

    #[test]
    fn number_readout_fraction_trims_zeros() {
        assert_eq!(format_number_general(1.5), "1.5");
        assert_eq!(format_number_general(1.50), "1.5");
        assert_eq!(format_number_general(0.0), "0");
        assert_eq!(format_number_general(-1.25), "-1.25");
    }

    #[test]
    fn number_readout_large_integer() {
        assert_eq!(format_number_general(1_000_000.0), "1000000");
    }

    #[test]
    fn from_source_empty() {
        assert_eq!(
            CellContents::from_source("", LabelPrefix::Apostrophe),
            CellContents::Empty
        );
    }

    #[test]
    fn from_source_default_label() {
        assert_eq!(
            CellContents::from_source("hello", LabelPrefix::Apostrophe),
            CellContents::Label {
                prefix: LabelPrefix::Apostrophe,
                text: "hello".into(),
            }
        );
    }

    #[test]
    fn from_source_explicit_prefix() {
        for (src, want_prefix, want_text) in [
            ("'hello", LabelPrefix::Apostrophe, "hello"),
            ("\"right", LabelPrefix::Quote, "right"),
            ("^center", LabelPrefix::Caret, "center"),
            ("\\-", LabelPrefix::Backslash, "-"),
        ] {
            let got = CellContents::from_source(src, LabelPrefix::Apostrophe);
            assert_eq!(
                got,
                CellContents::Label { prefix: want_prefix, text: want_text.into() },
                "source {src:?}"
            );
        }
    }

    #[test]
    fn from_source_number() {
        assert_eq!(
            CellContents::from_source("42", LabelPrefix::Apostrophe),
            CellContents::Constant(Value::Number(42.0))
        );
        assert_eq!(
            CellContents::from_source("-1.25", LabelPrefix::Apostrophe),
            CellContents::Constant(Value::Number(-1.25))
        );
    }

    #[test]
    fn from_source_unparseable_value_becomes_label() {
        // `+A1` starts with a value-starter but isn't a number; until M2
        // wires formulas it falls through to a default-prefix label.
        let got = CellContents::from_source("+A1", LabelPrefix::Apostrophe);
        assert!(matches!(got, CellContents::Label { .. }));
    }

    #[test]
    fn from_source_roundtrip_with_source_form() {
        for c in [
            CellContents::Empty,
            CellContents::Label {
                prefix: LabelPrefix::Apostrophe,
                text: "hello".into(),
            },
            CellContents::Label {
                prefix: LabelPrefix::Quote,
                text: "right".into(),
            },
            CellContents::Constant(Value::Number(1.25)),
        ] {
            let s = c.source_form();
            let back = CellContents::from_source(&s, LabelPrefix::Apostrophe);
            assert_eq!(back, c, "roundtrip via {s:?}");
        }
    }
}
