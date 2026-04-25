//! Cell contents.
//!
//! SPEC §16. `Label` preserves its prefix (for round-trip and display);
//! `Constant` carries a pure value; `Formula` is introduced in M2.

use crate::{format::Format, label::is_value_starter, LabelPrefix, Value};

#[derive(Clone, Debug, Default, PartialEq)]
pub enum CellContents {
    #[default]
    Empty,
    Label {
        prefix: LabelPrefix,
        text: String,
    },
    Constant(Value),
    /// A formula. `expr` is the Lotus-shape source as typed by the user
    /// (e.g. `@SUM(A1..A5)`); `cached_value` is the engine's most recent
    /// evaluation, or `None` until the next recalc.
    Formula {
        expr: String,
        cached_value: Option<Value>,
    },
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
            CellContents::Constant(v) => render_value_source(v),
            CellContents::Formula { expr, .. } => expr.clone(),
        }
    }

    /// Display text for the grid cell (NOT the control-panel readout).
    /// For a formula this is the cached value (or empty before first recalc).
    pub fn display_text(&self) -> String {
        match self {
            CellContents::Empty => String::new(),
            CellContents::Label { text, .. } => text.clone(),
            CellContents::Constant(v) => render_value_source(v),
            CellContents::Formula { cached_value, .. } => match cached_value {
                Some(v) => render_value_source(v),
                None => String::new(),
            },
        }
    }

    /// Extract the cached value of a formula, or the literal value of a
    /// constant. Returns `Value::Empty` for Empty / Label / unevaluated
    /// Formula.
    pub fn value(&self) -> Value {
        match self {
            CellContents::Empty | CellContents::Label { .. } => Value::Empty,
            CellContents::Constant(v) => v.clone(),
            CellContents::Formula { cached_value, .. } => {
                cached_value.clone().unwrap_or(Value::Empty)
            }
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
        Self::from_source_with_format(s, default_prefix).0
    }

    /// Like [`Self::from_source`] but also returns the inferred display
    /// format for typed-value patterns (e.g. `$1234` → Currency 0).
    /// Returns `None` for the format when no Lotus-style markers were
    /// present, so the caller can leave any pre-existing cell format
    /// intact.
    pub fn from_source_with_format(
        s: &str,
        default_prefix: LabelPrefix,
    ) -> (CellContents, Option<Format>) {
        let mut chars = s.chars();
        match chars.next() {
            None => (CellContents::Empty, None),
            Some(c) if matches!(c, '\'' | '"' | '^' | '\\') => {
                let prefix = LabelPrefix::from_char(c).expect("matched above");
                (
                    CellContents::Label {
                        prefix,
                        text: chars.collect(),
                    },
                    None,
                )
            }
            Some(c) if is_value_starter(c) => match parse_typed_value(s) {
                Some(iv) => (CellContents::Constant(Value::Number(iv.number)), iv.format),
                None => (
                    CellContents::Formula {
                        expr: s.to_string(),
                        cached_value: None,
                    },
                    None,
                ),
            },
            Some(_) => (
                CellContents::Label {
                    prefix: default_prefix,
                    text: s.to_string(),
                },
                None,
            ),
        }
    }
}

/// Outcome of [`parse_typed_value`]: a numeric value plus the display
/// format implied by the source markers (`$`, `%`, thousands `,`).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct InferredValue {
    pub number: f64,
    /// `None` when no format-bearing markers were present (plain number).
    pub format: Option<Format>,
}

/// Parse a Lotus-style typed value: plain number, currency (`$1234`),
/// percent (`12%`), thousands-separated (`1,234`), parenthesized
/// negative (`(123)`), or any combination thereof. Returns `None` for
/// inputs that aren't numeric (formulas, garbage, empty).
pub fn parse_typed_value(s: &str) -> Option<InferredValue> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }
    let mut t = s;
    let mut negate = false;
    let mut is_currency = false;
    let mut is_percent = false;
    let mut had_commas = false;

    // Outer parens → negate (handles `(123)` and similar).
    if let Some(inner) = t.strip_prefix('(').and_then(|x| x.strip_suffix(')')) {
        negate = !negate;
        t = inner.trim();
    }

    // Leading sign.
    if let Some(rest) = t.strip_prefix('-') {
        negate = !negate;
        t = rest.trim_start();
    } else if let Some(rest) = t.strip_prefix('+') {
        t = rest.trim_start();
    }

    // Leading `$`.
    if let Some(rest) = t.strip_prefix('$') {
        is_currency = true;
        t = rest.trim_start();
    }

    // A second paren strip after `$` to handle `$(50)`.
    if let Some(inner) = t.strip_prefix('(').and_then(|x| x.strip_suffix(')')) {
        negate = !negate;
        t = inner.trim();
    }

    // Trailing `%`.
    if let Some(rest) = t.strip_suffix('%') {
        is_percent = true;
        t = rest.trim_end();
    }

    // Any remaining parens at this point are unbalanced (e.g. `(123` or
    // `123)`); reject so `f64::from_str` doesn't have to lie about it.
    if t.contains('(') || t.contains(')') {
        return None;
    }

    // Strip thousands `,` separators. Reject if `,` appears in the
    // fractional part — that's neither a number nor a Lotus shape.
    let cleaned: String = if t.contains(',') {
        had_commas = true;
        let (int_part, frac_part) = match t.find('.') {
            Some(i) => (&t[..i], &t[i..]),
            None => (t, ""),
        };
        if frac_part.contains(',') {
            return None;
        }
        let mut s = String::with_capacity(t.len());
        s.extend(int_part.chars().filter(|c| *c != ','));
        s.push_str(frac_part);
        s
    } else {
        t.to_string()
    };

    if cleaned.is_empty() {
        return None;
    }
    let mut number: f64 = cleaned.parse().ok()?;
    if is_percent {
        number /= 100.0;
    }
    if negate {
        number = -number;
    }

    let decimals = match cleaned.find('.') {
        Some(i) => (cleaned.len() - i - 1).min(15) as u8,
        None => 0,
    };

    let format = if is_currency {
        Some(Format::currency(decimals))
    } else if is_percent {
        Some(Format::percent(decimals))
    } else if had_commas {
        Some(Format::comma(decimals))
    } else {
        None
    };

    Some(InferredValue { number, format })
}

fn render_value_source(v: &Value) -> String {
    match v {
        Value::Number(n) => format_number_general(*n),
        Value::Text(s) => s.clone(),
        Value::Bool(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Value::Error(e) => e.lotus_tag().into(),
        Value::Empty => String::new(),
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
    use crate::format::Format;

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
                CellContents::Label {
                    prefix: want_prefix,
                    text: want_text.into()
                },
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
    fn from_source_non_number_value_starter_is_formula() {
        let got = CellContents::from_source("+A1", LabelPrefix::Apostrophe);
        assert!(matches!(
            got,
            CellContents::Formula { expr, cached_value: None } if expr == "+A1"
        ));
        let got = CellContents::from_source("@SUM(A1..A5)", LabelPrefix::Apostrophe);
        assert!(matches!(
            got,
            CellContents::Formula { expr, cached_value: None }
                if expr == "@SUM(A1..A5)"
        ));
    }

    #[test]
    fn formula_source_form_roundtrip() {
        let f = CellContents::Formula {
            expr: "@SUM(A1..A5)".into(),
            cached_value: Some(Value::Number(150.0)),
        };
        assert_eq!(f.source_form(), "@SUM(A1..A5)");
        assert_eq!(f.display_text(), "150");
        assert_eq!(f.value(), Value::Number(150.0));
    }

    #[test]
    fn formula_without_cache_shows_empty_grid_text() {
        let f = CellContents::Formula {
            expr: "+A1+B1".into(),
            cached_value: None,
        };
        assert_eq!(f.display_text(), "");
        assert_eq!(f.value(), Value::Empty);
    }

    #[test]
    fn parse_typed_value_plain_number() {
        let iv = parse_typed_value("1234").unwrap();
        assert_eq!(iv.number, 1234.0);
        assert_eq!(iv.format, None);

        let iv = parse_typed_value("1234.5").unwrap();
        assert_eq!(iv.number, 1234.5);
        assert_eq!(iv.format, None);

        let iv = parse_typed_value("-1234").unwrap();
        assert_eq!(iv.number, -1234.0);
        assert_eq!(iv.format, None);
    }

    #[test]
    fn parse_typed_value_currency() {
        let iv = parse_typed_value("$1234").unwrap();
        assert_eq!(iv.number, 1234.0);
        assert_eq!(iv.format, Some(Format::currency(0)));

        let iv = parse_typed_value("$1234.56").unwrap();
        assert_eq!(iv.number, 1234.56);
        assert_eq!(iv.format, Some(Format::currency(2)));

        let iv = parse_typed_value("$1,234.50").unwrap();
        assert_eq!(iv.number, 1234.5);
        assert_eq!(iv.format, Some(Format::currency(2)));
    }

    #[test]
    fn parse_typed_value_percent() {
        let iv = parse_typed_value("12%").unwrap();
        assert!((iv.number - 0.12).abs() < 1e-9);
        assert_eq!(iv.format, Some(Format::percent(0)));

        let iv = parse_typed_value("12.5%").unwrap();
        assert!((iv.number - 0.125).abs() < 1e-9);
        assert_eq!(iv.format, Some(Format::percent(1)));
    }

    #[test]
    fn parse_typed_value_comma() {
        let iv = parse_typed_value("1,234").unwrap();
        assert_eq!(iv.number, 1234.0);
        assert_eq!(iv.format, Some(Format::comma(0)));

        let iv = parse_typed_value("1,234.567").unwrap();
        assert!((iv.number - 1234.567).abs() < 1e-9);
        assert_eq!(iv.format, Some(Format::comma(3)));
    }

    #[test]
    fn parse_typed_value_paren_negate() {
        let iv = parse_typed_value("(123)").unwrap();
        assert_eq!(iv.number, -123.0);
        assert_eq!(iv.format, None);
    }

    #[test]
    fn parse_typed_value_dollar_paren_negate() {
        let iv = parse_typed_value("$(50.00)").unwrap();
        assert_eq!(iv.number, -50.0);
        assert_eq!(iv.format, Some(Format::currency(2)));
    }

    #[test]
    fn parse_typed_value_neg_dollar() {
        let iv = parse_typed_value("-$50").unwrap();
        assert_eq!(iv.number, -50.0);
        assert_eq!(iv.format, Some(Format::currency(0)));
    }

    #[test]
    fn parse_typed_value_rejects_formulas_and_garbage() {
        assert!(parse_typed_value("+A1").is_none());
        assert!(parse_typed_value("=A1+B1").is_none());
        assert!(parse_typed_value("$$5").is_none());
        assert!(parse_typed_value("1.2.3").is_none());
        assert!(parse_typed_value("(123").is_none());
        assert!(parse_typed_value("hello").is_none());
        assert!(parse_typed_value("").is_none());
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
