//! Excel `num_fmt` string ↔ `l123_core::Format` translation.
//!
//! The xlsx style table stores per-cell number formats as strings
//! (`"$#,##0.00"`, `"0.00%"`, `"0.00E+00"`, …). L123's `Format` type is
//! a tagged kind + decimal count. This module converts between the two
//! so `/File Retrieve` preserves formats set in Excel and `/File Save`
//! round-trips formats set in l123.
//!
//! The parser is pattern-based rather than a full strptime-style scan:
//! it normalises the positive section (sections are separated by `;`),
//! strips Excel's literal-quoting and spacer escapes, and then
//! classifies by the surviving glyphs (`$`, `%`, `E`, `#,##0`, `0`).
//! Anything it can't classify maps to `General` — which matches the
//! current pre-fix behaviour and is harmless.
//!
//! Rendering to xlsx goes the other way: each `FormatKind` has a
//! canonical num_fmt string the adapter writes back.

use l123_core::{Format, FormatKind};

/// Parse an Excel `num_fmt` string to an L123 `Format`.
///
/// Returns `None` if the string is empty, `"general"`, or otherwise
/// unrecognisable — callers treat `None` as "cell inherits General and
/// should not carry an entry in `cell_formats`."
pub fn parse(raw: &str) -> Option<Format> {
    let trimmed = raw.trim();
    if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("general") {
        return None;
    }

    // Excel splits positive;negative;zero;text — the positive section
    // dictates the kind. The other sections are purely cosmetic.
    let positive = trimmed.split(';').next().unwrap_or("");
    let stripped = strip_cosmetic(positive);
    if stripped.is_empty() {
        return None;
    }

    let decimals = count_decimals(&stripped);

    if stripped.contains('%') {
        return Some(Format {
            kind: FormatKind::Percent,
            decimals,
        });
    }
    if stripped.contains('E') || stripped.contains('e') {
        return Some(Format {
            kind: FormatKind::Scientific,
            decimals,
        });
    }
    if stripped.contains('$')
        || stripped.contains('£')
        || stripped.contains('€')
        || stripped.contains('¥')
    {
        return Some(Format {
            kind: FormatKind::Currency,
            decimals,
        });
    }
    if stripped.contains('#') || stripped.contains(',') {
        return Some(Format {
            kind: FormatKind::Comma,
            decimals,
        });
    }
    if stripped.contains('0') {
        return Some(Format {
            kind: FormatKind::Fixed,
            decimals,
        });
    }

    None
}

/// Render an L123 `Format` as an Excel `num_fmt` string.
///
/// For `Format::GENERAL` (and other inherit-from-default formats) returns
/// `"general"` so callers can round-trip by writing the default string.
pub fn to_num_fmt(format: Format) -> String {
    let d = format.decimals as usize;
    match format.kind {
        FormatKind::Fixed => zeros_with_decimals("0", d),
        FormatKind::Scientific => format!("{base}E+00", base = zeros_with_decimals("0", d)),
        FormatKind::Currency => format!("\"$\"{}", zeros_with_decimals("#,##0", d)),
        FormatKind::Comma => zeros_with_decimals("#,##0", d),
        FormatKind::Percent => format!("{}%", zeros_with_decimals("0", d)),
        // Kinds we don't yet render to Excel fall back to General so the
        // cell at least opens without an error in Excel.
        FormatKind::General
        | FormatKind::PlusMinus
        | FormatKind::DateDmy
        | FormatKind::DateDm
        | FormatKind::DateMy
        | FormatKind::DateLongIntl
        | FormatKind::DateShortIntl
        | FormatKind::TimeHmsAmPm
        | FormatKind::TimeHmAmPm
        | FormatKind::TimeLongIntl
        | FormatKind::TimeShortIntl
        | FormatKind::Text
        | FormatKind::Hidden
        | FormatKind::Automatic
        | FormatKind::LabelOnly
        | FormatKind::Reset => "general".to_string(),
    }
}

fn zeros_with_decimals(integer_pattern: &str, decimals: usize) -> String {
    if decimals == 0 {
        integer_pattern.to_string()
    } else {
        let mut s = String::with_capacity(integer_pattern.len() + 1 + decimals);
        s.push_str(integer_pattern);
        s.push('.');
        for _ in 0..decimals {
            s.push('0');
        }
        s
    }
}

/// Remove Excel's literal-quote, escape, spacer, and color-tag markup
/// so the classifier sees only the format glyphs that matter.
fn strip_cosmetic(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            // "literal" — keep the inner chars so `"$"` still contributes `$`.
            '"' => {
                for inner in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                    out.push(inner);
                }
            }
            // Backslash-escape — keep the escaped char.
            '\\' => {
                if let Some(next) = chars.next() {
                    out.push(next);
                }
            }
            // `_X` is a width-spacer: consume the next char and drop it.
            '_' => {
                chars.next();
            }
            // `*X` is a fill-repeat: consume the next char and drop it.
            '*' => {
                chars.next();
            }
            // Bracketed tags. `[Red]` / `[h]` are cosmetic — skip.
            // `[$<symbol>-<locale>]` carries the currency symbol we
            // need for classification; emit the symbol portion.
            '[' => {
                let mut inside = String::new();
                for inner in chars.by_ref() {
                    if inner == ']' {
                        break;
                    }
                    inside.push(inner);
                }
                if let Some(rest) = inside.strip_prefix('$') {
                    // Symbol runs until `-` (locale code) or end.
                    let sym = rest.split('-').next().unwrap_or(rest);
                    // `[$-409]` — empty symbol, locale only. Emit `$`
                    // as a sentinel so currency is still detected only
                    // when a real symbol was present.
                    if sym.is_empty() {
                        // locale-only tag, no symbol: drop entirely.
                    } else {
                        out.push_str(sym);
                    }
                }
            }
            _ => out.push(c),
        }
    }
    out
}

/// Count the `0`s immediately following the first `.` in a stripped
/// format string — that's the decimal-place count. Stops at the first
/// non-digit so `"0.00E+00"` yields 2, not 4.
fn count_decimals(stripped: &str) -> u8 {
    let Some(dot) = stripped.find('.') else {
        return 0;
    };
    let mut n: u8 = 0;
    for c in stripped[dot + 1..].chars() {
        if c == '0' || c == '#' {
            n = n.saturating_add(1);
        } else {
            break;
        }
    }
    n.min(15)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn general_and_empty_return_none() {
        assert_eq!(parse(""), None);
        assert_eq!(parse("general"), None);
        assert_eq!(parse("General"), None);
        assert_eq!(parse("  General  "), None);
    }

    #[test]
    fn fixed_decimals() {
        assert_eq!(parse("0"), Some(Format::fixed(0)));
        assert_eq!(parse("0.00"), Some(Format::fixed(2)));
        assert_eq!(parse("0.000"), Some(Format::fixed(3)));
    }

    #[test]
    fn comma_with_thousands() {
        assert_eq!(parse("#,##0"), Some(Format::comma(0)));
        assert_eq!(parse("#,##0.00"), Some(Format::comma(2)));
    }

    #[test]
    fn percent() {
        assert_eq!(parse("0%"), Some(Format::percent(0)));
        assert_eq!(parse("0.00%"), Some(Format::percent(2)));
        assert_eq!(parse("0.0%"), Some(Format::percent(1)));
    }

    #[test]
    fn scientific() {
        assert_eq!(
            parse("0.00E+00"),
            Some(Format {
                kind: FormatKind::Scientific,
                decimals: 2
            })
        );
        assert_eq!(
            parse("0E+00"),
            Some(Format {
                kind: FormatKind::Scientific,
                decimals: 0
            })
        );
    }

    #[test]
    fn currency_with_dollar_sign() {
        assert_eq!(parse("$#,##0.00"), Some(Format::currency(2)));
        assert_eq!(parse("$#,##0"), Some(Format::currency(0)));
        // Quote-wrapped dollar sign.
        assert_eq!(parse("\"$\"#,##0.00"), Some(Format::currency(2)));
    }

    #[test]
    fn currency_excel_accounting_builtin_44() {
        // Built-in numFmtId 44 resolves to this string.
        let fmt = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
        assert_eq!(parse(fmt), Some(Format::currency(2)));
    }

    #[test]
    fn currency_with_locale_tag() {
        // Common form emitted by localised Excel.
        assert_eq!(
            parse("[$$-409]#,##0.00"),
            Some(Format::currency(2)),
            "bracket tag should be stripped"
        );
    }

    #[test]
    fn negative_section_ignored() {
        // Positive;negative;zero;text — kind is decided by the first section.
        assert_eq!(parse("$#,##0.00_);($#,##0.00)"), Some(Format::currency(2)));
    }

    #[test]
    fn to_num_fmt_round_trips_common_kinds() {
        for f in [
            Format::fixed(0),
            Format::fixed(2),
            Format::comma(0),
            Format::comma(2),
            Format::percent(0),
            Format::percent(2),
            Format::currency(0),
            Format::currency(2),
        ] {
            let s = to_num_fmt(f);
            assert_eq!(parse(&s), Some(f), "round-trip for {f:?} via {s:?}");
        }
    }

    #[test]
    fn to_num_fmt_general_is_string_general() {
        assert_eq!(to_num_fmt(Format::GENERAL), "general");
    }

    #[test]
    fn to_num_fmt_scientific_round_trips() {
        let f = Format {
            kind: FormatKind::Scientific,
            decimals: 2,
        };
        assert_eq!(parse(&to_num_fmt(f)), Some(f));
    }
}
