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

use l123_core::format::{parse_date_pattern, DateFormatTable};
use l123_core::{Format, FormatKind};

/// Parse an Excel `num_fmt` string to an L123 `Format`.
///
/// Returns `None` if the string is empty, `"general"`, or otherwise
/// unrecognisable — callers treat `None` as "cell inherits General and
/// should not carry an entry in `cell_formats`."
///
/// `dates` receives any non-canonical date pattern via
/// [`DateFormatTable::intern`], so the resulting `FormatKind::DateCustom(id)`
/// can later be rendered or round-tripped back to its original Excel
/// glyphs.
pub fn parse(raw: &str, dates: &mut DateFormatTable) -> Option<Format> {
    let trimmed = raw.trim();
    if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("general") {
        return None;
    }

    // Excel splits positive;negative;zero;text — the positive section
    // dictates the kind. The other sections are purely cosmetic.
    let positive = trimmed.split(';').next().unwrap_or("");

    // Date/time detection runs first because m/d/y/h/s glyphs do not
    // appear in numeric formats — once we've matched a date or time,
    // we don't fall through to the currency/scientific classifier.
    if let Some(kind) = classify_datetime(positive) {
        return Some(Format { kind, decimals: 0 });
    }

    // Non-canonical date pattern? Intern the token list and emit a
    // `DateCustom` so the original Excel glyphs survive load and save.
    if let Some(tokens) = parse_date_pattern(positive) {
        let id = dates.intern(tokens);
        return Some(Format {
            kind: FormatKind::DateCustom(id),
            decimals: 0,
        });
    }

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
/// `DateCustom` formats consult `dates` to emit the original Excel
/// pattern verbatim.
pub fn to_num_fmt(format: Format, dates: &DateFormatTable) -> String {
    let d = format.decimals as usize;
    match format.kind {
        FormatKind::Fixed => zeros_with_decimals("0", d),
        FormatKind::Scientific => format!("{base}E+00", base = zeros_with_decimals("0", d)),
        FormatKind::Currency => format!("\"$\"{}", zeros_with_decimals("#,##0", d)),
        FormatKind::Comma => zeros_with_decimals("#,##0", d),
        FormatKind::Percent => format!("{}%", zeros_with_decimals("0", d)),
        // Date / time kinds round-trip via canonical Excel format
        // strings that classify_datetime() will recognize on re-load.
        FormatKind::DateDmy => "dd-mmm-yy".to_string(),
        FormatKind::DateDm => "dd-mmm".to_string(),
        FormatKind::DateMy => "mmm-yy".to_string(),
        FormatKind::DateLongIntl => "m/d/yy".to_string(),
        FormatKind::DateShortIntl => "m/d".to_string(),
        FormatKind::TimeLongIntl => "h:mm:ss".to_string(),
        FormatKind::TimeShortIntl => "h:mm".to_string(),
        FormatKind::DateCustom(id) => emit_date_pattern(dates.get(id)),
        // Kinds we don't yet render to Excel fall back to General so the
        // cell at least opens without an error in Excel.
        FormatKind::General
        | FormatKind::PlusMinus
        | FormatKind::TimeHmsAmPm
        | FormatKind::TimeHmAmPm
        | FormatKind::Text
        | FormatKind::Hidden
        | FormatKind::Automatic
        | FormatKind::LabelOnly
        | FormatKind::Reset => "general".to_string(),
    }
}

/// Emit a `DateCustom` token list back as an Excel num_fmt string. Letter
/// literals get quoted so they don't re-tokenize as date glyphs on load.
fn emit_date_pattern(tokens: &[l123_core::format::DateToken]) -> String {
    use l123_core::format::DateToken::*;
    let mut s = String::new();
    for t in tokens {
        match t {
            Year2 => s.push_str("yy"),
            Year4 => s.push_str("yyyy"),
            MonthNum => s.push('m'),
            MonthNumPadded => s.push_str("mm"),
            MonthAbbrev => s.push_str("mmm"),
            MonthFull => s.push_str("mmmm"),
            Day => s.push('d'),
            DayPadded => s.push_str("dd"),
            Literal(lit) => {
                if lit
                    .chars()
                    .any(|c| matches!(c, 'y' | 'Y' | 'm' | 'M' | 'd' | 'D' | 'h' | 'H' | 's' | 'S'))
                {
                    s.push('"');
                    s.push_str(lit);
                    s.push('"');
                } else {
                    s.push_str(lit);
                }
            }
        }
    }
    s
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

/// Detect a date or time format by exact-match against the canonical
/// 1-2-3 D1..D5 / D8..D9 strings (after stripping cosmetic markup).
///
/// Times still use a loose heuristic — any `h` or `s` outside markup
/// counts — because time fidelity is a follow-up. Dates are strict:
/// only the exact strings emitted by [`to_num_fmt`] map to D1..D5;
/// every other date pattern (`m/yyyy`, `dd-mmm-yyyy`, …) falls
/// through here and is handled by [`parse_date_pattern`] →
/// `FormatKind::DateCustom`, which preserves the original glyphs.
fn classify_datetime(positive: &str) -> Option<l123_core::FormatKind> {
    use l123_core::FormatKind;

    let n = strip_for_datetime_match(positive);

    // Time path: keep the existing loose heuristic so formats like
    // `h:mm AM/PM` still classify as time. (Time fidelity is a separate
    // task.)
    if n.contains('h') || n.contains('s') {
        return Some(if n.contains('s') {
            FormatKind::TimeLongIntl
        } else {
            FormatKind::TimeShortIntl
        });
    }

    // Date path: exact match against canonical strings only.
    match n.as_str() {
        "dd-mmm-yy" => Some(FormatKind::DateDmy),
        "dd-mmm" => Some(FormatKind::DateDm),
        "mmm-yy" => Some(FormatKind::DateMy),
        "m/d/yy" => Some(FormatKind::DateLongIntl),
        "m/d" => Some(FormatKind::DateShortIntl),
        _ => None,
    }
}

/// Strip cosmetic markup from a format string, lowercasing the rest, so
/// `[$-409]M/D/YY` matches `m/d/yy`. Quoted literals, backslash escapes,
/// `_`/`*` spacers and `[...]` tags are all dropped.
fn strip_for_datetime_match(positive: &str) -> String {
    let mut out = String::new();
    let mut chars = positive.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for inner in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                }
            }
            '\\' => {
                chars.next();
            }
            '_' | '*' => {
                chars.next();
            }
            '[' => {
                for inner in chars.by_ref() {
                    if inner == ']' {
                        break;
                    }
                }
            }
            _ => out.push(c.to_ascii_lowercase()),
        }
    }
    out
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
    use l123_core::format::DateToken;

    fn p(s: &str) -> Option<Format> {
        parse(s, &mut DateFormatTable::new())
    }

    fn t(f: Format) -> String {
        to_num_fmt(f, &DateFormatTable::new())
    }

    #[test]
    fn general_and_empty_return_none() {
        assert_eq!(p(""), None);
        assert_eq!(p("general"), None);
        assert_eq!(p("General"), None);
        assert_eq!(p("  General  "), None);
    }

    #[test]
    fn fixed_decimals() {
        assert_eq!(p("0"), Some(Format::fixed(0)));
        assert_eq!(p("0.00"), Some(Format::fixed(2)));
        assert_eq!(p("0.000"), Some(Format::fixed(3)));
    }

    #[test]
    fn comma_with_thousands() {
        assert_eq!(p("#,##0"), Some(Format::comma(0)));
        assert_eq!(p("#,##0.00"), Some(Format::comma(2)));
    }

    #[test]
    fn percent() {
        assert_eq!(p("0%"), Some(Format::percent(0)));
        assert_eq!(p("0.00%"), Some(Format::percent(2)));
        assert_eq!(p("0.0%"), Some(Format::percent(1)));
    }

    #[test]
    fn scientific() {
        assert_eq!(
            p("0.00E+00"),
            Some(Format {
                kind: FormatKind::Scientific,
                decimals: 2
            })
        );
        assert_eq!(
            p("0E+00"),
            Some(Format {
                kind: FormatKind::Scientific,
                decimals: 0
            })
        );
    }

    #[test]
    fn currency_with_dollar_sign() {
        assert_eq!(p("$#,##0.00"), Some(Format::currency(2)));
        assert_eq!(p("$#,##0"), Some(Format::currency(0)));
        // Quote-wrapped dollar sign.
        assert_eq!(p("\"$\"#,##0.00"), Some(Format::currency(2)));
    }

    #[test]
    fn currency_excel_accounting_builtin_44() {
        // Built-in numFmtId 44 resolves to this string.
        let fmt = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
        assert_eq!(p(fmt), Some(Format::currency(2)));
    }

    #[test]
    fn currency_with_locale_tag() {
        // Common form emitted by localised Excel.
        assert_eq!(
            p("[$$-409]#,##0.00"),
            Some(Format::currency(2)),
            "bracket tag should be stripped"
        );
    }

    #[test]
    fn negative_section_ignored() {
        // Positive;negative;zero;text — kind is decided by the first section.
        assert_eq!(p("$#,##0.00_);($#,##0.00)"), Some(Format::currency(2)));
    }

    #[test]
    fn canonical_date_my_for_mmm_yy_only() {
        // Only `mmm-yy` (the to_num_fmt output for DateMy) maps to the
        // 1-2-3 enum. The other month/year shapes preserve fidelity via
        // DateCustom — covered by the date_custom_* tests below.
        assert_eq!(
            p("mmm-yy"),
            Some(Format {
                kind: FormatKind::DateMy,
                decimals: 0
            })
        );
    }

    #[test]
    fn canonical_date_dmy_for_dd_mmm_yy_only() {
        assert_eq!(
            p("dd-mmm-yy"),
            Some(Format {
                kind: FormatKind::DateDmy,
                decimals: 0
            })
        );
    }

    #[test]
    fn canonical_date_dm_for_dd_mmm_only() {
        assert_eq!(
            p("dd-mmm"),
            Some(Format {
                kind: FormatKind::DateDm,
                decimals: 0
            })
        );
    }

    #[test]
    fn canonical_date_long_intl_for_m_d_yy_only() {
        assert_eq!(
            p("m/d/yy"),
            Some(Format {
                kind: FormatKind::DateLongIntl,
                decimals: 0
            })
        );
    }

    #[test]
    fn canonical_date_short_intl_for_m_d_only() {
        assert_eq!(
            p("m/d"),
            Some(Format {
                kind: FormatKind::DateShortIntl,
                decimals: 0
            })
        );
    }

    #[test]
    fn canonical_form_matches_case_insensitively() {
        // Excel may emit upper- or lower-case glyphs. Case folds during
        // canonical lookup so DD-MMM-YY still maps to DateDmy.
        assert_eq!(
            p("DD-MMM-YY"),
            Some(Format {
                kind: FormatKind::DateDmy,
                decimals: 0
            })
        );
    }

    #[test]
    fn time_long_intl_for_h_m_s() {
        assert_eq!(
            p("h:mm:ss"),
            Some(Format {
                kind: FormatKind::TimeLongIntl,
                decimals: 0
            })
        );
    }

    #[test]
    fn time_short_intl_for_h_m_no_seconds() {
        assert_eq!(
            p("h:mm"),
            Some(Format {
                kind: FormatKind::TimeShortIntl,
                decimals: 0
            })
        );
    }

    #[test]
    fn quoted_letters_do_not_trigger_date_detection() {
        // Quoted "m" is a literal, not a month token. Should fall through
        // to fixed since the rest is `0`.
        assert_eq!(p("\"month\" 0"), Some(Format::fixed(0)));
    }

    #[test]
    fn date_custom_for_m_yyyy_preserves_tokens() {
        let mut tbl = DateFormatTable::new();
        let f = parse("m/yyyy", &mut tbl).unwrap();
        let FormatKind::DateCustom(id) = f.kind else {
            panic!("expected DateCustom, got {:?}", f.kind);
        };
        assert_eq!(
            tbl.get(id),
            &[
                DateToken::MonthNum,
                DateToken::Literal("/".into()),
                DateToken::Year4,
            ]
        );
    }

    #[test]
    fn date_custom_for_mmm_yyyy_preserves_tokens() {
        let mut tbl = DateFormatTable::new();
        let f = parse("mmm-yyyy", &mut tbl).unwrap();
        let FormatKind::DateCustom(id) = f.kind else {
            panic!("expected DateCustom, got {:?}", f.kind);
        };
        assert_eq!(
            tbl.get(id),
            &[
                DateToken::MonthAbbrev,
                DateToken::Literal("-".into()),
                DateToken::Year4,
            ]
        );
    }

    #[test]
    fn date_custom_for_dd_mmm_yyyy_preserves_tokens() {
        let mut tbl = DateFormatTable::new();
        let f = parse("dd-mmm-yyyy", &mut tbl).unwrap();
        let FormatKind::DateCustom(id) = f.kind else {
            panic!("expected DateCustom, got {:?}", f.kind);
        };
        assert_eq!(
            tbl.get(id),
            &[
                DateToken::DayPadded,
                DateToken::Literal("-".into()),
                DateToken::MonthAbbrev,
                DateToken::Literal("-".into()),
                DateToken::Year4,
            ]
        );
    }

    #[test]
    fn date_custom_for_m_d_yyyy_preserves_tokens() {
        let mut tbl = DateFormatTable::new();
        let f = parse("m/d/yyyy", &mut tbl).unwrap();
        let FormatKind::DateCustom(id) = f.kind else {
            panic!("expected DateCustom, got {:?}", f.kind);
        };
        assert_eq!(
            tbl.get(id),
            &[
                DateToken::MonthNum,
                DateToken::Literal("/".into()),
                DateToken::Day,
                DateToken::Literal("/".into()),
                DateToken::Year4,
            ]
        );
    }

    #[test]
    fn date_custom_dedupes_repeat_pattern_into_same_id() {
        let mut tbl = DateFormatTable::new();
        let a = parse("m/yyyy", &mut tbl).unwrap();
        let b = parse("m/yyyy", &mut tbl).unwrap();
        assert_eq!(a, b);
        assert_eq!(tbl.len(), 1);
    }

    #[test]
    fn date_custom_round_trips_through_to_num_fmt() {
        for raw in ["m/yyyy", "mmm-yyyy", "mmm yyyy", "dd-mmm-yyyy", "m/d/yyyy"] {
            let mut tbl = DateFormatTable::new();
            let f = parse(raw, &mut tbl).expect(raw);
            let emitted = to_num_fmt(f, &tbl);
            assert_eq!(emitted, raw, "round-trip preserves Excel pattern");
            // Re-parsing the emitted string into a fresh table yields
            // the same tokens (interned at id 0 in both tables).
            let mut tbl2 = DateFormatTable::new();
            let g = parse(&emitted, &mut tbl2).expect(raw);
            let (FormatKind::DateCustom(id_a), FormatKind::DateCustom(id_b)) = (f.kind, g.kind)
            else {
                panic!("expected DateCustom on both sides");
            };
            assert_eq!(tbl.get(id_a), tbl2.get(id_b));
        }
    }

    #[test]
    fn date_round_trips_via_to_num_fmt() {
        for kind in [
            FormatKind::DateDmy,
            FormatKind::DateDm,
            FormatKind::DateMy,
            FormatKind::DateLongIntl,
            FormatKind::DateShortIntl,
            FormatKind::TimeLongIntl,
            FormatKind::TimeShortIntl,
        ] {
            let f = Format { kind, decimals: 0 };
            let s = t(f);
            assert_eq!(p(&s), Some(f), "round-trip for {f:?} via {s:?}");
        }
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
            let s = t(f);
            assert_eq!(p(&s), Some(f), "round-trip for {f:?} via {s:?}");
        }
    }

    #[test]
    fn to_num_fmt_general_is_string_general() {
        assert_eq!(t(Format::GENERAL), "general");
    }

    #[test]
    fn to_num_fmt_scientific_round_trips() {
        let f = Format {
            kind: FormatKind::Scientific,
            decimals: 2,
        };
        assert_eq!(p(&t(f)), Some(f));
    }
}
