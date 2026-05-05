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

/// True for the `FormatKind` variants that represent a date or time
/// rendering (D1..D9). The override pipeline only stores raw Excel
/// format strings for these; numeric formats stay on the canonical
/// kind path.
pub fn is_date_or_time_kind(kind: FormatKind) -> bool {
    use FormatKind::*;
    matches!(
        kind,
        DateDmy
            | DateDm
            | DateMy
            | DateLongIntl
            | DateShortIntl
            | TimeHmsAmPm
            | TimeHmAmPm
            | TimeLongIntl
            | TimeShortIntl
    )
}

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

    // Date/time detection runs first because m/d/y/h/s glyphs do not
    // appear in numeric formats — once we've matched a date or time,
    // we don't fall through to the currency/scientific classifier.
    if let Some(kind) = classify_datetime(positive) {
        return Some(Format {
            kind,
            decimals: 0,
            parens: false,
            negative_color: None,
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
            parens: false,
            negative_color: None,
        });
    }
    if stripped.contains('E') || stripped.contains('e') {
        return Some(Format {
            kind: FormatKind::Scientific,
            decimals,
            parens: false,
            negative_color: None,
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
            parens: false,
            negative_color: None,
        });
    }
    if stripped.contains('#') || stripped.contains(',') {
        return Some(Format {
            kind: FormatKind::Comma,
            decimals,
            parens: false,
            negative_color: None,
        });
    }
    if stripped.contains('0') {
        return Some(Format {
            kind: FormatKind::Fixed,
            decimals,
            parens: false,
            negative_color: None,
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
        // Date / time kinds round-trip via canonical Excel format
        // strings that classify_datetime() will recognize on re-load.
        FormatKind::DateDmy => "dd-mmm-yy".to_string(),
        FormatKind::DateDm => "dd-mmm".to_string(),
        FormatKind::DateMy => "mmm-yy".to_string(),
        FormatKind::DateLongIntl => "m/d/yy".to_string(),
        FormatKind::DateShortIntl => "m/d".to_string(),
        FormatKind::TimeLongIntl => "h:mm:ss".to_string(),
        FormatKind::TimeShortIntl => "h:mm".to_string(),
        FormatKind::TimeHmsAmPm => "h:mm:ss AM/PM".to_string(),
        FormatKind::TimeHmAmPm => "h:mm AM/PM".to_string(),
        // Kinds we don't yet render to Excel fall back to General so the
        // cell at least opens without an error in Excel.
        FormatKind::General
        | FormatKind::PlusMinus
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

/// Detect a date or time format by scanning unquoted `m`, `d`, `y`,
/// `h`, `s` glyphs in the positive section.
///
/// 1-2-3 has five date kinds (D1..D5) and four time kinds (D6..D9).
/// Excel's format strings carry more variation than that — we collapse:
///
/// * Any format containing `h` or `s` → **time**
///   (`TimeLongIntl` if `s` is present, else `TimeShortIntl`).
/// * Date with year + month + day:
///   * letter month (`mmm`/`mmmm`) → `DateDmy` (D1)
///   * numeric month → `DateLongIntl` (D4)
/// * Date with month + day, no year:
///   * letter month → `DateDm` (D2)
///   * numeric month → `DateShortIntl` (D5)
/// * Date with month + year, no day → `DateMy` (D3)
fn classify_datetime(positive: &str) -> Option<l123_core::FormatKind> {
    use l123_core::FormatKind;

    let mut has_y = false;
    let mut has_d = false;
    let mut has_h = false;
    let mut has_s = false;
    let mut has_m = false;
    let mut has_mmm = false;
    let mut has_ampm = false;

    let bytes = positive.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let c = bytes[i] as char;
        match c {
            // Quoted literal — skip the inner chars entirely so a
            // literal "moon" doesn't trigger month detection.
            '"' => {
                i += 1;
                while i < bytes.len() && bytes[i] != b'"' {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            // Backslash-escape consumes the next char literally.
            '\\' => {
                i += 2;
            }
            // Width-spacer / fill-repeat — drop the next char.
            '_' | '*' => {
                i += 2;
            }
            // `[Red]`, `[h]`, `[$-409]` — cosmetic / locale tags. We
            // skip the entire bracketed run; the contents do not count
            // as date glyphs.
            '[' => {
                i += 1;
                while i < bytes.len() && bytes[i] != b']' {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            'm' | 'M' => {
                let mut run = 1;
                while i + run < bytes.len() && matches!(bytes[i + run], b'm' | b'M') {
                    run += 1;
                }
                has_m = true;
                if run >= 3 {
                    has_mmm = true;
                }
                i += run;
            }
            'd' | 'D' => {
                has_d = true;
                i += 1;
            }
            'y' | 'Y' => {
                has_y = true;
                i += 1;
            }
            'h' | 'H' => {
                has_h = true;
                i += 1;
            }
            's' | 'S' => {
                has_s = true;
                i += 1;
            }
            // `AM/PM` (or `am/pm`) and `A/P` (or `a/p`) — literal AM/PM
            // markers Excel uses to flag a 12-hour time format.
            'A' | 'a' => {
                let lower = positive[i..].to_ascii_lowercase();
                if lower.starts_with("am/pm") {
                    has_ampm = true;
                    i += 5;
                } else if lower.starts_with("a/p") {
                    has_ampm = true;
                    i += 3;
                } else {
                    i += 1;
                }
            }
            _ => {
                i += 1;
            }
        }
    }

    // Time wins over date. `[h]:mm:ss` (elapsed time) hides `h` inside
    // brackets — `s` alone is enough to mark it as time. AM/PM marker
    // selects the 12-hour kinds (D6/D7) over the international ones.
    if has_h || has_s {
        return Some(match (has_s, has_ampm) {
            (true, true) => FormatKind::TimeHmsAmPm,
            (false, true) => FormatKind::TimeHmAmPm,
            (true, false) => FormatKind::TimeLongIntl,
            (false, false) => FormatKind::TimeShortIntl,
        });
    }

    match (has_y, has_m, has_d) {
        (true, true, true) => Some(if has_mmm {
            FormatKind::DateDmy
        } else {
            FormatKind::DateLongIntl
        }),
        (false, true, true) => Some(if has_mmm {
            FormatKind::DateDm
        } else {
            FormatKind::DateShortIntl
        }),
        (true, true, false) => Some(FormatKind::DateMy),
        // Day-only (`d`, `dd`, `ddd`, `dddd`) and year-only (`yyyy`)
        // are valid date formats too — collapse to a sensible D-letter
        // for the canonical-tag fallback. The actual rendering goes
        // through the override path because `to_num_fmt(D5)` = "m/d"
        // doesn't match the original.
        (false, false, true) => Some(FormatKind::DateShortIntl),
        (true, false, false) => Some(FormatKind::DateMy),
        _ => None,
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
                decimals: 2,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("0E+00"),
            Some(Format {
                kind: FormatKind::Scientific,
                decimals: 0,
                parens: false,
                negative_color: None,
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
    fn date_my_for_month_year_formats() {
        // The atlas-model.xlsx fixture uses these three forms.
        assert_eq!(
            parse("m/yyyy"),
            Some(Format {
                kind: FormatKind::DateMy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("mmm-yyyy"),
            Some(Format {
                kind: FormatKind::DateMy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("mmm yyyy"),
            Some(Format {
                kind: FormatKind::DateMy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("mmm-yy"),
            Some(Format {
                kind: FormatKind::DateMy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn date_dmy_for_letter_month_full_dates() {
        assert_eq!(
            parse("dd-mmm-yy"),
            Some(Format {
                kind: FormatKind::DateDmy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("d-mmm-yyyy"),
            Some(Format {
                kind: FormatKind::DateDmy,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn date_dm_for_letter_month_no_year() {
        assert_eq!(
            parse("d-mmm"),
            Some(Format {
                kind: FormatKind::DateDm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("dd-mmm"),
            Some(Format {
                kind: FormatKind::DateDm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn date_long_intl_for_numeric_full_dates() {
        // Built-in numFmtId 14 = "m/d/yyyy"; 22 = "m/d/yyyy h:mm" (handled
        // separately as time). All-numeric dates map to D4.
        assert_eq!(
            parse("m/d/yyyy"),
            Some(Format {
                kind: FormatKind::DateLongIntl,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("m/d/yy"),
            Some(Format {
                kind: FormatKind::DateLongIntl,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn date_short_intl_for_numeric_no_year() {
        assert_eq!(
            parse("m/d"),
            Some(Format {
                kind: FormatKind::DateShortIntl,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn time_long_intl_for_h_m_s() {
        assert_eq!(
            parse("h:mm:ss"),
            Some(Format {
                kind: FormatKind::TimeLongIntl,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn time_short_intl_for_h_m_no_seconds() {
        assert_eq!(
            parse("h:mm"),
            Some(Format {
                kind: FormatKind::TimeShortIntl,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn time_hms_ampm_for_h_m_s_with_ampm_marker() {
        assert_eq!(
            parse("h:mm:ss AM/PM"),
            Some(Format {
                kind: FormatKind::TimeHmsAmPm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        // Lowercase variant.
        assert_eq!(
            parse("h:mm:ss am/pm"),
            Some(Format {
                kind: FormatKind::TimeHmsAmPm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        // Single-letter A/P short form Excel also accepts.
        assert_eq!(
            parse("h:mm:ss A/P"),
            Some(Format {
                kind: FormatKind::TimeHmsAmPm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn time_hm_ampm_for_h_m_with_ampm_marker() {
        assert_eq!(
            parse("h:mm AM/PM"),
            Some(Format {
                kind: FormatKind::TimeHmAmPm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
        assert_eq!(
            parse("h:mm am/pm"),
            Some(Format {
                kind: FormatKind::TimeHmAmPm,
                decimals: 0,
                parens: false,
                negative_color: None,
            })
        );
    }

    #[test]
    fn time_ampm_round_trips_via_to_num_fmt() {
        for kind in [FormatKind::TimeHmsAmPm, FormatKind::TimeHmAmPm] {
            let f = Format {
                kind,
                decimals: 0,
                parens: false,
                negative_color: None,
            };
            let s = to_num_fmt(f);
            assert_eq!(parse(&s), Some(f), "round-trip for {f:?} via {s:?}");
        }
    }

    #[test]
    fn quoted_letters_do_not_trigger_date_detection() {
        // Quoted "m" is a literal, not a month token. Should fall through
        // to None (or, here, to currency since the format also has $).
        assert_eq!(parse("\"month\" 0"), Some(Format::fixed(0)));
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
            let f = Format {
                kind,
                decimals: 0,
                parens: false,
                negative_color: None,
            };
            let s = to_num_fmt(f);
            assert_eq!(parse(&s), Some(f), "round-trip for {f:?} via {s:?}");
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
            parens: false,
            negative_color: None,
        };
        assert_eq!(parse(&to_num_fmt(f)), Some(f));
    }
}
