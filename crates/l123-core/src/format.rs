//! Cell display formats and their parenthesized tags as shown in the
//! control panel (e.g. `(C2)` = Currency 2dp).  See SPEC §12.

use crate::international::{CurrencyPosition, DateIntl, International, NegativeStyle, TimeIntl};
use std::fmt;

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum FormatKind {
    Fixed,
    Scientific,
    Currency,
    Comma,
    General,
    PlusMinus,
    Percent,
    DateDmy,       // D1: DD-MMM-YY
    DateDm,        // D2: DD-MMM
    DateMy,        // D3: MMM-YY
    DateLongIntl,  // D4
    DateShortIntl, // D5
    TimeHmsAmPm,   // D6
    TimeHmAmPm,    // D7
    TimeLongIntl,  // D8
    TimeShortIntl, // D9
    Text,          // Show formula, not value
    Hidden,
    Automatic,
    LabelOnly,
    /// Reset / inherit from global.
    Reset,
}

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub struct Format {
    pub kind: FormatKind,
    /// Decimal places, 0..=15. Ignored for kinds that don't use it.
    pub decimals: u8,
}

impl Format {
    pub const GENERAL: Format = Format {
        kind: FormatKind::General,
        decimals: 0,
    };
    pub const RESET: Format = Format {
        kind: FormatKind::Reset,
        decimals: 0,
    };

    pub fn fixed(d: u8) -> Self {
        Self {
            kind: FormatKind::Fixed,
            decimals: d.min(15),
        }
    }
    pub fn currency(d: u8) -> Self {
        Self {
            kind: FormatKind::Currency,
            decimals: d.min(15),
        }
    }
    pub fn percent(d: u8) -> Self {
        Self {
            kind: FormatKind::Percent,
            decimals: d.min(15),
        }
    }
    pub fn comma(d: u8) -> Self {
        Self {
            kind: FormatKind::Comma,
            decimals: d.min(15),
        }
    }

    /// Tag as shown in parentheses on control-panel line 1.
    /// Returns None for Reset (no tag shown when inheriting).
    pub fn tag(self) -> Option<String> {
        use FormatKind::*;
        let s = match self.kind {
            Fixed => format!("F{}", self.decimals),
            Scientific => format!("S{}", self.decimals),
            Currency => format!("C{}", self.decimals),
            Comma => format!(",{}", self.decimals),
            Percent => format!("P{}", self.decimals),
            General => "G".into(),
            PlusMinus => "+".into(),
            DateDmy => "D1".into(),
            DateDm => "D2".into(),
            DateMy => "D3".into(),
            DateLongIntl => "D4".into(),
            DateShortIntl => "D5".into(),
            TimeHmsAmPm => "D6".into(),
            TimeHmAmPm => "D7".into(),
            TimeLongIntl => "D8".into(),
            TimeShortIntl => "D9".into(),
            Text => "T".into(),
            Hidden => "H".into(),
            Automatic => "A".into(),
            LabelOnly => "L".into(),
            Reset => return None,
        };
        Some(s)
    }
}

impl Default for Format {
    fn default() -> Self {
        Format::GENERAL
    }
}

impl fmt::Display for Format {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self.tag() {
            Some(s) => write!(f, "({s})"),
            None => Ok(()),
        }
    }
}

/// Format a number per the given cell format. Falls back to General for
/// formats that are not yet implemented (Date/Time/Text).
///
/// This is the single source of truth for numeric display; both the grid
/// widget and the control panel's cell readout call it. `intl` supplies
/// punctuation (decimal point + thousands separator), negative style,
/// and currency symbol/position.
pub fn format_number(n: f64, format: Format, intl: &International) -> String {
    use FormatKind::*;
    let d = format.decimals as usize;
    let dec = intl.punctuation.decimal_char();
    let thou = intl.punctuation.thousands_sep();
    match format.kind {
        Fixed => swap_decimal(format!("{n:.d$}"), dec),
        Scientific => swap_decimal(format!("{n:.d$e}"), dec),
        Currency => {
            let abs = n.abs();
            let body = swap_decimal(format!("{abs:.d$}"), dec);
            let sym = &intl.currency.symbol;
            let with_sym = match intl.currency.position {
                CurrencyPosition::Prefix => format!("{sym}{body}"),
                CurrencyPosition::Suffix => format!("{body}{sym}"),
            };
            apply_negative(n < 0.0, with_sym, intl.negative_style)
        }
        Comma => {
            let body = with_thousands(&format!("{:.d$}", n.abs()), thou, dec);
            apply_negative(n < 0.0, body, intl.negative_style)
        }
        Percent => {
            let v = n * 100.0;
            swap_decimal(format!("{v:.d$}%"), dec)
        }
        PlusMinus => plus_minus_bar(n),
        General | Reset | Automatic => swap_decimal(crate::contents::format_number_general(n), dec),
        DateLongIntl => format_date_intl(n, intl.date_intl, true),
        DateShortIntl => format_date_intl(n, intl.date_intl, false),
        TimeLongIntl => format_time_intl(n, intl.time_intl, true),
        TimeShortIntl => format_time_intl(n, intl.time_intl, false),
        // D1/D2/D3 (DateDmy/DateDm/DateMy) and D6/D7 (TimeHmsAmPm/
        // HmAmPm): not yet wired. Display the underlying number until
        // their milestones land.
        DateDmy | DateDm | DateMy | TimeHmsAmPm | TimeHmAmPm | Text | Hidden | LabelOnly => {
            swap_decimal(crate::contents::format_number_general(n), dec)
        }
    }
}

/// Convert a Lotus serial date to (year, month, day). The serial is
/// days since 1899-12-30, with Lotus's R3 1900-leap-year quirk
/// preserved (serial 60 = "1900-02-29", though that day did not
/// exist). This matches Excel/IronCalc, so round-tripping through
/// xlsx stays exact.
///
/// Algorithm: Howard Hinnant's `civil_from_days` on the proleptic
/// Gregorian calendar, with the +1 quirk for serials > 60.
pub(crate) fn serial_to_ymd(serial: f64) -> (i32, u32, u32) {
    let days = serial.trunc() as i64;
    // Excel/Lotus convention: serial 1 = 1900-01-01. Hinnant's
    // civil_from_days uses 1970-01-01 = 0; days from 1970-01-01 back
    // to 1900-01-01 = 25567, so serial 1 → civil day -25567, hence
    // the epoch offset for serial 0 is -25568.
    const EPOCH_OFFSET: i64 = -25568;
    // Lotus pretends 1900-02-29 exists (serial 60). Real calendar
    // skips it. For serials >= 60 we shift back by 1 day so subsequent
    // dates land on the correct civil date. Serial 60 itself is
    // synthetic and reported as (1900, 2, 29) explicitly.
    let civil_days = if days >= 60 {
        days - 1 + EPOCH_OFFSET
    } else {
        days + EPOCH_OFFSET
    };
    let (y, m, d) = civil_from_days(civil_days);
    if days == 60 {
        (1900, 2, 29)
    } else {
        (y, m, d)
    }
}

/// Howard Hinnant's `civil_from_days`, returning (year, month, day)
/// for `days` since 1970-01-01 in the proleptic Gregorian calendar.
fn civil_from_days(days: i64) -> (i32, u32, u32) {
    let z = days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32;
    let m = (if mp < 10 { mp + 3 } else { mp - 9 }) as u32;
    let y_final = y + if m <= 2 { 1 } else { 0 };
    (y_final as i32, m, d)
}

/// Convert the fractional part of a Lotus serial to (h, m, s).
/// `0.0` → midnight, `0.5` → noon, `0.999988425926` ≈ 23:59:59.
/// Wraps at exactly 24:00:00 → 0:00:00 to match Lotus.
fn fraction_to_hms(serial: f64) -> (u32, u32, u32) {
    let frac = serial.fract().abs();
    let total_seconds = (frac * 86_400.0).round() as u64;
    // 86400 (the wraparound case) folds to 0 — display as 00:00:00.
    let total = total_seconds % 86_400;
    let h = ((total / 3_600) % 24) as u32;
    let m = ((total / 60) % 60) as u32;
    let s = (total % 60) as u32;
    (h, m, s)
}

/// Render the time-of-day portion of `serial` per `intl`. `long`
/// includes seconds (D8); `short` omits them (D9). Time D falls back
/// to colon glyphs (= Time A) until LICS letter glyphs land.
fn format_time_intl(serial: f64, intl: TimeIntl, long: bool) -> String {
    let (h, m, s) = fraction_to_hms(serial);
    match (intl, long) {
        (TimeIntl::A, true) | (TimeIntl::D, true) => format!("{h:02}:{m:02}:{s:02}"),
        (TimeIntl::A, false) | (TimeIntl::D, false) => format!("{h:02}:{m:02}"),
        (TimeIntl::B, true) => format!("{h:02}.{m:02}.{s:02}"),
        (TimeIntl::B, false) => format!("{h:02}.{m:02}"),
        (TimeIntl::C, true) => format!("{h:02},{m:02},{s:02}"),
        (TimeIntl::C, false) => format!("{h:02},{m:02}"),
    }
}

/// Render `serial` as an international date according to `intl`.
/// `long` toggles between D4 (long: with year) and D5 (short: no year).
fn format_date_intl(serial: f64, intl: DateIntl, long: bool) -> String {
    let (y, m, d) = serial_to_ymd(serial);
    let yy = (y % 100).unsigned_abs();
    match (intl, long) {
        (DateIntl::A, true) => format!("{m:02}/{d:02}/{yy:02}"),
        (DateIntl::A, false) => format!("{m:02}/{d:02}"),
        (DateIntl::B, true) => format!("{d:02}/{m:02}/{yy:02}"),
        (DateIntl::B, false) => format!("{d:02}/{m:02}"),
        (DateIntl::C, true) => format!("{d:02}.{m:02}.{yy:02}"),
        (DateIntl::C, false) => format!("{d:02}.{m:02}"),
        (DateIntl::D, true) => format!("{yy:02}-{m:02}-{d:02}"),
        (DateIntl::D, false) => format!("{m:02}-{d:02}"),
    }
}

/// Replace the canonical Rust-style `.` decimal point with `dec`.
/// No-op when `dec == '.'`. The `e` exponent indicator passes through.
fn swap_decimal(s: String, dec: char) -> String {
    if dec == '.' {
        s
    } else {
        s.replace('.', &dec.to_string())
    }
}

/// Wrap a positive-magnitude body with the configured negative style.
/// `Sign` produces `-body`; `Parens` produces `(body)`. Positives pass
/// through unchanged.
fn apply_negative(negative: bool, body: String, style: NegativeStyle) -> String {
    if !negative {
        return body;
    }
    match style {
        NegativeStyle::Sign => format!("-{body}"),
        NegativeStyle::Parens => format!("({body})"),
    }
}

fn with_thousands(s: &str, thou: char, dec: char) -> String {
    // Insert thousands separators in the integer part. Works on
    // Rust-style `"-1234.56"`. The Rust `.` is swapped to `dec` last.
    let (sign, rest) = if let Some(rest) = s.strip_prefix('-') {
        ("-", rest)
    } else {
        ("", s)
    };
    let (int_part, frac_part) = match rest.find('.') {
        Some(i) => (&rest[..i], &rest[i..]),
        None => (rest, ""),
    };
    let mut out = String::with_capacity(int_part.len() + int_part.len() / 3);
    for (i, ch) in int_part.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            out.push(thou);
        }
        out.push(ch);
    }
    let int_with_seps: String = out.chars().rev().collect();
    let frac_swapped = if dec == '.' {
        frac_part.to_string()
    } else {
        frac_part.replace('.', &dec.to_string())
    };
    format!("{sign}{int_with_seps}{frac_swapped}")
}

/// `(+)` bar chart format: each unit is represented as `+` (positive) or
/// `-` (negative). Truncated to a reasonable width.
fn plus_minus_bar(n: f64) -> String {
    let units = n.round() as i64;
    let ch = if units < 0 { '-' } else { '+' };
    let count = units.unsigned_abs().min(40) as usize;
    std::iter::repeat_n(ch, count).collect()
}

#[cfg(test)]
mod format_number_tests {
    use super::*;
    use crate::international::Punctuation;

    fn intl_default() -> International {
        International::default()
    }

    #[test]
    fn fixed_rounds_to_decimals() {
        let i = intl_default();
        assert_eq!(format_number(1.25, Format::fixed(2), &i), "1.25");
        // Rust uses banker's rounding (round half to even) for `{:.*}`.
        assert_eq!(format_number(1.25, Format::fixed(1), &i), "1.2");
        assert_eq!(format_number(1.35, Format::fixed(1), &i), "1.4");
        assert_eq!(format_number(1.23456, Format::fixed(2), &i), "1.23");
    }

    #[test]
    fn currency_has_dollar_and_decimals() {
        let i = intl_default();
        assert_eq!(format_number(1000.0, Format::currency(2), &i), "$1000.00");
        assert_eq!(format_number(42.5, Format::currency(0), &i), "$42"); // banker's rounding: 42
        assert_eq!(format_number(-42.5, Format::currency(0), &i), "-$42");
    }

    #[test]
    fn percent_multiplies_by_100() {
        let i = intl_default();
        assert_eq!(format_number(0.5, Format::percent(0), &i), "50%");
        assert_eq!(format_number(0.123, Format::percent(1), &i), "12.3%");
    }

    #[test]
    fn comma_inserts_thousands_separators() {
        let i = intl_default();
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i),
            "1,234,567"
        );
        assert_eq!(format_number(1000.5, Format::comma(2), &i), "1,000.50");
        assert_eq!(format_number(-1000.5, Format::comma(2), &i), "-1,000.50");
        assert_eq!(format_number(999.0, Format::comma(0), &i), "999");
    }

    #[test]
    fn general_matches_existing_behavior() {
        let i = intl_default();
        assert_eq!(format_number(123.0, Format::GENERAL, &i), "123");
        assert_eq!(format_number(1.5, Format::GENERAL, &i), "1.5");
    }

    #[test]
    fn scientific_uses_e_notation() {
        let i = intl_default();
        assert_eq!(
            format_number(
                1234.5,
                Format {
                    kind: FormatKind::Scientific,
                    decimals: 2
                },
                &i
            ),
            "1.23e3"
        );
    }

    #[test]
    fn plus_minus_bar_draws_bars() {
        let i = intl_default();
        assert_eq!(
            format_number(
                3.0,
                Format {
                    kind: FormatKind::PlusMinus,
                    decimals: 0
                },
                &i
            ),
            "+++"
        );
        assert_eq!(
            format_number(
                -2.0,
                Format {
                    kind: FormatKind::PlusMinus,
                    decimals: 0
                },
                &i
            ),
            "--"
        );
        assert_eq!(
            format_number(
                0.0,
                Format {
                    kind: FormatKind::PlusMinus,
                    decimals: 0
                },
                &i
            ),
            ""
        );
    }

    fn intl_with(p: Punctuation) -> International {
        International {
            punctuation: p,
            ..Default::default()
        }
    }

    #[test]
    fn fixed_under_punct_b_uses_comma_as_decimal() {
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(1.25, Format::fixed(2), &i), "1,25");
        assert_eq!(format_number(-1.25, Format::fixed(2), &i), "-1,25");
    }

    #[test]
    fn comma_under_punct_b_swaps_separators() {
        // Punct B: thousands `.`, decimal `,`.
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(1234.5, Format::comma(2), &i), "1.234,50");
        assert_eq!(format_number(-1234.5, Format::comma(2), &i), "-1.234,50");
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i),
            "1.234.567"
        );
    }

    #[test]
    fn comma_under_punct_c_uses_space_thousands() {
        // Punct C: thousands ' ' (space), decimal `.`.
        let i = intl_with(Punctuation::C);
        assert_eq!(format_number(1234.5, Format::comma(2), &i), "1 234.50");
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i),
            "1 234 567"
        );
    }

    #[test]
    fn currency_under_punct_b_swaps_decimal() {
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(1000.5, Format::currency(2), &i), "$1000,50");
        assert_eq!(format_number(-42.5, Format::currency(0), &i), "-$42");
    }

    #[test]
    fn percent_under_punct_b_swaps_decimal() {
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(0.123, Format::percent(1), &i), "12,3%");
    }

    #[test]
    fn scientific_under_punct_b_swaps_decimal() {
        let i = intl_with(Punctuation::B);
        // The 'e' is just an exponent marker; the decimal is what swaps.
        assert_eq!(
            format_number(
                1234.5,
                Format {
                    kind: FormatKind::Scientific,
                    decimals: 2
                },
                &i
            ),
            "1,23e3"
        );
    }

    use crate::international::{CurrencyConfig, CurrencyPosition, NegativeStyle};

    #[test]
    fn currency_suffix_renders_after_body() {
        let i = International {
            currency: CurrencyConfig {
                symbol: "€".into(),
                position: CurrencyPosition::Suffix,
            },
            ..Default::default()
        };
        assert_eq!(format_number(1234.5, Format::currency(2), &i), "1234.50€");
    }

    #[test]
    fn currency_suffix_with_parens_negative() {
        let i = International {
            currency: CurrencyConfig {
                symbol: "€".into(),
                position: CurrencyPosition::Suffix,
            },
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(
            format_number(-1234.5, Format::currency(2), &i),
            "(1234.50€)"
        );
    }

    #[test]
    fn currency_prefix_with_parens_negative() {
        let i = International {
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(
            format_number(-1234.5, Format::currency(2), &i),
            "($1234.50)"
        );
    }

    #[test]
    fn comma_with_parens_negative() {
        let i = International {
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(format_number(-1234.5, Format::comma(2), &i), "(1,234.50)");
        // Positive untouched.
        assert_eq!(format_number(1234.5, Format::comma(2), &i), "1,234.50");
    }

    #[test]
    fn comma_under_punct_b_with_parens_negative() {
        let i = International {
            punctuation: Punctuation::B,
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(format_number(-1234.5, Format::comma(2), &i), "(1.234,50)");
    }

    #[test]
    fn currency_uses_configured_symbol() {
        let i = International {
            currency: CurrencyConfig {
                symbol: "USD ".into(),
                position: CurrencyPosition::Prefix,
            },
            ..Default::default()
        };
        assert_eq!(
            format_number(1234.5, Format::currency(2), &i),
            "USD 1234.50"
        );
    }

    use crate::international::DateIntl;

    fn fmt_d4() -> Format {
        Format {
            kind: FormatKind::DateLongIntl,
            decimals: 0,
        }
    }
    fn fmt_d5() -> Format {
        Format {
            kind: FormatKind::DateShortIntl,
            decimals: 0,
        }
    }

    fn intl_with_date(d: DateIntl) -> International {
        International {
            date_intl: d,
            ..Default::default()
        }
    }

    #[test]
    fn date_intl_a_renders_us_long_short() {
        // Serial 36526 = 2000-01-01 (Excel/IronCalc convention).
        let i = intl_with_date(DateIntl::A);
        assert_eq!(format_number(36526.0, fmt_d4(), &i), "01/01/00");
        assert_eq!(format_number(36526.0, fmt_d5(), &i), "01/01");
    }

    #[test]
    fn date_intl_b_renders_dd_mm_yy() {
        // 2000-01-15 = 36540.
        let i = intl_with_date(DateIntl::B);
        assert_eq!(format_number(36540.0, fmt_d4(), &i), "15/01/00");
        assert_eq!(format_number(36540.0, fmt_d5(), &i), "15/01");
    }

    #[test]
    fn date_intl_c_uses_dot_separator() {
        let i = intl_with_date(DateIntl::C);
        assert_eq!(format_number(36540.0, fmt_d4(), &i), "15.01.00");
        assert_eq!(format_number(36540.0, fmt_d5(), &i), "15.01");
    }

    #[test]
    fn date_intl_d_renders_yy_mm_dd() {
        let i = intl_with_date(DateIntl::D);
        assert_eq!(format_number(36540.0, fmt_d4(), &i), "00-01-15");
        assert_eq!(format_number(36540.0, fmt_d5(), &i), "01-15");
    }

    #[test]
    fn lotus_1900_leap_quirk_preserved() {
        // Serial 60 = synthetic 1900-02-29 in Lotus/Excel.
        // Serial 61 = real 1900-03-01.
        let i = intl_with_date(DateIntl::A);
        assert_eq!(format_number(60.0, fmt_d4(), &i), "02/29/00");
        assert_eq!(format_number(61.0, fmt_d4(), &i), "03/01/00");
    }

    #[test]
    fn date_intl_year_2099_two_digits() {
        // Serial 73050 = 2099-12-31 (just under the wrap to 2100).
        let i = intl_with_date(DateIntl::A);
        assert_eq!(format_number(73050.0, fmt_d4(), &i), "12/31/99");
    }

    use crate::international::TimeIntl;

    fn fmt_d8() -> Format {
        Format {
            kind: FormatKind::TimeLongIntl,
            decimals: 0,
        }
    }
    fn fmt_d9() -> Format {
        Format {
            kind: FormatKind::TimeShortIntl,
            decimals: 0,
        }
    }

    fn intl_with_time(t: TimeIntl) -> International {
        International {
            time_intl: t,
            ..Default::default()
        }
    }

    #[test]
    fn time_intl_a_uses_colon_separator() {
        let i = intl_with_time(TimeIntl::A);
        // Noon = 0.5.
        assert_eq!(format_number(0.5, fmt_d8(), &i), "12:00:00");
        assert_eq!(format_number(0.5, fmt_d9(), &i), "12:00");
        // 6 AM = 0.25.
        assert_eq!(format_number(0.25, fmt_d8(), &i), "06:00:00");
    }

    #[test]
    fn time_intl_b_uses_dot_separator() {
        let i = intl_with_time(TimeIntl::B);
        assert_eq!(format_number(0.5, fmt_d8(), &i), "12.00.00");
        assert_eq!(format_number(0.5, fmt_d9(), &i), "12.00");
    }

    #[test]
    fn time_intl_c_uses_comma_separator() {
        let i = intl_with_time(TimeIntl::C);
        assert_eq!(format_number(0.5, fmt_d8(), &i), "12,00,00");
        assert_eq!(format_number(0.5, fmt_d9(), &i), "12,00");
    }

    #[test]
    fn time_intl_d_falls_back_to_colon() {
        let i = intl_with_time(TimeIntl::D);
        assert_eq!(format_number(0.5, fmt_d8(), &i), "12:00:00");
        assert_eq!(format_number(0.5, fmt_d9(), &i), "12:00");
    }

    #[test]
    fn time_drops_date_part_at_serial_with_time() {
        // 36526.5 = 2000-01-01 12:00 noon. D8 shows time only.
        let i = intl_with_time(TimeIntl::A);
        assert_eq!(format_number(36526.5, fmt_d8(), &i), "12:00:00");
    }

    #[test]
    fn time_wrap_at_midnight_shows_zero() {
        let i = intl_with_time(TimeIntl::A);
        // 23:59:59.5 should round to 24:00:00 → 00:00:00.
        let near_midnight = 86_399.5 / 86_400.0;
        assert_eq!(format_number(near_midnight, fmt_d8(), &i), "00:00:00");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tags() {
        assert_eq!(Format::currency(2).tag().as_deref(), Some("C2"));
        assert_eq!(Format::percent(1).tag().as_deref(), Some("P1"));
        assert_eq!(Format::comma(0).tag().as_deref(), Some(",0"));
        assert_eq!(Format::GENERAL.tag().as_deref(), Some("G"));
        assert_eq!(Format::RESET.tag(), None);
    }

    #[test]
    fn display() {
        assert_eq!(format!("{}", Format::currency(2)), "(C2)");
        assert_eq!(format!("{}", Format::GENERAL), "(G)");
        assert_eq!(format!("{}", Format::RESET), "");
    }
}
