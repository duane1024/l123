//! Cell display formats and their parenthesized tags as shown in the
//! control panel (e.g. `(C2)` = Currency 2dp).  See SPEC §12.

use crate::international::{CurrencyPosition, DateIntl, International, NegativeStyle, TimeIntl};
use std::fmt;

/// A single component of an Excel-style date pattern, preserved verbatim
/// so non-canonical formats (e.g. `m/yyyy`, `dd-mmm-yyyy`) round-trip
/// through xlsx without collapsing to a 1-2-3 D1..D5 enum.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum DateToken {
    /// `yy` — two-digit year, zero-padded.
    Year2,
    /// `yyyy` — four-digit year.
    Year4,
    /// `m` — numeric month, no padding.
    MonthNum,
    /// `mm` — numeric month, zero-padded.
    MonthNumPadded,
    /// `mmm` — three-letter month abbreviation, mixed-case (e.g. `Oct`).
    MonthAbbrev,
    /// `mmmm` — full month name, mixed-case (e.g. `October`).
    MonthFull,
    /// `d` — day of month, no padding.
    Day,
    /// `dd` — day of month, zero-padded.
    DayPadded,
    /// Any non-glyph character(s) between tokens — separators like `/`,
    /// `-`, `.`, ` `, or quoted literals.
    Literal(String),
}

/// Workbook-scoped intern table for `DateToken` lists. `FormatKind::DateCustom`
/// stores a `u16` index into this table so `Format` stays `Copy`.
#[derive(Default, Debug, Clone)]
pub struct DateFormatTable {
    entries: Vec<Vec<DateToken>>,
}

impl DateFormatTable {
    pub fn new() -> Self {
        Self::default()
    }

    /// Intern a token list. Equal lists return the same id; the table
    /// caps at `u16::MAX - 1` entries (panics if exceeded — xlsx files
    /// will never approach that).
    pub fn intern(&mut self, tokens: Vec<DateToken>) -> u16 {
        if let Some(i) = self.entries.iter().position(|e| e == &tokens) {
            return i as u16;
        }
        let id = self.entries.len();
        assert!(id < u16::MAX as usize, "date format table overflow");
        self.entries.push(tokens);
        id as u16
    }

    pub fn get(&self, id: u16) -> &[DateToken] {
        &self.entries[id as usize]
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

/// Tokenize an Excel date format string into a `Vec<DateToken>`.
///
/// Returns `None` if the string contains no date glyphs (`y`, `m`, `d`).
/// Time glyphs (`h`, `s`) currently disqualify the input — date-only
/// for now; combined date+time is a future extension.
pub fn parse_date_pattern(positive: &str) -> Option<Vec<DateToken>> {
    let mut tokens: Vec<DateToken> = Vec::new();
    let mut literal = String::new();
    let mut chars = positive.chars().peekable();
    let mut saw_date_glyph = false;

    let flush_literal = |literal: &mut String, tokens: &mut Vec<DateToken>| {
        if !literal.is_empty() {
            tokens.push(DateToken::Literal(std::mem::take(literal)));
        }
    };

    while let Some(c) = chars.next() {
        match c {
            'h' | 'H' | 's' | 'S' => return None,
            '"' => {
                for inner in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                    literal.push(inner);
                }
            }
            '\\' => {
                if let Some(next) = chars.next() {
                    literal.push(next);
                }
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
            'y' | 'Y' => {
                let mut run = 1;
                while matches!(chars.peek(), Some('y') | Some('Y')) {
                    chars.next();
                    run += 1;
                }
                flush_literal(&mut literal, &mut tokens);
                tokens.push(if run <= 2 {
                    DateToken::Year2
                } else {
                    DateToken::Year4
                });
                saw_date_glyph = true;
            }
            'm' | 'M' => {
                let mut run = 1;
                while matches!(chars.peek(), Some('m') | Some('M')) {
                    chars.next();
                    run += 1;
                }
                flush_literal(&mut literal, &mut tokens);
                tokens.push(match run {
                    1 => DateToken::MonthNum,
                    2 => DateToken::MonthNumPadded,
                    3 => DateToken::MonthAbbrev,
                    _ => DateToken::MonthFull,
                });
                saw_date_glyph = true;
            }
            'd' | 'D' => {
                let mut run = 1;
                while matches!(chars.peek(), Some('d') | Some('D')) {
                    chars.next();
                    run += 1;
                }
                flush_literal(&mut literal, &mut tokens);
                tokens.push(if run == 1 {
                    DateToken::Day
                } else {
                    DateToken::DayPadded
                });
                saw_date_glyph = true;
            }
            _ => literal.push(c),
        }
    }

    if !saw_date_glyph {
        return None;
    }
    flush_literal(&mut literal, &mut tokens);
    Some(tokens)
}

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
    /// Excel date pattern that doesn't reduce to D1..D5. The `u16` is
    /// an index into the workbook's `DateFormatTable` and is opaque to
    /// callers that don't have the table.
    DateCustom(u16),
    Text, // Show formula, not value
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
    ///
    /// For `DateCustom`, callers without a `DateFormatTable` see `D*` —
    /// callers that have one should use [`Format::tag_with_dates`] for
    /// the closest-1-2-3-fit tag (`D1*`, `D3*`, …).
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
            DateCustom(_) => "D*".into(),
            Text => "T".into(),
            Hidden => "H".into(),
            Automatic => "A".into(),
            LabelOnly => "L".into(),
            Reset => return None,
        };
        Some(s)
    }

    /// Like [`Format::tag`], but produces a richer tag for `DateCustom`
    /// by classifying the interned tokens against the 1-2-3 D1..D5
    /// shapes — `dd-mmm-yyyy` → `D1*`, `m/d/yyyy` → `D4*`, `mmm-yyyy`
    /// → `D3*`. The `*` suffix marks that the rendered string deviates
    /// from the canonical 1-2-3 form (different separator, year width,
    /// month casing). Falls back to `D*` when no shape matches.
    pub fn tag_with_dates(self, table: &DateFormatTable) -> Option<String> {
        if let FormatKind::DateCustom(id) = self.kind {
            let tokens = table.get(id);
            return Some(match closest_canonical_date_kind(tokens) {
                Some(FormatKind::DateDmy) => "D1*".into(),
                Some(FormatKind::DateDm) => "D2*".into(),
                Some(FormatKind::DateMy) => "D3*".into(),
                Some(FormatKind::DateLongIntl) => "D4*".into(),
                Some(FormatKind::DateShortIntl) => "D5*".into(),
                _ => "D*".into(),
            });
        }
        self.tag()
    }
}

/// Classify a `DateCustom` token list against the canonical D1..D5
/// shapes by glyph presence (year/month/day, letter vs numeric month).
/// Used to surface a useful control-panel tag for non-canonical
/// patterns. Literal separators are ignored.
pub fn closest_canonical_date_kind(tokens: &[DateToken]) -> Option<FormatKind> {
    let mut has_year = false;
    let mut has_day = false;
    let mut has_month = false;
    let mut letter_month = false;
    for t in tokens {
        match t {
            DateToken::Year2 | DateToken::Year4 => has_year = true,
            DateToken::Day | DateToken::DayPadded => has_day = true,
            DateToken::MonthNum | DateToken::MonthNumPadded => has_month = true,
            DateToken::MonthAbbrev | DateToken::MonthFull => {
                has_month = true;
                letter_month = true;
            }
            DateToken::Literal(_) => {}
        }
    }
    if !has_month {
        return None;
    }
    Some(match (has_year, has_day, letter_month) {
        (true, true, true) => FormatKind::DateDmy,
        (false, true, true) => FormatKind::DateDm,
        (true, false, true) => FormatKind::DateMy,
        (true, true, false) => FormatKind::DateLongIntl,
        (false, true, false) => FormatKind::DateShortIntl,
        // Year-only-numeric (e.g. `m/yyyy` — numeric month + year, no day)
        // has no clean 1-2-3 analog. D3 (MMM-YY) is the spirit but the
        // month form differs; mark as closest-fit anyway.
        (true, false, false) => FormatKind::DateMy,
        _ => return None,
    })
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
/// and currency symbol/position. `dates` carries the workbook's
/// interned `DateCustom` token lists; an empty table is fine when no
/// `DateCustom` formats are in use.
pub fn format_number(
    n: f64,
    format: Format,
    intl: &International,
    dates: &DateFormatTable,
) -> String {
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
        DateDmy => format_date_letter_month(n, true, true),
        DateDm => format_date_letter_month(n, false, true),
        DateMy => format_date_letter_month(n, true, false),
        DateCustom(id) => format_custom_date(n, dates.get(id)),
        // D6/D7 (TimeHmsAmPm/HmAmPm) not yet wired. Display the
        // underlying number until their milestones land.
        TimeHmsAmPm | TimeHmAmPm | Text | Hidden | LabelOnly => {
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

const MONTH_ABBREV_UPPER: [&str; 12] = [
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
];

const MONTH_ABBREV_MIXED: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

const MONTH_FULL_MIXED: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

/// Render a Lotus serial date as the literal Excel pattern represented by
/// `tokens`. Mixed-case month names match Excel's default casing.
pub fn format_custom_date(serial: f64, tokens: &[DateToken]) -> String {
    let (y, m, d) = serial_to_ymd(serial);
    let mi = (m.saturating_sub(1) as usize).min(11);
    let mut out = String::new();
    for t in tokens {
        match t {
            DateToken::Year2 => {
                let yy = (y % 100).unsigned_abs();
                out.push_str(&format!("{yy:02}"));
            }
            DateToken::Year4 => out.push_str(&format!("{y:04}")),
            DateToken::MonthNum => out.push_str(&format!("{m}")),
            DateToken::MonthNumPadded => out.push_str(&format!("{m:02}")),
            DateToken::MonthAbbrev => out.push_str(MONTH_ABBREV_MIXED[mi]),
            DateToken::MonthFull => out.push_str(MONTH_FULL_MIXED[mi]),
            DateToken::Day => out.push_str(&format!("{d}")),
            DateToken::DayPadded => out.push_str(&format!("{d:02}")),
            DateToken::Literal(s) => out.push_str(s),
        }
    }
    out
}

/// Render `serial` using a 1-2-3 letter-month date kind:
/// * `with_year && with_day` → D1 (`DD-MMM-YY`)
/// * `with_year && !with_day` → D3 (`MMM-YY`)
/// * `!with_year && with_day` → D2 (`DD-MMM`)
fn format_date_letter_month(serial: f64, with_year: bool, with_day: bool) -> String {
    let (y, m, d) = serial_to_ymd(serial);
    let mon = MONTH_ABBREV_UPPER[(m.saturating_sub(1) as usize).min(11)];
    let yy = (y % 100).unsigned_abs();
    match (with_day, with_year) {
        (true, true) => format!("{d:02}-{mon}-{yy:02}"),
        (true, false) => format!("{d:02}-{mon}"),
        (false, true) => format!("{mon}-{yy:02}"),
        (false, false) => mon.to_string(),
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
        assert_eq!(format_number(1.25, Format::fixed(2), &i, &DateFormatTable::default()), "1.25");
        // Rust uses banker's rounding (round half to even) for `{:.*}`.
        assert_eq!(format_number(1.25, Format::fixed(1), &i, &DateFormatTable::default()), "1.2");
        assert_eq!(format_number(1.35, Format::fixed(1), &i, &DateFormatTable::default()), "1.4");
        assert_eq!(format_number(1.23456, Format::fixed(2), &i, &DateFormatTable::default()), "1.23");
    }

    #[test]
    fn currency_has_dollar_and_decimals() {
        let i = intl_default();
        assert_eq!(format_number(1000.0, Format::currency(2), &i, &DateFormatTable::default()), "$1000.00");
        assert_eq!(format_number(42.5, Format::currency(0), &i, &DateFormatTable::default()), "$42"); // banker's rounding: 42
        assert_eq!(format_number(-42.5, Format::currency(0), &i, &DateFormatTable::default()), "-$42");
    }

    #[test]
    fn percent_multiplies_by_100() {
        let i = intl_default();
        assert_eq!(format_number(0.5, Format::percent(0), &i, &DateFormatTable::default()), "50%");
        assert_eq!(format_number(0.123, Format::percent(1), &i, &DateFormatTable::default()), "12.3%");
    }

    #[test]
    fn comma_inserts_thousands_separators() {
        let i = intl_default();
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i, &DateFormatTable::default()),
            "1,234,567"
        );
        assert_eq!(format_number(1000.5, Format::comma(2), &i, &DateFormatTable::default()), "1,000.50");
        assert_eq!(format_number(-1000.5, Format::comma(2), &i, &DateFormatTable::default()), "-1,000.50");
        assert_eq!(format_number(999.0, Format::comma(0), &i, &DateFormatTable::default()), "999");
    }

    #[test]
    fn general_matches_existing_behavior() {
        let i = intl_default();
        assert_eq!(format_number(123.0, Format::GENERAL, &i, &DateFormatTable::default()), "123");
        assert_eq!(format_number(1.5, Format::GENERAL, &i, &DateFormatTable::default()), "1.5");
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
                &i, &DateFormatTable::default()),
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
                &i, &DateFormatTable::default()),
            "+++"
        );
        assert_eq!(
            format_number(
                -2.0,
                Format {
                    kind: FormatKind::PlusMinus,
                    decimals: 0
                },
                &i, &DateFormatTable::default()),
            "--"
        );
        assert_eq!(
            format_number(
                0.0,
                Format {
                    kind: FormatKind::PlusMinus,
                    decimals: 0
                },
                &i, &DateFormatTable::default()),
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
        assert_eq!(format_number(1.25, Format::fixed(2), &i, &DateFormatTable::default()), "1,25");
        assert_eq!(format_number(-1.25, Format::fixed(2), &i, &DateFormatTable::default()), "-1,25");
    }

    #[test]
    fn comma_under_punct_b_swaps_separators() {
        // Punct B: thousands `.`, decimal `,`.
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(1234.5, Format::comma(2), &i, &DateFormatTable::default()), "1.234,50");
        assert_eq!(format_number(-1234.5, Format::comma(2), &i, &DateFormatTable::default()), "-1.234,50");
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i, &DateFormatTable::default()),
            "1.234.567"
        );
    }

    #[test]
    fn comma_under_punct_c_uses_space_thousands() {
        // Punct C: thousands ' ' (space), decimal `.`.
        let i = intl_with(Punctuation::C);
        assert_eq!(format_number(1234.5, Format::comma(2), &i, &DateFormatTable::default()), "1 234.50");
        assert_eq!(
            format_number(1_234_567.0, Format::comma(0), &i, &DateFormatTable::default()),
            "1 234 567"
        );
    }

    #[test]
    fn currency_under_punct_b_swaps_decimal() {
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(1000.5, Format::currency(2), &i, &DateFormatTable::default()), "$1000,50");
        assert_eq!(format_number(-42.5, Format::currency(0), &i, &DateFormatTable::default()), "-$42");
    }

    #[test]
    fn percent_under_punct_b_swaps_decimal() {
        let i = intl_with(Punctuation::B);
        assert_eq!(format_number(0.123, Format::percent(1), &i, &DateFormatTable::default()), "12,3%");
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
                &i, &DateFormatTable::default()),
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
        assert_eq!(format_number(1234.5, Format::currency(2), &i, &DateFormatTable::default()), "1234.50€");
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
            format_number(-1234.5, Format::currency(2), &i, &DateFormatTable::default()),
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
            format_number(-1234.5, Format::currency(2), &i, &DateFormatTable::default()),
            "($1234.50)"
        );
    }

    #[test]
    fn comma_with_parens_negative() {
        let i = International {
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(format_number(-1234.5, Format::comma(2), &i, &DateFormatTable::default()), "(1,234.50)");
        // Positive untouched.
        assert_eq!(format_number(1234.5, Format::comma(2), &i, &DateFormatTable::default()), "1,234.50");
    }

    #[test]
    fn comma_under_punct_b_with_parens_negative() {
        let i = International {
            punctuation: Punctuation::B,
            negative_style: NegativeStyle::Parens,
            ..Default::default()
        };
        assert_eq!(format_number(-1234.5, Format::comma(2), &i, &DateFormatTable::default()), "(1.234,50)");
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
            format_number(1234.5, Format::currency(2), &i, &DateFormatTable::default()),
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
        assert_eq!(format_number(36526.0, fmt_d4(), &i, &DateFormatTable::default()), "01/01/00");
        assert_eq!(format_number(36526.0, fmt_d5(), &i, &DateFormatTable::default()), "01/01");
    }

    #[test]
    fn date_intl_b_renders_dd_mm_yy() {
        // 2000-01-15 = 36540.
        let i = intl_with_date(DateIntl::B);
        assert_eq!(format_number(36540.0, fmt_d4(), &i, &DateFormatTable::default()), "15/01/00");
        assert_eq!(format_number(36540.0, fmt_d5(), &i, &DateFormatTable::default()), "15/01");
    }

    #[test]
    fn date_intl_c_uses_dot_separator() {
        let i = intl_with_date(DateIntl::C);
        assert_eq!(format_number(36540.0, fmt_d4(), &i, &DateFormatTable::default()), "15.01.00");
        assert_eq!(format_number(36540.0, fmt_d5(), &i, &DateFormatTable::default()), "15.01");
    }

    #[test]
    fn date_intl_d_renders_yy_mm_dd() {
        let i = intl_with_date(DateIntl::D);
        assert_eq!(format_number(36540.0, fmt_d4(), &i, &DateFormatTable::default()), "00-01-15");
        assert_eq!(format_number(36540.0, fmt_d5(), &i, &DateFormatTable::default()), "01-15");
    }

    #[test]
    fn lotus_1900_leap_quirk_preserved() {
        // Serial 60 = synthetic 1900-02-29 in Lotus/Excel.
        // Serial 61 = real 1900-03-01.
        let i = intl_with_date(DateIntl::A);
        assert_eq!(format_number(60.0, fmt_d4(), &i, &DateFormatTable::default()), "02/29/00");
        assert_eq!(format_number(61.0, fmt_d4(), &i, &DateFormatTable::default()), "03/01/00");
    }

    #[test]
    fn date_intl_year_2099_two_digits() {
        // Serial 73050 = 2099-12-31 (just under the wrap to 2100).
        let i = intl_with_date(DateIntl::A);
        assert_eq!(format_number(73050.0, fmt_d4(), &i, &DateFormatTable::default()), "12/31/99");
    }

    fn fmt_d1() -> Format {
        Format {
            kind: FormatKind::DateDmy,
            decimals: 0,
        }
    }
    fn fmt_d2() -> Format {
        Format {
            kind: FormatKind::DateDm,
            decimals: 0,
        }
    }
    fn fmt_d3() -> Format {
        Format {
            kind: FormatKind::DateMy,
            decimals: 0,
        }
    }

    #[test]
    fn date_dmy_renders_dd_mmm_yy_uppercase() {
        let i = intl_default();
        // Serial 36540 = 2000-01-15.
        assert_eq!(format_number(36540.0, fmt_d1(), &i, &DateFormatTable::default()), "15-JAN-00");
        // Serial 45931 = 2025-10-01 — atlas-model.xlsx Cost Model B1.
        assert_eq!(format_number(45931.0, fmt_d1(), &i, &DateFormatTable::default()), "01-OCT-25");
    }

    #[test]
    fn date_dm_renders_dd_mmm_no_year() {
        let i = intl_default();
        assert_eq!(format_number(36540.0, fmt_d2(), &i, &DateFormatTable::default()), "15-JAN");
        assert_eq!(format_number(45931.0, fmt_d2(), &i, &DateFormatTable::default()), "01-OCT");
    }

    #[test]
    fn date_my_renders_mmm_yy_no_day() {
        let i = intl_default();
        assert_eq!(format_number(36540.0, fmt_d3(), &i, &DateFormatTable::default()), "JAN-00");
        // Atlas-model.xlsx Cost Model B1 has format `m/yyyy` → DateMy.
        assert_eq!(format_number(45931.0, fmt_d3(), &i, &DateFormatTable::default()), "OCT-25");
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
        assert_eq!(format_number(0.5, fmt_d8(), &i, &DateFormatTable::default()), "12:00:00");
        assert_eq!(format_number(0.5, fmt_d9(), &i, &DateFormatTable::default()), "12:00");
        // 6 AM = 0.25.
        assert_eq!(format_number(0.25, fmt_d8(), &i, &DateFormatTable::default()), "06:00:00");
    }

    #[test]
    fn time_intl_b_uses_dot_separator() {
        let i = intl_with_time(TimeIntl::B);
        assert_eq!(format_number(0.5, fmt_d8(), &i, &DateFormatTable::default()), "12.00.00");
        assert_eq!(format_number(0.5, fmt_d9(), &i, &DateFormatTable::default()), "12.00");
    }

    #[test]
    fn time_intl_c_uses_comma_separator() {
        let i = intl_with_time(TimeIntl::C);
        assert_eq!(format_number(0.5, fmt_d8(), &i, &DateFormatTable::default()), "12,00,00");
        assert_eq!(format_number(0.5, fmt_d9(), &i, &DateFormatTable::default()), "12,00");
    }

    #[test]
    fn time_intl_d_falls_back_to_colon() {
        let i = intl_with_time(TimeIntl::D);
        assert_eq!(format_number(0.5, fmt_d8(), &i, &DateFormatTable::default()), "12:00:00");
        assert_eq!(format_number(0.5, fmt_d9(), &i, &DateFormatTable::default()), "12:00");
    }

    #[test]
    fn time_drops_date_part_at_serial_with_time() {
        // 36526.5 = 2000-01-01 12:00 noon. D8 shows time only.
        let i = intl_with_time(TimeIntl::A);
        assert_eq!(format_number(36526.5, fmt_d8(), &i, &DateFormatTable::default()), "12:00:00");
    }

    #[test]
    fn time_wrap_at_midnight_shows_zero() {
        let i = intl_with_time(TimeIntl::A);
        // 23:59:59.5 should round to 24:00:00 → 00:00:00.
        let near_midnight = 86_399.5 / 86_400.0;
        assert_eq!(format_number(near_midnight, fmt_d8(), &i, &DateFormatTable::default()), "00:00:00");
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

#[cfg(test)]
mod date_pattern_tests {
    use super::DateToken::*;
    use super::*;

    fn lit(s: &str) -> DateToken {
        Literal(s.to_string())
    }

    #[test]
    fn non_date_returns_none() {
        assert_eq!(parse_date_pattern("0.00"), None);
        assert_eq!(parse_date_pattern(""), None);
        assert_eq!(parse_date_pattern("hello"), None);
    }

    #[test]
    fn time_glyphs_disqualify() {
        // h/s are time, not date — caller decides time handling separately.
        assert_eq!(parse_date_pattern("h:mm:ss"), None);
    }

    #[test]
    fn month_year_numeric_with_slash() {
        assert_eq!(
            parse_date_pattern("m/yyyy"),
            Some(vec![MonthNum, lit("/"), Year4])
        );
    }

    #[test]
    fn month_year_letter_abbrev_with_dash() {
        assert_eq!(
            parse_date_pattern("mmm-yyyy"),
            Some(vec![MonthAbbrev, lit("-"), Year4])
        );
    }

    #[test]
    fn month_year_letter_abbrev_with_space() {
        assert_eq!(
            parse_date_pattern("mmm yyyy"),
            Some(vec![MonthAbbrev, lit(" "), Year4])
        );
    }

    #[test]
    fn dmy_letter_abbrev_four_digit_year() {
        assert_eq!(
            parse_date_pattern("dd-mmm-yyyy"),
            Some(vec![DayPadded, lit("-"), MonthAbbrev, lit("-"), Year4])
        );
    }

    #[test]
    fn mdy_numeric_four_digit_year() {
        assert_eq!(
            parse_date_pattern("m/d/yyyy"),
            Some(vec![MonthNum, lit("/"), Day, lit("/"), Year4])
        );
    }

    #[test]
    fn full_month_name() {
        assert_eq!(
            parse_date_pattern("mmmm d, yyyy"),
            Some(vec![MonthFull, lit(" "), Day, lit(", "), Year4])
        );
    }

    #[test]
    fn padded_numeric_month() {
        assert_eq!(
            parse_date_pattern("mm/dd/yyyy"),
            Some(vec![
                MonthNumPadded,
                lit("/"),
                DayPadded,
                lit("/"),
                Year4
            ])
        );
    }
}

#[cfg(test)]
mod format_custom_date_tests {
    use super::DateToken::*;
    use super::*;

    fn lit(s: &str) -> DateToken {
        Literal(s.to_string())
    }

    // Serial 45931 = 2025-10-01 (atlas-model.xlsx Cost Model B1).
    const OCT_1_2025: f64 = 45931.0;
    // Serial 36540 = 2000-01-15.
    const JAN_15_2000: f64 = 36540.0;

    #[test]
    fn numeric_month_slash_four_digit_year() {
        let tokens = vec![MonthNum, lit("/"), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "10/2025");
        assert_eq!(format_custom_date(JAN_15_2000, &tokens), "1/2000");
    }

    #[test]
    fn letter_month_dash_four_digit_year_mixed_case() {
        let tokens = vec![MonthAbbrev, lit("-"), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "Oct-2025");
        assert_eq!(format_custom_date(JAN_15_2000, &tokens), "Jan-2000");
    }

    #[test]
    fn letter_month_space_four_digit_year_mixed_case() {
        let tokens = vec![MonthAbbrev, lit(" "), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "Oct 2025");
    }

    #[test]
    fn dmy_with_padded_day_letter_month_long_year() {
        let tokens = vec![DayPadded, lit("-"), MonthAbbrev, lit("-"), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "01-Oct-2025");
        assert_eq!(format_custom_date(JAN_15_2000, &tokens), "15-Jan-2000");
    }

    #[test]
    fn mdy_numeric_unpadded_long_year() {
        let tokens = vec![MonthNum, lit("/"), Day, lit("/"), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "10/1/2025");
        assert_eq!(format_custom_date(JAN_15_2000, &tokens), "1/15/2000");
    }

    #[test]
    fn full_month_name_with_comma() {
        let tokens = vec![MonthFull, lit(" "), Day, lit(", "), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "October 1, 2025");
    }

    #[test]
    fn padded_numeric_month_and_day() {
        let tokens = vec![MonthNumPadded, lit("/"), DayPadded, lit("/"), Year4];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "10/01/2025");
        assert_eq!(format_custom_date(JAN_15_2000, &tokens), "01/15/2000");
    }

    #[test]
    fn two_digit_year() {
        let tokens = vec![DayPadded, lit("-"), MonthAbbrev, lit("-"), Year2];
        assert_eq!(format_custom_date(OCT_1_2025, &tokens), "01-Oct-25");
    }
}

#[cfg(test)]
mod date_format_table_tests {
    use super::DateToken::*;
    use super::*;

    #[test]
    fn intern_returns_zero_for_first_entry() {
        let mut t = DateFormatTable::new();
        let id = t.intern(vec![MonthNum, Literal("/".into()), Year4]);
        assert_eq!(id, 0);
        assert_eq!(t.len(), 1);
    }

    #[test]
    fn intern_dedupes_equal_token_lists() {
        let mut t = DateFormatTable::new();
        let a = t.intern(vec![MonthNum, Literal("/".into()), Year4]);
        let b = t.intern(vec![MonthNum, Literal("/".into()), Year4]);
        assert_eq!(a, b);
        assert_eq!(t.len(), 1);
    }

    #[test]
    fn intern_separates_distinct_token_lists() {
        let mut t = DateFormatTable::new();
        let a = t.intern(vec![MonthNum, Literal("/".into()), Year4]);
        let b = t.intern(vec![MonthAbbrev, Literal("-".into()), Year4]);
        assert_ne!(a, b);
        assert_eq!(t.len(), 2);
    }

    #[test]
    fn get_returns_interned_tokens() {
        let mut t = DateFormatTable::new();
        let tokens = vec![DayPadded, Literal("-".into()), MonthAbbrev, Literal("-".into()), Year4];
        let id = t.intern(tokens.clone());
        assert_eq!(t.get(id), tokens.as_slice());
    }
}
