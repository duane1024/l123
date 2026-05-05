//! Cell display formats and their parenthesized tags as shown in the
//! control panel (e.g. `(C2)` = Currency 2dp).  See SPEC §12.

use crate::color::RgbColor;
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
    /// `/Range Format Other Parentheses Yes` wraps the rendered numeric
    /// body in `()` regardless of sign. Independent of `kind` so it can
    /// layer on top of Fixed/Currency/Comma/etc. Ignored for the
    /// non-numeric kinds (General/Text/Hidden/LabelOnly/Automatic/Reset).
    pub parens: bool,
    /// `/Range Format Other Color Negative <color>` — when set, negative
    /// numeric values render with this foreground color, overriding any
    /// explicit per-cell font color. `None` means inherit normally.
    /// Ignored for non-numeric kinds.
    pub negative_color: Option<RgbColor>,
}

impl Format {
    pub const GENERAL: Format = Format {
        kind: FormatKind::General,
        decimals: 0,
        parens: false,
        negative_color: None,
    };
    pub const RESET: Format = Format {
        kind: FormatKind::Reset,
        decimals: 0,
        parens: false,
        negative_color: None,
    };

    pub fn fixed(d: u8) -> Self {
        Self {
            kind: FormatKind::Fixed,
            decimals: d.min(15),
            parens: false,
            negative_color: None,
        }
    }
    pub fn currency(d: u8) -> Self {
        Self {
            kind: FormatKind::Currency,
            decimals: d.min(15),
            parens: false,
            negative_color: None,
        }
    }
    pub fn percent(d: u8) -> Self {
        Self {
            kind: FormatKind::Percent,
            decimals: d.min(15),
            parens: false,
            negative_color: None,
        }
    }
    pub fn comma(d: u8) -> Self {
        Self {
            kind: FormatKind::Comma,
            decimals: d.min(15),
            parens: false,
            negative_color: None,
        }
    }

    /// Return a copy of `self` with the parentheses flag turned on.
    pub fn with_parens(self) -> Self {
        Self {
            parens: true,
            ..self
        }
    }

    /// Construct a format from `kind` with default modifiers
    /// (`decimals: 0`, no parens, no negative-color override).
    /// Convenience for kinds that don't carry a decimals count.
    pub fn from_kind(kind: FormatKind) -> Self {
        Self {
            kind,
            decimals: 0,
            parens: false,
            negative_color: None,
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
    let body = match format.kind {
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
        TimeHmsAmPm => format_time_12h(n, true),
        TimeHmAmPm => format_time_12h(n, false),
        Text | Hidden | LabelOnly => swap_decimal(crate::contents::format_number_general(n), dec),
    };
    if format.parens && parens_apply_to(format.kind) {
        format!("({body})")
    } else {
        body
    }
}

/// `parens` is meaningful only for the numeric format kinds — wrapping
/// dates/times or non-numeric tags would be visual noise.
fn parens_apply_to(kind: FormatKind) -> bool {
    use FormatKind::*;
    matches!(kind, Fixed | Scientific | Currency | Comma | Percent)
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

/// Render `serial` against an Excel date/time format string,
/// preserving the original Excel author's intent verbatim.
///
/// Used when an `.xlsx` cell carries a num_fmt string that doesn't
/// round-trip exactly through the canonical 1-2-3 D1..D9 mapping
/// (e.g. `"yyyy-mm-dd"`, `"d-mmm-yyyy"`, `"dddd"`). The override is
/// stored on the cell at load time and re-emitted at save time, so a
/// load → save with no edits leaves the format string untouched.
///
/// Recognised glyphs (case-insensitive except `AM/PM` casing rules):
/// * `y`, `yy` (2-digit year), `yyy`/`yyyy` (4-digit year)
/// * `m` runs: month or minute by context — minute when adjacent to
///   an hour or seconds token, else month. Lengths: `m`/`mm` numeric,
///   `mmm` abbrev, `mmmm` full, `mmmmm` first letter.
/// * `d`, `dd` (day of month); `ddd` (weekday abbrev), `dddd` (full).
/// * `h`, `hh` (hour: 12-hour iff section contains an AM/PM marker).
/// * `s`, `ss` (second).
/// * `AM/PM`, `am/pm`, `A/P`, `a/p` — 12-hour suffix; emitted with the
///   matched casing.
/// * `"text"` quoted literal, `\x` backslash-escape, `_x` width-spacer
///   (one space + drop next), `*x` fill-repeat (dropped).
/// * `[…]` cosmetic / locale brackets — dropped.
/// * `;` section separator — only the first section is rendered for
///   dates/times (negative/zero/text sections don't apply).
pub fn format_datetime_excel(serial: f64, fmt: &str) -> String {
    // Only the positive section drives date/time rendering.
    let section = first_section(fmt);
    let tokens = tokenize_excel_datetime(section);
    let has_ampm = tokens.iter().any(|t| matches!(t, ExcelToken::AmPm { .. }));
    render_excel_datetime(serial, &tokens, has_ampm)
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum ExcelToken {
    Year(usize),
    /// Month/minute is decided after tokenisation by neighbour scan.
    /// `len` is the run length; `is_minute` set during disambiguation.
    MorMin {
        len: usize,
        is_minute: bool,
    },
    Day(usize),
    Hour(usize),
    Second(usize),
    /// `lower=true` means render `am`/`pm`; `short=true` means `A`/`P`.
    AmPm {
        lower: bool,
        short: bool,
    },
    Literal(String),
}

/// Split `fmt` at unquoted/un-bracketed `;` and return the first
/// section. Mirrors how Excel sections are delimited.
fn first_section(fmt: &str) -> &str {
    let bytes = fmt.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                i += 1;
                while i < bytes.len() && bytes[i] != b'"' {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            b'[' => {
                while i < bytes.len() && bytes[i] != b']' {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            b'\\' => i += 2,
            b';' => return &fmt[..i],
            _ => i += 1,
        }
    }
    fmt
}

fn tokenize_excel_datetime(section: &str) -> Vec<ExcelToken> {
    let bytes = section.as_bytes();
    let mut tokens: Vec<ExcelToken> = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        let c = bytes[i];
        match c {
            b'"' => {
                i += 1;
                let start = i;
                while i < bytes.len() && bytes[i] != b'"' {
                    i += 1;
                }
                push_literal(&mut tokens, &section[start..i]);
                if i < bytes.len() {
                    i += 1;
                }
            }
            b'\\' => {
                if i + 1 < bytes.len() {
                    push_literal(&mut tokens, &section[i + 1..i + 2]);
                    i += 2;
                } else {
                    i += 1;
                }
            }
            b'_' => {
                push_literal(&mut tokens, " ");
                i += 2;
            }
            b'*' => {
                // Fill-repeat — consume next char and emit nothing.
                i += 2;
            }
            b'[' => {
                // Drop entire bracketed run; cosmetic / locale tag.
                while i < bytes.len() && bytes[i] != b']' {
                    i += 1;
                }
                if i < bytes.len() {
                    i += 1;
                }
            }
            b'y' | b'Y' => {
                let len = run_len(bytes, i, |b| matches!(b, b'y' | b'Y'));
                tokens.push(ExcelToken::Year(len));
                i += len;
            }
            b'm' | b'M' => {
                let len = run_len(bytes, i, |b| matches!(b, b'm' | b'M'));
                tokens.push(ExcelToken::MorMin {
                    len,
                    is_minute: false,
                });
                i += len;
            }
            b'd' | b'D' => {
                let len = run_len(bytes, i, |b| matches!(b, b'd' | b'D'));
                tokens.push(ExcelToken::Day(len));
                i += len;
            }
            b'h' | b'H' => {
                let len = run_len(bytes, i, |b| matches!(b, b'h' | b'H'));
                tokens.push(ExcelToken::Hour(len));
                i += len;
            }
            b's' | b'S' => {
                let len = run_len(bytes, i, |b| matches!(b, b's' | b'S'));
                tokens.push(ExcelToken::Second(len));
                i += len;
            }
            b'A' | b'a' | b'P' | b'p' => {
                if let Some((consumed, lower, short)) = match_ampm(&section[i..]) {
                    tokens.push(ExcelToken::AmPm { lower, short });
                    i += consumed;
                } else {
                    push_literal(&mut tokens, &section[i..i + 1]);
                    i += 1;
                }
            }
            _ => {
                push_literal(&mut tokens, &section[i..i + 1]);
                i += 1;
            }
        }
    }
    disambiguate_minutes(&mut tokens);
    tokens
}

fn run_len(bytes: &[u8], start: usize, pred: impl Fn(u8) -> bool) -> usize {
    let mut n = 0;
    while start + n < bytes.len() && pred(bytes[start + n]) {
        n += 1;
    }
    n
}

fn push_literal(tokens: &mut Vec<ExcelToken>, s: &str) {
    if s.is_empty() {
        return;
    }
    if let Some(ExcelToken::Literal(prev)) = tokens.last_mut() {
        prev.push_str(s);
    } else {
        tokens.push(ExcelToken::Literal(s.to_string()));
    }
}

/// Match `AM/PM`, `am/pm`, `A/P`, `a/p` at the head of `s`.
/// Returns `(consumed_bytes, lowercase, short)`.
fn match_ampm(s: &str) -> Option<(usize, bool, bool)> {
    let lower = s.to_ascii_lowercase();
    if lower.starts_with("am/pm") {
        // Casing is determined by the first byte of the actual input.
        let is_lower = s.as_bytes().first() == Some(&b'a');
        Some((5, is_lower, false))
    } else if lower.starts_with("a/p") {
        let is_lower = s.as_bytes().first() == Some(&b'a');
        Some((3, is_lower, true))
    } else {
        None
    }
}

/// Mark `MorMin` runs as minute when they neighbour a time token.
/// Excel rule: `m`/`mm` is a minute when the nearest non-literal
/// neighbour is `h`/`hh` (preceding) or `s`/`ss` (following).
fn disambiguate_minutes(tokens: &mut [ExcelToken]) {
    let n = tokens.len();
    for idx in 0..n {
        if !matches!(tokens[idx], ExcelToken::MorMin { .. }) {
            continue;
        }
        let prev = (0..idx).rev().find_map(|j| {
            if matches!(tokens[j], ExcelToken::Literal(_)) {
                None
            } else {
                Some(&tokens[j])
            }
        });
        let next = ((idx + 1)..n).find_map(|j| {
            if matches!(tokens[j], ExcelToken::Literal(_)) {
                None
            } else {
                Some(&tokens[j])
            }
        });
        let is_min = matches!(prev, Some(ExcelToken::Hour(_)))
            || matches!(next, Some(ExcelToken::Second(_)));
        if let ExcelToken::MorMin { is_minute, .. } = &mut tokens[idx] {
            *is_minute = is_min;
        }
    }
}

const MONTH_TITLE_ABBREV: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
const MONTH_TITLE_FULL: [&str; 12] = [
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
const WEEKDAY_TITLE_ABBREV: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const WEEKDAY_TITLE_FULL: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

fn render_excel_datetime(serial: f64, tokens: &[ExcelToken], has_ampm: bool) -> String {
    let (year, month, day) = serial_to_ymd(serial);
    let (h24, minute, second) = fraction_to_hms(serial);
    let weekday = weekday_index(serial);
    let mut out = String::new();
    for token in tokens {
        match token {
            ExcelToken::Year(len) => {
                if *len <= 2 {
                    let yy = (year % 100).unsigned_abs();
                    out.push_str(&format!("{yy:02}"));
                } else {
                    out.push_str(&format!("{year:04}"));
                }
            }
            ExcelToken::MorMin { len, is_minute } => {
                if *is_minute {
                    match len {
                        1 => out.push_str(&format!("{minute}")),
                        _ => out.push_str(&format!("{minute:02}")),
                    }
                } else {
                    let m = (month.saturating_sub(1) as usize).min(11);
                    match len {
                        1 => out.push_str(&format!("{month}")),
                        2 => out.push_str(&format!("{month:02}")),
                        3 => out.push_str(MONTH_TITLE_ABBREV[m]),
                        4 => out.push_str(MONTH_TITLE_FULL[m]),
                        _ => out.push(MONTH_TITLE_ABBREV[m].as_bytes()[0] as char),
                    }
                }
            }
            ExcelToken::Day(len) => {
                let w = weekday as usize;
                match len {
                    1 => out.push_str(&format!("{day}")),
                    2 => out.push_str(&format!("{day:02}")),
                    3 => out.push_str(WEEKDAY_TITLE_ABBREV[w]),
                    _ => out.push_str(WEEKDAY_TITLE_FULL[w]),
                }
            }
            ExcelToken::Hour(len) => {
                let h = if has_ampm {
                    match h24 {
                        0 => 12,
                        1..=12 => h24,
                        _ => h24 - 12,
                    }
                } else {
                    h24
                };
                match len {
                    1 => out.push_str(&format!("{h}")),
                    _ => out.push_str(&format!("{h:02}")),
                }
            }
            ExcelToken::Second(len) => match len {
                1 => out.push_str(&format!("{second}")),
                _ => out.push_str(&format!("{second:02}")),
            },
            ExcelToken::AmPm { lower, short } => {
                let pm = h24 >= 12;
                let s = match (*short, *lower, pm) {
                    (false, false, false) => "AM",
                    (false, false, true) => "PM",
                    (false, true, false) => "am",
                    (false, true, true) => "pm",
                    (true, false, false) => "A",
                    (true, false, true) => "P",
                    (true, true, false) => "a",
                    (true, true, true) => "p",
                };
                out.push_str(s);
            }
            ExcelToken::Literal(s) => out.push_str(s),
        }
    }
    out
}

/// Day-of-week index for a Lotus serial: 0 = Sunday, 6 = Saturday.
/// Serial 1 = 1900-01-01 (real-calendar Monday). The 1900 leap quirk
/// (serial 60 = fictional Feb 29) means we subtract 1 day for serials
/// ≥ 60 to align with the real-calendar weekday — matching Excel's
/// `WEEKDAY()` and the offset used by `serial_to_ymd`.
fn weekday_index(serial: f64) -> u32 {
    let days = serial.trunc() as i64;
    let adjusted = if days >= 60 { days - 1 } else { days };
    adjusted.rem_euclid(7) as u32
}

/// Render the time-of-day portion of `serial` as a 12-hour clock with
/// an `AM`/`PM` suffix. `with_seconds` toggles between D6 (`h:mm:ss
/// AM/PM`) and D7 (`h:mm AM/PM`). Hours are unpadded (1..12) to match
/// 1-2-3 R3.4 and Excel's `h:mm AM/PM`; minutes and seconds keep the
/// two-digit pad.
fn format_time_12h(serial: f64, with_seconds: bool) -> String {
    let (h24, m, s) = fraction_to_hms(serial);
    let (h12, suffix) = match h24 {
        0 => (12, "AM"),
        1..=11 => (h24, "AM"),
        12 => (12, "PM"),
        _ => (h24 - 12, "PM"),
    };
    if with_seconds {
        format!("{h12}:{m:02}:{s:02} {suffix}")
    } else {
        format!("{h12}:{m:02} {suffix}")
    }
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
        let sci2 = Format {
            kind: FormatKind::Scientific,
            decimals: 2,
            parens: false,
            negative_color: None,
        };
        assert_eq!(format_number(1234.5, sci2, &i), "1.23e3");
    }

    #[test]
    fn plus_minus_bar_draws_bars() {
        let i = intl_default();
        let pm = Format::from_kind(FormatKind::PlusMinus);
        assert_eq!(format_number(3.0, pm, &i), "+++");
        assert_eq!(format_number(-2.0, pm, &i), "--");
        assert_eq!(format_number(0.0, pm, &i), "");
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
        let sci2 = Format {
            kind: FormatKind::Scientific,
            decimals: 2,
            parens: false,
            negative_color: None,
        };
        assert_eq!(format_number(1234.5, sci2, &i), "1,23e3");
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
    fn parens_flag_wraps_currency_positive() {
        let i = intl_default();
        let f = Format::currency(2).with_parens();
        assert_eq!(format_number(1234.5, f, &i), "($1234.50)");
    }

    #[test]
    fn parens_flag_wraps_currency_negative_too() {
        // The flag wraps regardless of sign — the negative sign survives
        // inside the parens (still distinguishable from a positive).
        let i = intl_default();
        let f = Format::currency(2).with_parens();
        assert_eq!(format_number(-1234.5, f, &i), "(-$1234.50)");
    }

    #[test]
    fn parens_flag_wraps_fixed() {
        let i = intl_default();
        let f = Format::fixed(2).with_parens();
        assert_eq!(format_number(42.0, f, &i), "(42.00)");
    }

    #[test]
    fn automatic_renders_identically_to_general() {
        // SPEC §12 calls (A) "Automatic (type-sniffs)". L123's MVP
        // delegates that sniffing to entry-time inference (see
        // `parse_typed_value`'s `inferred_format`); the format-time
        // path treats Automatic as an alias of General. This test
        // locks that contract so a future change to either branch
        // keeps the two in lockstep until real sniffing lands.
        let i = intl_default();
        let auto = Format::from_kind(FormatKind::Automatic);
        for &n in &[0.0, 1.0, -1.0, 1.5, 1234567.89, 0.0001, -1e10, 36540.0] {
            assert_eq!(
                format_number(n, auto, &i),
                format_number(n, Format::GENERAL, &i),
                "Automatic must alias General for n={n}"
            );
        }
    }

    #[test]
    fn parens_flag_skips_general_and_text() {
        // General/Text/Hidden/LabelOnly/Automatic/Reset don't wrap —
        // parens is meaningful only for the numeric format kinds.
        let i = intl_default();
        let g = Format::GENERAL.with_parens();
        assert_eq!(format_number(42.0, g, &i), "42");
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
        Format::from_kind(FormatKind::DateLongIntl)
    }
    fn fmt_d5() -> Format {
        Format::from_kind(FormatKind::DateShortIntl)
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

    fn fmt_d1() -> Format {
        Format::from_kind(FormatKind::DateDmy)
    }
    fn fmt_d2() -> Format {
        Format::from_kind(FormatKind::DateDm)
    }
    fn fmt_d3() -> Format {
        Format::from_kind(FormatKind::DateMy)
    }

    #[test]
    fn date_dmy_renders_dd_mmm_yy_uppercase() {
        let i = intl_default();
        // Serial 36540 = 2000-01-15.
        assert_eq!(format_number(36540.0, fmt_d1(), &i), "15-JAN-00");
        // Serial 45931 = 2025-10-01 — atlas-model.xlsx Cost Model B1.
        assert_eq!(format_number(45931.0, fmt_d1(), &i), "01-OCT-25");
    }

    #[test]
    fn date_dm_renders_dd_mmm_no_year() {
        let i = intl_default();
        assert_eq!(format_number(36540.0, fmt_d2(), &i), "15-JAN");
        assert_eq!(format_number(45931.0, fmt_d2(), &i), "01-OCT");
    }

    #[test]
    fn date_my_renders_mmm_yy_no_day() {
        let i = intl_default();
        assert_eq!(format_number(36540.0, fmt_d3(), &i), "JAN-00");
        // Atlas-model.xlsx Cost Model B1 has format `m/yyyy` → DateMy.
        assert_eq!(format_number(45931.0, fmt_d3(), &i), "OCT-25");
    }

    use crate::international::TimeIntl;

    fn fmt_d8() -> Format {
        Format::from_kind(FormatKind::TimeLongIntl)
    }
    fn fmt_d9() -> Format {
        Format::from_kind(FormatKind::TimeShortIntl)
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

    fn fmt_d6() -> Format {
        Format::from_kind(FormatKind::TimeHmsAmPm)
    }
    fn fmt_d7() -> Format {
        Format::from_kind(FormatKind::TimeHmAmPm)
    }

    #[test]
    fn time_hms_ampm_renders_noon_as_12_pm() {
        let i = intl_default();
        assert_eq!(format_number(0.5, fmt_d6(), &i), "12:00:00 PM");
    }

    #[test]
    fn time_hm_ampm_renders_noon_as_12_pm_no_seconds() {
        let i = intl_default();
        assert_eq!(format_number(0.5, fmt_d7(), &i), "12:00 PM");
    }

    #[test]
    fn time_ampm_midnight_is_12_am() {
        let i = intl_default();
        assert_eq!(format_number(0.0, fmt_d6(), &i), "12:00:00 AM");
        assert_eq!(format_number(0.0, fmt_d7(), &i), "12:00 AM");
    }

    #[test]
    fn time_ampm_morning_hours_unpadded() {
        let i = intl_default();
        // 06:00:00 → 6:00:00 AM (no leading zero on hour).
        assert_eq!(format_number(0.25, fmt_d6(), &i), "6:00:00 AM");
        // 09:30:00 → 9:30:00 AM.
        let nine_thirty = (9.0 * 3_600.0 + 30.0 * 60.0) / 86_400.0;
        assert_eq!(format_number(nine_thirty, fmt_d6(), &i), "9:30:00 AM");
    }

    #[test]
    fn time_ampm_afternoon_subtracts_12() {
        let i = intl_default();
        // 13:30:00 → 1:30:00 PM.
        let one_thirty_pm = (13.0 * 3_600.0 + 30.0 * 60.0) / 86_400.0;
        assert_eq!(format_number(one_thirty_pm, fmt_d6(), &i), "1:30:00 PM");
        assert_eq!(format_number(one_thirty_pm, fmt_d7(), &i), "1:30 PM");
    }

    #[test]
    fn time_ampm_eleven_fifty_nine_pm() {
        let i = intl_default();
        // 23:59:59.
        let almost_midnight = (23.0 * 3_600.0 + 59.0 * 60.0 + 59.0) / 86_400.0;
        assert_eq!(format_number(almost_midnight, fmt_d6(), &i), "11:59:59 PM");
    }

    #[test]
    fn time_ampm_drops_date_part() {
        let i = intl_default();
        // 36526.5 = 2000-01-01 noon. D6 shows time only.
        assert_eq!(format_number(36526.5, fmt_d6(), &i), "12:00:00 PM");
    }

    #[test]
    fn time_ampm_wrap_to_midnight_shows_12_am() {
        let i = intl_default();
        // Just under 24:00:00 should round to 24:00:00 → 00:00:00 → 12:00:00 AM.
        let near_midnight = 86_399.5 / 86_400.0;
        assert_eq!(format_number(near_midnight, fmt_d6(), &i), "12:00:00 AM");
    }
}

#[cfg(test)]
mod excel_datetime_tests {
    use super::*;

    // Serial 36540 = 2000-01-15 (Saturday). Serial 45931 = 2025-10-01.
    const JAN_15_2000: f64 = 36540.0;
    const OCT_1_2025: f64 = 45931.0;

    #[test]
    fn iso_date_yyyy_mm_dd() {
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "yyyy-mm-dd"),
            "2000-01-15"
        );
    }

    #[test]
    fn unpadded_yyyy_m_d() {
        assert_eq!(format_datetime_excel(JAN_15_2000, "yyyy-m-d"), "2000-1-15");
    }

    #[test]
    fn padded_yy_renders_two_digits() {
        assert_eq!(format_datetime_excel(JAN_15_2000, "yy"), "00");
        assert_eq!(format_datetime_excel(OCT_1_2025, "yy"), "25");
    }

    #[test]
    fn d_mmm_yyyy_with_letter_month() {
        // Excel: "1-Oct-2025" — title-case month abbrev.
        assert_eq!(
            format_datetime_excel(OCT_1_2025, "d-mmm-yyyy"),
            "1-Oct-2025"
        );
    }

    #[test]
    fn mmmm_renders_full_month_name() {
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "mmmm d, yyyy"),
            "January 15, 2000"
        );
    }

    #[test]
    fn mmmmm_renders_initial_letter() {
        assert_eq!(format_datetime_excel(JAN_15_2000, "mmmmm"), "J");
        assert_eq!(format_datetime_excel(OCT_1_2025, "mmmmm"), "O");
    }

    #[test]
    fn dddd_renders_full_weekday() {
        // 2000-01-15 was a Saturday.
        assert_eq!(format_datetime_excel(JAN_15_2000, "dddd"), "Saturday");
    }

    #[test]
    fn ddd_renders_short_weekday() {
        assert_eq!(format_datetime_excel(JAN_15_2000, "ddd"), "Sat");
    }

    #[test]
    fn time_24_hour_hh_mm_ss() {
        // 0.5 = noon.
        assert_eq!(format_datetime_excel(0.5, "hh:mm:ss"), "12:00:00");
    }

    #[test]
    fn time_12_hour_h_mm_ampm() {
        assert_eq!(format_datetime_excel(0.5, "h:mm AM/PM"), "12:00 PM");
        // 0.25 = 6 AM.
        assert_eq!(format_datetime_excel(0.25, "h:mm AM/PM"), "6:00 AM");
    }

    #[test]
    fn time_12_hour_lowercase_ampm() {
        assert_eq!(format_datetime_excel(0.5, "h:mm am/pm"), "12:00 pm");
        assert_eq!(format_datetime_excel(0.25, "h:mm am/pm"), "6:00 am");
    }

    #[test]
    fn time_12_hour_short_a_p() {
        assert_eq!(format_datetime_excel(0.5, "h:mm A/P"), "12:00 P");
        assert_eq!(format_datetime_excel(0.25, "h:mm a/p"), "6:00 a");
    }

    #[test]
    fn quoted_literal_preserved() {
        // Quoted "-" in `yyyy"-"mm"-"dd` should pass through.
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "yyyy\"-\"mm\"-\"dd"),
            "2000-01-15"
        );
        // Multi-character literal.
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "mmm \"of\" yyyy"),
            "Jan of 2000"
        );
    }

    #[test]
    fn backslash_escape_passes_next_char() {
        // `\m` should render literal "m", not month.
        assert_eq!(format_datetime_excel(JAN_15_2000, "yyyy\\-mm"), "2000-01");
    }

    #[test]
    fn locale_bracket_is_dropped() {
        // `[$-409]` is a locale tag and should not appear.
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "[$-409]m/d/yyyy"),
            "1/15/2000"
        );
    }

    #[test]
    fn minute_disambiguation_after_hour() {
        // `mm` after `h:` is minute, not month.
        assert_eq!(format_datetime_excel(0.5, "h:mm"), "12:00");
    }

    #[test]
    fn minute_disambiguation_before_seconds() {
        // `mm` before `:ss` is minute.
        let almost_one = (3_600.0 + 30.0 * 60.0 + 45.0) / 86_400.0;
        assert_eq!(format_datetime_excel(almost_one, "mm:ss"), "30:45");
    }

    #[test]
    fn month_when_alone_with_date_glyphs() {
        // No h/s context — `m` is month.
        assert_eq!(format_datetime_excel(JAN_15_2000, "m/d/yyyy"), "1/15/2000");
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "mm/dd/yyyy"),
            "01/15/2000"
        );
    }

    #[test]
    fn underscore_spacer_emits_space_consumes_next() {
        // Excel `_x` reserves space the width of x — render as a single space.
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "mm/dd/yyyy_)"),
            "01/15/2000 "
        );
    }

    #[test]
    fn first_section_used_for_dates() {
        // Sections are positive;negative;zero;text. For dates, only positive matters.
        assert_eq!(
            format_datetime_excel(JAN_15_2000, "yyyy-mm-dd;@"),
            "2000-01-15"
        );
    }

    #[test]
    fn d3_style_mmm_yy_round_trip_format() {
        // The atlas-model.xlsx fixture uses `mmm-yyyy`. Display should match
        // Excel's intent: "Oct-2025".
        assert_eq!(format_datetime_excel(OCT_1_2025, "mmm-yyyy"), "Oct-2025");
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
