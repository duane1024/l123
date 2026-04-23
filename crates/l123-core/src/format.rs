//! Cell display formats and their parenthesized tags as shown in the
//! control panel (e.g. `(C2)` = Currency 2dp).  See SPEC §12.

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
    DateDmy,          // D1: DD-MMM-YY
    DateDm,           // D2: DD-MMM
    DateMy,           // D3: MMM-YY
    DateLongIntl,     // D4
    DateShortIntl,    // D5
    TimeHmsAmPm,      // D6
    TimeHmAmPm,       // D7
    TimeLongIntl,     // D8
    TimeShortIntl,    // D9
    Text,             // Show formula, not value
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
    pub const GENERAL: Format = Format { kind: FormatKind::General, decimals: 0 };
    pub const RESET: Format = Format { kind: FormatKind::Reset, decimals: 0 };

    pub fn fixed(d: u8) -> Self {
        Self { kind: FormatKind::Fixed, decimals: d.min(15) }
    }
    pub fn currency(d: u8) -> Self {
        Self { kind: FormatKind::Currency, decimals: d.min(15) }
    }
    pub fn percent(d: u8) -> Self {
        Self { kind: FormatKind::Percent, decimals: d.min(15) }
    }
    pub fn comma(d: u8) -> Self {
        Self { kind: FormatKind::Comma, decimals: d.min(15) }
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
/// widget and the control panel's cell readout call it.
pub fn format_number(n: f64, format: Format) -> String {
    use FormatKind::*;
    let d = format.decimals as usize;
    match format.kind {
        Fixed => format!("{n:.d$}"),
        Scientific => format!("{n:.d$e}"),
        Currency => {
            let sign = if n < 0.0 { "-" } else { "" };
            let abs = n.abs();
            format!("{sign}${abs:.d$}")
        }
        // Comma is "Fixed with thousands separators". Full locale-aware
        // separators land with /WGD International — for M3 we emit the
        // grouped form without locale awareness.
        Comma => with_thousands(&format!("{n:.d$}")),
        Percent => {
            let v = n * 100.0;
            format!("{v:.d$}%")
        }
        PlusMinus => plus_minus_bar(n),
        General | Reset | Automatic => crate::contents::format_number_general(n),
        // Date/Time/Text/Hidden/Label — display the underlying number
        // until later milestones teach the formatter about those.
        DateDmy | DateDm | DateMy | DateLongIntl | DateShortIntl | TimeHmsAmPm
        | TimeHmAmPm | TimeLongIntl | TimeShortIntl | Text | Hidden | LabelOnly => {
            crate::contents::format_number_general(n)
        }
    }
}

fn with_thousands(s: &str) -> String {
    // Insert commas in the integer part.  Works on Rust-style `"-1234.56"`.
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
            out.push(',');
        }
        out.push(ch);
    }
    let int_with_commas: String = out.chars().rev().collect();
    format!("{sign}{int_with_commas}{frac_part}")
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

    #[test]
    fn fixed_rounds_to_decimals() {
        assert_eq!(format_number(1.25, Format::fixed(2)), "1.25");
        // Rust uses banker's rounding (round half to even) for `{:.*}`.
        assert_eq!(format_number(1.25, Format::fixed(1)), "1.2");
        assert_eq!(format_number(1.35, Format::fixed(1)), "1.4");
        assert_eq!(format_number(1.23456, Format::fixed(2)), "1.23");
    }

    #[test]
    fn currency_has_dollar_and_decimals() {
        assert_eq!(format_number(1000.0, Format::currency(2)), "$1000.00");
        assert_eq!(format_number(42.5, Format::currency(0)), "$42"); // banker's rounding: 42
        assert_eq!(format_number(-42.5, Format::currency(0)), "-$42");
    }

    #[test]
    fn percent_multiplies_by_100() {
        assert_eq!(format_number(0.5, Format::percent(0)), "50%");
        assert_eq!(format_number(0.123, Format::percent(1)), "12.3%");
    }

    #[test]
    fn comma_inserts_thousands_separators() {
        assert_eq!(format_number(1_234_567.0, Format::comma(0)), "1,234,567");
        assert_eq!(format_number(1000.5, Format::comma(2)), "1,000.50");
        assert_eq!(format_number(-1000.5, Format::comma(2)), "-1,000.50");
        assert_eq!(format_number(999.0, Format::comma(0)), "999");
    }

    #[test]
    fn general_matches_existing_behavior() {
        assert_eq!(format_number(123.0, Format::GENERAL), "123");
        assert_eq!(format_number(1.5, Format::GENERAL), "1.5");
    }

    #[test]
    fn scientific_uses_e_notation() {
        assert_eq!(format_number(1234.5, Format { kind: FormatKind::Scientific, decimals: 2 }), "1.23e3");
    }

    #[test]
    fn plus_minus_bar_draws_bars() {
        assert_eq!(format_number(3.0, Format { kind: FormatKind::PlusMinus, decimals: 0 }), "+++");
        assert_eq!(format_number(-2.0, Format { kind: FormatKind::PlusMinus, decimals: 0 }), "--");
        assert_eq!(format_number(0.0, Format { kind: FormatKind::PlusMinus, decimals: 0 }), "");
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
