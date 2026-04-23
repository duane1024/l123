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
