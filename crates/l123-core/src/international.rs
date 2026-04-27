//! International settings: punctuation, dates, times, negative style,
//! currency. Configured via `/Worksheet Global Default Other International`.
//!
//! Source for Punctuation A-H: 1-2-3 R3.4a Reference Manual, "Selecting an
//! International Character Set" appendix.
//!
//! Persistence to `L123.CNF` via `/WGDU` (Update) is out of scope — the
//! setting is session-only until that lands.

/// Punctuation triple (decimal point, argument separator, thousands
/// separator) selected by `/WGDOI Punctuation A..H`.
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum Punctuation {
    #[default]
    A,
    B,
    C,
    D,
    E,
    F,
    G,
    H,
}

impl Punctuation {
    /// Character used as the decimal point.
    pub fn decimal_char(self) -> char {
        match self {
            Self::A | Self::C | Self::E | Self::G => '.',
            Self::B | Self::D | Self::F | Self::H => ',',
        }
    }

    /// Character separating function arguments in source form.
    /// Note: the Excel emission target always uses `,`.
    pub fn argument_sep(self) -> char {
        match self {
            Self::A | Self::C => ',',
            Self::B | Self::D => '.',
            Self::E | Self::F | Self::G | Self::H => ';',
        }
    }

    /// Character grouping thousands in display.
    pub fn thousands_sep(self) -> char {
        match self {
            Self::A | Self::E => ',',
            Self::B | Self::F => '.',
            Self::C | Self::D | Self::G | Self::H => ' ',
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::A => "A",
            Self::B => "B",
            Self::C => "C",
            Self::D => "D",
            Self::E => "E",
            Self::F => "F",
            Self::G => "G",
            Self::H => "H",
        }
    }
}

/// Date Intl A-D — selects the order/separator used by D4 (long) and
/// D5 (short).
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum DateIntl {
    #[default]
    A,
    B,
    C,
    D,
}

impl DateIntl {
    pub fn long_label(self) -> &'static str {
        match self {
            Self::A => "MM/DD/YY",
            Self::B => "DD/MM/YY",
            Self::C => "DD.MM.YY",
            Self::D => "YY-MM-DD",
        }
    }

    pub fn short_label(self) -> &'static str {
        match self {
            Self::A => "MM/DD",
            Self::B => "DD/MM",
            Self::C => "DD.MM",
            Self::D => "MM-DD",
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::A => "A",
            Self::B => "B",
            Self::C => "C",
            Self::D => "D",
        }
    }
}

/// Time Intl A-D — selects the separator used by D8 (long) and D9
/// (short). Time D renders with the colon fallback (= Time A) until
/// LICS letter glyphs (`HHhMMmSSs`) land.
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum TimeIntl {
    #[default]
    A,
    B,
    C,
    D,
}

impl TimeIntl {
    pub fn long_label(self) -> &'static str {
        match self {
            Self::A | Self::D => "HH:MM:SS",
            Self::B => "HH.MM.SS",
            Self::C => "HH,MM,SS",
        }
    }

    pub fn short_label(self) -> &'static str {
        match self {
            Self::A | Self::D => "HH:MM",
            Self::B => "HH.MM",
            Self::C => "HH,MM",
        }
    }

    pub fn label(self) -> &'static str {
        match self {
            Self::A => "A",
            Self::B => "B",
            Self::C => "C",
            Self::D => "D",
        }
    }
}

/// How negative numbers are displayed in Currency and Comma formats.
/// Other formats (Fixed, Sci, General, Percent) always use `Sign`.
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum NegativeStyle {
    #[default]
    Sign,
    Parens,
}

impl NegativeStyle {
    pub fn label(self) -> &'static str {
        match self {
            Self::Sign => "Sign",
            Self::Parens => "Parens",
        }
    }
}

/// Whether the currency symbol leads (`$1234`) or trails (`1234€`) the
/// numeric body.
#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum CurrencyPosition {
    #[default]
    Prefix,
    Suffix,
}

impl CurrencyPosition {
    pub fn label(self) -> &'static str {
        match self {
            Self::Prefix => "Prefix",
            Self::Suffix => "Suffix",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CurrencyConfig {
    pub symbol: String,
    pub position: CurrencyPosition,
}

impl Default for CurrencyConfig {
    fn default() -> Self {
        Self {
            symbol: "$".into(),
            position: CurrencyPosition::Prefix,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct International {
    pub punctuation: Punctuation,
    pub date_intl: DateIntl,
    pub time_intl: TimeIntl,
    pub negative_style: NegativeStyle,
    pub currency: CurrencyConfig,
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Pin the R3.4a Reference Manual mapping. If a row here is wrong,
    /// every cell display under non-default Punctuation is wrong.
    #[test]
    fn punctuation_table_matches_r3_4a_reference() {
        use Punctuation::*;
        let cases = [
            (A, '.', ',', ','),
            (B, ',', '.', '.'),
            (C, '.', ',', ' '),
            (D, ',', '.', ' '),
            (E, '.', ';', ','),
            (F, ',', ';', '.'),
            (G, '.', ';', ' '),
            (H, ',', ';', ' '),
        ];
        for (p, dec, arg, thou) in cases {
            assert_eq!(p.decimal_char(), dec, "decimal_char for {p:?}");
            assert_eq!(p.argument_sep(), arg, "argument_sep for {p:?}");
            assert_eq!(p.thousands_sep(), thou, "thousands_sep for {p:?}");
        }
    }

    #[test]
    fn defaults_match_us_layout() {
        let i = International::default();
        assert_eq!(i.punctuation, Punctuation::A);
        assert_eq!(i.date_intl, DateIntl::A);
        assert_eq!(i.time_intl, TimeIntl::A);
        assert_eq!(i.negative_style, NegativeStyle::Sign);
        assert_eq!(i.currency.symbol, "$");
        assert_eq!(i.currency.position, CurrencyPosition::Prefix);
    }

    #[test]
    fn date_intl_labels_cover_all_variants() {
        for d in [DateIntl::A, DateIntl::B, DateIntl::C, DateIntl::D] {
            assert!(!d.long_label().is_empty());
            assert!(!d.short_label().is_empty());
        }
    }

    #[test]
    fn time_intl_d_falls_back_to_colon() {
        assert_eq!(TimeIntl::D.long_label(), "HH:MM:SS");
        assert_eq!(TimeIntl::D.short_label(), "HH:MM");
    }
}
