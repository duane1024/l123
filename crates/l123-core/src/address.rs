//! Cell addresses, ranges, and sheet identifiers.
//!
//! 1-2-3 R3.x addresses sheets by letter (A, B, ..., Z, AA, ..., IV);
//! internally we store them as 0-based indices.
//! - Sheets: 0..256
//! - Columns: 0..256 (letters A..IV)
//! - Rows: 0..8192 (display as 1-based: 1..8192)

use std::fmt;

use thiserror::Error;

pub const MAX_SHEETS: u16 = 256;
pub const MAX_COLS: u16 = 256;
pub const MAX_ROWS: u32 = 8192;

#[derive(Copy, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct SheetId(pub u16);

impl SheetId {
    pub const A: SheetId = SheetId(0);

    pub fn letter(self) -> String {
        col_to_letters(self.0)
    }
}

impl fmt::Debug for SheetId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Sheet({})", self.letter())
    }
}

impl fmt::Display for SheetId {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.letter())
    }
}

#[derive(Copy, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Address {
    pub sheet: SheetId,
    pub col: u16,
    pub row: u32,
}

impl Address {
    pub const A1: Address = Address {
        sheet: SheetId::A,
        col: 0,
        row: 0,
    };

    pub fn new(sheet: SheetId, col: u16, row: u32) -> Self {
        Self { sheet, col, row }
    }

    /// Render as `A:A1` form (1-2-3 style, 1-based row).
    pub fn display_full(&self) -> String {
        format!(
            "{}:{}{}",
            self.sheet,
            col_to_letters(self.col),
            self.row + 1
        )
    }

    /// Render as `A1` form without sheet prefix.
    pub fn display_short(&self) -> String {
        format!("{}{}", col_to_letters(self.col), self.row + 1)
    }

    pub fn shifted(&self, d_col: i32, d_row: i32) -> Option<Address> {
        let c = (self.col as i32).checked_add(d_col)?;
        let r = (self.row as i32).checked_add(d_row)?;
        if c < 0 || c >= MAX_COLS as i32 || r < 0 || r >= MAX_ROWS as i32 {
            return None;
        }
        Some(Address {
            sheet: self.sheet,
            col: c as u16,
            row: r as u32,
        })
    }

    /// Parse `A:B5` or `B5` (no sheet prefix → sheet A).
    /// 3D range separators, file refs, `$` absolutes, and named ranges are
    /// not supported here — this is a plain address parser only.
    pub fn parse(s: &str) -> Result<Address, AddressError> {
        let (sheet_part, cell_part) = match s.find(':') {
            Some(i) => (&s[..i], &s[i + 1..]),
            None => ("A", s),
        };
        let sheet = SheetId(letters_to_col(sheet_part)?);
        // Split cell_part into letter and digit runs.
        let digit_start = cell_part
            .find(|c: char| c.is_ascii_digit())
            .ok_or_else(|| AddressError::Malformed(format!("no row digits: {s}")))?;
        let (col_str, row_str) = cell_part.split_at(digit_start);
        if col_str.is_empty() {
            return Err(AddressError::Malformed(format!("no column letters: {s}")));
        }
        let col = letters_to_col(col_str)?;
        let row_1based: u32 = row_str
            .parse()
            .map_err(|_| AddressError::Malformed(format!("bad row: {row_str}")))?;
        if row_1based == 0 {
            return Err(AddressError::Malformed("row must be >= 1".into()));
        }
        let row = row_1based - 1;
        if row >= MAX_ROWS {
            return Err(AddressError::Overflow);
        }
        Ok(Address { sheet, col, row })
    }
}

impl fmt::Debug for Address {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.display_full())
    }
}

/// Inclusive rectangular range. May span sheets within the same file (3D range).
#[derive(Copy, Clone, PartialEq, Eq, Hash, Debug)]
pub struct Range {
    pub start: Address,
    pub end: Address,
}

impl Range {
    pub fn single(a: Address) -> Self {
        Self { start: a, end: a }
    }

    /// Normalize so start <= end on every axis.
    pub fn normalized(&self) -> Self {
        let (s1, s2) = ord(self.start.sheet.0, self.end.sheet.0);
        let (c1, c2) = ord(self.start.col, self.end.col);
        let (r1, r2) = ord(self.start.row, self.end.row);
        Range {
            start: Address {
                sheet: SheetId(s1),
                col: c1,
                row: r1,
            },
            end: Address {
                sheet: SheetId(s2),
                col: c2,
                row: r2,
            },
        }
    }

    pub fn contains(&self, a: Address) -> bool {
        let n = self.normalized();
        a.sheet >= n.start.sheet
            && a.sheet <= n.end.sheet
            && a.col >= n.start.col
            && a.col <= n.end.col
            && a.row >= n.start.row
            && a.row <= n.end.row
    }

    pub fn is_single_sheet(&self) -> bool {
        self.start.sheet == self.end.sheet
    }
}

fn ord<T: Ord>(a: T, b: T) -> (T, T) {
    if a <= b {
        (a, b)
    } else {
        (b, a)
    }
}

/// Convert a 0-based column index to A..IV letters (1-2-3 R3.x limit is 256 cols).
pub fn col_to_letters(mut c: u16) -> String {
    let mut out = Vec::with_capacity(3);
    loop {
        out.push(b'A' + (c % 26) as u8);
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    out.reverse();
    String::from_utf8(out).expect("valid ASCII")
}

/// Parse A..IV letters to a 0-based column index.
pub fn letters_to_col(s: &str) -> Result<u16, AddressError> {
    if s.is_empty() {
        return Err(AddressError::Malformed("empty column".into()));
    }
    let mut n: u32 = 0;
    for (i, c) in s.bytes().enumerate() {
        if !c.is_ascii_alphabetic() {
            return Err(AddressError::Malformed(format!("bad column char at {i}")));
        }
        let v = c.to_ascii_uppercase() - b'A';
        n = n.checked_mul(26).ok_or(AddressError::Overflow)? + v as u32 + 1;
    }
    let n = n - 1;
    if n >= MAX_COLS as u32 {
        return Err(AddressError::Overflow);
    }
    Ok(n as u16)
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum AddressError {
    #[error("address overflow")]
    Overflow,
    #[error("malformed address: {0}")]
    Malformed(String),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_letters_roundtrip() {
        for c in [0u16, 1, 25, 26, 51, 52, 255] {
            assert_eq!(letters_to_col(&col_to_letters(c)).unwrap(), c);
        }
    }

    #[test]
    fn col_letters_specific() {
        assert_eq!(col_to_letters(0), "A");
        assert_eq!(col_to_letters(25), "Z");
        assert_eq!(col_to_letters(26), "AA");
        assert_eq!(col_to_letters(51), "AZ");
        assert_eq!(col_to_letters(52), "BA");
        assert_eq!(col_to_letters(255), "IV"); // 1-2-3 R3.x max
    }

    #[test]
    fn address_display() {
        let a = Address {
            sheet: SheetId(0),
            col: 1,
            row: 4,
        };
        assert_eq!(a.display_full(), "A:B5");
        assert_eq!(a.display_short(), "B5");

        let b = Address {
            sheet: SheetId(2),
            col: 27,
            row: 0,
        };
        assert_eq!(b.display_full(), "C:AB1");
    }

    #[test]
    fn address_shift() {
        let a = Address::A1;
        assert_eq!(a.shifted(1, 1).unwrap(), Address::new(SheetId::A, 1, 1));
        assert!(a.shifted(-1, 0).is_none());
        assert!(Address::new(SheetId::A, MAX_COLS - 1, 0)
            .shifted(1, 0)
            .is_none());
    }

    #[test]
    fn range_contains() {
        let r = Range {
            start: Address::new(SheetId::A, 0, 0),
            end: Address::new(SheetId::A, 4, 4),
        };
        assert!(r.contains(Address::new(SheetId::A, 2, 2)));
        assert!(r.contains(Address::new(SheetId::A, 0, 0)));
        assert!(r.contains(Address::new(SheetId::A, 4, 4)));
        assert!(!r.contains(Address::new(SheetId::A, 5, 5)));
    }

    #[test]
    fn range_normalize() {
        let r = Range {
            start: Address::new(SheetId::A, 4, 4),
            end: Address::new(SheetId::A, 0, 0),
        };
        let n = r.normalized();
        assert_eq!(n.start, Address::new(SheetId::A, 0, 0));
        assert_eq!(n.end, Address::new(SheetId::A, 4, 4));
    }

    #[test]
    fn parse_simple() {
        assert_eq!(Address::parse("A1").unwrap(), Address::A1);
        assert_eq!(
            Address::parse("B5").unwrap(),
            Address {
                sheet: SheetId(0),
                col: 1,
                row: 4
            }
        );
    }

    #[test]
    fn parse_with_sheet() {
        assert_eq!(
            Address::parse("A:A1").unwrap(),
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 0
            }
        );
        assert_eq!(
            Address::parse("C:AB100").unwrap(),
            Address {
                sheet: SheetId(2),
                col: 27,
                row: 99
            }
        );
    }

    #[test]
    fn parse_rejects_row_zero() {
        assert!(Address::parse("A0").is_err());
    }

    #[test]
    fn parse_rejects_bare_letter() {
        assert!(Address::parse("A").is_err());
        assert!(Address::parse("A:").is_err());
    }

    #[test]
    fn parse_roundtrip() {
        for s in ["A:A1", "A:B5", "B:Z99", "C:AA1", "A:IV8192"] {
            let a = Address::parse(s).unwrap();
            assert_eq!(a.display_full(), s);
        }
    }

    #[test]
    fn range_3d() {
        let r = Range {
            start: Address::new(SheetId(0), 1, 2),
            end: Address::new(SheetId(2), 3, 4),
        };
        assert!(!r.is_single_sheet());
        assert!(r.contains(Address::new(SheetId(1), 2, 3)));
    }
}
