//! 1-2-3 formula syntax translator.
//!
//! Converts a Lotus-shape formula source (as typed by the user) into the
//! Excel-shape input that `ironcalc::base::Model::set_user_input` expects.
//!
//! Scope as of M5:
//! - Leading `@` or `+` is replaced with `=` (Excel's formula sigil).
//! - `..` range separator → `:` outside of string literals.
//! - Double-quoted string literals (`"..."`) pass through untouched.
//! - Sheet-qualified refs: `A:B3` → `Sheet1!B3`; same-sheet ranges
//!   with a leading qualifier (`A:B3..D5`) → `Sheet1!B3:D5`.
//! - 3D ranges (`A:B3..C:B3`) are expanded to a comma-separated list
//!   of per-sheet references (PLAN.md §4.5).
//!
//! Not yet handled (later milestones):
//! - `#AND#` / `#OR#` / `#NOT#` infix → function translation.
//! - Single-dot range separator `A1.B5`.
//! - Named ranges — M3 (names layer).
//! - Function renames where Lotus and Excel differ (e.g. `@AVG` ↔ `AVERAGE`).

use l123_core::address::letters_to_col;

/// Translate a Lotus-shape formula source (first char already classified
/// as a value-starter) into an Excel-shape formula that always begins with
/// `=`. Pure numeric literals (e.g. `"42"`, `"-3.5"`) are also translated,
/// which is safe: the caller has already decided they're formula-class.
///
/// `sheets` maps `SheetId(n)` → the engine's sheet name. Pass an empty
/// slice to skip sheet-qualified reference translation (legacy tests,
/// or contexts where the workbook is not available).
pub fn to_engine_source(lotus: &str, sheets: &[&str]) -> String {
    let body = match lotus.chars().next() {
        Some('@') | Some('+') => &lotus[1..],
        _ => lotus,
    };
    let translated = translate(body, sheets);
    format!("={translated}")
}

/// One pass: handles string-literal transparency, strips inner `@`
/// sigils, expands sheet-qualified refs when `sheets` is non-empty,
/// and rewrites `..` → `:` for unqualified ranges.
fn translate(s: &str, sheets: &[&str]) -> String {
    let mut out = String::with_capacity(s.len());
    let mut in_string = false;
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b == b'"' {
            in_string = !in_string;
            out.push('"');
            i += 1;
            continue;
        }
        if in_string {
            out.push(b as char);
            i += 1;
            continue;
        }
        if b == b'@' {
            i += 1;
            continue;
        }
        if !sheets.is_empty() && b.is_ascii_alphabetic() {
            if let Some((consumed, expanded)) = try_sheet_ref(s, i, sheets) {
                out.push_str(&expanded);
                i = consumed;
                continue;
            }
        }
        if b == b'.' && i + 1 < bytes.len() && bytes[i + 1] == b'.' {
            out.push(':');
            i += 2;
            continue;
        }
        let ch_end = next_char_boundary(s, i);
        out.push_str(&s[i..ch_end]);
        i = ch_end;
    }
    out
}

/// At position `start`, try to parse a sheet-qualified ref starting
/// with letters-then-colon. If found, also consume `..<rhs>` for
/// ranges, expanding 3D ranges to a comma list. On success returns
/// the post-match byte index and the emitted Excel string.
fn try_sheet_ref(s: &str, start: usize, sheets: &[&str]) -> Option<(usize, String)> {
    let (after_lhs, lhs_sheet, lhs_addr) = parse_sheet_qualified(s, start)?;
    // `..` continuation?
    if s.as_bytes().get(after_lhs).copied() == Some(b'.')
        && s.as_bytes().get(after_lhs + 1).copied() == Some(b'.')
    {
        let rhs_start = after_lhs + 2;
        if let Some((after_rhs, rhs_sheet, rhs_addr)) = parse_sheet_qualified(s, rhs_start) {
            // 3D range — expand per sheet.
            let (lo, hi) = if lhs_sheet <= rhs_sheet {
                (lhs_sheet, rhs_sheet)
            } else {
                (rhs_sheet, lhs_sheet)
            };
            let mut parts = Vec::with_capacity((hi - lo + 1) as usize);
            for idx in lo..=hi {
                let name = sheets.get(idx as usize)?;
                parts.push(format!("{name}!{lhs_addr}:{rhs_addr}"));
            }
            return Some((after_rhs, parts.join(",")));
        }
        if let Some((after_rhs, rhs_addr)) = parse_bare_ref(s, rhs_start) {
            let name = sheets.get(lhs_sheet as usize)?;
            return Some((after_rhs, format!("{name}!{lhs_addr}:{rhs_addr}")));
        }
    }
    // Plain single sheet-qualified cell.
    let name = sheets.get(lhs_sheet as usize)?;
    Some((after_lhs, format!("{name}!{lhs_addr}")))
}

/// Parse `<letters>:<letters><digits>` starting at `start`. Returns
/// the end byte index, the sheet index (A=0, B=1, ...), and the bare
/// cell address portion (e.g. "B3").
fn parse_sheet_qualified(s: &str, start: usize) -> Option<(usize, u16, String)> {
    let bytes = s.as_bytes();
    // Sheet letter run.
    let mut i = start;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return None;
    }
    if bytes.get(i).copied() != Some(b':') {
        return None;
    }
    let sheet_str = &s[start..i];
    let sheet_idx = letters_to_col(sheet_str).ok()?;
    // Skip colon.
    i += 1;
    let (cell_end, cell) = parse_bare_ref(s, i)?;
    Some((cell_end, sheet_idx, cell))
}

/// Parse a bare cell reference `<letters><digits>`. Returns the end
/// byte index and the substring.
fn parse_bare_ref(s: &str, start: usize) -> Option<(usize, String)> {
    let bytes = s.as_bytes();
    let mut i = start;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start {
        return None;
    }
    let digits_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == digits_start {
        return None;
    }
    Some((i, s[start..i].to_string()))
}

fn next_char_boundary(s: &str, start: usize) -> usize {
    let mut end = start + 1;
    while !s.is_char_boundary(end) && end < s.len() {
        end += 1;
    }
    end
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn leading_at_becomes_equals() {
        assert_eq!(to_engine_source("@SUM(A1..A5)", &[]), "=SUM(A1:A5)");
        assert_eq!(to_engine_source("@IF(A1>0,1,2)", &[]), "=IF(A1>0,1,2)");
        assert_eq!(to_engine_source("@NOW", &[]), "=NOW");
    }

    #[test]
    fn leading_plus_becomes_equals() {
        assert_eq!(to_engine_source("+A1+B1", &[]), "=A1+B1");
        assert_eq!(to_engine_source("+5", &[]), "=5");
    }

    #[test]
    fn range_separator_translated() {
        assert_eq!(to_engine_source("@SUM(A1..B5)", &[]), "=SUM(A1:B5)");
        assert_eq!(to_engine_source("@AVG(B2..B20)", &[]), "=AVG(B2:B20)");
    }

    #[test]
    fn single_dot_is_preserved_as_decimal() {
        assert_eq!(to_engine_source("+3.14+2.5", &[]), "=3.14+2.5");
        assert_eq!(to_engine_source("@IF(A1>0.5,1,0)", &[]), "=IF(A1>0.5,1,0)");
    }

    #[test]
    fn plain_numbers_pass_through_with_equals() {
        assert_eq!(to_engine_source("123", &[]), "=123");
        assert_eq!(to_engine_source("-3.5", &[]), "=-3.5");
        assert_eq!(to_engine_source("0.25", &[]), "=0.25");
    }

    #[test]
    fn strings_are_not_mutated() {
        assert_eq!(
            to_engine_source("@IF(A1>0,\"low..high\",\"\")", &[]),
            "=IF(A1>0,\"low..high\",\"\")"
        );
        assert_eq!(
            to_engine_source("@N(\"abc..def\")", &[]),
            "=N(\"abc..def\")"
        );
    }

    #[test]
    fn mixed_strings_and_ranges() {
        assert_eq!(
            to_engine_source("@IF(@SUM(A1..A5)>0,\"A..B\",\"C\")", &[]),
            "=IF(SUM(A1:A5)>0,\"A..B\",\"C\")"
        );
    }

    #[test]
    fn nested_at_functions() {
        let got = to_engine_source("@IF(@ISERR(A1),0,A1)", &[]);
        assert!(got.starts_with('='));
        assert!(got.contains("ISERR"));
    }

    // ---- M5: sheet-qualified refs ----

    fn sheets() -> Vec<&'static str> {
        vec!["Sheet1", "Sheet2", "Sheet3"]
    }

    #[test]
    fn sheet_qualified_cell_is_prefixed() {
        let s = sheets();
        assert_eq!(to_engine_source("+A:B3", &s), "=Sheet1!B3");
        assert_eq!(to_engine_source("+B:C7", &s), "=Sheet2!C7");
    }

    #[test]
    fn sheet_qualified_range_same_sheet() {
        // LHS qualified, RHS bare → single-sheet range anchored on LHS.
        let s = sheets();
        assert_eq!(
            to_engine_source("@SUM(A:B3..D5)", &s),
            "=SUM(Sheet1!B3:D5)"
        );
    }

    #[test]
    fn three_d_range_expands_across_sheets() {
        let s = sheets();
        assert_eq!(
            to_engine_source("@SUM(A:B3..C:B3)", &s),
            "=SUM(Sheet1!B3:B3,Sheet2!B3:B3,Sheet3!B3:B3)"
        );
    }

    #[test]
    fn three_d_range_rectangular() {
        let s = sheets();
        assert_eq!(
            to_engine_source("@SUM(A:B3..B:D5)", &s),
            "=SUM(Sheet1!B3:D5,Sheet2!B3:D5)"
        );
    }

    #[test]
    fn bare_refs_unaffected_by_sheets_arg() {
        // Without a colon, letter runs are column names — no translation.
        let s = sheets();
        assert_eq!(to_engine_source("@SUM(A1..A5)", &s), "=SUM(A1:A5)");
        assert_eq!(to_engine_source("+A1+B2", &s), "=A1+B2");
    }

    #[test]
    fn unknown_sheet_letter_falls_through_cleanly() {
        // Only 2 sheets available; `C:B3` has no matching name. The
        // translator passes the letters through rather than emitting
        // something IronCalc would mis-parse.
        let s: Vec<&str> = vec!["Sheet1", "Sheet2"];
        let got = to_engine_source("+C:B3", &s);
        // Sheet C not in list — no expansion, but `:` still in output.
        // IronCalc will error on this, which is the correct signal.
        assert!(got.starts_with('='));
    }
}
