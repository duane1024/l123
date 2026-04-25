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

use l123_core::address::{col_to_letters, letters_to_col, MAX_COLS, MAX_ROWS};

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
                parts.push(format!(
                    "{name}!{lhs_addr}:{rhs_addr}",
                    name = quote_sheet_name(name)
                ));
            }
            return Some((after_rhs, parts.join(",")));
        }
        if let Some((after_rhs, rhs_addr)) = parse_bare_ref(s, rhs_start) {
            let name = sheets.get(lhs_sheet as usize)?;
            return Some((
                after_rhs,
                format!(
                    "{name}!{lhs_addr}:{rhs_addr}",
                    name = quote_sheet_name(name)
                ),
            ));
        }
    }
    // Plain single sheet-qualified cell.
    let name = sheets.get(lhs_sheet as usize)?;
    Some((
        after_lhs,
        format!("{name}!{lhs_addr}", name = quote_sheet_name(name)),
    ))
}

/// Wrap a sheet name in single quotes per Excel's formula grammar if
/// it contains any character outside the ASCII-identifier set
/// (letters, digits, underscore). Embedded `'` is doubled.
fn quote_sheet_name(name: &str) -> String {
    let needs_quote = name
        .chars()
        .any(|c| !(c.is_ascii_alphanumeric() || c == '_'));
    if !needs_quote {
        return name.to_string();
    }
    let mut out = String::with_capacity(name.len() + 2);
    out.push('\'');
    for c in name.chars() {
        if c == '\'' {
            out.push_str("''");
        } else {
            out.push(c);
        }
    }
    out.push('\'');
    out
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

/// Shift every relative cell reference in a Lotus-form `formula` by
/// `(dx, dy)` columns and rows. Used by `/Copy` to update formula
/// references when a cell is pasted to a new location.
///
/// Reference shapes recognised: bare `A1`, column-absolute `$A1`,
/// row-absolute `A$1`, full-absolute `$A$1`, sheet-qualified `A:B5`
/// (with `$` permitted on the cell part), and `..` ranges (each end
/// shifted independently). Sheet parts are preserved verbatim.
///
/// String literals (`"..."`) are passed through unchanged. Identifiers
/// without trailing digits (function names, named ranges) are not
/// treated as refs. References that would shift out of the addressable
/// space (`MAX_COLS` / `MAX_ROWS`) emit `ERR` — Lotus's error sigil.
pub fn shift_refs(formula: &str, dx: i32, dy: i32) -> String {
    let mut out = String::with_capacity(formula.len());
    let mut in_string = false;
    let bytes = formula.as_bytes();
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
        if let Some((shifted, end)) = try_shift_ref_at(formula, i, dx, dy) {
            out.push_str(&shifted);
            i = end;
            continue;
        }
        let ch_end = next_char_boundary(formula, i);
        out.push_str(&formula[i..ch_end]);
        i = ch_end;
    }
    out
}

/// Try to parse a cell reference (possibly sheet-qualified, with
/// optional `$` absolutes) starting at byte `start` and emit it
/// shifted by `(dx, dy)`. Returns `None` if no ref matches — caller
/// then emits one character verbatim.
fn try_shift_ref_at(s: &str, start: usize, dx: i32, dy: i32) -> Option<(String, usize)> {
    // First try sheet-qualified, then bare. Each branch fully validates
    // before consuming so the failure of one doesn't leak.
    if let Some(r) = parse_qualified_ref(s, start) {
        return Some(emit_shifted(r, dx, dy));
    }
    let r = parse_bare_cell_ref(s, start)?;
    Some(emit_shifted(r, dx, dy))
}

#[derive(Debug)]
struct ParsedRef {
    sheet: Option<String>,
    col_abs: bool,
    col: u16,
    row_abs: bool,
    row: u32,
    end: usize,
}

/// Parse `<letters>:[$]<letters>[$]<digits>` starting at `start`.
fn parse_qualified_ref(s: &str, start: usize) -> Option<ParsedRef> {
    let bytes = s.as_bytes();
    // Sheet letter run.
    let mut i = start;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start || bytes.get(i).copied() != Some(b':') {
        return None;
    }
    let sheet_str = &s[start..i];
    // Validate sheet letters (overflow → not a ref).
    letters_to_col(sheet_str).ok()?;
    i += 1; // consume colon
    let mut cell = parse_bare_cell_ref(s, i)?;
    cell.sheet = Some(sheet_str.to_string());
    Some(cell)
}

/// Parse `[$]<letters>[$]<digits>` starting at `start`.
fn parse_bare_cell_ref(s: &str, start: usize) -> Option<ParsedRef> {
    let bytes = s.as_bytes();
    let mut i = start;
    let col_abs = bytes.get(i).copied() == Some(b'$');
    if col_abs {
        i += 1;
    }
    let col_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == col_start {
        return None;
    }
    let col_letters = &s[col_start..i];
    let col = letters_to_col(col_letters).ok()?;
    let row_abs = bytes.get(i).copied() == Some(b'$');
    if row_abs {
        i += 1;
    }
    let row_start = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == row_start {
        return None;
    }
    let row: u32 = s[row_start..i].parse().ok()?;
    if row == 0 {
        return None;
    }
    Some(ParsedRef {
        sheet: None,
        col_abs,
        col,
        row_abs,
        row,
        end: i,
    })
}

fn emit_shifted(r: ParsedRef, dx: i32, dy: i32) -> (String, usize) {
    let new_col = if r.col_abs {
        r.col as i32
    } else {
        r.col as i32 + dx
    };
    let new_row = if r.row_abs {
        r.row as i32
    } else {
        r.row as i32 + dy
    };
    if new_col < 0 || new_col >= MAX_COLS as i32 || new_row < 1 || new_row > MAX_ROWS as i32 {
        return ("ERR".to_string(), r.end);
    }
    let mut out = String::new();
    if let Some(sheet) = r.sheet {
        out.push_str(&sheet);
        out.push(':');
    }
    if r.col_abs {
        out.push('$');
    }
    out.push_str(&col_to_letters(new_col as u16));
    if r.row_abs {
        out.push('$');
    }
    out.push_str(&new_row.to_string());
    (out, r.end)
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
        assert_eq!(to_engine_source("@SUM(A:B3..D5)", &s), "=SUM(Sheet1!B3:D5)");
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
    fn sheet_name_with_space_is_single_quoted() {
        // Excel requires names containing spaces / punctuation to be
        // wrapped in single quotes: `'Q1 Sales'!B3`. Exercises the path
        // that matters after loading an xlsx authored in Excel.
        let s: Vec<&str> = vec!["Q1 Sales", "Q2 Budget"];
        assert_eq!(to_engine_source("+A:B3", &s), "='Q1 Sales'!B3");
        assert_eq!(
            to_engine_source("@SUM(A:B3..D5)", &s),
            "=SUM('Q1 Sales'!B3:D5)"
        );
        assert_eq!(
            to_engine_source("@SUM(A:B3..B:B3)", &s),
            "=SUM('Q1 Sales'!B3:B3,'Q2 Budget'!B3:B3)"
        );
    }

    #[test]
    fn sheet_name_with_apostrophe_is_escaped() {
        // Single quotes inside a quoted name are doubled per Excel's
        // formula grammar: `'It''s Here'!A1`.
        let s: Vec<&str> = vec!["It's Here"];
        assert_eq!(to_engine_source("+A:A1", &s), "='It''s Here'!A1");
    }

    #[test]
    fn sheet_name_without_special_chars_is_not_quoted() {
        // Plain identifiers (letters/digits/underscore) don't need
        // quoting — keep output tight.
        let s: Vec<&str> = vec!["Sheet1", "Q2"];
        assert_eq!(to_engine_source("+A:B3", &s), "=Sheet1!B3");
        assert_eq!(to_engine_source("+B:B3", &s), "=Q2!B3");
    }

    // ---- shift_refs ----

    #[test]
    fn shift_simple_relative_cell() {
        assert_eq!(shift_refs("+A1", 1, 0), "+B1");
        assert_eq!(shift_refs("+A1", 0, 1), "+A2");
        assert_eq!(shift_refs("+A1", 3, 2), "+D3");
    }

    #[test]
    fn shift_preserves_leading_sigil() {
        assert_eq!(shift_refs("@SUM(A1..A5)", 1, 0), "@SUM(B1..B5)");
        assert_eq!(shift_refs("+A1+B2", 0, 2), "+A3+B4");
    }

    #[test]
    fn shift_absolute_col_keeps_col() {
        // $ before letter freezes the column.
        assert_eq!(shift_refs("+$A1", 5, 0), "+$A1");
        assert_eq!(shift_refs("+$A1", 0, 3), "+$A4");
    }

    #[test]
    fn shift_absolute_row_keeps_row() {
        // $ before digits freezes the row.
        assert_eq!(shift_refs("+A$1", 0, 5), "+A$1");
        assert_eq!(shift_refs("+A$1", 3, 0), "+D$1");
    }

    #[test]
    fn shift_full_absolute_no_change() {
        assert_eq!(shift_refs("+$A$1", 5, 5), "+$A$1");
        assert_eq!(shift_refs("@SUM($A$1..$B$3)", 10, 10), "@SUM($A$1..$B$3)");
    }

    #[test]
    fn shift_range_both_ends() {
        assert_eq!(shift_refs("@SUM(A1..C3)", 2, 0), "@SUM(C1..E3)");
        assert_eq!(shift_refs("@SUM(A1..C3)", 0, 5), "@SUM(A6..C8)");
    }

    #[test]
    fn shift_skips_string_literals() {
        assert_eq!(
            shift_refs("@IF(A1>0,\"A1 is big\",\"low\")", 1, 0),
            "@IF(B1>0,\"A1 is big\",\"low\")"
        );
    }

    #[test]
    fn shift_function_names_unchanged() {
        // SUM, AVG, NOW, etc. are letter-runs without trailing digits and
        // shouldn't be confused with cell refs.
        assert_eq!(shift_refs("@NOW", 5, 5), "@NOW");
        assert_eq!(shift_refs("@SUM(A1)+@AVG(B1)", 1, 0), "@SUM(B1)+@AVG(C1)");
    }

    #[test]
    fn shift_sheet_qualified_keeps_sheet() {
        assert_eq!(shift_refs("+A:B5", 1, 1), "+A:C6");
        assert_eq!(shift_refs("@SUM(A:B5..A:D10)", 1, 0), "@SUM(A:C5..A:E10)");
    }

    #[test]
    fn shift_3d_range_keeps_sheet_span() {
        // Sheets stay at A..C; only col/row shift.
        assert_eq!(shift_refs("@SUM(A:B5..C:D10)", 1, 0), "@SUM(A:C5..C:E10)");
    }

    #[test]
    fn shift_underflow_emits_err() {
        // Shifting A1 left by 1 would yield col -1 → ERR.
        assert_eq!(shift_refs("+A1", -1, 0), "+ERR");
        assert_eq!(shift_refs("+A1", 0, -1), "+ERR");
    }

    #[test]
    fn shift_negative_within_bounds() {
        assert_eq!(shift_refs("+D5", -1, -2), "+C3");
    }

    #[test]
    fn shift_named_range_passthrough() {
        // 5-letter "name" overflows column space — emit unchanged.
        assert_eq!(shift_refs("@SUM(sales)", 1, 0), "@SUM(sales)");
    }

    #[test]
    fn shift_decimals_not_misparsed() {
        // 3.14 contains no letters → not a ref.
        assert_eq!(shift_refs("+3.14+A1", 1, 0), "+3.14+B1");
    }

    #[test]
    fn shift_zero_is_identity() {
        let s = "@IF(A1>0,$A$1+B$2..C$3,\"none\")";
        assert_eq!(shift_refs(s, 0, 0), s);
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
