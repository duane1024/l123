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
//! - `@`-name → IronCalc-name translation table (renames, arg-fixes,
//!   and emulations). Catalog in `docs/AT_FUNCTIONS.md`.

use l123_core::address::{col_to_letters, letters_to_col, MAX_COLS, MAX_ROWS};

/// Locale-specific punctuation the parser must honor when translating
/// to Excel input. Defaults to Punct A: `,` arg-sep, `.` decimal —
/// Excel's native shape, and a no-op translation.
///
/// Built from `l123_core::Punctuation` by the caller; kept as a small
/// owned struct here so `l123-parse` doesn't take a dependency on
/// `l123-core::International` itself.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub struct ParseConfig {
    /// Character separating function arguments in Lotus source.
    pub argument_sep: char,
    /// Character used as the decimal point in numeric literals.
    pub decimal_point: char,
}

impl Default for ParseConfig {
    fn default() -> Self {
        Self {
            argument_sep: ',',
            decimal_point: '.',
        }
    }
}

/// Translate a Lotus-shape formula source (first char already classified
/// as a value-starter) into an Excel-shape formula that always begins with
/// `=`. Pure numeric literals (e.g. `"42"`, `"-3.5"`) are also translated,
/// which is safe: the caller has already decided they're formula-class.
///
/// `sheets` maps `SheetId(n)` → the engine's sheet name. Pass an empty
/// slice to skip sheet-qualified reference translation (legacy tests,
/// or contexts where the workbook is not available).
///
/// Uses default `ParseConfig` (Punct A: `,` arg-sep, `.` decimal). For
/// non-default Punctuation, call [`to_engine_source_with_config`].
pub fn to_engine_source(lotus: &str, sheets: &[&str]) -> String {
    to_engine_source_with_config(lotus, sheets, &ParseConfig::default())
}

/// Like [`to_engine_source`], but threads a [`ParseConfig`] for
/// non-default Punctuation A-H. The user's argument separator is
/// translated to Excel's `,`; the user's decimal point is translated
/// to Excel's `.`.
pub fn to_engine_source_with_config(lotus: &str, sheets: &[&str], cfg: &ParseConfig) -> String {
    // `+` is just a value-starter sigil; strip it before translating.
    // `@` is left in place so the function-name translator inside
    // `translate` can recognise it.
    let body = match lotus.chars().next() {
        Some('+') => &lotus[1..],
        _ => lotus,
    };
    let translated = translate(body, sheets, cfg);
    format!("={translated}")
}

/// One pass: handles string-literal transparency, translates `@name`
/// function calls (renames, niladic-paren completion, `@@` → INDIRECT),
/// expands sheet-qualified refs when `sheets` is non-empty, rewrites
/// `..` → `:` for unqualified ranges, and translates the user's
/// argument-separator and decimal-point to Excel's `,` and `.`.
fn translate(s: &str, sheets: &[&str], cfg: &ParseConfig) -> String {
    let mut out = String::with_capacity(s.len());
    let mut in_string = false;
    let bytes = s.as_bytes();
    let mut i = 0;
    let arg_sep_byte = cfg.argument_sep as u32;
    let dec_byte = cfg.decimal_point as u32;
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
            if let Some((emitted, end)) = try_at_function(s, i, sheets, cfg) {
                out.push_str(&emitted);
                i = end;
                continue;
            }
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
        // Range separator `..` always wins over single-`.` arg-sep.
        // Preserves the design where `A1..A5` is a range under every
        // Punctuation, even when arg-sep is `.` (Punct B/D).
        if b == b'.' && i + 1 < bytes.len() && bytes[i + 1] == b'.' {
            out.push(':');
            i += 2;
            continue;
        }
        // Translate the user's argument separator to Excel's `,`.
        // Skipped when arg-sep is already `,` (Punct A/C).
        if cfg.argument_sep != ',' && (b as u32) == arg_sep_byte {
            out.push(',');
            i += 1;
            continue;
        }
        // Translate the user's decimal point to Excel's `.`. Skipped
        // when decimal is already `.` (Punct A/C/E/G).
        if cfg.decimal_point != '.' && (b as u32) == dec_byte {
            out.push('.');
            i += 1;
            continue;
        }
        let ch_end = next_char_boundary(s, i);
        out.push_str(&s[i..ch_end]);
        i = ch_end;
    }
    out
}

/// 1-2-3 → IronCalc function-name renames where the names diverge.
///
/// Catalog and rationale in `docs/AT_FUNCTIONS.md`. Lookup keys are
/// upper-case; the user's source casing is irrelevant. Names not in
/// this table pass through unchanged (the `@` is stripped by the
/// caller).
const FN_RENAMES: &[(&str, &str)] = &[
    ("AVG", "AVERAGE"),
    ("COUNT", "COUNTA"),
    ("STD", "STDEV.P"),
    ("STDS", "STDEV.S"),
    ("VAR", "VAR.P"),
    ("VARS", "VAR.S"),
    ("ISSTRING", "ISTEXT"),
    ("LENGTH", "LEN"),
    ("REPEAT", "REPT"),
    ("COLS", "COLUMNS"),
    // Database family — same swap pattern: 1-2-3 `@D…S` is sample,
    // bare `@D…` is population. IronCalc's names invert the suffixing
    // (DSTDEV is sample, DSTDEVP is population).
    ("DAVG", "DAVERAGE"),
    ("DSTD", "DSTDEVP"),
    ("DSTDS", "DSTDEV"),
    ("DVAR", "DVARP"),
    ("DVARS", "DVAR"),
    // 1-2-3 @CODE returns the LICS/ASCII code of the first char.
    // IronCalc 0.7.1 ships only `UNICODE`, which is identical for
    // ASCII; close-enough mapping until/unless a `CODE` shim lands.
    ("CODE", "UNICODE"),
];

/// 1-2-3 niladic functions — written without parens in 1-2-3, but
/// Excel requires `()`. The translator appends `()` unless the user
/// already wrote them. `@ERR` is also niladic but has no Excel
/// equivalent, so it's emulated in PR3 rather than completed here.
const NILADIC: &[&str] = &["PI", "NOW", "TODAY", "RAND", "NA", "TRUE", "FALSE"];

fn lookup_rename(upper_name: &str) -> Option<&'static str> {
    FN_RENAMES
        .iter()
        .find(|(k, _)| *k == upper_name)
        .map(|(_, v)| *v)
}

fn is_niladic(upper_name: &str) -> bool {
    NILADIC.contains(&upper_name)
}

/// At byte `i`, where `s.as_bytes()[i] == b'@'`, try to recognise a
/// 1-2-3 function call and emit the IronCalc-shape replacement.
///
/// Handles four cases:
/// 1. `@@(...)` — Lotus indirect → `INDIRECT(...)`.
/// 2. `@NAME` (no `(`) where `NAME` is niladic → `NAME()`.
/// 3. `@NAME(...)` where `NAME` needs argument-shape rewriting (e.g.
///    `@MID` 0→1-based start) — fully consume the call and emit the
///    rewritten Excel form.
/// 4. `@NAME(...)` — apply the rename table; pass through unchanged
///    if not in the table. The opening `(` is left for the caller's
///    next iteration to emit normally.
///
/// Returns `(emitted_text, byte_index_after_match)`. Returns `None`
/// only when `@` is followed by no name letters at all (in which
/// case the caller falls back to plain `@`-strip).
fn try_at_function(
    s: &str,
    i: usize,
    sheets: &[&str],
    cfg: &ParseConfig,
) -> Option<(String, usize)> {
    debug_assert_eq!(s.as_bytes().get(i).copied(), Some(b'@'));
    let bytes = s.as_bytes();

    // Case 1: `@@` → INDIRECT. The opening `(` (if any) is left for
    // the main loop — same as a normal function call.
    if bytes.get(i + 1).copied() == Some(b'@') {
        return Some(("INDIRECT".to_string(), i + 2));
    }

    // Scan name run: [A-Za-z][A-Za-z0-9]*. Function names like
    // `ATAN2` end in a digit; that's the only digit-bearing 1-2-3
    // built-in but the rule is general.
    let name_start = i + 1;
    let mut j = name_start;
    while j < bytes.len() && bytes[j].is_ascii_alphabetic() {
        j += 1;
    }
    while j < bytes.len() && bytes[j].is_ascii_alphanumeric() {
        j += 1;
    }
    if j == name_start {
        return None;
    }

    let name = &s[name_start..j];
    let upper = name.to_ascii_uppercase();
    let next = bytes.get(j).copied();

    // Case 2.5: `@ERR` — Excel has no equivalent function; emit a
    // `#VALUE!` literal so the cell evaluates to an error. 1-2-3
    // `@ERR` is niladic and is the only emulated niladic — the
    // generic niladic-paren completion below would emit `ERR()`,
    // which IronCalc rejects.
    if upper == "ERR" && next != Some(b'(') {
        return Some(("#VALUE!".to_string(), j));
    }

    // Case 3: arg-fix. Functions where the engine name is the same
    // (or trivially renamed) but argument shape needs rewriting.
    // Only triggers when `(` follows; otherwise the call is malformed
    // and we fall through to the rename-table path.
    if next == Some(b'(') {
        if let Some(emitted_end) = try_arg_fix(&upper, s, j, sheets, cfg) {
            return Some(emitted_end);
        }
    }

    // Resolved engine-side name: rename-table hit wins; otherwise
    // preserve the user's original casing (the engine accepts both,
    // and we avoid a gratuitous re-case for `@SUM` etc.).
    let resolved: String = match lookup_rename(&upper) {
        Some(repl) => repl.to_string(),
        None => name.to_string(),
    };

    // Niladic completion only applies when the user did NOT supply
    // `(...)`; otherwise we'd emit `PI()()`.
    let emitted = if is_niladic(&upper) && next != Some(b'(') {
        format!("{resolved}()")
    } else {
        resolved
    };
    Some((emitted, j))
}

/// Per-function argument-shape rewrites. Called with the upper-cased
/// 1-2-3 name and the index of `(` immediately following it. Returns
/// the rewritten Excel-form call plus the byte index after the
/// matching `)`, or `None` if the name has no arg-fix or the call is
/// malformed (unbalanced parens, wrong arity).
///
/// The args text is split at top-level by the user's argument
/// separator, then each arg is recursively translated through
/// [`translate`] so nested 1-2-3 syntax (`..` ranges, `@` calls,
/// non-default punctuation) is handled before the rewrite combines
/// them.
fn try_arg_fix(
    name: &str,
    s: &str,
    paren_idx: usize,
    sheets: &[&str],
    cfg: &ParseConfig,
) -> Option<(String, usize)> {
    let close = find_matching_paren(s, paren_idx)?;
    let args_text = &s[paren_idx + 1..close];
    let raw_args = split_top_level_args(args_text, cfg);
    let translated: Vec<String> = raw_args.iter().map(|a| translate(a, sheets, cfg)).collect();
    let emitted = match name {
        "MID" => rewrite_mid(&translated)?,
        "FIND" => rewrite_find(&translated)?,
        "STRING" => rewrite_string(&translated)?,
        "CTERM" => rewrite_cterm(&translated)?,
        "TERM" => rewrite_term(&translated)?,
        "SUMPRODUCT" => rewrite_sumproduct(&translated)?,
        _ => return None,
    };
    Some((emitted, close + 1))
}

/// `@MID(s, start, count)` — 1-2-3 `start` is 0-based, Excel
/// `MID` is 1-based. Wrap `start` to add 1.
fn rewrite_mid(args: &[String]) -> Option<String> {
    if args.len() != 3 {
        return None;
    }
    Some(format!("MID({},({})+1,{})", args[0], args[1], args[2]))
}

/// `@FIND(needle, haystack, start)` — both `start` arg and returned
/// position are 0-based in 1-2-3; Excel `FIND` is 1-based on both.
/// Translate `start` via `+1`; wrap whole call to subtract 1 from the
/// result. `#VALUE!` propagates through the subtraction unchanged.
fn rewrite_find(args: &[String]) -> Option<String> {
    if args.len() != 3 {
        return None;
    }
    Some(format!("FIND({},{},({})+1)-1", args[0], args[1], args[2]))
}

/// `@CTERM(rate, fv, pv)` → `LN(fv/pv)/LN(1+rate)`. Periods to grow
/// `pv` to `fv` at constant `rate`. IronCalc has no equivalent.
fn rewrite_cterm(args: &[String]) -> Option<String> {
    if args.len() != 3 {
        return None;
    }
    let (rate, fv, pv) = (&args[0], &args[1], &args[2]);
    Some(format!("LN({fv}/{pv})/LN(1+{rate})"))
}

/// `@TERM(pmt, rate, fv)` → `LN(1+(fv*rate)/pmt)/LN(1+rate)`.
/// Periods of payments to reach `fv` at constant `rate`. IronCalc
/// has no equivalent.
fn rewrite_term(args: &[String]) -> Option<String> {
    if args.len() != 3 {
        return None;
    }
    let (pmt, rate, fv) = (&args[0], &args[1], &args[2]);
    Some(format!("LN(1+({fv}*{rate})/{pmt})/LN(1+{rate})"))
}

/// `@SUMPRODUCT(a, b, ...)` → `SUM((a)*(b)*...)`. IronCalc 0.7.1
/// lacks `SUMPRODUCT` but its `SUM` broadcasts arrays element-wise,
/// so the rewrite is mechanical.
fn rewrite_sumproduct(args: &[String]) -> Option<String> {
    if args.is_empty() {
        return None;
    }
    let parts: Vec<String> = args.iter().map(|a| format!("({a})")).collect();
    Some(format!("SUM({})", parts.join("*")))
}

/// `@STRING(n, decimals)` → Excel `TEXT(n, fmt)`. When `decimals` is
/// a non-negative integer literal, build `fmt` statically (`"0"`,
/// `"0.00"`, ...). Otherwise emit a runtime form using `REPT` and
/// branch on `decimals=0` so the format never has a trailing `.`.
fn rewrite_string(args: &[String]) -> Option<String> {
    if args.len() != 2 {
        return None;
    }
    let n = &args[0];
    let dec = &args[1];
    if let Ok(d) = dec.trim().parse::<u32>() {
        let fmt = if d == 0 {
            "0".to_string()
        } else {
            format!("0.{}", "0".repeat(d as usize))
        };
        return Some(format!("TEXT({n},\"{fmt}\")"));
    }
    Some(format!(
        "IF({dec}>0,TEXT({n},\"0.\"&REPT(\"0\",{dec})),TEXT({n},\"0\"))"
    ))
}

/// Find the byte index of the `)` matching the `(` at `s[open_idx]`.
/// Honors string literals and nested parens. Returns `None` if
/// unbalanced.
fn find_matching_paren(s: &str, open_idx: usize) -> Option<usize> {
    debug_assert_eq!(s.as_bytes().get(open_idx).copied(), Some(b'('));
    let bytes = s.as_bytes();
    let mut depth = 0i32;
    let mut in_string = false;
    let mut i = open_idx;
    while i < bytes.len() {
        let b = bytes[i];
        if b == b'"' {
            in_string = !in_string;
            i += 1;
            continue;
        }
        if in_string {
            i += 1;
            continue;
        }
        if b == b'(' {
            depth += 1;
        } else if b == b')' {
            depth -= 1;
            if depth == 0 {
                return Some(i);
            }
        }
        i += 1;
    }
    None
}

/// Split the body of a function call (the substring strictly between
/// the outer `(` and `)`) into top-level argument slices using the
/// user's argument separator. Honors nested parens and string
/// literals. The `..` range separator is never a split, even when
/// the arg separator is `.` (Punctuation B/D).
///
/// An empty `args_text` yields `vec![]` (zero-arg call). A trailing
/// separator yields a final empty slice — callers requiring a fixed
/// arity should reject that explicitly.
fn split_top_level_args<'a>(args_text: &'a str, cfg: &ParseConfig) -> Vec<&'a str> {
    if args_text.is_empty() {
        return Vec::new();
    }
    let bytes = args_text.as_bytes();
    let arg_sep_byte = cfg.argument_sep as u32;
    let mut depth = 0i32;
    let mut in_string = false;
    let mut last = 0;
    let mut parts: Vec<&str> = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        if b == b'"' {
            in_string = !in_string;
            i += 1;
            continue;
        }
        if in_string {
            i += 1;
            continue;
        }
        if b == b'(' {
            depth += 1;
            i += 1;
            continue;
        }
        if b == b')' {
            depth -= 1;
            i += 1;
            continue;
        }
        // `..` always wins over single-`.` arg-sep.
        if b == b'.' && bytes.get(i + 1).copied() == Some(b'.') {
            i += 2;
            continue;
        }
        if depth == 0 && (b as u32) == arg_sep_byte {
            parts.push(&args_text[last..i]);
            last = i + 1;
        }
        i += 1;
    }
    parts.push(&args_text[last..]);
    parts
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
        // Niladic: 1-2-3 omits the parens; Excel requires them.
        assert_eq!(to_engine_source("@NOW", &[]), "=NOW()");
    }

    #[test]
    fn leading_plus_becomes_equals() {
        assert_eq!(to_engine_source("+A1+B1", &[]), "=A1+B1");
        assert_eq!(to_engine_source("+5", &[]), "=5");
    }

    #[test]
    fn range_separator_translated() {
        assert_eq!(to_engine_source("@SUM(A1..B5)", &[]), "=SUM(A1:B5)");
        assert_eq!(to_engine_source("@AVG(B2..B20)", &[]), "=AVERAGE(B2:B20)");
    }

    // ---- M5+: function-name renames ----
    //
    // 1-2-3 names that map to a different IronCalc/Excel name.
    // Catalog in docs/AT_FUNCTIONS.md.

    #[test]
    fn rename_avg_to_average() {
        assert_eq!(to_engine_source("@AVG(A1..A5)", &[]), "=AVERAGE(A1:A5)");
    }

    #[test]
    fn rename_count_to_counta() {
        // 1-2-3 @COUNT counts non-empty cells (incl. labels). Excel COUNT
        // only counts numbers; COUNTA matches 1-2-3 semantics.
        assert_eq!(to_engine_source("@COUNT(A1..A5)", &[]), "=COUNTA(A1:A5)");
    }

    #[test]
    fn rename_std_var_family() {
        assert_eq!(to_engine_source("@STD(A1..A5)", &[]), "=STDEV.P(A1:A5)");
        assert_eq!(to_engine_source("@STDS(A1..A5)", &[]), "=STDEV.S(A1:A5)");
        assert_eq!(to_engine_source("@VAR(A1..A5)", &[]), "=VAR.P(A1:A5)");
        assert_eq!(to_engine_source("@VARS(A1..A5)", &[]), "=VAR.S(A1:A5)");
    }

    #[test]
    fn rename_isstring_to_istext() {
        assert_eq!(to_engine_source("@ISSTRING(A1)", &[]), "=ISTEXT(A1)");
    }

    #[test]
    fn rename_length_to_len() {
        assert_eq!(to_engine_source("@LENGTH(A1)", &[]), "=LEN(A1)");
    }

    #[test]
    fn rename_repeat_to_rept() {
        assert_eq!(to_engine_source("@REPEAT(\"-\",5)", &[]), "=REPT(\"-\",5)");
    }

    #[test]
    fn rename_cols_to_columns() {
        assert_eq!(to_engine_source("@COLS(A1..C5)", &[]), "=COLUMNS(A1:C5)");
    }

    #[test]
    fn rename_database_family() {
        assert_eq!(
            to_engine_source("@DAVG(A1..C5,2,E1..E2)", &[]),
            "=DAVERAGE(A1:C5,2,E1:E2)"
        );
        assert_eq!(
            to_engine_source("@DSTD(A1..C5,2,E1..E2)", &[]),
            "=DSTDEVP(A1:C5,2,E1:E2)"
        );
        assert_eq!(
            to_engine_source("@DSTDS(A1..C5,2,E1..E2)", &[]),
            "=DSTDEV(A1:C5,2,E1:E2)"
        );
        assert_eq!(
            to_engine_source("@DVAR(A1..C5,2,E1..E2)", &[]),
            "=DVARP(A1:C5,2,E1:E2)"
        );
        assert_eq!(
            to_engine_source("@DVARS(A1..C5,2,E1..E2)", &[]),
            "=DVAR(A1:C5,2,E1:E2)"
        );
    }

    #[test]
    fn at_at_becomes_indirect() {
        // 1-2-3 indirect: @@(B5) where B5 holds e.g. "A1".
        assert_eq!(to_engine_source("@@(B5)", &[]), "=INDIRECT(B5)");
        assert_eq!(
            to_engine_source("@@(\"A\"&\"1\")", &[]),
            "=INDIRECT(\"A\"&\"1\")"
        );
    }

    #[test]
    fn niladic_functions_get_parens() {
        // 1-2-3 omits parens on niladic functions; Excel requires them.
        assert_eq!(to_engine_source("@PI", &[]), "=PI()");
        assert_eq!(to_engine_source("@TODAY", &[]), "=TODAY()");
        assert_eq!(to_engine_source("@NOW", &[]), "=NOW()");
        assert_eq!(to_engine_source("@RAND", &[]), "=RAND()");
        assert_eq!(to_engine_source("@NA", &[]), "=NA()");
        assert_eq!(to_engine_source("@TRUE", &[]), "=TRUE()");
        assert_eq!(to_engine_source("@FALSE", &[]), "=FALSE()");
    }

    #[test]
    fn niladic_with_explicit_parens_not_doubled() {
        // If the user wrote `@PI()`, we must not emit `PI()()`.
        assert_eq!(to_engine_source("@PI()", &[]), "=PI()");
        assert_eq!(to_engine_source("@TODAY()", &[]), "=TODAY()");
    }

    #[test]
    fn niladic_in_expression_context() {
        // The post-paren-completion must not interfere with surrounding
        // operators or commas.
        assert_eq!(to_engine_source("@PI*2", &[]), "=PI()*2");
        assert_eq!(
            to_engine_source("@IF(@RAND>0.5,1,0)", &[]),
            "=IF(RAND()>0.5,1,0)"
        );
        assert_eq!(to_engine_source("2*@PI+1", &[]), "=2*PI()+1");
    }

    #[test]
    fn passthrough_function_names_preserved() {
        // Functions not in the rename table pass through unchanged.
        assert_eq!(to_engine_source("@SUM(A1..A5)", &[]), "=SUM(A1:A5)");
        assert_eq!(
            to_engine_source("@VLOOKUP(A1,B1..C5,2,0)", &[]),
            "=VLOOKUP(A1,B1:C5,2,0)"
        );
        assert_eq!(to_engine_source("@LEN(A1)", &[]), "=LEN(A1)");
    }

    #[test]
    fn renames_inside_nested_calls() {
        assert_eq!(
            to_engine_source("@IF(@AVG(A1..A5)>0,@STDS(A1..A5),0)", &[]),
            "=IF(AVERAGE(A1:A5)>0,STDEV.S(A1:A5),0)"
        );
    }

    // ---- M5+: arg-fix rewriters (per AT_FUNCTIONS.md) ----
    //
    // Functions that share a name with their IronCalc target but have
    // a different argument shape (e.g. 0-based vs 1-based indices).

    #[test]
    fn mid_arg_zero_to_one_based() {
        // 1-2-3 @MID(s, start, count): start is 0-based.
        // Excel MID(s, start, count): start is 1-based. Add 1.
        assert_eq!(
            to_engine_source("@MID(\"hello\",2,3)", &[]),
            "=MID(\"hello\",(2)+1,3)"
        );
        assert_eq!(
            to_engine_source("@MID(A1,B1,C1)", &[]),
            "=MID(A1,(B1)+1,C1)"
        );
    }

    #[test]
    fn mid_with_inner_call_in_start_arg() {
        assert_eq!(
            to_engine_source("@MID(A1,@FIND(\"x\",A1,0),3)", &[]),
            "=MID(A1,(FIND(\"x\",A1,(0)+1)-1)+1,3)"
        );
    }

    #[test]
    fn find_arg_zero_to_one_based_and_result_minus_one() {
        // 1-2-3 @FIND(needle, haystack, start): both start arg and
        // returned position are 0-based. Excel FIND is 1-based on
        // both. Add 1 to start, subtract 1 from result.
        assert_eq!(
            to_engine_source("@FIND(\"l\",\"hello\",0)", &[]),
            "=FIND(\"l\",\"hello\",(0)+1)-1"
        );
        assert_eq!(
            to_engine_source("@FIND(A1,B1,C1)", &[]),
            "=FIND(A1,B1,(C1)+1)-1"
        );
    }

    #[test]
    fn find_inside_iserr_still_works() {
        // FIND failure → #VALUE! in Excel; subtracting 1 propagates
        // the error, so ISERR still detects it.
        assert_eq!(
            to_engine_source("@ISERR(@FIND(\"x\",A1,0))", &[]),
            "=ISERR(FIND(\"x\",A1,(0)+1)-1)"
        );
    }

    #[test]
    fn string_to_text_with_literal_decimals() {
        // @STRING(n, decimals) — for a numeric-literal `decimals`,
        // build the format string statically: `0`, `0.0`, `0.00`, ...
        assert_eq!(
            to_engine_source("@STRING(1234.5,2)", &[]),
            "=TEXT(1234.5,\"0.00\")"
        );
        assert_eq!(to_engine_source("@STRING(A1,0)", &[]), "=TEXT(A1,\"0\")");
        assert_eq!(
            to_engine_source("@STRING(A1,5)", &[]),
            "=TEXT(A1,\"0.00000\")"
        );
    }

    #[test]
    fn string_with_dynamic_decimals_uses_if_rept() {
        // When `decimals` isn't a numeric literal, build the format
        // string at runtime using REPT — and special-case 0 so we
        // don't get a trailing `.`.
        assert_eq!(
            to_engine_source("@STRING(A1,B1)", &[]),
            "=IF(B1>0,TEXT(A1,\"0.\"&REPT(\"0\",B1)),TEXT(A1,\"0\"))"
        );
    }

    // ---- M5+: parse-time emulations (PR3) ----
    //
    // Functions IronCalc 0.7.1 lacks but that compose cleanly to
    // existing primitives.

    #[test]
    fn err_literal() {
        // 1-2-3 @ERR → Excel #VALUE! literal. Verified IronCalc
        // 0.7.1 accepts `=#VALUE!` as a formula.
        assert_eq!(to_engine_source("@ERR", &[]), "=#VALUE!");
        // Inside expressions: the literal participates as an error.
        assert_eq!(
            to_engine_source("@IF(A1<0,@ERR,A1)", &[]),
            "=IF(A1<0,#VALUE!,A1)"
        );
    }

    #[test]
    fn cterm_emulated_via_log() {
        // @CTERM(rate, fv, pv) → LN(fv/pv)/LN(1+rate)
        assert_eq!(
            to_engine_source("@CTERM(0.05,1000,500)", &[]),
            "=LN(1000/500)/LN(1+0.05)"
        );
        assert_eq!(
            to_engine_source("@CTERM(A1,B1,C1)", &[]),
            "=LN(B1/C1)/LN(1+A1)"
        );
    }

    #[test]
    fn term_emulated_via_log() {
        // @TERM(pmt, rate, fv) → LN(1+(fv*rate)/pmt)/LN(1+rate)
        assert_eq!(
            to_engine_source("@TERM(100,0.05,1000)", &[]),
            "=LN(1+(1000*0.05)/100)/LN(1+0.05)"
        );
        assert_eq!(
            to_engine_source("@TERM(A1,B1,C1)", &[]),
            "=LN(1+(C1*B1)/A1)/LN(1+B1)"
        );
    }

    #[test]
    fn sumproduct_via_array_in_sum() {
        // 1-2-3 (post-MVP) @SUMPRODUCT — IronCalc lacks SUMPRODUCT but
        // accepts SUM with array broadcast: SUM(A1:A3*B1:B3) = 140
        // for {1,2,3}*{10,20,30}.
        assert_eq!(
            to_engine_source("@SUMPRODUCT(A1..A3,B1..B3)", &[]),
            "=SUM((A1:A3)*(B1:B3))"
        );
        // Three-arg form composes the same way.
        assert_eq!(
            to_engine_source("@SUMPRODUCT(A1..A3,B1..B3,C1..C3)", &[]),
            "=SUM((A1:A3)*(B1:B3)*(C1:C3))"
        );
    }

    #[test]
    fn rename_code_to_unicode() {
        // @CODE returns the codepoint of the first char of a string.
        // Excel's `CODE` returns ASCII; IronCalc only ships `UNICODE`,
        // which is identical to `CODE` for ASCII input.
        assert_eq!(to_engine_source("@CODE(A1)", &[]), "=UNICODE(A1)");
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

    // ---- Punctuation A-H — non-default ParseConfig ----

    fn cfg_punct_b() -> ParseConfig {
        // Punct B: argument `.`, decimal `,`.
        ParseConfig {
            argument_sep: '.',
            decimal_point: ',',
        }
    }

    fn cfg_punct_e() -> ParseConfig {
        // Punct E: argument `;`, decimal `.`.
        ParseConfig {
            argument_sep: ';',
            decimal_point: '.',
        }
    }

    #[test]
    fn punct_e_translates_semicolon_arg_sep_to_comma() {
        let cfg = cfg_punct_e();
        assert_eq!(
            to_engine_source_with_config("@SUM(A1;A2;A3)", &[], &cfg),
            "=SUM(A1,A2,A3)"
        );
        assert_eq!(
            to_engine_source_with_config("@IF(A1>0;1;2)", &[], &cfg),
            "=IF(A1>0,1,2)"
        );
    }

    #[test]
    fn punct_e_keeps_double_dot_as_range() {
        // `..` greedy-eat fires before any single-`.` translation,
        // so range syntax always means range regardless of Punctuation.
        let cfg = cfg_punct_e();
        assert_eq!(
            to_engine_source_with_config("@SUM(A1..A5)", &[], &cfg),
            "=SUM(A1:A5)"
        );
    }

    #[test]
    fn punct_b_translates_dot_arg_sep_to_comma() {
        let cfg = cfg_punct_b();
        assert_eq!(
            to_engine_source_with_config("@SUM(A1.A2.A3)", &[], &cfg),
            "=SUM(A1,A2,A3)"
        );
    }

    #[test]
    fn punct_b_translates_decimal_comma_to_dot() {
        let cfg = cfg_punct_b();
        // `1,5` is the decimal `1.5` under Punct B; translate so
        // IronCalc parses it correctly.
        assert_eq!(
            to_engine_source_with_config("+1,5+2,5", &[], &cfg),
            "=1.5+2.5"
        );
    }

    #[test]
    fn punct_b_keeps_double_dot_as_range_even_with_dot_arg_sep() {
        let cfg = cfg_punct_b();
        // Range `A1..A5` wins over single-`.` arg sep.
        assert_eq!(
            to_engine_source_with_config("@SUM(A1..A5)", &[], &cfg),
            "=SUM(A1:A5)"
        );
    }

    #[test]
    fn punct_e_strings_are_not_mutated() {
        // Even with `;` as the arg-sep, semicolons inside strings
        // pass through untouched.
        let cfg = cfg_punct_e();
        assert_eq!(
            to_engine_source_with_config("@IF(A1>0;\"a;b\";\"c\")", &[], &cfg),
            "=IF(A1>0,\"a;b\",\"c\")"
        );
    }

    #[test]
    fn punct_a_default_is_unchanged_from_two_arg_form() {
        let cfg = ParseConfig::default();
        let s = sheets();
        assert_eq!(
            to_engine_source_with_config("@SUM(A1..A5)", &s, &cfg),
            to_engine_source("@SUM(A1..A5)", &s)
        );
    }
}
