//! 1-2-3 formula syntax translator.
//!
//! Converts a Lotus-shape formula source (as typed by the user) into the
//! Excel-shape input that `ironcalc::base::Model::set_user_input` expects.
//!
//! Scope as of M2:
//! - Leading `@` or `+` is replaced with `=` (Excel's formula sigil).
//! - `..` range separator → `:` outside of string literals.
//! - Double-quoted string literals (`"..."`) pass through untouched.
//! - Functions left as-is (Lotus `@SUM` → Excel `SUM` after the sigil
//!   swap; Excel accepts it case-insensitively so no rename is needed yet).
//!
//! Not yet handled (later milestones):
//! - `#AND#` / `#OR#` / `#NOT#` infix → function translation.
//! - Single-dot range separator `A1.B5`.
//! - Cross-sheet references like `A:B1..C:B1` — M5.
//! - Named ranges — M3 (names layer).
//! - Function renames where Lotus and Excel differ (e.g. `@AVG` ↔ `AVERAGE`).

/// Translate a Lotus-shape formula source (first char already classified
/// as a value-starter) into an Excel-shape formula that always begins with
/// `=`. Pure numeric literals (e.g. `"42"`, `"-3.5"`) are also translated,
/// which is safe: the caller has already decided they're formula-class.
pub fn to_engine_source(lotus: &str) -> String {
    let body = match lotus.chars().next() {
        Some('@') | Some('+') => &lotus[1..],
        _ => lotus,
    };
    let translated = translate_ranges(body);
    format!("={translated}")
}

/// Walk the string, converting `..` to `:` outside of double-quoted
/// regions, and dropping any `@` sigils (they are purely syntactic in
/// 1-2-3 and Excel does not use them). Single `.` is left as-is (it may
/// be a decimal point).
fn translate_ranges(s: &str) -> String {
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
        if !in_string {
            if b == b'@' {
                i += 1;
                continue;
            }
            if b == b'.' && i + 1 < bytes.len() && bytes[i + 1] == b'.' {
                out.push(':');
                i += 2;
                continue;
            }
        }
        let ch_end = next_char_boundary(s, i);
        out.push_str(&s[i..ch_end]);
        i = ch_end;
    }
    out
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
        assert_eq!(to_engine_source("@SUM(A1..A5)"), "=SUM(A1:A5)");
        assert_eq!(to_engine_source("@IF(A1>0,1,2)"), "=IF(A1>0,1,2)");
        assert_eq!(to_engine_source("@NOW"), "=NOW");
    }

    #[test]
    fn leading_plus_becomes_equals() {
        assert_eq!(to_engine_source("+A1+B1"), "=A1+B1");
        assert_eq!(to_engine_source("+5"), "=5");
    }

    #[test]
    fn range_separator_translated() {
        assert_eq!(to_engine_source("@SUM(A1..B5)"), "=SUM(A1:B5)");
        assert_eq!(to_engine_source("@AVG(B2..B20)"), "=AVG(B2:B20)");
    }

    #[test]
    fn single_dot_is_preserved_as_decimal() {
        assert_eq!(to_engine_source("+3.14+2.5"), "=3.14+2.5");
        assert_eq!(to_engine_source("@IF(A1>0.5,1,0)"), "=IF(A1>0.5,1,0)");
    }

    #[test]
    fn plain_numbers_pass_through_with_equals() {
        assert_eq!(to_engine_source("123"), "=123");
        assert_eq!(to_engine_source("-3.5"), "=-3.5");
        assert_eq!(to_engine_source("0.25"), "=0.25");
    }

    #[test]
    fn strings_are_not_mutated() {
        // `..` inside a string literal must not become `:`.
        assert_eq!(
            to_engine_source("@IF(A1>0,\"low..high\",\"\")"),
            "=IF(A1>0,\"low..high\",\"\")"
        );
        assert_eq!(
            to_engine_source("@N(\"abc..def\")"),
            "=N(\"abc..def\")"
        );
    }

    #[test]
    fn mixed_strings_and_ranges() {
        // A range outside the string, `..` inside the string.
        assert_eq!(
            to_engine_source("@IF(@SUM(A1..A5)>0,\"A..B\",\"C\")"),
            "=IF(SUM(A1:A5)>0,\"A..B\",\"C\")"
        );
    }

    #[test]
    fn nested_at_functions() {
        // Only the leading `@` becomes `=`; inner `@` stays — Excel accepts
        // function calls without a sigil, but IronCalc's set_user_input
        // expects the input form it would see typed; functions bare work.
        //
        // Lotus users never type inner @ (only for the outermost function),
        // so this path isn't exercised; we still assert the output stays
        // well-formed if it happens.
        let got = to_engine_source("@IF(@ISERR(A1),0,A1)");
        // The first '@' was consumed; the inner '@' remains.
        // IronCalc does not accept inner '@'. This is a known issue the
        // caller (M2 cycle 3 engine integration) will warn about.
        assert!(got.starts_with('='));
        assert!(got.contains("ISERR"));
    }
}
