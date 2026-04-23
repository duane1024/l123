//! Minimal RFC-4180-ish CSV parser for `/File Import Numbers`.
//!
//! Scope: rows separated by `\n` (tolerates `\r\n`), fields by `,`.
//! Double-quoted fields may contain commas and newlines; `""` inside a
//! quoted field is an escaped quote. No configurable delimiters — 1-2-3
//! /File Import Numbers settings are a post-MVP concern.

/// Parse a CSV document into rows of fields.
pub fn parse(src: &str) -> Vec<Vec<String>> {
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut row: Vec<String> = Vec::new();
    let mut field = String::new();
    let mut in_quotes = false;
    let mut chars = src.chars().peekable();
    while let Some(c) = chars.next() {
        if in_quotes {
            if c == '"' {
                if chars.peek() == Some(&'"') {
                    field.push('"');
                    chars.next();
                } else {
                    in_quotes = false;
                }
            } else {
                field.push(c);
            }
            continue;
        }
        match c {
            '"' if field.is_empty() => in_quotes = true,
            ',' => row.push(std::mem::take(&mut field)),
            '\n' => {
                row.push(std::mem::take(&mut field));
                rows.push(std::mem::take(&mut row));
            }
            '\r' => { /* tolerate CRLF */ }
            _ => field.push(c),
        }
    }
    // Trailing row without a final newline.
    if !field.is_empty() || !row.is_empty() {
        row.push(field);
        rows.push(row);
    }
    rows
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn plain_numbers() {
        let got = parse("10,20,30\n40,50,60\n");
        assert_eq!(got, vec![
            vec!["10", "20", "30"],
            vec!["40", "50", "60"],
        ]);
    }

    #[test]
    fn quoted_fields() {
        let got = parse("\"foo\",\"bar\",\"baz\"\n");
        assert_eq!(got, vec![vec!["foo", "bar", "baz"]]);
    }

    #[test]
    fn quoted_field_with_comma_and_escape() {
        let got = parse("\"a,b\",\"he said \"\"hi\"\"\"\n");
        assert_eq!(got, vec![vec!["a,b", "he said \"hi\""]]);
    }

    #[test]
    fn trailing_row_without_newline() {
        let got = parse("1,2\n3,4");
        assert_eq!(got, vec![vec!["1", "2"], vec!["3", "4"]]);
    }

    #[test]
    fn tolerates_crlf() {
        let got = parse("1,2\r\n3,4\r\n");
        assert_eq!(got, vec![vec!["1", "2"], vec!["3", "4"]]);
    }

    #[test]
    fn empty_fields_preserved() {
        let got = parse(",a,,b,\n");
        assert_eq!(got, vec![vec!["", "a", "", "b", ""]]);
    }
}
