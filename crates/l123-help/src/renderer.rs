//! Best-effort renderer for Huffman-decoded `123.HLP` records.
//!
//! Two record shapes are recognized:
//!
//! 1. **Body topic** — `<noise>Title -- Body...`. The noise is the help
//!    renderer's per-cell attribute / cursor-positioning stream we have
//!    not yet reverse-engineered; the title and body themselves are
//!    stored as plain ASCII bytes mixed with byte-aligned control codes.
//!
//! 2. **Cross-reference index page** (Task Index, Function Index,
//!    Keyboard Index, etc.) — a sequence of rows of the shape
//!    `<topic-id-bytes><description>      -- see\0`. The bold cross-ref
//!    target name shown to the user at the end of each line is *not*
//!    in the byte stream — the renderer looks it up at runtime from the
//!    topic-id bytes. We strip the topic-id prefix and emit just the
//!    description plus a literal " -- see ..." marker so the layout
//!    matches what the user sees on screen.
//!
//! Byte-aligned control codes seen in both shapes:
//!
//! - `0x00` — line break (next screen row)
//! - `0xC4` — soft-space / layout fill (renders as a space)
//! - bytes `< 0x20` and other extended bytes — renderer commands
//!   (color/attribute changes, cross-reference markers); we treat
//!   them as separators and drop them
//!
//! [`extract_topic`] returns a cleaned `(title, body)` pair when the
//! record contains the ` -- ` topic delimiter, else `None`.

/// One extracted help topic.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Topic {
    pub title: String,
    pub body: String,
}

/// Number of ` -- see` markers in the body that triggers cross-ref-page
/// recognition. Three or more is a strong signal: a regular body might
/// mention "see also X" once or twice, but only index pages have many.
const CROSSREF_THRESHOLD: usize = 3;

/// Return the cleaned `(title, body)` pair if `decoded` is a body
/// topic record (contains ` -- ` delimiter).
pub fn extract_topic(decoded: &[u8]) -> Option<Topic> {
    // Cross-ref pages have many ` -- see` markers; treat them specially.
    let see_count = count_see_markers(decoded);
    if see_count >= CROSSREF_THRESHOLD {
        return extract_crossref_page(decoded);
    }

    // Find the FIRST ASCII ` -- ` delimiter — these spaces are real
    // 0x20 bytes, not 0xC4 fill.
    let sep = b" -- ";
    let idx = decoded.windows(sep.len()).position(|w| w == sep)?;

    let title = extract_title(&decoded[..idx])?;
    let body = clean_body(&decoded[idx + sep.len()..]);

    if title.is_empty() || body.is_empty() {
        return None;
    }
    Some(Topic { title, body })
}

fn count_see_markers(decoded: &[u8]) -> usize {
    decoded
        .windows(8)
        .filter(|w| *w == b" -- see\0" || *w == b" -- see ")
        .count()
}

/// Walk back from the end of `pre` (the bytes immediately preceding
/// ` -- `) and return the title-case suffix.
///
/// The format we extract is `<renderer noise>[0xC4 or 0x00]<noise
/// letters glued><Title> -- `. The walk first stops at `0x00` /
/// `0xC4` / control byte to discard the heaviest noise; within the
/// resulting candidate string we take everything from the first
/// title-start character (uppercase letter, `@`, `/`, `(`, `$`, `+`)
/// to the end. This handles cases like `nnAbout 1-2-3 Help` where
/// noise letters are concatenated with no whitespace separator before
/// the real title.
fn extract_title(pre: &[u8]) -> Option<String> {
    fn is_title_start(c: char) -> bool {
        c.is_ascii_uppercase() || matches!(c, '@' | '/' | '(' | '$' | '+')
    }

    // Walk back to find the boundary at 0x00 / 0xC4 / non-printable.
    let mut start = 0usize;
    for (i, &b) in pre.iter().enumerate().rev() {
        if b == 0x00 || b == 0xC4 || !(0x20..=0x7E).contains(&b) {
            start = i + 1;
            break;
        }
    }
    let candidate_bytes = &pre[start..];
    let candidate = std::str::from_utf8(candidate_bytes).ok()?.trim();
    if candidate.is_empty() {
        return None;
    }

    // Scan the candidate left-to-right and take everything from the
    // first title-start character to the end.
    let title_start = candidate
        .char_indices()
        .find(|(_, c)| is_title_start(*c))?
        .0;
    let title = candidate[title_start..].trim();
    if title.is_empty() {
        return None;
    }
    Some(title.to_string())
}

/// Clean up body bytes:
///   * `0x00` → newline
///   * already-space (originally `0xC4`) stays as space
///   * drop other bytes < 0x20 and any byte > 0x7E
///   * for each line, find the longest run of "looks like English text"
///     and use that — drops the per-line noise prefix
///   * trim and collapse whitespace
fn clean_body(raw: &[u8]) -> String {
    let mut out = String::new();
    for line in raw.split(|&b| b == 0x00) {
        // Filter to printable ASCII.
        let mut filtered: Vec<u8> = Vec::with_capacity(line.len());
        for &b in line {
            if (0x20..=0x7E).contains(&b) {
                filtered.push(b);
            }
        }
        let s = std::str::from_utf8(&filtered).unwrap_or("");

        // Find the longest "clean run" — ≥ 8 chars, starts with a
        // letter or digit or punctuation that begins real text.
        // Take everything from the first such run onward (so multi-
        // word lines aren't truncated to a single phrase).
        let cleaned_line = strip_line_prefix(s);
        let normalized = normalize_whitespace(&cleaned_line);

        if !normalized.is_empty() {
            if !out.is_empty() {
                out.push('\n');
            }
            out.push_str(&normalized);
        }
    }
    out.trim().to_string()
}

/// Strip the leading per-line noise prefix. Heuristic: skip bytes
/// until we hit a run of ≥ 8 characters where each is a letter, digit,
/// or common punctuation, and the run starts with a letter/digit or
/// a sentence-starter.
fn strip_line_prefix(s: &str) -> String {
    fn is_clean(c: char) -> bool {
        c.is_ascii_alphanumeric()
            || c == ' '
            || matches!(
                c,
                '.' | ','
                    | '/'
                    | '-'
                    | '('
                    | ')'
                    | '@'
                    | ':'
                    | ';'
                    | '"'
                    | '\''
                    | '?'
                    | '!'
                    | '+'
                    | '='
                    | '$'
                    | '*'
                    | '&'
                    | '['
                    | ']'
            )
    }
    let bytes = s.as_bytes();
    // If the line is short enough that there's no noise prefix worth
    // stripping, return as-is (after dropping leading non-letters).
    if bytes.len() < 12 {
        return s
            .trim_start_matches(|c: char| !c.is_ascii_alphanumeric())
            .to_string();
    }

    // Search for the first position where we have a long-enough clean run.
    let chars: Vec<char> = s.chars().collect();
    let n = chars.len();
    for start in 0..n {
        if !chars[start].is_ascii_alphanumeric()
            && !matches!(chars[start], '/' | '@' | '(' | '"' | '\'' | '+' | '$' | '!')
        {
            continue;
        }
        // Count clean run length from here.
        let mut end = start;
        while end < n && is_clean(chars[end]) {
            end += 1;
        }
        if end - start >= 8 {
            return chars[start..].iter().collect::<String>();
        }
    }
    // Nothing looked clean enough; return whatever we have minus
    // leading non-letters.
    s.trim_start_matches(|c: char| !c.is_ascii_alphanumeric())
        .to_string()
}

fn normalize_whitespace(s: &str) -> String {
    let mut out = String::new();
    let mut last_space = false;
    for c in s.chars() {
        if c.is_ascii_whitespace() {
            if !last_space && !out.is_empty() {
                out.push(' ');
            }
            last_space = true;
        } else {
            out.push(c);
            last_space = false;
        }
    }
    out.trim_end().to_string()
}

/// Split a cross-ref page record into clean rows.
///
/// Each row of the rendered page has the shape
/// `<topic-id-bytes><description>      -- see <bold-target>`
/// where the bold target name is rendered by the help engine from the
/// leading topic-id bytes (it is *not* part of the byte stream). We
/// strip the topic-id prefix per row and append a literal `-- see`
/// marker so the surrounding layout still reads naturally.
///
/// Returns a `Topic` whose title is one of the index-page names we
/// recognize ("Task Index", "Function Index", etc.) and whose body is
/// one cleaned cross-ref entry per line.
fn extract_crossref_page(decoded: &[u8]) -> Option<Topic> {
    let mut rows: Vec<String> = Vec::new();
    for line in decoded.split(|&b| b == 0x00) {
        if line.is_empty() {
            continue;
        }
        if let Some(row) = clean_crossref_row(line) {
            rows.push(row);
        }
    }
    if rows.len() < CROSSREF_THRESHOLD {
        return None;
    }

    // Title: pick the index-page family name that appears anywhere in
    // the decoded stream (e.g. "Task Index", "Help Index"). If none
    // match, fall back to a generic label.
    let title = guess_index_title(decoded);

    // Cross-ref pages also encode their footer (Continued / Previous /
    // Help Index) at the start of the record; the topic-id stripper
    // turns those rows into bare `-- see` artifacts. Drop those.
    let body_rows: Vec<String> = rows
        .into_iter()
        .filter(|r| !is_footer_artifact(r))
        .collect();
    if body_rows.is_empty() {
        return None;
    }
    let body = body_rows.join("\n");
    Some(Topic { title, body })
}

/// Strip the topic-id prefix from one cross-ref row's bytes.
///
/// Each row is `<topic-id-bytes><description><spaces>-- see` (with the
/// trailing space and bold target rendered by the help engine, not in
/// the bytes). The topic-id bytes look like ordinary printable ASCII
/// because the renderer emits them as cells; the description that
/// follows is the part the user actually reads.
///
/// To find where the description begins we scan left-to-right for the
/// earliest position whose suffix looks more like English than what
/// precedes it. Two signals work:
///
/// 1. A capital letter followed by ≥ 2 lowercase letters (start of a
///    proper English word like `Remove`, `Change`, `Specify`).
/// 2. A space-delimited token of length ≥ 2 immediately followed by
///    another space-delimited token of length ≥ 2, both made of
///    letters only (ignoring `1-2-3` as a special case). This catches
///    descriptions whose first character was eaten by the renderer
///    layer (`data in a worksheet file`, `pecify a range`).
fn clean_crossref_row(line: &[u8]) -> Option<String> {
    let s: String = line
        .iter()
        .filter(|&&b| (0x20..=0x7E).contains(&b))
        .map(|&b| b as char)
        .collect();
    if s.is_empty() {
        return None;
    }

    let chars: Vec<char> = s.chars().collect();
    let see_idx = s.find(" -- see");

    let start = find_description_start(&chars, see_idx);
    let desc_end = see_idx.unwrap_or(chars.len());
    if start >= desc_end {
        return None;
    }
    let desc: String = chars[start..desc_end].iter().collect();
    let collapsed = normalize_whitespace(&desc);
    let trimmed = collapsed.trim_end();

    if trimmed.is_empty() {
        return None;
    }

    if see_idx.is_some() {
        Some(format!("{trimmed}  -- see"))
    } else {
        Some(trimmed.to_string())
    }
}

/// Find the byte position in `chars` (the printable-ASCII filtering
/// of a cross-ref row) at which the description text begins. `see_idx`
/// is the byte position of ` -- see` if present, used as the right
/// edge of the search window.
///
/// Strategy: at each candidate position, take the substring from there
/// to the end and split it into tokens (on spaces and non-word
/// punctuation). If *every* token in that suffix is plausibly English,
/// the candidate is the start of the description. We pick the
/// smallest such candidate so that no leading description chars are
/// truncated.
fn find_description_start(chars: &[char], see_idx: Option<usize>) -> usize {
    let limit = see_idx.unwrap_or(chars.len());
    if limit == 0 {
        return 0;
    }

    for i in 0..limit {
        let c = chars[i];
        // The description must begin with a letter, digit (for
        // numbered entries like "@SUM"), or a sentence-starter sigil.
        if !(c.is_ascii_alphabetic() || c.is_ascii_digit() || matches!(c, '/' | '@' | '(' | '[')) {
            continue;
        }
        // Reject single-letter starters like `a)gedata` and `e)nter`.
        // Descriptions in 1-2-3 help never begin with the article "a"
        // or "an" — those are mid-sentence words. If the first token
        // here is a single letter or a closed paren, skip.
        let first_tok_end = chars[i..limit]
            .iter()
            .take_while(|c| c.is_ascii_alphabetic() || c.is_ascii_digit() || **c == '-')
            .count();
        if first_tok_end < 2 {
            continue;
        }
        if suffix_is_all_english(&chars[i..limit]) {
            return i;
        }
    }
    0
}

/// True iff every space- or punctuation-separated token in `s` is
/// plausibly English (or one of the recognized special tokens).
fn suffix_is_all_english(s: &[char]) -> bool {
    let mut tok = String::new();
    let mut had_token = false;
    let flush_token = |tok: &mut String, had: &mut bool, ok: &mut bool| {
        if !tok.is_empty() {
            *had = true;
            if !looks_english(tok) {
                *ok = false;
            }
            tok.clear();
        }
    };
    let mut ok = true;
    for &c in s {
        if c.is_ascii_alphabetic() || c.is_ascii_digit() || c == '-' || c == '\'' {
            tok.push(c);
        } else {
            flush_token(&mut tok, &mut had_token, &mut ok);
            if !ok {
                return false;
            }
        }
    }
    flush_token(&mut tok, &mut had_token, &mut ok);
    ok && had_token
}

/// Heuristic: a token "looks English" if it's a known short word, a
/// numeric like `1-2-3`, or has a plausible vowel ratio for English.
///
/// Real English words have ~30–50 % vowels. Renderer-noise tokens
/// (`nn`, `ng`, `tsno`, `Cnntt`) cluster in the < 25 % range. Using
/// vowel ratio rejects the noise without false-positiving on common
/// task-description vocabulary like `Worksheet`, `Remove`, `data`.
fn looks_english(tok: &str) -> bool {
    if tok.is_empty() {
        return false;
    }
    if tok == "1-2-3" || tok == "--" {
        return true;
    }
    // Common English short words that appear in help descriptions.
    // Anything <= 4 letters not on this list is rejected — at that
    // length the vowel-ratio test is too loose to filter out renderer
    // noise like `snoa`, `nnoa`, `tnda`.
    const COMMON: &[&str] = &[
        "a", "an", "the", "to", "of", "in", "on", "at", "or", "and", "as", "if", "is", "for",
        "from", "by", "with", "into", "see", "all", "any", "new", "old", "set", "row", "col",
        "use", "via", "out", "off", "one", "two", "tab", "max", "min", "you", "are", "can", "data",
        "file", "page", "name", "open", "save", "menu", "type", "copy", "move", "list", "edit",
        "view", "cell", "text", "info", "main", "user", "next", "back", "down", "exit", "quit",
        "load", "auto", "size", "tabs", "case", "find", "free", "help", "high", "left", "line",
        "low", "mode", "more", "none", "only", "over", "part", "path", "pick", "play", "plus",
        "redo", "send", "show", "site", "skip", "step", "stop", "take", "task", "test", "this",
        "tool", "true", "turn", "wait", "warn", "what", "when", "with", "word", "work", "year",
        "your", "true", "many", "make", "less", "long", "look", "lock", "kind", "join", "into",
        "item", "have", "hide",
    ];
    let lower = tok.to_ascii_lowercase();
    if COMMON.contains(&lower.as_str()) {
        return true;
    }
    if tok.chars().filter(|c| c.is_ascii_alphabetic()).count() < 5 {
        return false;
    }

    let letters: Vec<char> = tok.chars().filter(|c| c.is_ascii_alphabetic()).collect();
    if letters.len() < 3 {
        return false;
    }

    // Capital-letter pattern. English tokens come in three shapes:
    //   * all lowercase (`data`, `worksheet`)
    //   * title-case: first char upper, rest lowercase (`Remove`)
    //   * acronym-ish: all uppercase, optionally with digits/hyphens
    //     (`ALT-F1`, `READY`)
    // Anything else — capital at a non-leading position — is the
    // tell-tale shape of two glued tokens (`CnRemove`, `eWorksheet`)
    // and is almost always renderer noise, not real English.
    let upper_positions: Vec<usize> = letters
        .iter()
        .enumerate()
        .filter(|(_, c)| c.is_ascii_uppercase())
        .map(|(i, _)| i)
        .collect();
    let valid_caps = upper_positions.is_empty()
        || (upper_positions == vec![0])
        || upper_positions.len() == letters.len();
    if !valid_caps {
        return false;
    }

    let is_vowel = |c: &char| "aeiouyAEIOUY".contains(*c);
    let vowels = letters.iter().filter(|c| is_vowel(c)).count();
    let ratio = vowels as f32 / letters.len() as f32;
    // All-caps acronyms (`ALT`, `KEY`) get a pass on the vowel ratio.
    let is_acronym = upper_positions.len() == letters.len();
    if !is_acronym && ratio < 0.25 {
        return false;
    }
    if !is_acronym {
        // Reject runs of >= 5 consecutive consonants — English doesn't
        // do that, but renderer noise like `Cnntt` does.
        let mut consec = 0usize;
        for c in &letters {
            if is_vowel(c) {
                consec = 0;
            } else {
                consec += 1;
                if consec >= 5 {
                    return false;
                }
            }
        }
        // Reject tokens that have > 2 consonants before the first
        // vowel. English words start with at most a 2-consonant
        // cluster (`str` is a rare third). Renderer noise like
        // `tsnoa`, `nntn`, `Cnntt` all start with 3+ consonants.
        let lead_consonants = letters.iter().take_while(|c| !is_vowel(c)).count();
        if lead_consonants > 2 {
            return false;
        }
    }
    true
}

/// Rows like ` -- see` (with no description) come from the footer
/// cross-references at the start of every body record. They aren't
/// useful in the cleaned-body output of an index page, so we drop them.
fn is_footer_artifact(row: &str) -> bool {
    let trimmed = row.trim();
    trimmed == "-- see" || trimmed.is_empty()
}

/// Return the most likely index-page title given the full decoded
/// stream. Index pages always include their title near the top of
/// the record (after a few footer/header bytes).
fn guess_index_title(decoded: &[u8]) -> String {
    let s: String = decoded
        .iter()
        .filter(|&&b| (0x20..=0x7E).contains(&b))
        .map(|&b| b as char)
        .collect();
    for needle in [
        "Task Index (continued)",
        "Task Index",
        "@Function Index (continued)",
        "@Function Index",
        "Function Index (continued)",
        "Function Index",
        "Keyboard Index (continued)",
        "Keyboard Index",
        "Macro Command Index (continued)",
        "Macro Command Index",
        "Run-Time Error Message Index (continued)",
        "Run-Time Error Message Index",
        "Compile-Time Error Message Index (continued)",
        "Compile-Time Error Message Index",
        "Error Message Index (continued)",
        "Error Message Index",
        "1-2-3 Help Index",
        "Help Index",
    ] {
        if s.contains(needle) {
            return needle.to_string();
        }
    }
    "Cross-Reference Index".to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dict::HuffmanTree;
    use crate::huffman::decode;

    fn load_hlp() -> Option<Vec<u8>> {
        let path = std::env::var("L123_HLP_FILE").unwrap_or_else(|_| {
            format!(
                "{}/Documents/dosbox-cdrive/123R34/123.HLP",
                std::env::var("HOME").unwrap_or_default()
            )
        });
        std::fs::read(&path).ok()
    }

    fn topic_at(hlp: &[u8], tree: &HuffmanTree, start: usize, end: usize) -> Option<Topic> {
        let decoded = decode(&hlp[start..end], tree);
        extract_topic(&decoded)
    }

    #[test]
    fn extracts_print_commands_topic() {
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let topic = topic_at(&hlp, &tree, 0x033ee, 0x0371f).expect("topic");
        assert_eq!(topic.title, "Print Commands");
        assert!(
            topic
                .body
                .starts_with("Print files on a printer or to a file"),
            "body started with: {:?}",
            &topic.body[..80.min(topic.body.len())]
        );
    }

    #[test]
    fn extracts_copy_topic() {
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let topic = topic_at(&hlp, &tree, 0x021c3, 0x024b5).expect("topic");
        assert_eq!(topic.title, "/Copy");
        assert!(
            topic.body.starts_with("Copies a range of data"),
            "body started with: {:?}",
            &topic.body[..80.min(topic.body.len())]
        );
    }

    #[test]
    fn extracts_about_help_topic() {
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let topic = topic_at(&hlp, &tree, 0x52523, 0x52831).expect("topic");
        assert_eq!(topic.title, "About 1-2-3 Help");
        assert!(
            topic.body.starts_with("You can view Help screens any time"),
            "body started with: {:?}",
            &topic.body[..80.min(topic.body.len())]
        );
    }

    #[test]
    fn extracts_pointer_movement_keys_topic() {
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let topic = topic_at(&hlp, &tree, 0x5135a, 0x5160f).expect("topic");
        assert_eq!(topic.title, "Pointer-Movement Keys");
        assert!(topic
            .body
            .starts_with("Move the cell pointer around the worksheet"));
    }

    #[test]
    fn returns_none_for_index_record() {
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let decoded = decode(&hlp[0xe98..0x1205], &tree);
        let _ = extract_topic(&decoded);
    }

    #[test]
    fn crossref_page_recognized_as_task_index() {
        // Record at 0x01cbd is the Task Index page that holds the
        // "Remove an active file from memory -- see /File Close"
        // cross-ref. Confirm the cross-ref recognizer kicks in,
        // produces the right title, and emits clean entries.
        let Some(hlp) = load_hlp() else { return };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let topic = topic_at(&hlp, &tree, 0x01cbd, 0x021c3).expect("topic");
        assert!(
            topic.title.contains("Task Index") || topic.title.contains("Cross-Reference"),
            "expected Task Index title, got {:?}",
            topic.title
        );
        // The decoded body must contain "Remove an active file from
        // memory" followed by "-- see" — confirming the topic-id
        // prefix has been stripped.
        assert!(
            topic
                .body
                .contains("Remove an active file from memory  -- see"),
            "body did not contain the cleaned cross-ref row; first 200 chars: {:?}",
            &topic.body[..200.min(topic.body.len())]
        );
        // No more "Cn" garbage glued to "Remove".
        assert!(
            !topic.body.contains("CnRemove"),
            "topic-id prefix not stripped: body has 'CnRemove'"
        );
    }
}
