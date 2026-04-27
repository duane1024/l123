//! Tiny HTML extractor for the Lotus help corpus.
//!
//! The corpus is regular: each file is `<title>X</title>` followed by
//! `<h1>X</h1><pre>BODY</pre>`. The body is plain text with inline
//! `<a href="...">...</a>` hyperlinks and a fixed set of HTML entities
//! (`&amp; &lt; &gt; &quot; &#x27; &#39;`). No nested tags, no scripts,
//! no styles inside the pre block.
//!
//! [`parse`] returns `(title, body, links)`:
//! - `title` and `body` are plain text with entities decoded.
//! - Each `<a>` tag is replaced in `body` by its display text, and a
//!   [`super::HelpLink`] records the byte range plus target filename.

use super::HelpLink;

pub(crate) fn parse(html: &str) -> (String, String, Vec<HelpLink>) {
    let title = extract_tag(html, "title")
        .map(decode_entities)
        .unwrap_or_default();
    let pre = extract_tag(html, "pre").unwrap_or("");
    let (body, links) = extract_body_and_links(pre);
    (title, body, links)
}

fn extract_tag<'a>(html: &'a str, tag: &str) -> Option<&'a str> {
    let open = format!("<{tag}>");
    let close = format!("</{tag}>");
    let start = html.find(&open)? + open.len();
    let end = html[start..].find(&close)? + start;
    Some(&html[start..end])
}

fn extract_body_and_links(pre: &str) -> (String, Vec<HelpLink>) {
    let mut body = String::with_capacity(pre.len());
    let mut links = Vec::new();
    let bytes = pre.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'<' {
            // Look for `<a href="...">...</a>`.
            if let Some((href, text_in_html, after)) = parse_anchor(&pre[i..]) {
                let start = body.len();
                let decoded = decode_entities(text_in_html);
                body.push_str(&decoded);
                let end = body.len();
                links.push(HelpLink {
                    start,
                    end,
                    target: href,
                });
                i += after;
                continue;
            }
            // Some other tag — skip until `>`.
            if let Some(rel) = pre[i..].find('>') {
                i += rel + 1;
                continue;
            }
            // Malformed: emit literal `<` and move on.
            body.push('<');
            i += 1;
            continue;
        }
        if bytes[i] == b'&' {
            if let Some((decoded_char, len)) = decode_one_entity(&pre[i..]) {
                body.push(decoded_char);
                i += len;
                continue;
            }
        }
        // Push the next UTF-8 char.
        let ch_end = next_char_boundary(pre, i);
        body.push_str(&pre[i..ch_end]);
        i = ch_end;
    }
    (body, links)
}

/// Try to parse `<a href="X">TEXT</a>` starting at the beginning of
/// `s`. Returns `(href, raw text, bytes consumed)` on a hit.
fn parse_anchor(s: &str) -> Option<(String, &str, usize)> {
    let s_lower = s.as_bytes();
    if !s_lower.starts_with(b"<a ") && !s_lower.starts_with(b"<A ") {
        return None;
    }
    // Find the closing `>` of the open tag.
    let open_end = s.find('>')?;
    let open_tag = &s[..open_end];
    let href = extract_attr(open_tag, "href")?;
    // Find the matching `</a>`.
    let after_open = open_end + 1;
    let close_rel = find_case_insensitive(&s[after_open..], "</a>")?;
    let text = &s[after_open..after_open + close_rel];
    let consumed = after_open + close_rel + "</a>".len();
    Some((href, text, consumed))
}

fn extract_attr(tag: &str, name: &str) -> Option<String> {
    let lower = tag.to_ascii_lowercase();
    let key = format!("{name}=\"");
    let pos = lower.find(&key)?;
    let value_start = pos + key.len();
    let rel_end = tag[value_start..].find('"')?;
    Some(tag[value_start..value_start + rel_end].to_string())
}

fn find_case_insensitive(haystack: &str, needle: &str) -> Option<usize> {
    let h = haystack.as_bytes();
    let n = needle.as_bytes();
    if n.is_empty() || h.len() < n.len() {
        return None;
    }
    'outer: for i in 0..=h.len() - n.len() {
        for j in 0..n.len() {
            if !h[i + j].eq_ignore_ascii_case(&n[j]) {
                continue 'outer;
            }
        }
        return Some(i);
    }
    None
}

fn next_char_boundary(s: &str, mut i: usize) -> usize {
    i += 1;
    while i < s.len() && !s.is_char_boundary(i) {
        i += 1;
    }
    i
}

fn decode_entities(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'&' {
            if let Some((c, len)) = decode_one_entity(&s[i..]) {
                out.push(c);
                i += len;
                continue;
            }
        }
        let ch_end = next_char_boundary(s, i);
        out.push_str(&s[i..ch_end]);
        i = ch_end;
    }
    out
}

/// If `s` begins with a recognized HTML entity (`&...;`), return the
/// decoded char and the byte length consumed.
fn decode_one_entity(s: &str) -> Option<(char, usize)> {
    let bytes = s.as_bytes();
    if bytes.first() != Some(&b'&') {
        return None;
    }
    let semi = s[1..].find(';')?;
    let name = &s[1..1 + semi];
    let consumed = 2 + semi; // '&' + name + ';'
    let c = match name {
        "amp" => '&',
        "lt" => '<',
        "gt" => '>',
        "quot" => '"',
        "apos" => '\'',
        "nbsp" => '\u{00A0}',
        n if n.starts_with("#x") || n.starts_with("#X") => {
            let hex = &n[2..];
            let code = u32::from_str_radix(hex, 16).ok()?;
            char::from_u32(code)?
        }
        n if n.starts_with('#') => {
            let dec = &n[1..];
            let code: u32 = dec.parse().ok()?;
            char::from_u32(code)?
        }
        _ => return None,
    };
    Some((c, consumed))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_title_and_body_text() {
        let html = "<title>Hello</title><h1>Hello</h1><pre>One\nTwo</pre>";
        let (title, body, links) = parse(html);
        assert_eq!(title, "Hello");
        assert_eq!(body, "One\nTwo");
        assert!(links.is_empty());
    }

    #[test]
    fn anchor_text_inlines_with_byte_range() {
        let html = r#"<title>X</title><pre>see <a href="0007-copy-continued.html">Continued</a> end</pre>"#;
        let (_, body, links) = parse(html);
        assert_eq!(body, "see Continued end");
        assert_eq!(links.len(), 1);
        assert_eq!(links[0].target, "0007-copy-continued.html");
        assert_eq!(&body[links[0].start..links[0].end], "Continued");
    }

    #[test]
    fn entities_decode() {
        let html = "<title>&amp;</title><pre>&lt;a&gt; &quot;hi&quot; it&#x27;s</pre>";
        let (title, body, _) = parse(html);
        assert_eq!(title, "&");
        assert_eq!(body, "<a> \"hi\" it's");
    }

    #[test]
    fn missing_pre_yields_empty_body() {
        let html = "<title>X</title>";
        let (title, body, links) = parse(html);
        assert_eq!(title, "X");
        assert_eq!(body, "");
        assert!(links.is_empty());
    }

    #[test]
    fn unknown_entity_passes_through() {
        let html = "<title>X</title><pre>&unknown;</pre>";
        let (_, body, _) = parse(html);
        assert_eq!(body, "&unknown;");
    }

    #[test]
    fn entity_inside_anchor_text_decodes() {
        let html = r#"<pre><a href="x.html">a&amp;b</a></pre>"#;
        let (_, body, links) = parse(html);
        assert_eq!(body, "a&b");
        assert_eq!(&body[links[0].start..links[0].end], "a&b");
    }
}
