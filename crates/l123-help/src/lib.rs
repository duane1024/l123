//! HTML-backed help content for L123.
//!
//! The original Lotus 1-2-3 R3.4a help pages have been transcribed into
//! plain HTML files (one per topic) and live under `assets/help_html/`.
//! At build time the `build.rs` script generates a sorted
//! `(filename, file_contents)` slice that is `include_str!`'d into the
//! binary, so the help corpus ships embedded with no runtime I/O.
//!
//! [`load_page`] turns one of those static strings into a [`HelpPage`]:
//! plain-text body (HTML entities decoded, `<a>` text inlined) plus a
//! list of [`HelpLink`] byte-range spans into the body, each pointing
//! at the filename of another embedded page. The renderer in
//! `l123-ui` paints those spans as hyperlinks and follows them when
//! the user presses ENTER.

mod parser;

include!(concat!(env!("OUT_DIR"), "/help_files.rs"));

/// Filename of the help index — the page F1 opens by default.
pub const INDEX_FILENAME: &str = "0000-1-2-3-help-index.html";

/// One hyperlink inside a [`HelpPage`] body. `start..end` is a byte
/// range into [`HelpPage::body`] covering the link's display text;
/// `target` is the filename of another embedded page.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HelpLink {
    pub start: usize,
    pub end: usize,
    pub target: String,
}

/// A parsed help page ready to render.
#[derive(Debug, Clone)]
pub struct HelpPage {
    /// The filename this page was loaded from (e.g. `"0006-copy.html"`).
    pub filename: &'static str,
    /// Page title from `<title>...</title>`. HTML entities decoded.
    pub title: String,
    /// Plain-text body (the inside of `<pre>...</pre>`), with `<a>` tags
    /// flattened into their display text and entities decoded. Byte
    /// ranges inside [`HelpLink`] index into this string.
    pub body: String,
    /// Hyperlink spans, in the order they appear in the body.
    pub links: Vec<HelpLink>,
}

impl HelpPage {
    /// `(row, col)` of each link's first display character, in the same
    /// order as [`HelpPage::links`]. Row is the count of preceding `\n`
    /// in [`HelpPage::body`]; column is the byte offset since the last
    /// `\n` (the body is ASCII-equivalent after entity decoding, so
    /// byte offset == column on a monospace TUI). Used by 2D arrow-key
    /// navigation.
    pub fn link_positions(&self) -> Vec<(usize, usize)> {
        // `row_starts[r]` is the byte offset where row `r` begins.
        let mut row_starts: Vec<usize> = vec![0];
        for (i, b) in self.body.bytes().enumerate() {
            if b == b'\n' {
                row_starts.push(i + 1);
            }
        }
        self.links
            .iter()
            .map(|link| {
                let row = match row_starts.binary_search(&link.start) {
                    Ok(i) => i,
                    Err(i) => i.saturating_sub(1),
                };
                let col = link.start - row_starts[row];
                (row, col)
            })
            .collect()
    }
}

/// Return the raw embedded HTML for `filename`, or `None` if no such
/// page exists. Filenames are matched literally against the
/// `assets/help_html/` directory (e.g. `"0006-copy.html"`).
pub fn raw(filename: &str) -> Option<&'static str> {
    HELP_FILES
        .binary_search_by_key(&filename, |(name, _)| *name)
        .ok()
        .map(|i| HELP_FILES[i].1)
}

/// Load and parse the named help page.
pub fn load_page(filename: &str) -> Option<HelpPage> {
    let i = HELP_FILES
        .binary_search_by_key(&filename, |(name, _)| *name)
        .ok()?;
    let (name, html) = HELP_FILES[i];
    let (title, body, links) = parser::parse(html);
    Some(HelpPage {
        filename: name,
        title,
        body,
        links,
    })
}

/// Total number of embedded help pages. Useful for sanity tests.
pub fn page_count() -> usize {
    HELP_FILES.len()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn index_page_loads() {
        let page = load_page(INDEX_FILENAME).expect("index page must be embedded");
        assert_eq!(page.filename, INDEX_FILENAME);
        assert_eq!(page.title, "1-2-3 Help Index");
        assert!(!page.body.is_empty());
        assert!(!page.links.is_empty(), "index has hyperlinks");
    }

    #[test]
    fn copy_page_has_three_footer_links() {
        let page = load_page("0006-copy.html").expect("copy page");
        assert_eq!(page.title, "/Copy");
        assert!(page.body.contains("/Copy -- Copies a range"));
        let targets: Vec<&str> = page.links.iter().map(|l| l.target.as_str()).collect();
        assert!(
            targets.contains(&"0007-copy-continued.html"),
            "links: {targets:?}"
        );
        assert!(targets.contains(&"0365-specifying-ranges.html"));
        assert!(targets.contains(&"0000-1-2-3-help-index.html"));
    }

    #[test]
    fn unknown_page_returns_none() {
        assert!(load_page("nope.html").is_none());
        assert!(load_page("").is_none());
    }

    #[test]
    fn link_byte_ranges_are_valid_into_body() {
        let page = load_page("0006-copy.html").unwrap();
        for link in &page.links {
            assert!(
                page.body.is_char_boundary(link.start),
                "link start {} not on char boundary",
                link.start
            );
            assert!(
                page.body.is_char_boundary(link.end),
                "link end {} not on char boundary",
                link.end
            );
            assert!(link.start < link.end);
            assert!(
                !page.body[link.start..link.end].is_empty(),
                "link text empty"
            );
        }
    }

    #[test]
    fn index_first_link_is_help_index_self_reference() {
        let page = load_page(INDEX_FILENAME).unwrap();
        let first = &page.links[0];
        assert_eq!(first.target, "0201-help-index.html");
        assert_eq!(&page.body[first.start..first.end], "Help Index");
    }

    #[test]
    fn entities_in_body_are_decoded() {
        // pages with `&#x27;` for apostrophes should decode.
        let page = load_page("0006-copy.html").unwrap();
        // `0006-copy.html` itself doesn't have entities, but many do —
        // pick a known one. Walking every page sanity-checks the
        // decoder against the corpus.
        for (name, _) in HELP_FILES {
            if let Some(p) = load_page(name) {
                assert!(
                    !p.body.contains("&amp;")
                        && !p.body.contains("&lt;")
                        && !p.body.contains("&gt;")
                        && !p.body.contains("&quot;")
                        && !p.body.contains("&#x27;")
                        && !p.body.contains("&#39;"),
                    "page {name} has undecoded HTML entity in body"
                );
            }
        }
        let _ = page; // suppress unused-binding if the loop short-circuits
    }

    #[test]
    fn page_count_matches_corpus() {
        assert_eq!(page_count(), HELP_FILES.len());
        assert!(page_count() > 800);
    }

    #[test]
    fn link_positions_align_with_three_column_index() {
        let page = load_page(INDEX_FILENAME).unwrap();
        let positions = page.link_positions();
        assert_eq!(positions.len(), page.links.len());
        // The index page lays its links out in three columns. Find
        // the rows that have three links and verify the column triple
        // is roughly (left, middle, right).
        use std::collections::BTreeMap;
        let mut by_row: BTreeMap<usize, Vec<usize>> = BTreeMap::new();
        for &(r, c) in &positions {
            by_row.entry(r).or_default().push(c);
        }
        let three_link_rows: Vec<&Vec<usize>> =
            by_row.values().filter(|cols| cols.len() == 3).collect();
        assert!(
            three_link_rows.len() >= 5,
            "expected several 3-link rows, got {}",
            three_link_rows.len()
        );
        for cols in three_link_rows {
            let mut sorted = cols.clone();
            sorted.sort();
            assert!(sorted[0] < sorted[1] && sorted[1] < sorted[2]);
            // Sanity: gap between columns >= 5 chars on the index.
            assert!(sorted[1] - sorted[0] >= 5);
            assert!(sorted[2] - sorted[1] >= 5);
        }
    }

    #[test]
    fn link_positions_match_body_offsets() {
        let page = load_page("0006-copy.html").unwrap();
        let positions = page.link_positions();
        for (link, &(row, col)) in page.links.iter().zip(&positions) {
            // Re-derive the same (row, col) the slow way.
            let prefix = &page.body[..link.start];
            let expected_row = prefix.bytes().filter(|b| *b == b'\n').count();
            let last_nl = prefix.rfind('\n').map(|i| i + 1).unwrap_or(0);
            let expected_col = link.start - last_nl;
            assert_eq!((row, col), (expected_row, expected_col));
        }
    }

    #[test]
    fn every_link_target_is_embedded() {
        for (name, _) in HELP_FILES {
            let page = load_page(name).unwrap();
            for link in &page.links {
                assert!(
                    raw(&link.target).is_some(),
                    "{name} links to unknown {}",
                    link.target
                );
            }
        }
    }
}
