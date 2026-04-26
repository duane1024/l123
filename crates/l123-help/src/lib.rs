//! Decoder for the original Lotus 1-2-3 R3.4a help file (`123.HLP`).
//!
//! The file uses static Huffman compression over a 106-node tree
//! stored at the start of the file. Each compressed record is a
//! stream of MSB-first bits with the branch sense inverted; tree
//! nodes are pairs of `i16` where a non-negative child is a literal
//! output byte and a negative child is an internal-node reference.
//!
//! See `docs/HLP_DECODE_NOTES.md` for the format reconnaissance.

pub mod dict;
pub mod huffman;
pub mod renderer;

use std::collections::HashSet;
use thiserror::Error;

pub use renderer::Topic;

#[derive(Debug, Error)]
pub enum HlpError {
    #[error("file too short: have {got} bytes, need at least {need}")]
    Truncated { got: usize, need: usize },

    #[error(transparent)]
    Tree(#[from] dict::TreeError),
}

/// Decode every body topic from a `123.HLP` byte slice.
///
/// Walks the offset table, Huffman-decodes each record, and runs the
/// renderer's title/body extractor. The result is deduplicated by
/// title (first occurrence wins) so callers see one entry per unique
/// topic even though many records appear multiple times in the file.
pub fn topics(hlp_bytes: &[u8]) -> Result<Vec<Topic>, HlpError> {
    // Header sanity: need at least the tree + four bytes for the first
    // entry of the offset table.
    let need = dict::TREE_OFFSET + dict::TREE_BYTES + 4;
    if hlp_bytes.len() < need {
        return Err(HlpError::Truncated {
            got: hlp_bytes.len(),
            need,
        });
    }

    let tree = dict::HuffmanTree::parse(hlp_bytes)?;

    // Offset-table location is stored as a u32 LE at file offset 0x0a.
    let ot_start = u32::from_le_bytes(hlp_bytes[0x0a..0x0e].try_into().unwrap()) as usize;
    if ot_start + 4 > hlp_bytes.len() {
        return Err(HlpError::Truncated {
            got: hlp_bytes.len(),
            need: ot_start + 4,
        });
    }
    // The first dword in the table points at the first compressed
    // record — walk the table from `ot_start` until we reach that
    // record.
    let first_record =
        u32::from_le_bytes(hlp_bytes[ot_start..ot_start + 4].try_into().unwrap()) as usize;
    if first_record > hlp_bytes.len() || first_record < ot_start {
        return Ok(Vec::new());
    }

    let mut offsets: Vec<usize> = Vec::new();
    let mut off = ot_start;
    while off + 4 <= first_record {
        let v = u32::from_le_bytes(hlp_bytes[off..off + 4].try_into().unwrap()) as usize;
        if v > 0 && v <= hlp_bytes.len() {
            offsets.push(v);
        }
        off += 4;
    }
    offsets.sort();
    offsets.dedup();

    let mut seen_titles: HashSet<String> = HashSet::new();
    let mut out: Vec<Topic> = Vec::new();
    for w in offsets.windows(2) {
        let (s, e) = (w[0], w[1]);
        if e <= s {
            continue;
        }
        let decoded = huffman::decode(&hlp_bytes[s..e], &tree);
        if let Some(t) = renderer::extract_topic(&decoded) {
            if seen_titles.insert(t.title.clone()) {
                out.push(t);
            }
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn load_hlp() -> Option<Vec<u8>> {
        let path = std::env::var("L123_HLP_FILE").unwrap_or_else(|_| {
            format!(
                "{}/Documents/dosbox-cdrive/123R34/123.HLP",
                std::env::var("HOME").unwrap_or_default()
            )
        });
        std::fs::read(&path).ok()
    }

    #[test]
    fn topics_handles_truncated_input() {
        let too_short = vec![0u8; 16];
        let err = topics(&too_short).unwrap_err();
        match err {
            HlpError::Truncated { .. } => {}
            other => panic!("expected Truncated, got {other:?}"),
        }
    }

    #[test]
    fn topics_extracts_at_least_two_hundred_unique_topics() {
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found");
            return;
        };
        let ts = topics(&hlp).expect("topics ok");
        assert!(
            ts.len() >= 200,
            "expected ≥200 unique decoded topics, got {}",
            ts.len()
        );
        // No duplicate titles in output.
        let mut titles: Vec<&str> = ts.iter().map(|t| t.title.as_str()).collect();
        titles.sort();
        let n = titles.len();
        titles.dedup();
        assert_eq!(titles.len(), n, "duplicate titles in topics()");
    }

    #[test]
    fn topics_includes_about_help() {
        let Some(hlp) = load_hlp() else { return };
        let ts = topics(&hlp).expect("topics ok");
        let about = ts
            .iter()
            .find(|t| t.title == "About 1-2-3 Help")
            .expect("'About 1-2-3 Help' present");
        assert!(
            about.body.starts_with("You can view Help screens any time"),
            "body started with: {:?}",
            &about.body[..80.min(about.body.len())]
        );
    }
}
