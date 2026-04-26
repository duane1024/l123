//! Huffman decoder for `123.HLP` compressed records.
//!
//! Each record is a stream of MSB-first bits. At each bit, the decoder
//! traverses the static tree from [`crate::dict::HuffmanTree`]. A
//! non-negative node child is a literal output byte; the decoder emits
//! it and restarts at node 0. A negative child is an internal-node
//! reference: jump to node index `-value` and consume the next bit.
//!
//! The branch sense is **inverted**: at bit `k` of an input byte,
//! select the `right` child when `((byte >> k) & 1) ^ 1` is 1, else
//! `left`. Equivalent to `right` when bit `k` is 0.

use crate::dict::HuffmanTree;

/// Decode `input` as a Huffman-compressed bitstream against `tree`.
///
/// The decoder stops when `input` is exhausted (mid-codeword tails are
/// silently truncated — the file's record framing controls record
/// length so the last byte may have unused trailing bits).
pub fn decode(input: &[u8], tree: &HuffmanTree) -> Vec<u8> {
    let mut out = Vec::with_capacity(input.len() * 2);
    let mut node: usize = 0;
    let nodes = tree.nodes();
    for &byte in input {
        for k in (0..8).rev() {
            let bit = ((byte >> k) & 1) ^ 1;
            let v = if bit == 0 {
                nodes[node].left
            } else {
                nodes[node].right
            };
            if v < 0 {
                node = (-v) as usize;
            } else {
                out.push((v as u16 & 0xff) as u8);
                node = 0;
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dict::HuffmanTree;

    fn load_hlp() -> Option<Vec<u8>> {
        let path = std::env::var("L123_HLP_FILE").unwrap_or_else(|_| {
            format!(
                "{}/Documents/dosbox-cdrive/123R34/123.HLP",
                std::env::var("HOME").unwrap_or_default()
            )
        });
        std::fs::read(&path).ok()
    }

    fn contains(haystack: &[u8], needle: &[u8]) -> bool {
        haystack.windows(needle.len()).any(|w| w == needle)
    }

    #[test]
    fn decodes_first_record_to_help_index() {
        // First record per the offset table: 0xe98..0x1205 — the
        // "1-2-3 Help Index" page (first thing the user sees on F1).
        // Cross-reference rows on this page have renderer-control
        // prefixes, so the topic *titles* don't appear contiguously
        // in raw decoded bytes; the footer text does.
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found");
            return;
        };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let decoded = decode(&hlp[0xe98..0x1205], &tree);
        assert!(contains(&decoded, b"Help Index"));
        assert!(contains(
            &decoded,
            b"select a topic, press a pointer-movement key to highlight the topic"
        ));
        assert!(contains(&decoded, b"press ENTER"));
        assert!(contains(&decoded, b"BACKSPACE"));
    }

    #[test]
    fn decodes_record_with_record_feature_help() {
        // Record 0x3ae0c..0x3b0a3 contains a clean body fragment about
        // ALT-F2 (RECORD) — confirmed by Python decoder.
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found");
            return;
        };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let decoded = decode(&hlp[0x3ae0c..0x3b0a3], &tree);
        assert!(contains(
            &decoded,
            b"With 1-2-3 in READY mode, press ALT-F2 (RECORD) and select Trace to turn on"
        ));
    }

    #[test]
    fn decodes_record_with_print_margins_help() {
        // Record 0x48482..0x487a8 contains a clean body fragment about
        // /Print [B,E,F,P] Options Margins.
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found");
            return;
        };
        let tree = HuffmanTree::parse(&hlp).unwrap();
        let decoded = decode(&hlp[0x48482..0x487a8], &tree);
        assert!(contains(
            &decoded,
            b"corresponding /Print command is /Print [B,E,F,P] Options Margins [L,R,T,B]."
        ));
    }
}
