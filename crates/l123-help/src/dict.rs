//! Static Huffman tree embedded in the `123.HLP` header.
//!
//! Layout: 106 nodes, each a pair of `i16` little-endian values, at
//! file offsets `0x10..0x1b8` (424 bytes total; `0x1a8` is also stored
//! at file offset `0x0e` as the tree byte length).
//!
//! Each pair `(left, right)` is one internal node:
//!   * a non-negative value is a **literal output byte** to emit when
//!     that branch is selected; after emitting, the decoder restarts
//!     at the root.
//!   * a negative value is an **internal-node reference**: jump to
//!     node index `-value` and consume another bit.
//!
//! Bit order is MSB-first with the branch sense inverted: at each
//! input byte, walk bits 7..=0 and pick `right` when `((byte >> bit)
//! & 1) ^ 1 == 1`, else `left`.

use thiserror::Error;

/// Number of node entries in the tree.
pub const TREE_NODES: usize = 106;

/// Byte offset of the tree inside `123.HLP`.
pub const TREE_OFFSET: usize = 0x10;

/// Byte length of the tree block (106 nodes × 2 × 2 bytes).
pub const TREE_BYTES: usize = TREE_NODES * 4;

#[derive(Debug, Error)]
pub enum TreeError {
    #[error("file too short for huffman tree: have {got} bytes, need at least {need}")]
    Truncated { got: usize, need: usize },
}

/// One Huffman node: `(left, right)` children.
///
/// Either child may be a literal byte (non-negative) or a reference
/// to another node (negative; the absolute value is the node index).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Node {
    pub left: i16,
    pub right: i16,
}

/// Static Huffman tree parsed from a `123.HLP` byte slice.
#[derive(Debug, Clone)]
pub struct HuffmanTree {
    nodes: [Node; TREE_NODES],
}

impl HuffmanTree {
    /// Parse the tree out of a `123.HLP` byte slice. Reads `TREE_BYTES`
    /// bytes starting at `TREE_OFFSET`.
    pub fn parse(bytes: &[u8]) -> Result<Self, TreeError> {
        let need = TREE_OFFSET + TREE_BYTES;
        if bytes.len() < need {
            return Err(TreeError::Truncated {
                got: bytes.len(),
                need,
            });
        }
        let mut nodes = [Node { left: 0, right: 0 }; TREE_NODES];
        for (i, node) in nodes.iter_mut().enumerate() {
            let off = TREE_OFFSET + i * 4;
            let l = i16::from_le_bytes([bytes[off], bytes[off + 1]]);
            let r = i16::from_le_bytes([bytes[off + 2], bytes[off + 3]]);
            *node = Node { left: l, right: r };
        }
        Ok(Self { nodes })
    }

    pub fn nodes(&self) -> &[Node; TREE_NODES] {
        &self.nodes
    }

    /// Return the node at `index` (0-based).
    pub fn node(&self, index: usize) -> Node {
        self.nodes[index]
    }
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
    fn truncated_input_errors() {
        let too_short = vec![0u8; TREE_OFFSET + 10];
        let err = HuffmanTree::parse(&too_short).unwrap_err();
        match err {
            TreeError::Truncated { got, need } => {
                assert_eq!(got, TREE_OFFSET + 10);
                assert_eq!(need, TREE_OFFSET + TREE_BYTES);
            }
        }
    }

    #[test]
    fn synthetic_tree_round_trips() {
        let mut bytes = vec![0u8; TREE_OFFSET + TREE_BYTES];
        // Node 0: (-1, -26)
        bytes[TREE_OFFSET..TREE_OFFSET + 2].copy_from_slice(&(-1i16).to_le_bytes());
        bytes[TREE_OFFSET + 2..TREE_OFFSET + 4].copy_from_slice(&(-26i16).to_le_bytes());
        // Node 4 (= slot 5 in 1-indexed terms): (104, -5)
        let n5 = TREE_OFFSET + 4 * 4;
        bytes[n5..n5 + 2].copy_from_slice(&104i16.to_le_bytes());
        bytes[n5 + 2..n5 + 4].copy_from_slice(&(-5i16).to_le_bytes());

        let t = HuffmanTree::parse(&bytes).unwrap();
        assert_eq!(
            t.node(0),
            Node {
                left: -1,
                right: -26
            }
        );
        assert_eq!(
            t.node(4),
            Node {
                left: 104,
                right: -5
            }
        );
    }

    #[test]
    fn real_hlp_tree_matches_recon() {
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found at $L123_HLP_FILE or default path");
            return;
        };
        let t = HuffmanTree::parse(&hlp).unwrap();
        // First few nodes from docs/HLP_DECODE_NOTES.md (1-indexed slot
        // numbers there → 0-indexed node numbers here):
        //   slot 1  (-1,  -26)   slot 2  (-2, 196)   slot 3  (-3, -8)
        //   slot 4  (-4,  97)    slot 5  (104, -5)   slot 6  (-6, 112)
        //   slot 7  (-7,  119)   slot 8  (84, 107)   slot 9  (-9, 116)
        //   slot 10 (-10, 108)
        assert_eq!(
            t.node(0),
            Node {
                left: -1,
                right: -26
            }
        );
        assert_eq!(
            t.node(1),
            Node {
                left: -2,
                right: 196
            }
        );
        assert_eq!(
            t.node(2),
            Node {
                left: -3,
                right: -8
            }
        );
        assert_eq!(
            t.node(3),
            Node {
                left: -4,
                right: 97
            }
        );
        assert_eq!(
            t.node(4),
            Node {
                left: 104,
                right: -5
            }
        );
        assert_eq!(
            t.node(5),
            Node {
                left: -6,
                right: 112
            }
        );
        assert_eq!(
            t.node(6),
            Node {
                left: -7,
                right: 119
            }
        );
        assert_eq!(
            t.node(7),
            Node {
                left: 84,
                right: 107
            }
        );
        assert_eq!(
            t.node(8),
            Node {
                left: -9,
                right: 116
            }
        );
        assert_eq!(
            t.node(9),
            Node {
                left: -10,
                right: 108
            }
        );
    }
}
