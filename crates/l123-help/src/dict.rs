//! BPE dictionary embedded in the `123.HLP` header.
//!
//! Layout: 106 entries, each a pair of `i16` little-endian values, at
//! file offsets `0x10..0x1b8`. Each entry expands a "code" into two
//! tokens — either positive byte literals or negative recursive
//! references to other dictionary slots.

use thiserror::Error;

/// Number of pair entries in the dictionary.
pub const DICT_LEN: usize = 106;

/// Byte offset of the dictionary inside `123.HLP`.
pub const DICT_OFFSET: usize = 0x10;

/// Byte length of the dictionary block (106 pairs × 2 × 2 bytes).
pub const DICT_BYTES: usize = DICT_LEN * 4;

#[derive(Debug, Error)]
pub enum DictError {
    #[error("file too short for dictionary: have {got} bytes, need at least {need}")]
    Truncated { got: usize, need: usize },
}

/// One BPE expansion: a code maps to a `(left, right)` token pair.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Pair {
    pub left: i16,
    pub right: i16,
}

/// The 106-entry dictionary parsed from a `123.HLP` byte slice.
#[derive(Debug, Clone)]
pub struct Dictionary {
    pairs: [Pair; DICT_LEN],
}

impl Dictionary {
    /// Parse the dictionary out of a `123.HLP` byte slice. Reads
    /// `DICT_BYTES` bytes starting at `DICT_OFFSET`.
    pub fn parse(bytes: &[u8]) -> Result<Self, DictError> {
        let need = DICT_OFFSET + DICT_BYTES;
        if bytes.len() < need {
            return Err(DictError::Truncated {
                got: bytes.len(),
                need,
            });
        }
        let mut pairs = [Pair { left: 0, right: 0 }; DICT_LEN];
        for (i, pair) in pairs.iter_mut().enumerate() {
            let off = DICT_OFFSET + i * 4;
            let l = i16::from_le_bytes([bytes[off], bytes[off + 1]]);
            let r = i16::from_le_bytes([bytes[off + 2], bytes[off + 3]]);
            *pair = Pair { left: l, right: r };
        }
        Ok(Self { pairs })
    }

    pub fn pairs(&self) -> &[Pair; DICT_LEN] {
        &self.pairs
    }

    /// Return the pair at slot `n` (1-based).
    pub fn pair(&self, slot_one_based: usize) -> Pair {
        self.pairs[slot_one_based - 1]
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
        let too_short = vec![0u8; DICT_OFFSET + 10];
        let err = Dictionary::parse(&too_short).unwrap_err();
        match err {
            DictError::Truncated { got, need } => {
                assert_eq!(got, DICT_OFFSET + 10);
                assert_eq!(need, DICT_OFFSET + DICT_BYTES);
            }
        }
    }

    #[test]
    fn synthetic_dictionary_round_trips() {
        let mut bytes = vec![0u8; DICT_OFFSET + DICT_BYTES];
        // Slot 1: (-1, -26)
        bytes[DICT_OFFSET..DICT_OFFSET + 2].copy_from_slice(&(-1i16).to_le_bytes());
        bytes[DICT_OFFSET + 2..DICT_OFFSET + 4].copy_from_slice(&(-26i16).to_le_bytes());
        // Slot 5: (104, -5)
        let s5 = DICT_OFFSET + 4 * 4;
        bytes[s5..s5 + 2].copy_from_slice(&104i16.to_le_bytes());
        bytes[s5 + 2..s5 + 4].copy_from_slice(&(-5i16).to_le_bytes());

        let d = Dictionary::parse(&bytes).unwrap();
        assert_eq!(
            d.pair(1),
            Pair {
                left: -1,
                right: -26
            }
        );
        assert_eq!(
            d.pair(5),
            Pair {
                left: 104,
                right: -5
            }
        );
    }

    #[test]
    fn real_hlp_dictionary_matches_recon() {
        let Some(hlp) = load_hlp() else {
            eprintln!("skipping: 123.HLP not found at $L123_HLP_FILE or default path");
            return;
        };
        let d = Dictionary::parse(&hlp).unwrap();
        // From docs/HLP_DECODE_NOTES.md:
        //   (-1, -26)   (-2, 196)   (-3, -8)    (-4, 97)    (104, -5)
        //   (-6, 112)   (-7, 119)   (84, 107)   (-9, 116)   (-10, 108)
        assert_eq!(
            d.pair(1),
            Pair {
                left: -1,
                right: -26
            }
        );
        assert_eq!(
            d.pair(2),
            Pair {
                left: -2,
                right: 196
            }
        );
        assert_eq!(
            d.pair(3),
            Pair {
                left: -3,
                right: -8
            }
        );
        assert_eq!(
            d.pair(4),
            Pair {
                left: -4,
                right: 97
            }
        );
        assert_eq!(
            d.pair(5),
            Pair {
                left: 104,
                right: -5
            }
        );
        assert_eq!(
            d.pair(6),
            Pair {
                left: -6,
                right: 112
            }
        );
        assert_eq!(
            d.pair(7),
            Pair {
                left: -7,
                right: 119
            }
        );
        assert_eq!(
            d.pair(8),
            Pair {
                left: 84,
                right: 107
            }
        );
        assert_eq!(
            d.pair(9),
            Pair {
                left: -9,
                right: 116
            }
        );
        assert_eq!(
            d.pair(10),
            Pair {
                left: -10,
                right: 108
            }
        );
    }
}
