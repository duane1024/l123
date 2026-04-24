//! Decoder for Section 4 of Lotus 1-2-3 R3.4a's `ICONS3.DAT`: the
//! 8-color RLE bitmaps that accompany the mono catalog in Section 3.
//!
//! ## Wire format (reverse-engineered)
//!
//! Every icon record is a sequence of exactly 72 opcodes. Each opcode
//! produces one 8-pixel chunk; 72 chunks × 8 pixels fill the same
//! 24×24 frame as the mono catalog. "Higher-resolution" was a
//! misnomer: the improvement is colour depth (8-entry palette), not
//! spatial resolution.
//!
//! The opcode byte is a **plane bitmask**. Bit `k` set means "plane
//! `k` has a data byte following this opcode", so the number of data
//! bytes equals `popcount(opcode)`. For the current 8-pixel chunk,
//! bit `p` of plane `k`'s data byte says "the pixel at column `7−p`
//! of this chunk is colour `k`". The file's invariant — verified on
//! all 105 catalog entries — is that the OR of every opcode's plane
//! bytes is `0xFF` and their pairwise AND is zero, so every pixel is
//! claimed by exactly one plane.
//!
//! ```text
//!   01 FF          → plane 0 = 0xFF: all 8 pixels colour 0 (white)
//!   02 FF          → plane 1 = 0xFF: all 8 pixels colour 1 (black)
//!   03 EF 10       → plane 0 = EF (7 cols white) + plane 1 = 10 (1 col black)
//!   43 C7 28 10    → planes 0, 1, 6 populated (5 cols white, 2 black, 1 grey)
//! ```
//!
//! ## Palette
//!
//! The 8 palette slots were recovered by sampling a DOSBox-X
//! screenshot of the original WYSIWYG panel at positions the decoder
//! says carry each plane. The result is a specific subset of the EGA
//! 16-colour palette, in plane order: light-grey, black, yellow,
//! blue, cyan, red, magenta, bright-green. See [`LOTUS_PALETTE_RGB`].

use thiserror::Error;

/// Width and height of every icon in the catalog, in pixels.
pub const ICON_DIM: usize = 24;
/// Total pixels per icon (row-major, 0..=7 palette index).
pub const ICON_PIXELS: usize = ICON_DIM * ICON_DIM;
/// Number of opcodes in every Section-4 record: one 8-pixel chunk per
/// 3-byte mono row, 24 rows.
pub const OPCODES_PER_RECORD: usize = ICON_DIM * 3;

/// The 8 colours Lotus 1-2-3 R3.4a WYSIWYG draws its icon panel in.
///
/// Derived by sampling a DOSBox-X screenshot of the panel at pixel
/// positions the decoder says carry each plane. Values are IBM
/// EGA/VGA defaults, in `[R, G, B]` order:
///
/// | Plane | EGA # | Colour                |
/// |------:|------:|-----------------------|
/// | 0     | 7     | light grey (panel bg) |
/// | 1     | 0     | black (outline / ink) |
/// | 2     | 14    | yellow                |
/// | 3     | 1     | blue                  |
/// | 4     | 3     | cyan                  |
/// | 5     | 4     | red                   |
/// | 6     | 5     | magenta               |
/// | 7     | 10    | bright green          |
pub const LOTUS_PALETTE_RGB: [[u8; 3]; 8] = [
    [0xAA, 0xAA, 0xAA],
    [0x00, 0x00, 0x00],
    [0xFF, 0xFF, 0x55],
    [0x00, 0x00, 0xAA],
    [0x00, 0xAA, 0xAA],
    [0xAA, 0x00, 0x00],
    [0xAA, 0x00, 0xAA],
    [0x00, 0xFF, 0x00],
];

/// 24×24 bitmap of palette indices, one byte per pixel.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ColorBitmap {
    pub pixels: [u8; ICON_PIXELS],
}

impl ColorBitmap {
    /// Colour at `(row, col)` as a palette index 0..=7. Panics on OOB.
    pub fn get(&self, row: usize, col: usize) -> u8 {
        self.pixels[row * ICON_DIM + col]
    }
}

#[derive(Debug, Error, PartialEq, Eq)]
pub enum RleError {
    #[error("truncated record at byte {at}: opcode 0x{op:02x} needs {need} more bytes")]
    Truncated { at: usize, op: u8, need: usize },
    #[error("unexpected opcode 0x00 at byte {at}")]
    ZeroOpcode { at: usize },
    #[error(
        "wrong opcode count: record yielded {got}, expected {}",
        OPCODES_PER_RECORD
    )]
    WrongCount { got: usize },
    #[error("trailing bytes after {opcodes} opcodes: {remaining} unused")]
    Trailing { opcodes: usize, remaining: usize },
    #[error("planes overlap at byte {at}: bits 0x{overlap:02x} set in multiple planes")]
    PlaneOverlap { at: usize, overlap: u8 },
    #[error("planes leave pixels uncovered at byte {at}: bits 0x{gap:02x} have no colour")]
    PlaneGap { at: usize, gap: u8 },
}

/// Decode one Section-4 RLE record into a 24×24 colour bitmap.
pub fn decode_color_bitmap(rec: &[u8]) -> Result<ColorBitmap, RleError> {
    let mut pixels = [0u8; ICON_PIXELS];
    let mut cursor = 0usize;
    let mut opcode_idx = 0usize;

    while cursor < rec.len() {
        let op_at = cursor;
        let op = rec[cursor];
        if op == 0 {
            return Err(RleError::ZeroOpcode { at: op_at });
        }
        let n_planes = op.count_ones() as usize;
        let needed = 1 + n_planes;
        if rec.len() - cursor < needed {
            return Err(RleError::Truncated {
                at: op_at,
                op,
                need: needed - (rec.len() - cursor),
            });
        }

        // Gather (plane_index, data_byte) pairs. Plane `k` is present
        // iff bit `k` of the opcode is set, walked LSB-first so the
        // plane ordering in the byte stream matches plane index order.
        let mut plane_data: [Option<u8>; 8] = [None; 8];
        let mut data_i = 0;
        for (k, slot) in plane_data.iter_mut().enumerate() {
            if (op >> k) & 1 == 1 {
                *slot = Some(rec[cursor + 1 + data_i]);
                data_i += 1;
            }
        }

        // Invariant: every pixel belongs to exactly one plane.
        let mut union: u8 = 0;
        let mut inter: u8 = 0;
        for b in plane_data.iter().flatten() {
            inter |= union & *b;
            union |= *b;
        }
        if inter != 0 {
            return Err(RleError::PlaneOverlap {
                at: op_at,
                overlap: inter,
            });
        }
        if union != 0xFF {
            return Err(RleError::PlaneGap {
                at: op_at,
                gap: !union,
            });
        }

        if opcode_idx >= OPCODES_PER_RECORD {
            return Err(RleError::WrongCount {
                got: opcode_idx + 1,
            });
        }

        let row = opcode_idx / 3;
        let byte_col = opcode_idx % 3;
        for bit in 0..8 {
            let col = byte_col * 8 + bit;
            let mask = 1u8 << (7 - bit);
            let mut color = 0u8;
            for (k, maybe) in plane_data.iter().enumerate() {
                if let Some(v) = maybe {
                    if v & mask != 0 {
                        color = k as u8;
                        break;
                    }
                }
            }
            pixels[row * ICON_DIM + col] = color;
        }

        cursor += needed;
        opcode_idx += 1;
    }

    if opcode_idx != OPCODES_PER_RECORD {
        return Err(RleError::WrongCount { got: opcode_idx });
    }
    if cursor != rec.len() {
        return Err(RleError::Trailing {
            opcodes: opcode_idx,
            remaining: rec.len() - cursor,
        });
    }

    Ok(ColorBitmap { pixels })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Icon 0 from `ICONS3.DAT`: a stubby left-pointing arrow with the
    /// full-width ink bar on row 23 (the catalog's icon separator).
    const ICON_0: &[u8] = &[
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x01, 0xff, 0x03, 0xef, 0x10, 0x01, 0xff, 0x01, 0xff, 0x03, 0xcf, 0x30, //
        0x01, 0xff, 0x01, 0xff, 0x03, 0x8f, 0x70, 0x01, 0xff, 0x01, 0xff, 0x03, //
        0x0f, 0xf0, 0x01, 0xff, 0x03, 0xfe, 0x01, 0x02, 0xff, 0x03, 0x7f, 0x80, //
        0x03, 0xfc, 0x03, 0x02, 0xff, 0x03, 0x7f, 0x80, 0x03, 0xf8, 0x07, 0x02, //
        0xff, 0x03, 0x7f, 0x80, 0x03, 0xfc, 0x03, 0x02, 0xff, 0x03, 0x7f, 0x80, //
        0x03, 0xfe, 0x01, 0x02, 0xff, 0x03, 0x7f, 0x80, 0x01, 0xff, 0x03, 0x0f, //
        0xf0, 0x01, 0xff, 0x01, 0xff, 0x03, 0x8f, 0x70, 0x01, 0xff, 0x01, 0xff, //
        0x03, 0xcf, 0x30, 0x01, 0xff, 0x01, 0xff, 0x03, 0xef, 0x10, 0x01, 0xff, //
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x02, 0xff, 0x02, 0xff, 0x02, 0xff,
    ];

    /// Icon 5: multi-colour "file save" glyph exercising the 0x13 and
    /// 0x43 opcodes with three populated planes (colours 0, 1, 4, 6).
    const ICON_5: &[u8] = &[
        0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, //
        0x03, 0xef, 0x10, 0x01, 0xff, 0x01, 0xff, 0x43, 0xc7, 0x28, 0x10, 0x01, //
        0xff, 0x01, 0xff, 0x43, 0x83, 0x44, 0x38, 0x03, 0xbf, 0x40, 0x01, 0xff, //
        0x43, 0xc1, 0x22, 0x1c, 0x03, 0x3f, 0xc0, 0x01, 0xff, 0x43, 0xe0, 0x11, //
        0x0e, 0x43, 0x30, 0x4f, 0x80, 0x03, 0x1f, 0xe0, 0x43, 0xf0, 0x08, 0x07, //
        0x43, 0x3f, 0x40, 0x80, 0x03, 0xdf, 0x20, 0x43, 0xf8, 0x04, 0x03, 0x52, //
        0x40, 0x3f, 0x80, 0x13, 0x5f, 0x20, 0x80, 0x43, 0xf0, 0x08, 0x07, 0x52, //
        0x40, 0x3f, 0x80, 0x13, 0x1f, 0x20, 0xc0, 0x03, 0xe0, 0x1f, 0x12, 0xc0, //
        0x3f, 0x13, 0x5f, 0x20, 0x80, 0x11, 0xfe, 0x01, 0x12, 0x18, 0xe7, 0x13, //
        0x1f, 0x20, 0xc0, 0x13, 0xfa, 0x04, 0x01, 0x13, 0x1c, 0x20, 0xc3, 0x13, //
        0x5f, 0x20, 0x80, 0x13, 0xfa, 0x04, 0x01, 0x13, 0x1c, 0x20, 0xc3, 0x13, //
        0x1f, 0x20, 0xc0, 0x13, 0xfa, 0x04, 0x01, 0x11, 0x18, 0xe7, 0x13, 0x5f, //
        0x20, 0x80, 0x13, 0xfa, 0x04, 0x01, 0x10, 0xff, 0x13, 0x1f, 0x20, 0xc0, //
        0x13, 0xfa, 0x04, 0x01, 0x13, 0x08, 0x10, 0xe7, 0x13, 0x5f, 0x20, 0x80, //
        0x13, 0xfa, 0x04, 0x01, 0x13, 0x08, 0x10, 0xe7, 0x13, 0x1f, 0x20, 0xc0, //
        0x03, 0xfb, 0x04, 0x13, 0x4d, 0x10, 0xa2, 0x13, 0x5f, 0x20, 0x80, 0x03, //
        0xf8, 0x07, 0x02, 0xff, 0x03, 0x1f, 0xe0, 0x01, 0xff, 0x01, 0xff, 0x01, //
        0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, 0xff, 0x01, //
        0xff, 0x02, 0xff, 0x02, 0xff, 0x02, 0xff,
    ];

    #[test]
    fn decoded_icon_has_576_pixels() {
        let bm = decode_color_bitmap(ICON_0).unwrap();
        assert_eq!(bm.pixels.len(), 576);
    }

    #[test]
    fn icon_0_uses_only_colours_0_and_1() {
        let bm = decode_color_bitmap(ICON_0).unwrap();
        let mut seen = [false; 8];
        for p in bm.pixels {
            seen[p as usize] = true;
        }
        assert_eq!(seen, [true, true, false, false, false, false, false, false]);
    }

    #[test]
    fn icon_0_arrow_tip_lands_at_row_4_col_11() {
        let bm = decode_color_bitmap(ICON_0).unwrap();
        // The single inked pixel that forms the arrow tip.
        assert_eq!(bm.get(4, 11), 1);
        // Every other pixel on that row is background.
        for col in 0..ICON_DIM {
            if col == 11 {
                continue;
            }
            assert_eq!(bm.get(4, col), 0, "row 4 col {col} should be bg");
        }
    }

    #[test]
    fn icon_0_bottom_separator_row_is_solid_ink() {
        let bm = decode_color_bitmap(ICON_0).unwrap();
        for col in 0..ICON_DIM {
            assert_eq!(bm.get(23, col), 1, "row 23 col {col} must be ink");
        }
    }

    #[test]
    fn icon_5_uses_colours_beyond_0_and_1() {
        let bm = decode_color_bitmap(ICON_5).unwrap();
        let mut seen = [false; 8];
        for p in bm.pixels {
            seen[p as usize] = true;
        }
        assert!(seen[0] && seen[1], "icon 5 must include bg + ink");
        assert!(seen[4], "icon 5 must include the body-shade colour 4");
        assert!(seen[6], "icon 5 must include the detail-shade colour 6");
    }

    #[test]
    fn icon_5_outline_corner_is_black() {
        // Icon 5 is a "save to file" icon whose top-left outline
        // corner lands at row 2 col 3 (from the mono reference dump).
        let bm = decode_color_bitmap(ICON_5).unwrap();
        assert_eq!(bm.get(2, 3), 1);
        assert_eq!(bm.get(3, 2), 1);
        assert_eq!(bm.get(3, 4), 1);
    }

    #[test]
    fn reject_truncated_record() {
        // 0x03 needs 2 data bytes; give it 1.
        let err = decode_color_bitmap(&[0x03, 0xff]).unwrap_err();
        assert!(matches!(err, RleError::Truncated { .. }), "{err:?}");
    }

    #[test]
    fn reject_zero_opcode() {
        let err = decode_color_bitmap(&[0x00]).unwrap_err();
        assert!(matches!(err, RleError::ZeroOpcode { .. }), "{err:?}");
    }

    #[test]
    fn reject_plane_overlap() {
        // 03 FF FF → both planes claim bit 7; overlap = 0xFF.
        let mut rec = vec![0x03, 0xff, 0xff];
        // Pad to 72 opcodes so parsing reaches the bad one first.
        for _ in 0..(OPCODES_PER_RECORD - 1) {
            rec.push(0x01);
            rec.push(0xff);
        }
        let err = decode_color_bitmap(&rec).unwrap_err();
        assert!(matches!(err, RleError::PlaneOverlap { .. }), "{err:?}");
    }

    #[test]
    fn reject_plane_gap() {
        // 03 00 00 → union 0, every pixel uncovered.
        let mut rec = vec![0x03, 0x00, 0x00];
        for _ in 0..(OPCODES_PER_RECORD - 1) {
            rec.push(0x01);
            rec.push(0xff);
        }
        let err = decode_color_bitmap(&rec).unwrap_err();
        assert!(matches!(err, RleError::PlaneGap { .. }), "{err:?}");
    }

    #[test]
    fn reject_wrong_opcode_count() {
        // A short but otherwise valid record: only one opcode.
        let err = decode_color_bitmap(&[0x01, 0xff]).unwrap_err();
        assert!(matches!(err, RleError::WrongCount { got: 1 }), "{err:?}");
    }

    #[test]
    fn palette_pinned_to_canonical_ega_subset() {
        // These RGB values were read straight off a DOSBox-X screenshot of
        // the real WYSIWYG panel (see reverse-engineering notes). If the
        // table drifts, the rendered panel stops looking like 1-2-3.
        assert_eq!(LOTUS_PALETTE_RGB[0], [0xAA, 0xAA, 0xAA]); // panel bg
        assert_eq!(LOTUS_PALETTE_RGB[1], [0x00, 0x00, 0x00]); // outline
        assert_eq!(LOTUS_PALETTE_RGB[2], [0xFF, 0xFF, 0x55]); // yellow
        assert_eq!(LOTUS_PALETTE_RGB[3], [0x00, 0x00, 0xAA]); // blue
        assert_eq!(LOTUS_PALETTE_RGB[4], [0x00, 0xAA, 0xAA]); // cyan
        assert_eq!(LOTUS_PALETTE_RGB[5], [0xAA, 0x00, 0x00]); // red
        assert_eq!(LOTUS_PALETTE_RGB[6], [0xAA, 0x00, 0xAA]); // magenta
        assert_eq!(LOTUS_PALETTE_RGB[7], [0x00, 0xFF, 0x00]); // bright green
    }
}
