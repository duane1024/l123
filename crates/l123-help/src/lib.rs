//! Decoder for the original Lotus 1-2-3 R3.4a help file (`123.HLP`).
//!
//! The file uses a byte-pair-encoding (BPE) compression scheme over a
//! 106-entry dictionary stored at the start of the file. Topic bodies
//! live as variable-length bit-packed code streams pointed at by an
//! offset index.
//!
//! See `docs/HLP_DECODE_NOTES.md` for the format reconnaissance.

pub mod dict;
