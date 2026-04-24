//! Dumps the 17 icons in the default v3.4 WYSIWYG icon panel as
//! ASCII art plus their help descriptions, using ICONS3.DAT and
//! ICONS3.CNF from a user-supplied install.
//!
//! Usage:
//!   cargo run -p l123-graph --example dump_default_panel -- \
//!     <path-to-ICONS3.DAT> <path-to-ICONS3.CNF>

use std::env;
use std::fs;

const BITMAP_BYTES: usize = 72;
const DIM: usize = 24;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 3 {
        eprintln!("usage: {} ICONS3.DAT ICONS3.CNF", args[0]);
        std::process::exit(1);
    }
    let dat = fs::read(&args[1]).expect("read DAT");
    let cnf = fs::read(&args[2]).expect("read CNF");

    let ptr_mono = u16::from_le_bytes([dat[0x08], dat[0x09]]) as usize;
    let ptr_strings_end = u16::from_le_bytes([dat[0x04], dat[0x05]]) as usize;

    // Strings: null-terminated, starting at 0xB6 (end of 81-entry offset
    // table at 0x14..0xB5).
    let mut strings = Vec::new();
    let mut i = 0xB6;
    while i < ptr_strings_end {
        let s = i;
        while i < ptr_strings_end && dat[i] != 0 {
            i += 1;
        }
        strings.push(std::str::from_utf8(&dat[s..i]).unwrap_or("<?>").to_string());
        i += 1;
    }

    // ICONS3.CNF header: 2-byte magic (0xCF32) then u16 LE count or
    // a u8 + u8 (version?, count). Empirically, byte 3 is the icon
    // count (0x11 = 17) and IDs start at offset 4. Trailing bytes
    // after the IDs are 0xFF padding.
    let count = cnf[3] as usize;
    let ids: Vec<u8> = (0..count).map(|i| cnf[4 + i]).collect();

    println!("Default panel ({count} icons):");
    println!();
    for (slot, &id) in ids.iter().enumerate() {
        let desc = strings.get(id as usize).map(String::as_str).unwrap_or("<out of range>");
        let off = ptr_mono + u16::from_le_bytes([
            dat[ptr_mono + (id as usize) * 2],
            dat[ptr_mono + (id as usize) * 2 + 1],
        ]) as usize;
        let bits = &dat[off..off + BITMAP_BYTES];
        println!(
            "slot {slot:2}  id {id:3}  {desc}"
        );
        for y in 0..DIM {
            let mut row = String::new();
            for x in 0..DIM {
                let byte = bits[y * 3 + x / 8];
                let bit = 7 - (x % 8);
                row.push(if (byte >> bit) & 1 == 0 { '#' } else { '.' });
            }
            println!("    {row}");
        }
        println!();
    }
}
