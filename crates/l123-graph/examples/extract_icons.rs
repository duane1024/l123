//! Extractor for Lotus 1-2-3 R3.4 WYSIWYG icon data (ICONS3.DAT).
//!
//! Usage:
//!   cargo run -p l123-graph --example extract_icons -- \
//!     <path-to-ICONS3.DAT> <output-dir>
//!
//! Dumps 105 monochrome 24×24 icons as PNGs (scaled 8×) plus a text
//! file listing the engine's help-line descriptions. We use the output
//! as *visual reference* for redrawing the v3.4 icon panel from
//! scratch in plotters — the original bitmaps stay on the user's disk.
//!
//! File-format notes (reverse-engineered):
//!
//! ```text
//! ICONS3.DAT layout
//! -----------------
//! 0x0000  u16  magic              0xDF32 (DAT-v1)
//! 0x0002  u16  version / subkind  0x118B
//! 0x0004  u16  ptr_strings_end    0x121C (end of string descriptions)
//! 0x0006  u16  ptr_section2       0x12B3
//! 0x0008  u16  ptr_mono_bitmaps   0x2386 (section 3: offset table + 24x24 mono)
//! 0x000A  u16  ptr_rle_bitmaps    0x41E2 (section 4: RLE color/hires, not parsed)
//! 0x000C  u16  ptr_eof            0x8FA2 (= file size)
//! 0x000E  u16  count              0x00A8 (hint; 105 actual mono entries)
//! 0x0010..0x0013                  more header fields
//!
//! Strings (0x0014 .. 0x121C):
//!   u16 offset table, then null-terminated ASCII description strings.
//!   Descriptions here are the single-line help text shown when the
//!   mouse hovers over an icon in 1-2-3's worksheet area.
//!
//! Section 3 (0x2386 .. 0x41E2):  monochrome bitmaps
//!   105 × u16 LE offsets relative to section start, then
//!   105 × 72-byte bitmaps (MSB-first, 3 bytes per row, 24 rows).
//!   A bit value of 0 = pixel ON (ink), 1 = pixel OFF (background).
//!
//! Section 4 (0x41E2 .. EOF):  higher-resolution RLE bitmaps
//!   Variable-length records. Not decoded here.
//! ```

use std::env;
use std::fs;
use std::path::PathBuf;

const MONO_BITMAP_BYTES: usize = 72; // 24 × 24 / 8
const BITMAP_DIM: u32 = 24;
const ICON_COUNT: usize = 105;
const PNG_SCALE: u32 = 8;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() < 3 {
        eprintln!(
            "usage: {} <ICONS3.DAT> <output-dir>",
            args.first().map(String::as_str).unwrap_or("extract_icons")
        );
        std::process::exit(1);
    }
    let input = PathBuf::from(&args[1]);
    let outdir = PathBuf::from(&args[2]);
    fs::create_dir_all(&outdir).expect("mkdir output");

    let data = fs::read(&input).expect("read ICONS3.DAT");

    let magic = u16_le(&data, 0x0000);
    let ptr_strings_end = u16_le(&data, 0x0004) as usize;
    let ptr_mono = u16_le(&data, 0x0008) as usize;
    println!("magic       = 0x{magic:04x}");
    println!("strings_end = 0x{ptr_strings_end:04x}");
    println!("mono start  = 0x{ptr_mono:04x}");

    // ---- descriptions ----
    let descriptions = extract_strings(&data, 0x00B3, ptr_strings_end);
    println!("found {} description strings", descriptions.len());
    let mut catalog = String::new();
    for (i, d) in descriptions.iter().enumerate() {
        catalog.push_str(&format!("{i:3}: {d}\n"));
    }
    fs::write(outdir.join("descriptions.txt"), catalog).expect("write descriptions");

    // ---- monochrome bitmaps ----
    let mut offsets = Vec::with_capacity(ICON_COUNT);
    for i in 0..ICON_COUNT {
        let rel = u16_le(&data, ptr_mono + i * 2) as usize;
        offsets.push(ptr_mono + rel);
    }

    for (idx, &off) in offsets.iter().enumerate() {
        let bits = &data[off..off + MONO_BITMAP_BYTES];
        let png_path = outdir.join(format!("icon_{idx:03}.png"));
        save_bitmap_as_png(bits, &png_path);
        // Also emit an ASCII-art dump alongside so the whole catalog
        // is easy to skim in a terminal.
        if idx < 32 {
            println!("---- icon {idx} ----");
            print_ascii(bits);
        }
    }

    println!(
        "Wrote {} PNGs + descriptions.txt to {}",
        ICON_COUNT,
        outdir.display()
    );
}

fn u16_le(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn extract_strings(data: &[u8], start: usize, end: usize) -> Vec<String> {
    let mut out = Vec::new();
    let mut i = start.min(data.len());
    while i < end.min(data.len()) {
        let s = i;
        while i < end && data[i] != 0 {
            i += 1;
        }
        if i > s {
            if let Ok(text) = std::str::from_utf8(&data[s..i]) {
                if !text.trim().is_empty() {
                    out.push(text.to_string());
                }
            }
        }
        i += 1;
    }
    out
}

fn pixel_on(bits: &[u8], x: u32, y: u32) -> bool {
    let byte = bits[(y * 3 + x / 8) as usize];
    let bit = 7 - (x % 8);
    // 1-2-3 mono convention: bit == 0 → pixel ON (ink).
    (byte >> bit) & 1 == 0
}

fn save_bitmap_as_png(bits: &[u8], path: &std::path::Path) {
    let w = BITMAP_DIM * PNG_SCALE;
    let h = BITMAP_DIM * PNG_SCALE;
    let mut img = image::RgbImage::new(w, h);
    for y in 0..BITMAP_DIM {
        for x in 0..BITMAP_DIM {
            let color = if pixel_on(bits, x, y) {
                image::Rgb([0x10, 0x10, 0x10])
            } else {
                image::Rgb([0xC0, 0xC0, 0xC0])
            };
            for sy in 0..PNG_SCALE {
                for sx in 0..PNG_SCALE {
                    img.put_pixel(x * PNG_SCALE + sx, y * PNG_SCALE + sy, color);
                }
            }
        }
    }
    img.save(path).expect("write PNG");
}

fn print_ascii(bits: &[u8]) {
    for y in 0..BITMAP_DIM {
        let mut row = String::new();
        for x in 0..BITMAP_DIM {
            row.push(if pixel_on(bits, x, y) { '#' } else { '.' });
        }
        println!("{row}");
    }
}
