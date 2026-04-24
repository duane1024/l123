//! Extractor for Lotus 1-2-3 R3.4 WYSIWYG icon data (ICONS3.DAT).
//!
//! Usage:
//!   cargo run -p l123-graph --example extract_icons -- \
//!     <path-to-ICONS3.DAT> <output-dir>
//!
//! Dumps 105 24×24 icons as PNGs (scaled 8×) from both Section 3 (the
//! monochrome catalog) and Section 4 (the 8-colour RLE catalog), plus
//! the hover-help descriptions. We use the output as *visual
//! reference* for the L123 icon panel — the original bitmaps stay on
//! the user's disk.
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
//! 0x0008  u16  ptr_mono_bitmaps   0x2386 (section 3: offset table + 24×24 mono)
//! 0x000A  u16  ptr_rle_bitmaps    0x41E2 (section 4: plane-bitmap RLE)
//! 0x000C  u16  ptr_eof            0x8FA2 (= file size)
//! 0x000E  u16  count              0x00A8 (hint; 105 actual mono entries)
//! 0x0010..0x0013                  more header fields
//!
//! Strings (0x0014 .. 0x121C):
//!   u16 offset table, then null-terminated ASCII description strings.
//!
//! Section 3 (0x2386 .. 0x41E2):  monochrome bitmaps
//!   105 × u16 LE offsets relative to section start, then
//!   105 × 72-byte bitmaps (MSB-first, 3 bytes per row, 24 rows).
//!   A bit value of 0 = pixel ON (ink), 1 = pixel OFF (background).
//!
//! Section 4 (0x41E2 .. EOF):  8-colour RLE bitmaps
//!   105 × u16 LE offsets relative to section start, then 105
//!   variable-length records. Each record is exactly 72 opcodes.
//!   An opcode byte is a plane bitmask: bit `k` set means "plane `k`
//!   has a data byte following this opcode"; `popcount(op)` gives the
//!   number of trailing data bytes. For the current 8-pixel chunk,
//!   each data byte's 1-bits say which columns take the corresponding
//!   plane's colour. The OR of every opcode's planes is 0xFF with
//!   pairwise AND 0, so every pixel lands in exactly one palette slot
//!   0..=7. 72 opcodes × 8 pixels → the same 24×24 frame as Section 3,
//!   but with an 8-entry palette instead of 1-bit ink. See
//!   `l123-graph::icon_rle` for the Rust decoder.
//! ```

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use l123_graph::{decode_color_bitmap, icon_rle::PALETTE_INTENSITY};

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
    let ptr_rle = u16_le(&data, 0x000A) as usize;
    let ptr_eof = u16_le(&data, 0x000C) as usize;
    println!("magic       = 0x{magic:04x}");
    println!("strings_end = 0x{ptr_strings_end:04x}");
    println!("mono start  = 0x{ptr_mono:04x}");
    println!("rle  start  = 0x{ptr_rle:04x}");

    // ---- descriptions ----
    let descriptions = extract_strings(&data, 0x00B3, ptr_strings_end);
    println!("found {} description strings", descriptions.len());
    let mut catalog = String::new();
    for (i, d) in descriptions.iter().enumerate() {
        catalog.push_str(&format!("{i:3}: {d}\n"));
    }
    fs::write(outdir.join("descriptions.txt"), catalog).expect("write descriptions");

    // ---- monochrome bitmaps (Section 3) ----
    let mut mono_offsets = Vec::with_capacity(ICON_COUNT);
    for i in 0..ICON_COUNT {
        let rel = u16_le(&data, ptr_mono + i * 2) as usize;
        mono_offsets.push(ptr_mono + rel);
    }
    for (idx, &off) in mono_offsets.iter().enumerate() {
        let bits = &data[off..off + MONO_BITMAP_BYTES];
        save_bitmap_as_png(bits, &outdir.join(format!("mono_{idx:03}.png")));
        if idx < 10 {
            println!("---- mono icon {idx} ----");
            print_ascii(bits);
        }
    }

    // ---- 8-colour RLE bitmaps (Section 4) ----
    let mut rle_offsets = Vec::with_capacity(ICON_COUNT + 1);
    for i in 0..ICON_COUNT {
        rle_offsets.push(u16_le(&data, ptr_rle + i * 2) as usize);
    }
    rle_offsets.push(ptr_eof - ptr_rle);
    let mut color_fail = 0usize;
    for idx in 0..ICON_COUNT {
        let start = ptr_rle + rle_offsets[idx];
        let end = ptr_rle + rle_offsets[idx + 1];
        match decode_color_bitmap(&data[start..end]) {
            Ok(bm) => {
                save_color_bitmap_as_png(&bm.pixels, &outdir.join(format!("color_{idx:03}.png")));
            }
            Err(e) => {
                color_fail += 1;
                eprintln!("icon {idx}: RLE decode failed: {e}");
            }
        }
    }
    println!(
        "Wrote {} mono + {} colour PNGs + descriptions.txt to {} ({} RLE failures)",
        ICON_COUNT,
        ICON_COUNT - color_fail,
        outdir.display(),
        color_fail,
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

fn save_bitmap_as_png(bits: &[u8], path: &Path) {
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

fn save_color_bitmap_as_png(pixels: &[u8; 576], path: &Path) {
    // Render as a ramp from silver (palette 0) to near-black
    // (palette 1) using PALETTE_INTENSITY as the blend weight. This
    // is the same mapping the L123 panel renderer uses, so the PNG
    // dump matches what you'd see in-app.
    let bg = [0xC0u8, 0xC0, 0xC0];
    let ink = [0x10u8, 0x10, 0x10];
    let w = BITMAP_DIM * PNG_SCALE;
    let h = BITMAP_DIM * PNG_SCALE;
    let mut img = image::RgbImage::new(w, h);
    for y in 0..BITMAP_DIM {
        for x in 0..BITMAP_DIM {
            let c = pixels[(y * BITMAP_DIM + x) as usize] as usize;
            let w_blend = PALETTE_INTENSITY[c.min(7)];
            let mix = |a: u8, b: u8| ((a as f32) * w_blend + (b as f32) * (1.0 - w_blend)) as u8;
            let color = image::Rgb([mix(ink[0], bg[0]), mix(ink[1], bg[1]), mix(ink[2], bg[2])]);
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
