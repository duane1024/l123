//! Exploration binary for reverse-engineering 123.HLP. Prints the
//! dictionary and several speculative parses of the index/content
//! sections so we can decide on a decoder design.
//!
//! Run with:  cargo run -p l123-help --example explore -- [path] [--decode N]
//! Default path: $HOME/Documents/dosbox-cdrive/123R34/123.HLP

use l123_help::dict::Dictionary;
use std::path::PathBuf;

fn default_path() -> PathBuf {
    let home = std::env::var("HOME").unwrap_or_default();
    PathBuf::from(format!("{home}/Documents/dosbox-cdrive/123R34/123.HLP"))
}

fn read_u16_le(b: &[u8], off: usize) -> u16 {
    u16::from_le_bytes([b[off], b[off + 1]])
}
fn read_u32_le(b: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([b[off], b[off + 1], b[off + 2], b[off + 3]])
}

/// Reinterpret each i16 dictionary side as the low byte of its
/// signed-i16 storage. The i16 storage was a 16-bit DOS-era artifact;
/// values >= 128 got sign-extended into negative i16, but the
/// underlying token is just a u8.
fn dict_as_byte_pairs(d: &Dictionary) -> Vec<(u8, u8)> {
    d.pairs()
        .iter()
        .map(|p| ((p.left as i32 & 0xff) as u8, (p.right as i32 & 0xff) as u8))
        .collect()
}

fn print_dict_byte_pairs(pairs: &[(u8, u8)]) {
    println!("=== dictionary as raw byte pairs ===");
    for (i, &(a, b)) in pairs.iter().enumerate() {
        let a_disp = if a.is_ascii_graphic() || a == b' ' {
            format!("'{}'", a as char)
        } else {
            format!("{a:#04x}")
        };
        let b_disp = if b.is_ascii_graphic() || b == b' ' {
            format!("'{}'", b as char)
        } else {
            format!("{b:#04x}")
        };
        println!(
            "  code {:>3} (slot {:>3}): ({:>5}, {:>5})",
            256 + i,
            i + 1,
            a_disp,
            b_disp
        );
    }
}

fn print_chunk_records(label: &str, chunk: &[u8]) {
    println!("\n=== record parse of {label} (size {}) ===", chunk.len());
    if chunk.len() < 5 {
        println!("  too short");
        return;
    }
    let prefix0 = read_u16_le(chunk, 0);
    let prefix1 = read_u16_le(chunk, 2);
    let flag = chunk[4];
    println!(
        "  prefix: u16={} u16={} flag={:#04x}",
        prefix0, prefix1, flag
    );
    let mut p = 5usize;
    let mut rec = 0usize;
    while p < chunk.len() && rec < 60 {
        let sep = chunk[p];
        if sep != 0x2c && sep != 0x2b {
            // try other separators
            println!(
                "  rec {rec}: unexpected separator {:#04x} at {p}, stopping",
                sep
            );
            break;
        }
        if p + 4 > chunk.len() {
            break;
        }
        let len_a = chunk[p + 1] as usize;
        let len_b = chunk[p + 2] as usize;
        let pad = chunk[p + 3];
        let body_end = (p + 4 + len_a).min(chunk.len());
        let body = &chunk[p + 4..body_end];
        let hex: String = body
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect::<Vec<_>>()
            .join(" ");
        println!(
            "  rec {:>2}  pos={:<4} sep={:#04x} cmp={:<3} unc={:<3} pad={:#04x}  body({}): {}",
            rec,
            p,
            sep,
            len_a,
            len_b,
            pad,
            body.len(),
            hex
        );
        p = body_end;
        rec += 1;
    }
    if p < chunk.len() {
        let tail: String = chunk[p..]
            .iter()
            .take(16)
            .map(|b| format!("{b:02x}"))
            .collect::<Vec<_>>()
            .join(" ");
        println!(
            "  ...remaining {} bytes after rec {}: {}",
            chunk.len() - p,
            rec,
            tail
        );
    }
}

/// LSB-first bit reader. Returns Some(code) or None at EOF.
struct BitReader<'a> {
    bytes: &'a [u8],
    bit_pos: usize, // bit index from start of bytes
    bits_per_code: u8,
}
impl<'a> BitReader<'a> {
    fn new(bytes: &'a [u8], bits_per_code: u8) -> Self {
        Self {
            bytes,
            bit_pos: 0,
            bits_per_code,
        }
    }
    fn read(&mut self) -> Option<u32> {
        let need = self.bits_per_code as usize;
        if self.bit_pos + need > self.bytes.len() * 8 {
            return None;
        }
        let mut code = 0u32;
        for i in 0..need {
            let bp = self.bit_pos + i;
            let b = self.bytes[bp / 8];
            let bit = (b >> (bp % 8)) & 1;
            code |= (bit as u32) << i;
        }
        self.bit_pos += need;
        Some(code)
    }
}

/// MSB-first bit reader.
struct BitReaderMsb<'a> {
    bytes: &'a [u8],
    bit_pos: usize,
    bits_per_code: u8,
}
impl<'a> BitReaderMsb<'a> {
    fn new(bytes: &'a [u8], bits_per_code: u8) -> Self {
        Self {
            bytes,
            bit_pos: 0,
            bits_per_code,
        }
    }
    fn read(&mut self) -> Option<u32> {
        let need = self.bits_per_code as usize;
        if self.bit_pos + need > self.bytes.len() * 8 {
            return None;
        }
        let mut code = 0u32;
        for _ in 0..need {
            let bp = self.bit_pos;
            let b = self.bytes[bp / 8];
            let bit = (b >> (7 - bp % 8)) & 1;
            code = (code << 1) | bit as u32;
            self.bit_pos += 1;
        }
        Some(code)
    }
}

fn decode_record_lsb(body: &[u8], pairs: &[(u8, u8)], bits: u8) -> Vec<u8> {
    let mut out = Vec::new();
    let mut br = BitReader::new(body, bits);
    while let Some(code) = br.read() {
        if code < 256 {
            out.push(code as u8);
        } else {
            let i = code as usize - 256;
            if i < pairs.len() {
                out.push(pairs[i].0);
                out.push(pairs[i].1);
            } else {
                out.push(b'?');
            }
        }
    }
    out
}

fn decode_record_msb(body: &[u8], pairs: &[(u8, u8)], bits: u8) -> Vec<u8> {
    let mut out = Vec::new();
    let mut br = BitReaderMsb::new(body, bits);
    while let Some(code) = br.read() {
        if code < 256 {
            out.push(code as u8);
        } else {
            let i = code as usize - 256;
            if i < pairs.len() {
                out.push(pairs[i].0);
                out.push(pairs[i].1);
            } else {
                out.push(b'?');
            }
        }
    }
    out
}

fn english_score(bytes: &[u8]) -> f64 {
    let total = bytes.len() as f64;
    if total == 0.0 {
        return 0.0;
    }
    let printable = bytes
        .iter()
        .filter(|&&b| b.is_ascii_graphic() || b == b' ' || b == b'\n')
        .count() as f64;
    let lowers = bytes.iter().filter(|&&b| b.is_ascii_lowercase()).count() as f64;
    printable / total * 1.0 + lowers / total * 0.5
}

fn show_decoded(label: &str, data: &[u8]) {
    let mut s = String::new();
    for &b in data.iter().take(120) {
        if b.is_ascii_graphic() || b == b' ' {
            s.push(b as char);
        } else {
            s.push_str(&format!("\\x{:02x}", b));
        }
    }
    println!("    {label}: score={:.3}  {}", english_score(data), s);
}

fn try_decoders(label: &str, body: &[u8], pairs: &[(u8, u8)]) {
    println!(
        "  --- decoder trials for {label} ({} bytes) ---",
        body.len()
    );
    for &bits in &[8u8, 9, 10, 11, 12] {
        let lsb = decode_record_lsb(body, pairs, bits);
        show_decoded(&format!("{}b LSB", bits), &lsb);
        let msb = decode_record_msb(body, pairs, bits);
        show_decoded(&format!("{}b MSB", bits), &msb);
    }
}

fn main() {
    let path = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(default_path);
    println!("loading {}", path.display());
    let bytes = std::fs::read(&path).expect("read .HLP");
    println!("size = {} bytes", bytes.len());

    let dict = Dictionary::parse(&bytes).unwrap();
    let pairs = dict_as_byte_pairs(&dict);
    print_dict_byte_pairs(&pairs);

    // Inspect chunk 0 and a few neighbours.
    for i in 0..4 {
        let entry_off = 0x1bc + i * 4;
        let chunk_off = read_u32_le(&bytes, entry_off) as usize;
        let next_off = read_u32_le(&bytes, entry_off + 4) as usize;
        let chunk = &bytes[chunk_off..next_off];
        print_chunk_records(&format!("chunk {} @ {:#x}", i, chunk_off), chunk);

        // Try decoding each parsed record's body with various widths.
        // Currently we don't know which bytes within a record are the
        // bit-stream payload, so we test on a candidate body extracted
        // by skipping the first 2 control bytes of each record body.
        if i < 1 {
            if chunk.len() < 5 {
                continue;
            }
            let mut p = 5usize;
            let mut rec = 0usize;
            while p < chunk.len() && rec < 5 {
                if chunk[p] != 0x2c {
                    break;
                }
                let len_a = chunk[p + 1] as usize;
                let body = &chunk[p + 4..p + 4 + len_a.min(chunk.len() - p - 4)];
                if body.len() >= 4 {
                    try_decoders(&format!("rec {rec} of chunk {i}"), &body[2..], &pairs);
                }
                p += 4 + len_a;
                rec += 1;
            }
        }
    }
}
