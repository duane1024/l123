//! Walk all records of a `123.HLP` file and print extracted (title, body) pairs.
//!
//! Run with:
//!   cargo run -p l123-help --example dump_topics -- [path]
//! Default path: $HOME/Documents/dosbox-cdrive/123R34/123.HLP

use l123_help::dict::HuffmanTree;
use l123_help::huffman::decode;
use l123_help::renderer::extract_topic;
use std::path::PathBuf;

fn main() {
    let path = std::env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| {
            PathBuf::from(format!(
                "{}/Documents/dosbox-cdrive/123R34/123.HLP",
                std::env::var("HOME").unwrap_or_default()
            ))
        });
    let bytes = std::fs::read(&path).expect("read .HLP");
    let tree = HuffmanTree::parse(&bytes).expect("parse tree");

    let ot_start = u32::from_le_bytes(bytes[0x0a..0x0e].try_into().unwrap()) as usize;
    let first_record =
        u32::from_le_bytes(bytes[ot_start..ot_start + 4].try_into().unwrap()) as usize;

    let mut offsets = Vec::new();
    let mut off = ot_start;
    while off + 4 <= first_record {
        let v = u32::from_le_bytes(bytes[off..off + 4].try_into().unwrap()) as usize;
        if v > 0 && v <= bytes.len() {
            offsets.push(v);
        }
        off += 4;
    }
    offsets.sort();
    offsets.dedup();
    if !offsets.contains(&first_record) {
        offsets.insert(0, first_record);
    }

    let mut total = 0usize;
    let mut topics = 0usize;
    for w in offsets.windows(2) {
        let (s, e) = (w[0], w[1]);
        if e <= s {
            continue;
        }
        total += 1;
        let decoded = decode(&bytes[s..e], &tree);
        if let Some(topic) = extract_topic(&decoded) {
            topics += 1;
            let body_preview: String = topic.body.lines().take(2).collect::<Vec<_>>().join(" ");
            let preview: String = body_preview.chars().take(110).collect();
            println!("0x{s:05x} {:50}  {preview}", topic.title);
        }
    }
    eprintln!("\n{topics}/{total} records produced topics");
}
