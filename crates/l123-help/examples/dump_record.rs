use l123_help::dict::HuffmanTree;
use l123_help::huffman::decode;

fn main() {
    let path = format!(
        "{}/Documents/dosbox-cdrive/123R34/123.HLP",
        std::env::var("HOME").unwrap()
    );
    let bytes = std::fs::read(&path).expect("read .HLP");
    let tree = HuffmanTree::parse(&bytes).expect("parse tree");
    let arg: &str = &std::env::args().nth(1).unwrap_or("0x021c3..0x024b5".into());
    let parts: Vec<&str> = arg.split("..").collect();
    let s = usize::from_str_radix(parts[0].trim_start_matches("0x"), 16).unwrap();
    let e = usize::from_str_radix(parts[1].trim_start_matches("0x"), 16).unwrap();
    let decoded = decode(&bytes[s..e], &tree);

    // Print as hex+ascii rows of 24 bytes
    for (row, chunk) in decoded.chunks(24).enumerate() {
        print!("{:04}: ", row * 24);
        for b in chunk {
            print!("{:02x} ", b);
        }
        for _ in 0..(24 - chunk.len()) {
            print!("   ");
        }
        print!(" |");
        for &b in chunk {
            if (0x20..=0x7e).contains(&b) {
                print!("{}", b as char);
            } else if b == 0 {
                print!(".");
            } else {
                print!("·");
            }
        }
        println!("|");
    }
}
