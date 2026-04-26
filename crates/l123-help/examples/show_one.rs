use l123_help::dict::HuffmanTree;
use l123_help::huffman::decode;
use l123_help::renderer::extract_topic;

fn main() {
    let path = format!(
        "{}/Documents/dosbox-cdrive/123R34/123.HLP",
        std::env::var("HOME").unwrap()
    );
    let bytes = std::fs::read(&path).expect("read .HLP");
    let tree = HuffmanTree::parse(&bytes).expect("parse tree");
    let arg = std::env::args().nth(1).unwrap_or("0x01cbd..0x021c3".into());
    let parts: Vec<&str> = arg.split("..").collect();
    let s = usize::from_str_radix(parts[0].trim_start_matches("0x"), 16).unwrap();
    let e = usize::from_str_radix(parts[1].trim_start_matches("0x"), 16).unwrap();
    let decoded = decode(&bytes[s..e], &tree);
    if let Some(t) = extract_topic(&decoded) {
        println!("TITLE: {}", t.title);
        println!("BODY:");
        println!("{}", t.body);
    } else {
        println!("No topic extracted");
    }
}
