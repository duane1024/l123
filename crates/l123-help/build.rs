use std::env;
use std::fs;
use std::io::Write;
use std::path::PathBuf;

fn main() {
    let manifest = PathBuf::from(env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let assets = manifest.join("assets").join("help_html");
    println!("cargo:rerun-if-changed={}", assets.display());

    let mut entries: Vec<String> = fs::read_dir(&assets)
        .unwrap_or_else(|e| panic!("read help_html dir {}: {e}", assets.display()))
        .filter_map(|d| d.ok())
        .map(|d| d.file_name().to_string_lossy().into_owned())
        .filter(|n| n.ends_with(".html"))
        .collect();
    entries.sort();

    let out_dir = PathBuf::from(env::var_os("OUT_DIR").unwrap());
    let mut f = fs::File::create(out_dir.join("help_files.rs")).unwrap();
    writeln!(f, "pub static HELP_FILES: &[(&str, &str)] = &[").unwrap();
    for name in &entries {
        let path = assets.join(name);
        // Re-fetch on each file's mtime so include_str! changes propagate.
        println!("cargo:rerun-if-changed={}", path.display());
        writeln!(
            f,
            "    ({:?}, include_str!({:?})),",
            name,
            path.display().to_string()
        )
        .unwrap();
    }
    writeln!(f, "];").unwrap();
}
