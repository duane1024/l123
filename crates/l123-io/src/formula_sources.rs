//! Formula-source sidecar inside .xlsx files.
//!
//! When the user types `@CTERM(0.05,1000,500)`, `l123-parse` rewrites
//! it to `LN(1000/500)/LN(1+0.05)` so IronCalc can evaluate it. The
//! engine has no way to recover the `@CTERM` shape on reload — the
//! decomposition is irreversible. This sidecar stores the original
//! Lotus source per formula cell, embedded inside the xlsx zip, so
//! save → reload preserves the user-typed form.
//!
//! The sidecar lives at `l123/sources.tsv` inside the zip. Vanilla
//! Excel ignores unknown parts; round-tripping through Excel will
//! drop the sidecar, falling back to the cosmetic reverse translator
//! in `l123-parse`. That's the explicit limitation: full preservation
//! requires the file stays in L123.
//!
//! ## Format
//!
//! ```text
//! # l123 formula sources v1
//! 0\t0\t5\t@CTERM(0.05,1000,500)
//! 0\t0\t6\t@AVG(A1..A5)
//! ```
//!
//! Tab-separated: `sheet\tcol\trow\tsource`, all 0-based. Lotus
//! formulas can't contain tabs or newlines, so no escaping is
//! needed.
use std::collections::HashMap;
use std::fs::File;
use std::io::{self, Read, Write};
use std::path::Path;

use l123_core::{Address, SheetId};

const SIDECAR_PATH: &str = "l123/sources.tsv";
const SIDECAR_HEADER: &str = "# l123 formula sources v1\n";

/// Read the formula-sources sidecar from an xlsx file. Returns an
/// empty map when the file has no sidecar (vanilla Excel xlsx, or
/// an L123 file saved before this feature). I/O errors are
/// propagated; a malformed line is skipped silently — rationale: a
/// best-effort sidecar shouldn't fail-stop a load.
pub fn read_from_xlsx(path: &Path) -> io::Result<HashMap<Address, String>> {
    let f = File::open(path)?;
    let mut zip = match zip::ZipArchive::new(f) {
        Ok(z) => z,
        Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    };
    let mut entry = match zip.by_name(SIDECAR_PATH) {
        Ok(e) => e,
        Err(zip::result::ZipError::FileNotFound) => return Ok(HashMap::new()),
        Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    };
    let mut text = String::new();
    entry.read_to_string(&mut text)?;
    Ok(parse_tsv(&text))
}

/// Write the formula-sources sidecar into an xlsx file in place. If
/// the file already has a sidecar entry it is replaced; all other
/// entries are copied through unmodified using raw (compressed)
/// copy, so the round-trip cost is just one full re-zip rather than
/// re-encoding every part.
///
/// Implementation: read the original zip, stream every entry into a
/// tmp file, append our sidecar, then atomic-rename over the
/// original.
pub fn write_to_xlsx(path: &Path, sources: &HashMap<Address, String>) -> io::Result<()> {
    let tmp_path = path.with_extension("xlsx.l123tmp");
    {
        let in_file = File::open(path)?;
        let mut in_zip = match zip::ZipArchive::new(in_file) {
            Ok(z) => z,
            Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
        };
        let out_file = File::create(&tmp_path)?;
        let mut out_zip = zip::ZipWriter::new(out_file);
        for i in 0..in_zip.len() {
            let entry = match in_zip.by_index_raw(i) {
                Ok(e) => e,
                Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
            };
            if entry.name() == SIDECAR_PATH {
                continue;
            }
            if let Err(e) = out_zip.raw_copy_file(entry) {
                return Err(io::Error::other(e));
            }
        }
        let opts = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        if let Err(e) = out_zip.start_file(SIDECAR_PATH, opts) {
            return Err(io::Error::other(e));
        }
        out_zip.write_all(serialize_tsv(sources).as_bytes())?;
        if let Err(e) = out_zip.finish() {
            return Err(io::Error::other(e));
        }
    }
    std::fs::rename(&tmp_path, path)?;
    Ok(())
}

fn serialize_tsv(sources: &HashMap<Address, String>) -> String {
    let mut entries: Vec<(&Address, &String)> = sources.iter().collect();
    // Stable order so the sidecar diffs cleanly across saves.
    entries.sort_by_key(|(a, _)| (a.sheet.0, a.row, a.col));
    let mut out = String::with_capacity(SIDECAR_HEADER.len() + entries.len() * 32);
    out.push_str(SIDECAR_HEADER);
    for (addr, src) in entries {
        // Skip sources containing the framing chars; they shouldn't
        // appear in well-formed Lotus formulas, but a corrupt source
        // mustn't be allowed to break the sidecar.
        if src.contains('\t') || src.contains('\n') {
            continue;
        }
        out.push_str(&format!(
            "{}\t{}\t{}\t{}\n",
            addr.sheet.0, addr.col, addr.row, src
        ));
    }
    out
}

fn parse_tsv(text: &str) -> HashMap<Address, String> {
    let mut out = HashMap::new();
    for line in text.lines() {
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let mut parts = line.splitn(4, '\t');
        let sheet = parts.next().and_then(|s| s.parse::<u16>().ok());
        let col = parts.next().and_then(|s| s.parse::<u16>().ok());
        let row = parts.next().and_then(|s| s.parse::<u32>().ok());
        let src = parts.next();
        if let (Some(sheet), Some(col), Some(row), Some(src)) = (sheet, col, row, src) {
            out.insert(
                Address {
                    sheet: SheetId(sheet),
                    col,
                    row,
                },
                src.to_string(),
            );
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn serialize_then_parse_round_trips() {
        let mut sources = HashMap::new();
        sources.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 5,
            },
            "@CTERM(0.05,1000,500)".to_string(),
        );
        sources.insert(
            Address {
                sheet: SheetId(0),
                col: 1,
                row: 3,
            },
            "@SUMPRODUCT(A1..A3,B1..B3)".to_string(),
        );
        let tsv = serialize_tsv(&sources);
        assert!(tsv.starts_with("# l123 formula sources v1"));
        let parsed = parse_tsv(&tsv);
        assert_eq!(parsed, sources);
    }

    #[test]
    fn parse_skips_garbled_lines() {
        let tsv = "# l123 formula sources v1\n\
                   not a real line\n\
                   0\t0\t5\t@AVG(A1..A5)\n\
                   only\ttwo\n\
                   abc\tdef\tghi\tnope\n";
        let parsed = parse_tsv(tsv);
        assert_eq!(parsed.len(), 1);
        assert_eq!(
            parsed.get(&Address {
                sheet: SheetId(0),
                col: 0,
                row: 5
            }),
            Some(&"@AVG(A1..A5)".to_string())
        );
    }

    #[test]
    fn missing_sidecar_returns_empty_map() {
        // Build a minimal zip with one unrelated entry.
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("hello.txt", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"world").unwrap();
            zw.finish().unwrap();
        }
        let got = read_from_xlsx(&path).unwrap();
        assert!(got.is_empty());
    }

    #[test]
    fn write_then_read_round_trips_through_zip() {
        // Start with a zip that has unrelated entries — proves we
        // preserve them across a sidecar write.
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("xl/workbook.xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"<workbook/>").unwrap();
            zw.start_file("[Content_Types].xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"<types/>").unwrap();
            zw.finish().unwrap();
        }

        let mut sources = HashMap::new();
        sources.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 5,
            },
            "@CTERM(0.05,1000,500)".to_string(),
        );
        write_to_xlsx(&path, &sources).unwrap();

        // Sidecar reads back.
        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, sources);

        // Original entries still present.
        let f = File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(f).unwrap();
        let mut wb = String::new();
        zip.by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut wb)
            .unwrap();
        assert_eq!(wb, "<workbook/>");
    }

    #[test]
    fn write_replaces_prior_sidecar() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("xl/workbook.xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"<workbook/>").unwrap();
            zw.finish().unwrap();
        }
        let mut first = HashMap::new();
        first.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 0,
            },
            "@FIRST".to_string(),
        );
        write_to_xlsx(&path, &first).unwrap();

        let mut second = HashMap::new();
        second.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 0,
            },
            "@SECOND".to_string(),
        );
        write_to_xlsx(&path, &second).unwrap();

        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, second);
    }
}
