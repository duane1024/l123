//! `/Data External` source registry sidecar inside .xlsx files
//! (M12 v0.4 slice 3).
//!
//! Slice 1 / 2 registered external data sources in
//! `Workbook::external_sources` for the lifetime of the session.
//! This sidecar persists that registry across a `/File Save` →
//! `/File Retrieve` round-trip so a workbook on disk remembers
//! *which* SQL sources it was bound to.
//!
//! The sidecar lives at `l123/external_sources.json` inside the
//! xlsx zip. Vanilla Excel ignores unknown parts, so a non-L123
//! consumer just sees the workbook without the bindings — the cell
//! values are still there because `/Data External Use` wrote them
//! directly to the grid.
//!
//! ## Format
//!
//! One JSON document, version-tagged for future evolution:
//!
//! ```json
//! {
//!   "version": 1,
//!   "sources": [
//!     {
//!       "name": "sales",
//!       "connection": "sqlite:tests/fixtures/inventory.db",
//!       "last_query": "SELECT * FROM items",
//!       "last_range": {"sheet": 0, "start_col": 0, "start_row": 0, "end_col": 3, "end_row": 4},
//!       "last_refreshed_at": 1747590000
//!     }
//!   ]
//! }
//! ```
//!
//! JSON lets us carry arbitrary SQL bodies (which may contain tabs,
//! newlines, quotes) without an escaping pass; the formula-sources
//! sidecar uses TSV because well-formed Lotus formulas can't contain
//! those chars, but SQL has no such constraint.
//!
//! ## Credentials
//!
//! Connection strings are stored *as the user typed them*. The slice
//! 4 postgres driver will strip passwords from `postgres://user:pass@…`
//! URLs before persisting; sqlite strings (`sqlite:<path>`) have no
//! credentials to strip and are written verbatim.

use std::collections::HashMap;
use std::fs::File;
use std::io::{self, Read, Write};
use std::path::Path;

use serde_json::Value as JsonValue;

const SIDECAR_PATH: &str = "l123/external_sources.json";

/// Address-range pair as it lives in the sidecar — kept local to
/// `l123-io` so the on-disk format doesn't depend on `l123-core`'s
/// `Range`. The `l123-ui` caller translates to / from this snapshot
/// at the boundary.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeSnapshot {
    pub sheet: u16,
    pub start_col: u16,
    pub start_row: u32,
    pub end_col: u16,
    pub end_row: u32,
}

/// One row of the persisted source registry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExternalSourceSnapshot {
    pub name: String,
    pub connection: String,
    pub last_query: Option<String>,
    pub last_range: Option<RangeSnapshot>,
    pub last_refreshed_at: Option<u64>,
}

/// Read the external-sources sidecar from an xlsx file. Returns an
/// empty registry when the file has no sidecar (vanilla Excel xlsx,
/// or an L123 file saved before M12 slice 3). I/O errors propagate;
/// a malformed JSON body is treated as an empty registry — a
/// best-effort sidecar shouldn't fail-stop a load.
pub fn read_from_xlsx(path: &Path) -> io::Result<HashMap<String, ExternalSourceSnapshot>> {
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
    Ok(parse_json(&text).unwrap_or_default())
}

/// Write the external-sources sidecar into an xlsx file in place.
/// If the file already has a sidecar entry it is replaced; all other
/// entries are copied through unmodified (raw / compressed copy) so
/// the round-trip cost is one re-zip pass, not a full re-encode.
pub fn write_to_xlsx(
    path: &Path,
    sources: &HashMap<String, ExternalSourceSnapshot>,
) -> io::Result<()> {
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
        if !sources.is_empty() {
            let opts = zip::write::FileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            if let Err(e) = out_zip.start_file(SIDECAR_PATH, opts) {
                return Err(io::Error::other(e));
            }
            out_zip.write_all(serialize_json(sources).as_bytes())?;
        }
        if let Err(e) = out_zip.finish() {
            return Err(io::Error::other(e));
        }
    }
    std::fs::rename(&tmp_path, path)?;
    Ok(())
}

fn serialize_json(sources: &HashMap<String, ExternalSourceSnapshot>) -> String {
    let mut entries: Vec<&ExternalSourceSnapshot> = sources.values().collect();
    entries.sort_by_key(|s| s.name.to_ascii_lowercase());
    let list: Vec<JsonValue> = entries.iter().map(|s| snapshot_to_json(s)).collect();
    serde_json::json!({
        "version": 1,
        "sources": list,
    })
    .to_string()
}

fn snapshot_to_json(s: &ExternalSourceSnapshot) -> JsonValue {
    let range = s.last_range.as_ref().map(|r| {
        serde_json::json!({
            "sheet": r.sheet,
            "start_col": r.start_col,
            "start_row": r.start_row,
            "end_col": r.end_col,
            "end_row": r.end_row,
        })
    });
    serde_json::json!({
        "name": s.name,
        "connection": s.connection,
        "last_query": s.last_query,
        "last_range": range,
        "last_refreshed_at": s.last_refreshed_at,
    })
}

fn parse_json(text: &str) -> Option<HashMap<String, ExternalSourceSnapshot>> {
    let parsed: JsonValue = serde_json::from_str(text).ok()?;
    let list = parsed.get("sources")?.as_array()?;
    let mut out = HashMap::new();
    for item in list {
        let Some(name) = item.get("name").and_then(|v| v.as_str()) else {
            continue;
        };
        let Some(connection) = item.get("connection").and_then(|v| v.as_str()) else {
            continue;
        };
        let last_query = item
            .get("last_query")
            .and_then(|v| v.as_str())
            .map(String::from);
        let last_range = item.get("last_range").and_then(parse_range);
        let last_refreshed_at = item.get("last_refreshed_at").and_then(|v| v.as_u64());
        out.insert(
            name.to_ascii_lowercase(),
            ExternalSourceSnapshot {
                name: name.into(),
                connection: connection.into(),
                last_query,
                last_range,
                last_refreshed_at,
            },
        );
    }
    Some(out)
}

fn parse_range(v: &JsonValue) -> Option<RangeSnapshot> {
    Some(RangeSnapshot {
        sheet: v.get("sheet")?.as_u64()? as u16,
        start_col: v.get("start_col")?.as_u64()? as u16,
        start_row: v.get("start_row")?.as_u64()? as u32,
        end_col: v.get("end_col")?.as_u64()? as u16,
        end_row: v.get("end_row")?.as_u64()? as u32,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> HashMap<String, ExternalSourceSnapshot> {
        let mut m = HashMap::new();
        m.insert(
            "sales".into(),
            ExternalSourceSnapshot {
                name: "sales".into(),
                connection: "sqlite:tests/fixtures/inventory.db".into(),
                last_query: Some("SELECT * FROM items".into()),
                last_range: Some(RangeSnapshot {
                    sheet: 0,
                    start_col: 0,
                    start_row: 0,
                    end_col: 3,
                    end_row: 4,
                }),
                last_refreshed_at: Some(1_747_590_000),
            },
        );
        m.insert(
            "stock".into(),
            ExternalSourceSnapshot {
                name: "stock".into(),
                connection: "sqlite:tests/fixtures/inventory.db".into(),
                last_query: None,
                last_range: None,
                last_refreshed_at: None,
            },
        );
        m
    }

    #[test]
    fn serialize_then_parse_round_trips() {
        let m = sample();
        let body = serialize_json(&m);
        let parsed = parse_json(&body).expect("re-parses");
        assert_eq!(parsed, m);
    }

    #[test]
    fn parse_empty_body_is_none() {
        assert!(parse_json("").is_none());
        assert!(parse_json("not-json").is_none());
    }

    #[test]
    fn parse_skips_entries_missing_required_fields() {
        let body = r#"{"version":1,"sources":[
            {"connection":"sqlite:x"},
            {"name":"only-name"},
            {"name":"ok","connection":"sqlite:y"}
        ]}"#;
        let parsed = parse_json(body).expect("re-parses");
        assert_eq!(parsed.len(), 1);
        assert!(parsed.contains_key("ok"));
    }

    #[test]
    fn parse_optional_fields_default_to_none() {
        let body = r#"{"version":1,"sources":[
            {"name":"sales","connection":"sqlite:x"}
        ]}"#;
        let parsed = parse_json(body).expect("re-parses");
        let entry = parsed.get("sales").unwrap();
        assert!(entry.last_query.is_none());
        assert!(entry.last_range.is_none());
        assert!(entry.last_refreshed_at.is_none());
    }

    #[test]
    fn missing_sidecar_returns_empty_map() {
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

        let m = sample();
        write_to_xlsx(&path, &m).unwrap();
        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, m);

        // Pre-existing entries survive.
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
            "old".into(),
            ExternalSourceSnapshot {
                name: "old".into(),
                connection: "sqlite:first".into(),
                last_query: None,
                last_range: None,
                last_refreshed_at: None,
            },
        );
        write_to_xlsx(&path, &first).unwrap();

        let mut second = HashMap::new();
        second.insert(
            "new".into(),
            ExternalSourceSnapshot {
                name: "new".into(),
                connection: "sqlite:second".into(),
                last_query: None,
                last_range: None,
                last_refreshed_at: None,
            },
        );
        write_to_xlsx(&path, &second).unwrap();

        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, second);
    }

    #[test]
    fn writing_empty_registry_omits_sidecar() {
        // No sources → no sidecar entry in the zip. Mirrors the
        // cell_formats behavior so an L123 user with no /DE bindings
        // doesn't pay a sidecar entry in the file.
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
        write_to_xlsx(&path, &HashMap::new()).unwrap();
        let f = File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(f).unwrap();
        assert!(zip.by_name(SIDECAR_PATH).is_err());
    }
}
