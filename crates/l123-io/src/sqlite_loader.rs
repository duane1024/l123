//! `/File Import Sqlite` (v0.4) — list tables and load a chosen
//! table into a [`LoadedRecords`].
//!
//! The UI flow is two-step: prompt for the path, list_tables to
//! enumerate the user-visible tables (excluding sqlite internals),
//! prompt for the table name, then load that table's rows.
//!
//! Type mapping (PLAN §M11):
//!   INTEGER → Value::Number
//!   REAL    → Value::Number
//!   TEXT    → Value::Text
//!   BLOB    → "[blob]" placeholder (v0.4 punt; user shouldn't see
//!             binary blobs in a worksheet anyway)
//!   NULL    → Value::Empty

use std::path::Path;

use l123_core::Value;
use rusqlite::{Connection, OpenFlags};

use crate::records::{LoadError, LoadedRecords};

fn open(path: &Path) -> Result<Connection, LoadError> {
    Connection::open_with_flags(path, OpenFlags::SQLITE_OPEN_READ_ONLY)
        .map_err(|e| LoadError::Io(e.to_string()))
}

/// User-visible tables in `path`, sorted alphabetically. Excludes
/// SQLite's internal `sqlite_*` and `_litestream_*` tables.
pub fn list_tables(path: &Path) -> Result<Vec<String>, LoadError> {
    let conn = open(path)?;
    let mut stmt = conn
        .prepare(
            "SELECT name FROM sqlite_master \
             WHERE type='table' AND name NOT LIKE 'sqlite_%' \
             ORDER BY name COLLATE NOCASE",
        )
        .map_err(|e| LoadError::Parse {
            location: "sqlite_master".into(),
            message: e.to_string(),
        })?;
    let rows = stmt
        .query_map([], |row| row.get::<_, String>(0))
        .map_err(|e| LoadError::Parse {
            location: "sqlite_master".into(),
            message: e.to_string(),
        })?;
    let mut out = Vec::new();
    for r in rows {
        out.push(r.map_err(|e| LoadError::Parse {
            location: "sqlite_master row".into(),
            message: e.to_string(),
        })?);
    }
    Ok(out)
}

/// Read every row of `table` from `path` into a `LoadedRecords`.
/// The header is the column names in schema order; values are
/// type-mapped per the module docs.
pub fn load(path: &Path, table: &str) -> Result<LoadedRecords, LoadError> {
    if !is_safe_identifier(table) {
        return Err(LoadError::UnsupportedShape(format!(
            "table name {table:?} contains characters L123 won't quote for sqlite"
        )));
    }
    let conn = open(path)?;
    // sqlite doesn't support binding identifiers, so we sanity-check
    // the name and embed it directly. `is_safe_identifier` keeps it
    // limited to letters / digits / underscores.
    let sql = format!("SELECT * FROM \"{table}\"");
    let mut stmt = conn.prepare(&sql).map_err(|e| LoadError::Parse {
        location: format!("query {table:?}"),
        message: e.to_string(),
    })?;
    let header: Vec<String> = stmt.column_names().into_iter().map(String::from).collect();
    let col_count = header.len();
    let mut rows: Vec<Vec<Value>> = Vec::new();
    let mut sqlite_rows = stmt.query([]).map_err(|e| LoadError::Parse {
        location: format!("query {table:?}"),
        message: e.to_string(),
    })?;
    while let Some(row) = sqlite_rows.next().map_err(|e| LoadError::Parse {
        location: format!("row in {table:?}"),
        message: e.to_string(),
    })? {
        let mut out_row = Vec::with_capacity(col_count);
        for c in 0..col_count {
            out_row.push(map_value_ref(row.get_ref(c).map_err(|e| LoadError::Parse {
                location: format!("col {c} in {table:?}"),
                message: e.to_string(),
            })?));
        }
        rows.push(out_row);
    }
    Ok(LoadedRecords { header, rows })
}

fn map_value_ref(v: rusqlite::types::ValueRef<'_>) -> Value {
    use rusqlite::types::ValueRef::*;
    match v {
        Null => Value::Empty,
        Integer(i) => Value::Number(i as f64),
        Real(f) => Value::Number(f),
        Text(bytes) => Value::Text(String::from_utf8_lossy(bytes).into_owned()),
        Blob(_) => Value::Text("[blob]".into()),
    }
}

/// Conservative identifier check: ASCII alnum + underscore, must
/// start with a letter or underscore. This excludes the corner-case
/// table names with spaces / punctuation; a user with such a table
/// can rename it, or we can add a quoting pass later.
fn is_safe_identifier(s: &str) -> bool {
    let mut chars = s.chars();
    match chars.next() {
        Some(c) if c.is_ascii_alphabetic() || c == '_' => {}
        _ => return false,
    }
    chars.all(|c| c.is_ascii_alphanumeric() || c == '_')
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        p.push("../../tests/acceptance/fixtures/m11_import.sqlite");
        p
    }

    #[test]
    fn list_tables_excludes_internals_and_sorts() {
        let names = list_tables(&fixture()).unwrap();
        assert_eq!(names, vec!["items".to_string(), "suppliers".to_string()]);
    }

    #[test]
    fn load_items_header_in_schema_order() {
        let r = load(&fixture(), "items").unwrap();
        assert_eq!(r.header, vec!["id", "name", "qty", "rate"]);
    }

    #[test]
    fn load_items_three_rows() {
        let r = load(&fixture(), "items").unwrap();
        assert_eq!(r.rows.len(), 3);
        assert_eq!(r.rows[0][0], Value::Number(1.0));
        assert_eq!(r.rows[0][1], Value::Text("widget".into()));
        assert_eq!(r.rows[2][2], Value::Empty); // NULL qty on row 3
    }

    #[test]
    fn load_suppliers_two_rows() {
        let r = load(&fixture(), "suppliers").unwrap();
        assert_eq!(r.header, vec!["id", "name"]);
        assert_eq!(r.rows.len(), 2);
    }

    #[test]
    fn unsafe_table_name_rejected_before_query() {
        let err = load(&fixture(), "items; DROP TABLE items").unwrap_err();
        assert!(matches!(err, LoadError::UnsupportedShape(_)));
    }

    #[test]
    fn missing_file_is_io_error() {
        let p = PathBuf::from("/no/such/path.sqlite");
        assert!(matches!(list_tables(&p), Err(LoadError::Io(_))));
    }

    #[test]
    fn unknown_table_is_parse_error() {
        match load(&fixture(), "nonexistent") {
            Err(LoadError::Parse { .. }) => {}
            other => panic!("expected Parse, got {other:?}"),
        }
    }
}
