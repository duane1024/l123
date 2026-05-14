//! `/Data External` drivers (M12 v0.4).
//!
//! An [`ExternalSource`] is a named, refreshable backing for a worksheet
//! range. SPEC §10 / PLAN §M12 calls for `Connect`, `Use`, `Refresh`,
//! `List`, `Reset`, `Disconnect`. Slice 1 ships the `Connect` /
//! `Use` read path with the sqlite driver only; postgres and refresh
//! land in later slices.
//!
//! Connection-string scheme:
//!   `sqlite:<path>` — local file. `path` is relative to CWD.
//!   `postgres://…`  — reserved for slice 4; today returns an
//!                     [`ExtSourceError::UnsupportedScheme`].
//!
//! Drivers reuse the existing record-loader infrastructure: each
//! `query` returns a [`LoadedRecords`] so the UI walks it exactly
//! like a `/File Import` result.

use std::path::{Path, PathBuf};

use rusqlite::{Connection, OpenFlags};
use thiserror::Error;

use crate::records::{LoadError, LoadedRecords};
use crate::sqlite_loader;

#[derive(Debug, Error)]
pub enum ExtSourceError {
    #[error("unsupported scheme {0:?} (v0.4 supports `sqlite:`)")]
    UnsupportedScheme(String),
    #[error("malformed connection string: {0}")]
    Malformed(String),
    #[error("connect to {0}: {1}")]
    Connect(String, String),
    #[error("query: {0}")]
    Query(#[from] LoadError),
}

/// A live external data source. Implementations test their connection
/// when registered (`/DEC`) and run a SQL `query` when bound to a
/// range (`/DEU`).
pub trait DataSource: Send {
    fn test_connection(&self) -> Result<(), ExtSourceError>;
    fn query(&self, sql: &str) -> Result<LoadedRecords, ExtSourceError>;
}

/// Sqlite-file source. The connection string is `sqlite:<path>`;
/// `path` is resolved against CWD at every `query` call so a `/FD`
/// (change directory) in the middle of a session doesn't strand
/// the source against a stale absolute path.
pub struct SqliteSource {
    path: PathBuf,
}

impl SqliteSource {
    pub fn new(path: PathBuf) -> Self {
        Self { path }
    }
    pub fn path(&self) -> &Path {
        &self.path
    }
}

impl DataSource for SqliteSource {
    fn test_connection(&self) -> Result<(), ExtSourceError> {
        Connection::open_with_flags(&self.path, OpenFlags::SQLITE_OPEN_READ_ONLY)
            .map_err(|e| ExtSourceError::Connect(self.path.display().to_string(), e.to_string()))?;
        Ok(())
    }

    fn query(&self, sql: &str) -> Result<LoadedRecords, ExtSourceError> {
        Ok(sqlite_loader::query_raw(&self.path, sql)?)
    }
}

/// Decode a connection string into a typed [`DataSource`]. The leading
/// scheme determines the driver; nothing else about the string is
/// inspected at parse time (each driver validates its own tail when
/// [`DataSource::test_connection`] runs).
pub fn parse_connection_string(s: &str) -> Result<Box<dyn DataSource>, ExtSourceError> {
    let trimmed = s.trim();
    let Some((scheme, rest)) = trimmed.split_once(':') else {
        return Err(ExtSourceError::Malformed(
            "expected `<scheme>:<details>`".into(),
        ));
    };
    match scheme {
        "sqlite" => {
            if rest.is_empty() {
                return Err(ExtSourceError::Malformed(
                    "sqlite: connection string needs a path".into(),
                ));
            }
            Ok(Box::new(SqliteSource::new(PathBuf::from(rest))))
        }
        "postgres" | "postgresql" => Err(ExtSourceError::UnsupportedScheme(scheme.into())),
        other => Err(ExtSourceError::UnsupportedScheme(other.into())),
    }
}

/// Connection-string name validation. M12 mirrors the named-range
/// rules: 1..=15 ASCII chars, must start with a letter or underscore,
/// remainder alphanumeric / underscore / dash. Bare ASCII keeps
/// xlsx custom-property keys clean across the upcoming roundtrip
/// (slice 3).
pub fn is_valid_source_name(s: &str) -> bool {
    let n = s.chars().count();
    if !(1..=15).contains(&n) {
        return false;
    }
    let mut iter = s.chars();
    let Some(first) = iter.next() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }
    iter.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-')
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
    fn parses_sqlite_scheme() {
        let s = parse_connection_string("sqlite:foo.db").unwrap();
        assert!(s.test_connection().is_err()); // foo.db doesn't exist
    }

    #[test]
    fn parses_sqlite_against_real_fixture() {
        let conn = format!("sqlite:{}", fixture().display());
        let s = parse_connection_string(&conn).unwrap();
        s.test_connection().expect("fixture connects");
    }

    #[test]
    fn rejects_postgres_in_v04_slice1() {
        let err = parse_connection_string("postgres://localhost/db")
            .err()
            .expect("postgres should be rejected");
        let ExtSourceError::UnsupportedScheme(scheme) = err else {
            panic!("expected UnsupportedScheme");
        };
        assert_eq!(scheme, "postgres");
    }

    #[test]
    fn rejects_missing_scheme() {
        let err = parse_connection_string("just-a-path")
            .err()
            .expect("rejected");
        assert!(matches!(err, ExtSourceError::Malformed(_)));
    }

    #[test]
    fn rejects_empty_sqlite_path() {
        let err = parse_connection_string("sqlite:").err().expect("rejected");
        assert!(matches!(err, ExtSourceError::Malformed(_)));
    }

    #[test]
    fn sqlite_query_runs_arbitrary_sql() {
        let s = SqliteSource::new(fixture());
        let records = s.query("SELECT id, name FROM items ORDER BY id").unwrap();
        assert_eq!(records.header, vec!["id", "name"]);
        assert_eq!(records.rows.len(), 3);
    }

    #[test]
    fn sqlite_query_bad_sql_is_error() {
        let s = SqliteSource::new(fixture());
        assert!(s.query("SELECT * FROM no_such_table").is_err());
    }

    #[test]
    fn source_name_validation() {
        assert!(is_valid_source_name("sales"));
        assert!(is_valid_source_name("_internal"));
        assert!(is_valid_source_name("q1-fy24"));
        assert!(is_valid_source_name("a"));
        assert!(is_valid_source_name("a23456789bcdefg")); // 15 chars max

        assert!(!is_valid_source_name(""));
        assert!(!is_valid_source_name("0starts_with_digit"));
        assert!(!is_valid_source_name("has space"));
        assert!(!is_valid_source_name("a23456789bcdefgh")); // 16 chars
        assert!(!is_valid_source_name("has!bang"));
    }
}
