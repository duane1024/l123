//! `/Data External` drivers (M12 v0.4).
//!
//! An [`ExternalSource`] is a named, refreshable backing for a worksheet
//! range. SPEC §10 / PLAN §M12 calls for `Connect`, `Use`, `Refresh`,
//! `List`, `Reset`, `Disconnect`. Slices 1-5 ship the sqlite driver
//! + the full menu surface; slice 4b adds the postgres driver.
//!
//! Connection-string scheme:
//!   `sqlite:<path>`           — local file. `path` is resolved
//!                               against CWD at every `query` call.
//!   `postgres://<libpq url>`  — postgres server. Honors libpq env
//!                               vars (`PGPASSWORD`, `PGHOST`, …) so
//!                               a sidecar-restored URL stripped of
//!                               its password can still connect.
//!   `postgresql://<libpq url>` — alias for `postgres://`.
//!
//! Drivers reuse the existing record-loader infrastructure: each
//! `query` returns a [`LoadedRecords`] so the UI walks it exactly
//! like a `/File Import` result.
//!
//! ## Credentials
//!
//! The xlsx sidecar (`l123-io::external_sources`) stores connection
//! strings with their password component stripped via
//! [`strip_credentials`]. On reconnect the user supplies the
//! password through the libpq `PGPASSWORD` env var (or by re-running
//! `/Data External Connect` with a fresh URL). The
//! `~/.l123/credentials` keyfile mentioned in PLAN §M12 is a
//! follow-up.

use std::path::{Path, PathBuf};

use rusqlite::{Connection, OpenFlags};
use thiserror::Error;

use crate::records::{LoadError, LoadedRecords};
use crate::sqlite_loader;

#[derive(Debug, Error)]
pub enum ExtSourceError {
    #[error("unsupported scheme {0:?}")]
    UnsupportedScheme(String),
    #[error("malformed connection string: {0}")]
    Malformed(String),
    #[error("connect to {0}: {1}")]
    Connect(String, String),
    #[error("query: {0}")]
    Query(#[from] LoadError),
    #[error("postgres: {0}")]
    Postgres(String),
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

/// Postgres source (M12 v0.4 slice 4b). The connection string is
/// the libpq URL form (`postgres://user:pass@host:port/db?options`).
/// Uses the `postgres` crate's sync API; queries run from a
/// `spawn_blocking` worker so the UI thread doesn't pay the
/// network-latency cost.
pub struct PostgresSource {
    conn_str: String,
}

impl PostgresSource {
    pub fn new(conn_str: String) -> Self {
        Self { conn_str }
    }
}

impl DataSource for PostgresSource {
    fn test_connection(&self) -> Result<(), ExtSourceError> {
        // `connect` opens the TCP socket, performs the libpq
        // handshake, and confirms credentials. Drop the client
        // immediately — we just want to validate the URL.
        let display = strip_credentials(&self.conn_str);
        postgres::Client::connect(&self.conn_str, postgres::NoTls)
            .map(|_| ())
            .map_err(|e| ExtSourceError::Connect(display, e.to_string()))
    }

    fn query(&self, sql: &str) -> Result<LoadedRecords, ExtSourceError> {
        let mut client = postgres::Client::connect(&self.conn_str, postgres::NoTls)
            .map_err(|e| ExtSourceError::Postgres(e.to_string()))?;
        let rows = client
            .query(sql, &[])
            .map_err(|e| ExtSourceError::Postgres(e.to_string()))?;
        let header: Vec<String> = if let Some(first) = rows.first() {
            first.columns().iter().map(|c| c.name().to_string()).collect()
        } else {
            // Empty result — query against the same SQL just to get
            // the column metadata. Postgres's `prepare` returns
            // typed column info without executing.
            let stmt = client
                .prepare(sql)
                .map_err(|e| ExtSourceError::Postgres(e.to_string()))?;
            stmt.columns().iter().map(|c| c.name().to_string()).collect()
        };
        let mut out_rows: Vec<Vec<l123_core::Value>> = Vec::with_capacity(rows.len());
        for row in &rows {
            let mut out = Vec::with_capacity(header.len());
            for (i, col) in row.columns().iter().enumerate() {
                out.push(pg_column_to_value(row, i, col.type_()));
            }
            out_rows.push(out);
        }
        Ok(LoadedRecords::new(header, out_rows))
    }
}

/// Map one postgres cell to a [`Value`]. NULLs become `Empty`;
/// recognized scalars become `Number` / `Text` per the same widening
/// rules the other v0.4 loaders use; anything we don't recognize
/// falls back to its textual form (or `[<typename>]` when even text
/// conversion fails).
fn pg_column_to_value(
    row: &postgres::Row,
    idx: usize,
    col_type: &postgres::types::Type,
) -> l123_core::Value {
    use l123_core::Value;
    use postgres::types::Type;
    macro_rules! try_typed {
        ($t:ty, $to_value:expr) => {{
            match row.try_get::<_, Option<$t>>(idx) {
                Ok(Some(v)) => return $to_value(v),
                Ok(None) => return Value::Empty,
                Err(_) => {}
            }
        }};
    }
    match *col_type {
        Type::BOOL => try_typed!(bool, |b: bool| Value::Number(if b { 1.0 } else { 0.0 })),
        Type::INT2 => try_typed!(i16, |n: i16| Value::Number(n as f64)),
        Type::INT4 => try_typed!(i32, |n: i32| Value::Number(n as f64)),
        Type::INT8 => try_typed!(i64, |n: i64| Value::Number(n as f64)),
        Type::FLOAT4 => try_typed!(f32, |n: f32| Value::Number(n as f64)),
        Type::FLOAT8 => try_typed!(f64, Value::Number),
        Type::TEXT | Type::VARCHAR | Type::BPCHAR | Type::NAME => {
            try_typed!(String, Value::Text)
        }
        _ => {}
    }
    // Fallback: try a generic string conversion (covers NUMERIC,
    // UUID, JSON, dates returned as text by some drivers, etc.).
    match row.try_get::<_, Option<String>>(idx) {
        Ok(Some(s)) => Value::Text(s),
        Ok(None) => Value::Empty,
        Err(_) => Value::Text(format!("[{}]", col_type.name())),
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
        "postgres" | "postgresql" => Ok(Box::new(PostgresSource::new(trimmed.to_string()))),
        other => Err(ExtSourceError::UnsupportedScheme(other.into())),
    }
}

/// Strip the password from a connection string before persisting it
/// to the xlsx sidecar (M12 v0.4 slice 4b). Sqlite paths pass
/// through unchanged. Postgres URLs `postgres://user:pw@host/db`
/// become `postgres://user@host/db`; URLs without a password are
/// left alone.
///
/// On reconnect the password comes from the libpq `PGPASSWORD` env
/// var or from the user re-running `/Data External Connect` with a
/// fresh URL. The `~/.l123/credentials` keyfile is a follow-up.
pub fn strip_credentials(s: &str) -> String {
    let trimmed = s.trim();
    // Only postgres / postgresql URLs carry credentials in-band.
    let prefix = if let Some(rest) = trimmed.strip_prefix("postgres://") {
        ("postgres://", rest)
    } else if let Some(rest) = trimmed.strip_prefix("postgresql://") {
        ("postgresql://", rest)
    } else {
        return trimmed.to_string();
    };
    let (scheme, rest) = prefix;
    let Some(at) = rest.find('@') else {
        return trimmed.to_string(); // no userinfo segment
    };
    let (userinfo, tail) = rest.split_at(at);
    // `userinfo` is `user` or `user:password`; strip the password.
    let user = userinfo.split(':').next().unwrap_or("");
    if user.is_empty() {
        // Bizarre but possible: `postgres://:password@host`. Drop
        // both colon and password.
        format!("{scheme}{tail}", tail = &tail[1..])
    } else {
        format!("{scheme}{user}{tail}")
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
    fn parses_postgres_scheme() {
        // The URL parses; test_connection will fail because the
        // host doesn't resolve / port isn't listening. The point is
        // the dispatch reaches PostgresSource, not UnsupportedScheme.
        let s = parse_connection_string("postgres://localhost:1/nonexistent_db_v04")
            .expect("scheme parses");
        let err = s.test_connection().expect_err("nothing listening");
        assert!(matches!(err, ExtSourceError::Connect(_, _)));
    }

    #[test]
    fn parses_postgresql_alias() {
        let s = parse_connection_string("postgresql://localhost:1/x").expect("alias parses");
        assert!(s.test_connection().is_err());
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
    fn rejects_unknown_scheme() {
        let err = parse_connection_string("mysql://x").err().expect("rejected");
        assert!(matches!(err, ExtSourceError::UnsupportedScheme(_)));
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

    #[test]
    fn strip_credentials_removes_postgres_password() {
        assert_eq!(
            strip_credentials("postgres://alice:s3cret@db.local:5432/sales"),
            "postgres://alice@db.local:5432/sales"
        );
        assert_eq!(
            strip_credentials("postgresql://alice:s3cret@db.local/sales"),
            "postgresql://alice@db.local/sales"
        );
    }

    #[test]
    fn strip_credentials_leaves_username_only_alone() {
        assert_eq!(
            strip_credentials("postgres://alice@db.local/sales"),
            "postgres://alice@db.local/sales"
        );
    }

    #[test]
    fn strip_credentials_handles_no_userinfo() {
        assert_eq!(
            strip_credentials("postgres://db.local/sales"),
            "postgres://db.local/sales"
        );
    }

    #[test]
    fn strip_credentials_handles_empty_user_with_password() {
        // Edge case: `:pw@host` — keep only the host segment.
        assert_eq!(
            strip_credentials("postgres://:s3cret@db.local/sales"),
            "postgres://db.local/sales"
        );
    }

    #[test]
    fn strip_credentials_passes_through_sqlite() {
        assert_eq!(
            strip_credentials("sqlite:tests/fixtures/inventory.db"),
            "sqlite:tests/fixtures/inventory.db"
        );
    }

    #[test]
    fn strip_credentials_passes_through_other_schemes() {
        assert_eq!(strip_credentials("mysql://user:pw@host/db"), "mysql://user:pw@host/db");
    }
}
