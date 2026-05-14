//! Shared shape for `/File Import {Json,Parquet,Sqlite}` loaders (v0.4).
//!
//! Each loader reads a typed source into a header row + value rows.
//! The UI writes the header at the cell pointer and the rows below.

use l123_core::{Format, Value};
use thiserror::Error;

/// One column header + one value per data row. The header is the
/// authoritative column count; rows are padded to its length by the
/// loader (missing keys land as `Value::Empty`).
///
/// `column_formats` carries a per-column format hint — the parquet
/// loader uses it to tag `Date32` / `Date64` / `Timestamp` columns
/// with `Format::from_kind(FormatKind::DateDmy)` (the 1-2-3 `(D1)`
/// format) so cells render as dates instead of raw serial numbers.
/// JSON and SQLite loaders fill it with `None`s today.
#[derive(Debug, Clone, PartialEq)]
pub struct LoadedRecords {
    pub header: Vec<String>,
    pub rows: Vec<Vec<Value>>,
    pub column_formats: Vec<Option<Format>>,
}

impl LoadedRecords {
    /// Construct a record set with no per-column format hints — the
    /// common case for JSON / SQLite / anything where dates aren't
    /// type-annotated in the source.
    pub fn new(header: Vec<String>, rows: Vec<Vec<Value>>) -> Self {
        let column_formats = vec![None; header.len()];
        Self {
            header,
            rows,
            column_formats,
        }
    }
}

#[derive(Debug, Error)]
pub enum LoadError {
    #[error("I/O error: {0}")]
    Io(String),
    #[error("parse error at {location}: {message}")]
    Parse { location: String, message: String },
    #[error("unsupported shape: {0}")]
    UnsupportedShape(String),
}
