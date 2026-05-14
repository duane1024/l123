//! Shared shape for `/File Import {Json,Parquet,Sqlite}` loaders (v0.4).
//!
//! Each loader reads a typed source into a header row + value rows.
//! The UI writes the header at the cell pointer and the rows below.

use l123_core::Value;
use thiserror::Error;

/// One column header + one value per data row. The header is the
/// authoritative column count; rows are padded to its length by the
/// loader (missing keys land as `Value::Empty`).
#[derive(Debug, Clone, PartialEq)]
pub struct LoadedRecords {
    pub header: Vec<String>,
    pub rows: Vec<Vec<Value>>,
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
