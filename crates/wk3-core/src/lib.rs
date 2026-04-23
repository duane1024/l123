//! Core types for WK3: addresses, ranges, values, modes, formats.
//!
//! This crate has no I/O dependencies and no engine dependencies.
//! Everything here is pure data types + pure functions.

pub mod address;
pub mod contents;
pub mod format;
pub mod label;
pub mod mode;
pub mod value;

pub use address::{Address, Range, SheetId};
pub use contents::{format_number_general, CellContents};
pub use format::{Format, FormatKind};
pub use label::LabelPrefix;
pub use mode::Mode;
pub use value::{ErrKind, Value};
