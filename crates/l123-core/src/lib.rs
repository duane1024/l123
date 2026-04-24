//! Core types for L123: addresses, ranges, values, modes, formats.
//!
//! This crate has no I/O dependencies and no engine dependencies.
//! Everything here is pure data types + pure functions.

pub mod address;
pub mod cell_render;
pub mod contents;
pub mod format;
pub mod label;
pub mod mode;
pub mod text_style;
pub mod value;

pub use address::{Address, Range, SheetId};
pub use cell_render::{
    center_pad, plan_row_spill, render_label, render_value_in_cell, repeat_to_width, right_pad,
    PaintedSlot, SpillSlot,
};
pub use contents::{format_number_general, CellContents};
pub use format::{format_number, Format, FormatKind};
pub use label::LabelPrefix;
pub use mode::Mode;
pub use text_style::TextStyle;
pub use value::{ErrKind, Value};
