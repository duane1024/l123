//! Printing for L123: turn a range of cells into a `PageGrid` and
//! encode that grid to ASCII (`.prn`), a CUPS `lp` pipe, or PDF.
//!
//! The crate sits above `l123-core` and below `l123-ui`. It has no
//! engine or UI dependencies — `WorkbookView` is the input boundary.

pub mod encode;
pub mod grid;
pub mod render;
pub mod settings;
pub mod view;

pub use encode::ascii::to_ascii;
pub use grid::{Page, PageGrid};
pub use render::render;
pub use settings::{PrintContentMode, PrintFormatMode, PrintSettings};
pub use view::WorkbookView;
