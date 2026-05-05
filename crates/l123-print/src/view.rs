//! The narrow seam between the print layer and whoever owns the
//! workbook (today: `l123-ui::Workbook`). Keeps this crate off
//! `Workbook` and off the engine.

use l123_core::{Address, CellContents, Format, International, SheetId};

/// Read-only access to the cells, column widths, and per-cell formats
/// the renderer needs. Implementors are free to back this with a cache,
/// the engine directly, or a `HashMap` in tests.
pub trait WorkbookView {
    fn cell(&self, addr: Address) -> Option<&CellContents>;
    fn col_width(&self, sheet: SheetId, col: u16) -> u8;
    fn format_for_cell(&self, addr: Address) -> Format;
    fn international(&self) -> &International;
    /// Raw Excel `num_fmt` string preserved verbatim from xlsx import.
    /// `Some(_)` means render numeric values through the Excel format
    /// evaluator instead of the canonical 1-2-3 D1..D9 renderer.
    /// Default impl returns `None` for backends that don't track xlsx
    /// originals (tests, simple in-memory views).
    fn format_override_for_cell(&self, _addr: Address) -> Option<&str> {
        None
    }
}
