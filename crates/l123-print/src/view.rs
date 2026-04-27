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
}
