//! The `Engine` trait — the contract between L123's upper layers and whatever
//! compute engine is backing us.  SPEC §17.

use std::path::Path;

use l123_core::{
    Address, Alignment, Border, Comment, Fill, FontStyle, Format, Merge, Range, RgbColor, SheetId,
    SheetState, Table, TextStyle, Value,
};
use thiserror::Error;

#[derive(Debug, Error)]
pub enum EngineError {
    #[error("engine backend error: {0}")]
    Backend(String),
    #[error("unsupported operation: {0}")]
    Unsupported(&'static str),
    #[error("bad address: {0}")]
    BadAddress(String),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, EngineError>;

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum RecalcMode {
    Automatic,
    Manual,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellView {
    pub value: Value,
    pub formula: Option<String>,
    /// Formatted string as the engine would display (used as a cross-check;
    /// L123 re-formats via its own Format type for on-screen rendering).
    pub formatted: Option<String>,
}

impl CellView {
    pub fn empty() -> Self {
        Self {
            value: Value::Empty,
            formula: None,
            formatted: None,
        }
    }
}

/// The full Engine surface; the M0 smoke test exercises only a small subset,
/// but we declare the whole contract here so the adapter has to keep up.
pub trait Engine {
    /// Set a cell from a user-entered string. The input is Excel-shaped
    /// (`=SUM(A1:B2)`, `123`, `'text`); L123's parse layer translates 1-2-3
    /// syntax (`@SUM(A1..B2)`) to this form before calling.
    fn set_user_input(&mut self, addr: Address, input: &str) -> Result<()>;

    /// Read a cell.  Returns Value::Empty for unset cells.
    fn get_cell(&self, addr: Address) -> Result<CellView>;

    /// Clear the cell's contents (both value and formula). Style/format
    /// are not guaranteed to survive.
    fn clear_cell(&mut self, addr: Address) -> Result<()>;

    /// Trigger a full recalc.
    fn recalc(&mut self);

    /// Ensure a sheet exists at the given index.
    fn ensure_sheet(&mut self, id: SheetId) -> Result<()>;

    // ---- the following are stubs on the trait until their milestones land ----

    fn insert_rows(&mut self, _sheet: SheetId, _at: u32, _n: u32) -> Result<()> {
        Err(EngineError::Unsupported("insert_rows"))
    }
    fn delete_rows(&mut self, _sheet: SheetId, _at: u32, _n: u32) -> Result<()> {
        Err(EngineError::Unsupported("delete_rows"))
    }
    fn insert_cols(&mut self, _sheet: SheetId, _at: u16, _n: u16) -> Result<()> {
        Err(EngineError::Unsupported("insert_cols"))
    }
    fn delete_cols(&mut self, _sheet: SheetId, _at: u16, _n: u16) -> Result<()> {
        Err(EngineError::Unsupported("delete_cols"))
    }
    /// Insert a new empty worksheet at `at`; existing sheets at `at..`
    /// shift forward one position.
    fn insert_sheet_at(&mut self, _at: u16) -> Result<()> {
        Err(EngineError::Unsupported("insert_sheet_at"))
    }
    /// Drop the worksheet at `at`; existing sheets after it shift back
    /// one position. Refused when `at` is the only remaining sheet —
    /// every workbook must have at least one sheet.
    fn delete_sheet_at(&mut self, _at: u16) -> Result<()> {
        Err(EngineError::Unsupported("delete_sheet_at"))
    }
    /// Number of worksheets in the current workbook.
    fn sheet_count(&self) -> u16 {
        0
    }
    fn copy_range(&mut self, _src: Range, _dst: Address) -> Result<()> {
        Err(EngineError::Unsupported("copy_range"))
    }
    fn move_range(&mut self, _src: Range, _dst: Address) -> Result<()> {
        Err(EngineError::Unsupported("move_range"))
    }
    fn define_name(&mut self, _name: &str, _range: Range) -> Result<()> {
        Err(EngineError::Unsupported("define_name"))
    }
    fn delete_name(&mut self, _name: &str) -> Result<()> {
        Err(EngineError::Unsupported("delete_name"))
    }
    /// Name of the worksheet at this sheet ID, as the engine knows it.
    /// Used when constructing Excel-shape formulas that include a sheet
    /// qualifier (e.g. `Sheet1!$A$1:$A$5`).
    fn sheet_name(&self, _id: SheetId) -> Option<String> {
        None
    }
    fn save_xlsx(&self, _path: &Path) -> Result<()> {
        Err(EngineError::Unsupported("save_xlsx"))
    }
    fn load_xlsx(&mut self, _path: &Path) -> Result<()> {
        Err(EngineError::Unsupported("load_xlsx"))
    }
    /// Read a Lotus 1-2-3 R3 `.WK3` file. Read-only — the engine has no
    /// `save_wk3`; downstream callers convert to xlsx on save.
    fn load_wk3(&mut self, _path: &Path) -> Result<()> {
        Err(EngineError::Unsupported("load_wk3"))
    }
    /// Set a column's width in L123 character units (1..240). Stored so
    /// `save_xlsx` round-trips it; UI layers mirror it for on-screen
    /// geometry.
    fn set_column_width(&mut self, _sheet: SheetId, _col: u16, _width: u8) -> Result<()> {
        Err(EngineError::Unsupported("set_column_width"))
    }
    /// Read an explicit column-width override in L123 character units.
    /// Returns `None` when the column inherits the backend default.
    fn get_column_width(&self, _sheet: SheetId, _col: u16) -> Result<Option<u8>> {
        Err(EngineError::Unsupported("get_column_width"))
    }

    /// Set the text-style bits (bold / italic / underline) on a cell.
    /// Passing [`TextStyle::PLAIN`] clears the override.  Used by the
    /// xlsx round-trip: the UI pushes its per-cell `cell_text_styles`
    /// map into the engine before save.
    fn set_cell_text_style(&mut self, _addr: Address, _style: TextStyle) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_text_style"))
    }

    /// Attach a number format to a cell. Translated to an Excel
    /// `num_fmt` string on the underlying IronCalc style; the xlsx
    /// round-trip reads it back via the adapter's `used_cell_formats`.
    /// Passing [`Format::GENERAL`] / [`Format::RESET`] clears the
    /// override.
    fn set_cell_format(&mut self, _addr: Address, _format: Format) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_format"))
    }

    /// Attach a cell alignment (horizontal, vertical, wrap) to a cell.
    /// Passing [`Alignment::DEFAULT`] clears the override.  Used by the
    /// xlsx round-trip: the UI pushes its per-cell `cell_alignments`
    /// map into the engine before save.
    fn set_cell_alignment(&mut self, _addr: Address, _alignment: Alignment) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_alignment"))
    }

    /// Attach a cell background fill to a cell.  Passing [`Fill::DEFAULT`]
    /// clears the override.  Used by the xlsx round-trip: the UI pushes
    /// its per-cell `cell_fills` map into the engine before save.
    fn set_cell_fill(&mut self, _addr: Address, _fill: Fill) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_fill"))
    }

    /// Set a sheet tab color.  Passing `None` clears the override so
    /// the sheet's tab renders with the terminal default.
    fn set_sheet_color(&mut self, _sheet: SheetId, _color: Option<RgbColor>) -> Result<()> {
        Err(EngineError::Unsupported("set_sheet_color"))
    }

    /// Attach an xlsx-derived font style (color, size, strike) to a
    /// cell.  Passing [`FontStyle::DEFAULT`] clears the override.
    /// `TextStyle` (bold / italic / underline — the 1-2-3 WYSIWYG
    /// contract) is handled separately by [`Self::set_cell_text_style`];
    /// these two methods write different bits of the same underlying
    /// cell style and coexist on the same cell.
    fn set_cell_font_style(&mut self, _addr: Address, _style: FontStyle) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_font_style"))
    }

    /// Attach cell borders (left / right / top / bottom with style and
    /// color) to a cell.  Passing [`Border::NONE`] clears all four
    /// edges.  Diagonals are not represented on L123's `Border` and
    /// are wiped by this setter.
    fn set_cell_border(&mut self, _addr: Address, _border: Border) -> Result<()> {
        Err(EngineError::Unsupported("set_cell_border"))
    }

    /// Attach (or replace) a comment on `comment.addr`.  Replaces
    /// any existing comment at that address with the new author/text
    /// pair.
    fn set_comment(&mut self, _comment: Comment) -> Result<()> {
        Err(EngineError::Unsupported("set_comment"))
    }

    /// Drop the comment at `addr` if any.  No-op when none was set.
    fn delete_comment(&mut self, _addr: Address) -> Result<()> {
        Err(EngineError::Unsupported("delete_comment"))
    }

    /// Mark a rectangular range as a merged cell.  Idempotent: setting
    /// the same range twice leaves only one entry.  Single-cell merges
    /// (anchor == end) are silently dropped — they have no semantic
    /// effect and just bloat the xlsx.
    fn set_merged_range(&mut self, _merge: Merge) -> Result<()> {
        Err(EngineError::Unsupported("set_merged_range"))
    }

    /// Drop a previously-set merge.  No-op when the range isn't
    /// currently merged.  Looks up by exact anchor+end match —
    /// partial overlaps don't unmerge anything.
    fn unset_merged_range(&mut self, _merge: Merge) -> Result<()> {
        Err(EngineError::Unsupported("unset_merged_range"))
    }

    /// Set the frozen-row and frozen-column counts for a sheet.
    /// `(0, 0)` clears the freeze.  Existing freeze is replaced
    /// wholesale — the setter takes the absolute count, not a delta.
    fn set_frozen_panes(&mut self, _sheet: SheetId, _rows: u32, _cols: u16) -> Result<()> {
        Err(EngineError::Unsupported("set_frozen_panes"))
    }

    /// Set a sheet's visibility state (`Visible`, `Hidden`,
    /// `VeryHidden`).  Round-trips through xlsx via IronCalc's
    /// native `<sheet state="..."/>` writer.
    fn set_sheet_state(&mut self, _sheet: SheetId, _state: SheetState) -> Result<()> {
        Err(EngineError::Unsupported("set_sheet_state"))
    }

    /// Add (or replace) a table on `sheet`.  Replaces any existing
    /// table sharing `table.name`.  IronCalc 0.7's xlsx exporter does
    /// NOT write `xl/tables/*.xml` — see the
    /// `tables_are_dropped_on_xlsx_save_upstream_gap` engine test.
    fn set_table(&mut self, _sheet: SheetId, _table: Table) -> Result<()> {
        Err(EngineError::Unsupported("set_table"))
    }

    /// Drop a table by `name`.  No-op when no such table exists.
    fn unset_table(&mut self, _name: &str) -> Result<()> {
        Err(EngineError::Unsupported("unset_table"))
    }
}
