//! The `Engine` trait — the contract between L123's upper layers and whatever
//! compute engine is backing us.  SPEC §17.

use std::path::Path;

use thiserror::Error;
use l123_core::{Address, Range, SheetId, Value};

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
        Self { value: Value::Empty, formula: None, formatted: None }
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
}
