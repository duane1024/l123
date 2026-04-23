//! The `Engine` trait — the contract between WK3's upper layers and whatever
//! compute engine is backing us.  SPEC §17.

use std::path::Path;

use thiserror::Error;
use wk3_core::{Address, Range, SheetId, Value};

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
    /// WK3 re-formats via its own Format type for on-screen rendering).
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
    /// (`=SUM(A1:B2)`, `123`, `'text`); WK3's parse layer translates 1-2-3
    /// syntax (`@SUM(A1..B2)`) to this form before calling.
    fn set_user_input(&mut self, addr: Address, input: &str) -> Result<()>;

    /// Read a cell.  Returns Value::Empty for unset cells.
    fn get_cell(&self, addr: Address) -> Result<CellView>;

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
    fn copy_range(&mut self, _src: Range, _dst: Address) -> Result<()> {
        Err(EngineError::Unsupported("copy_range"))
    }
    fn move_range(&mut self, _src: Range, _dst: Address) -> Result<()> {
        Err(EngineError::Unsupported("move_range"))
    }
    fn define_name(&mut self, _name: &str, _range: Range) -> Result<()> {
        Err(EngineError::Unsupported("define_name"))
    }
    fn save_xlsx(&self, _path: &Path) -> Result<()> {
        Err(EngineError::Unsupported("save_xlsx"))
    }
    fn load_xlsx(&mut self, _path: &Path) -> Result<()> {
        Err(EngineError::Unsupported("load_xlsx"))
    }
}
