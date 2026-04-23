//! IronCalc implementation of the `Engine` trait.
//!
//! IronCalc uses 1-based (row, column) coordinates and Excel syntax
//! (`=SUM(A1:B2)`). This adapter bridges to WK3's 0-based addressing.
//! WK3's upper layers are responsible for translating 1-2-3 formula
//! syntax (`@SUM(A1..B2)`) to Excel shape *before* calling
//! `set_user_input` — that's the wk3-parse crate's job.

use std::path::Path;

use ironcalc::base::{expressions::utils::number_to_column, Model};
use ironcalc::export::save_to_xlsx;

use wk3_core::{Address, SheetId, Value};

use crate::engine::{CellView, Engine, EngineError, Result};

pub struct IronCalcEngine {
    model: Model<'static>,
}

impl IronCalcEngine {
    /// Create a fresh, empty workbook with a single sheet.
    pub fn new() -> Result<Self> {
        let model = Model::new_empty("workbook", "en", "UTC", "en")
            .map_err(EngineError::Backend)?;
        Ok(Self { model })
    }

    fn sheet_index(&self, id: SheetId) -> u32 {
        id.0 as u32
    }

    fn row_1based(addr: Address) -> i32 {
        (addr.row as i32) + 1
    }

    fn col_1based(addr: Address) -> i32 {
        (addr.col as i32) + 1
    }

    /// Assert a sheet at `id` exists by appending until it does. Used by
    /// `ensure_sheet`.
    fn extend_sheets_to(&mut self, id: SheetId) -> Result<()> {
        let want = self.sheet_index(id) + 1;
        let have = self.model.workbook.worksheets.len() as u32;
        for _ in have..want {
            self.model.new_sheet();
        }
        Ok(())
    }
}

impl Engine for IronCalcEngine {
    fn set_user_input(&mut self, addr: Address, input: &str) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        self.model
            .set_user_input(sheet, Self::row_1based(addr), Self::col_1based(addr), input.to_string())
            .map_err(EngineError::Backend)
    }

    fn get_cell(&self, addr: Address) -> Result<CellView> {
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let cv = self
            .model
            .get_cell_value_by_index(sheet, row, col)
            .map_err(EngineError::Backend)?;
        let value = match cv {
            ironcalc::base::cell::CellValue::None => Value::Empty,
            ironcalc::base::cell::CellValue::String(s) => Value::Text(s),
            ironcalc::base::cell::CellValue::Number(n) => Value::Number(n),
            ironcalc::base::cell::CellValue::Boolean(b) => Value::Bool(b),
        };
        // Formula retrieval is optional for M0; attempted but non-fatal.
        let formula = self.model.get_cell_formula(sheet, row, col).ok().flatten();
        Ok(CellView { value, formula, formatted: None })
    }

    fn recalc(&mut self) {
        self.model.evaluate();
    }

    fn ensure_sheet(&mut self, id: SheetId) -> Result<()> {
        self.extend_sheets_to(id)
    }

    fn save_xlsx(&self, path: &Path) -> Result<()> {
        let path_str = path.to_str().ok_or_else(|| {
            EngineError::Backend(format!("non-UTF8 path: {}", path.display()))
        })?;
        save_to_xlsx(&self.model, path_str).map_err(|e| EngineError::Backend(e.to_string()))
    }
}

/// Convenience: IronCalc column index (1-based) → letters. Only used in tests.
#[allow(dead_code)]
pub(crate) fn col_letters_1based(c: i32) -> Option<String> {
    number_to_column(c)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn create_empty() {
        let mut e = IronCalcEngine::new().unwrap();
        let cv = e.get_cell(Address::A1).unwrap();
        assert_eq!(cv.value, Value::Empty);
        // exercise recalc path on an empty book
        e.recalc();
    }

    #[test]
    fn set_numbers_and_formula_m0_smoke() {
        let mut e = IronCalcEngine::new().unwrap();
        // A1=1, A2=2, A3=+A1+A2 → expected A3=3 after recalc
        e.set_user_input(Address::new(SheetId::A, 0, 0), "1").unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "2").unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 2), "=A1+A2").unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
        assert_eq!(cv.value, Value::Number(3.0));
    }

    #[test]
    fn text_label() {
        let mut e = IronCalcEngine::new().unwrap();
        // '-prefixed means "force label"
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'hello").unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap();
        assert_eq!(cv.value, Value::Text("hello".into()));
    }

    #[test]
    fn sum_range() {
        let mut e = IronCalcEngine::new().unwrap();
        // Fill A1..A5 = 10,20,30,40,50  → C1 = SUM = 150
        for (row, n) in [(0, 10), (1, 20), (2, 30), (3, 40), (4, 50)] {
            e.set_user_input(Address::new(SheetId::A, 0, row), &n.to_string()).unwrap();
        }
        e.set_user_input(Address::new(SheetId::A, 2, 0), "=SUM(A1:A5)").unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 2, 0)).unwrap();
        assert_eq!(cv.value, Value::Number(150.0));
    }
}
