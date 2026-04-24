//! IronCalc implementation of the `Engine` trait.
//!
//! IronCalc uses 1-based (row, column) coordinates and Excel syntax
//! (`=SUM(A1:B2)`). This adapter bridges to L123's 0-based addressing.
//! L123's upper layers are responsible for translating 1-2-3 formula
//! syntax (`@SUM(A1..B2)`) to Excel shape *before* calling
//! `set_user_input` — that's the l123-parse crate's job.

use std::path::Path;

use ironcalc::base::{expressions::utils::number_to_column, types::Cell, Model};
use ironcalc::export::save_to_xlsx;
use ironcalc::import::load_from_xlsx;

use l123_core::{address::col_to_letters, Address, Range, SheetId, Value};

use crate::engine::{CellView, Engine, EngineError, Result};

pub struct IronCalcEngine {
    model: Model<'static>,
}

impl IronCalcEngine {
    /// Create a fresh, empty workbook with a single sheet.
    pub fn new() -> Result<Self> {
        let model =
            Model::new_empty("workbook", "en", "UTC", "en").map_err(EngineError::Backend)?;
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
            .set_user_input(
                sheet,
                Self::row_1based(addr),
                Self::col_1based(addr),
                input.to_string(),
            )
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
        Ok(CellView {
            value,
            formula,
            formatted: None,
        })
    }

    fn clear_cell(&mut self, addr: Address) -> Result<()> {
        let sheet = self.sheet_index(addr.sheet);
        self.model
            .cell_clear_contents(sheet, Self::row_1based(addr), Self::col_1based(addr))
            .map_err(EngineError::Backend)
    }

    fn insert_rows(&mut self, sheet: SheetId, at: u32, n: u32) -> Result<()> {
        if n == 0 {
            return Ok(());
        }
        self.extend_sheets_to(sheet)?;
        self.model
            .insert_rows(self.sheet_index(sheet), (at as i32) + 1, n as i32)
            .map_err(EngineError::Backend)
    }

    fn delete_rows(&mut self, sheet: SheetId, at: u32, n: u32) -> Result<()> {
        if n == 0 {
            return Ok(());
        }
        self.extend_sheets_to(sheet)?;
        self.model
            .delete_rows(self.sheet_index(sheet), (at as i32) + 1, n as i32)
            .map_err(EngineError::Backend)
    }

    fn insert_cols(&mut self, sheet: SheetId, at: u16, n: u16) -> Result<()> {
        if n == 0 {
            return Ok(());
        }
        self.extend_sheets_to(sheet)?;
        self.model
            .insert_columns(self.sheet_index(sheet), (at as i32) + 1, n as i32)
            .map_err(EngineError::Backend)
    }

    fn delete_cols(&mut self, sheet: SheetId, at: u16, n: u16) -> Result<()> {
        if n == 0 {
            return Ok(());
        }
        self.extend_sheets_to(sheet)?;
        self.model
            .delete_columns(self.sheet_index(sheet), (at as i32) + 1, n as i32)
            .map_err(EngineError::Backend)
    }

    fn recalc(&mut self) {
        self.model.evaluate();
    }

    fn ensure_sheet(&mut self, id: SheetId) -> Result<()> {
        self.extend_sheets_to(id)
    }

    fn insert_sheet_at(&mut self, at: u16) -> Result<()> {
        let current = self.model.workbook.worksheets.len();
        if at as usize > current {
            return Err(EngineError::BadAddress(format!(
                "insert_sheet_at: index {at} > current count {current}"
            )));
        }
        // IronCalc requires a unique sheet name regardless of insert
        // position; pick "Sheet<N>" where N avoids any existing.
        let existing: Vec<String> = self
            .model
            .workbook
            .get_worksheet_names()
            .iter()
            .map(|s| s.to_uppercase())
            .collect();
        let mut n = current + 1;
        let name = loop {
            let candidate = format!("Sheet{n}");
            if !existing.contains(&candidate.to_uppercase()) {
                break candidate;
            }
            n += 1;
        };
        self.model
            .insert_sheet(&name, at as u32, None)
            .map_err(EngineError::Backend)
    }

    fn sheet_count(&self) -> u16 {
        self.model.workbook.worksheets.len() as u16
    }

    fn define_name(&mut self, name: &str, range: Range) -> Result<()> {
        let r = range.normalized();
        let sheet_name = self
            .sheet_name(r.start.sheet)
            .ok_or_else(|| EngineError::BadAddress(format!("unknown sheet {:?}", r.start.sheet)))?;
        let formula = format!(
            "{}!${}${}:${}${}",
            sheet_name,
            col_to_letters(r.start.col),
            r.start.row + 1,
            col_to_letters(r.end.col),
            r.end.row + 1,
        );
        self.model
            .new_defined_name(name, None, &formula)
            .map_err(EngineError::Backend)
    }

    fn delete_name(&mut self, name: &str) -> Result<()> {
        self.model
            .delete_defined_name(name, None)
            .map_err(EngineError::Backend)
    }

    fn sheet_name(&self, id: SheetId) -> Option<String> {
        let idx = self.sheet_index(id) as usize;
        self.model
            .workbook
            .worksheets
            .get(idx)
            .map(|ws| ws.get_name())
    }

    fn save_xlsx(&self, path: &Path) -> Result<()> {
        let path_str = path
            .to_str()
            .ok_or_else(|| EngineError::Backend(format!("non-UTF8 path: {}", path.display())))?;
        save_to_xlsx(&self.model, path_str).map_err(|e| EngineError::Backend(e.to_string()))
    }

    fn load_xlsx(&mut self, path: &Path) -> Result<()> {
        let path_str = path
            .to_str()
            .ok_or_else(|| EngineError::Backend(format!("non-UTF8 path: {}", path.display())))?;
        let mut model = load_from_xlsx(path_str, "en", "UTC", "en").map_err(|e| {
            tracing::error!(path = %path.display(), err = %e, "ironcalc load_from_xlsx failed");
            EngineError::Backend(e.to_string())
        })?;
        model.evaluate();
        self.model = model;
        Ok(())
    }

    fn set_column_width(&mut self, sheet: SheetId, col: u16, width: u8) -> Result<()> {
        self.extend_sheets_to(sheet)?;
        // IronCalc's `set_column_width` expects pixels; 1 Excel character
        // width = `COLUMN_WIDTH_FACTOR` (12) pixels in IronCalc's model.
        let pixels = (width as f64) * COLUMN_WIDTH_FACTOR;
        self.model
            .set_column_width(self.sheet_index(sheet), (col as i32) + 1, pixels)
            .map_err(EngineError::Backend)
    }

    fn get_column_width(&self, sheet: SheetId, col: u16) -> Result<Option<u8>> {
        let idx = self.sheet_index(sheet) as usize;
        let Some(ws) = self.model.workbook.worksheets.get(idx) else {
            return Ok(None);
        };
        let col_1b = (col as i32) + 1;
        for c in &ws.cols {
            if c.min <= col_1b && col_1b <= c.max {
                if !c.custom_width {
                    return Ok(None);
                }
                // `c.width` is already stored in character units
                // (`pixels / COLUMN_WIDTH_FACTOR`).
                let chars = c.width.round().clamp(0.0, 240.0) as u8;
                return Ok(Some(chars));
            }
        }
        Ok(None)
    }
}

/// IronCalc stores column widths internally in character units; the
/// `set_column_width` / `get_column_width` APIs take and return pixels
/// via this factor.
const COLUMN_WIDTH_FACTOR: f64 = 12.0;

impl IronCalcEngine {
    /// All sheet names in workbook order, ready to index by SheetId.0.
    /// Used by the formula translator to expand sheet-qualified refs.
    pub fn all_sheet_names(&self) -> Vec<String> {
        self.model.workbook.get_worksheet_names()
    }

    /// Enumerate every column that carries a custom width override, in
    /// L123 character units. Used after `load_xlsx` to repopulate the
    /// UI's `col_widths` cache.
    pub fn used_column_widths(&self) -> Vec<(Address, u8)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for c in &ws.cols {
                if !c.custom_width {
                    continue;
                }
                let chars = c.width.round().clamp(0.0, 240.0) as u8;
                for col_1b in c.min.max(1)..=c.max {
                    let col_0b = (col_1b - 1) as u16;
                    // Address's row field is unused here; we key by
                    // (sheet, col) in the UI — supply row 0 for shape.
                    out.push((Address::new(sheet, col_0b, 0), chars));
                }
            }
        }
        out
    }

    /// Enumerate every non-empty cell in the workbook. Used after
    /// `load_xlsx` to repopulate the UI's `cells` cache.
    pub fn used_cells(&self) -> Vec<(Address, CellView)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for (&col_1b, cell) in row_cells {
                    if col_1b < 1 {
                        continue;
                    }
                    if matches!(cell, Cell::EmptyCell { .. }) {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    if let Ok(cv) = self.get_cell(addr) {
                        out.push((addr, cv));
                    }
                }
            }
        }
        out
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
        e.set_user_input(Address::new(SheetId::A, 0, 0), "1")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "2")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 2), "=A1+A2")
            .unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
        assert_eq!(cv.value, Value::Number(3.0));
    }

    #[test]
    fn text_label() {
        let mut e = IronCalcEngine::new().unwrap();
        // '-prefixed means "force label"
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'hello")
            .unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap();
        assert_eq!(cv.value, Value::Text("hello".into()));
    }

    #[test]
    fn named_range_can_be_defined_and_used_in_formula() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "10")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "20")
            .unwrap();
        let r = Range {
            start: Address::new(SheetId::A, 0, 0),
            end: Address::new(SheetId::A, 0, 1),
        };
        e.define_name("revenue", r).unwrap();
        e.set_user_input(Address::new(SheetId::A, 1, 0), "=SUM(revenue)")
            .unwrap();
        e.recalc();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 1, 0)).unwrap().value,
            Value::Number(30.0)
        );
    }

    #[test]
    fn delete_name_removes_it() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "42")
            .unwrap();
        let r = Range::single(Address::new(SheetId::A, 0, 0));
        e.define_name("tax", r).unwrap();
        e.delete_name("tax").unwrap();
        // Re-defining the same name should now succeed.
        e.define_name("tax", r).unwrap();
    }

    #[test]
    fn sheet_name_is_sheet1_by_default() {
        let e = IronCalcEngine::new().unwrap();
        assert_eq!(e.sheet_name(SheetId::A).as_deref(), Some("Sheet1"));
    }

    #[test]
    fn insert_sheet_at_end_appends() {
        let mut e = IronCalcEngine::new().unwrap();
        assert_eq!(e.sheet_count(), 1);
        e.insert_sheet_at(1).unwrap();
        assert_eq!(e.sheet_count(), 2);
        assert!(e.sheet_name(SheetId(1)).is_some());
    }

    #[test]
    fn insert_sheet_at_start_shifts_existing() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId(0), 0, 0), "42")
            .unwrap();
        let original_name = e.sheet_name(SheetId(0)).unwrap();
        e.insert_sheet_at(0).unwrap();
        assert_eq!(e.sheet_count(), 2);
        // The original sheet now sits at index 1; the value follows it.
        assert_eq!(
            e.sheet_name(SheetId(1)).as_deref(),
            Some(original_name.as_str())
        );
        assert_eq!(
            e.get_cell(Address::new(SheetId(1), 0, 0)).unwrap().value,
            Value::Number(42.0)
        );
        assert_eq!(
            e.get_cell(Address::new(SheetId(0), 0, 0)).unwrap().value,
            Value::Empty
        );
    }

    #[test]
    fn insert_sheet_at_out_of_range_fails() {
        let mut e = IronCalcEngine::new().unwrap();
        assert!(e.insert_sheet_at(5).is_err());
    }

    #[test]
    fn insert_rows_shifts_data_down() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "42")
            .unwrap();
        e.insert_rows(SheetId::A, 0, 1).unwrap();
        e.recalc();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap().value,
            Value::Empty
        );
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 1)).unwrap().value,
            Value::Number(42.0)
        );
    }

    #[test]
    fn delete_rows_shifts_data_up() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "A")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "B")
            .unwrap();
        e.delete_rows(SheetId::A, 0, 1).unwrap();
        e.recalc();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap().value,
            Value::Text("B".into())
        );
    }

    #[test]
    fn insert_cols_shifts_data_right() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "7")
            .unwrap();
        e.insert_cols(SheetId::A, 0, 1).unwrap();
        e.recalc();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 1, 0)).unwrap().value,
            Value::Number(7.0)
        );
    }

    #[test]
    fn clear_cell_removes_value_and_unreferences_formula() {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "10")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "=A1*2")
            .unwrap();
        e.recalc();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 1)).unwrap().value,
            Value::Number(20.0)
        );
        e.clear_cell(Address::new(SheetId::A, 0, 0)).unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap();
        assert_eq!(cv.value, Value::Empty);
    }

    #[test]
    fn sum_range() {
        let mut e = IronCalcEngine::new().unwrap();
        // Fill A1..A5 = 10,20,30,40,50  → C1 = SUM = 150
        for (row, n) in [(0, 10), (1, 20), (2, 30), (3, 40), (4, 50)] {
            e.set_user_input(Address::new(SheetId::A, 0, row), &n.to_string())
                .unwrap();
        }
        e.set_user_input(Address::new(SheetId::A, 2, 0), "=SUM(A1:A5)")
            .unwrap();
        e.recalc();
        let cv = e.get_cell(Address::new(SheetId::A, 2, 0)).unwrap();
        assert_eq!(cv.value, Value::Number(150.0));
    }

    #[test]
    fn column_width_round_trips_through_xlsx() {
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_colw_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("colw.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        // Column A → 15 chars wide (L123 character units).
        e.set_column_width(SheetId::A, 0, 15).unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(e2.get_column_width(SheetId::A, 0).unwrap(), Some(15));
        // Untouched column reads as None (i.e. default).
        assert_eq!(e2.get_column_width(SheetId::A, 1).unwrap(), None);

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn save_then_load_xlsx_round_trips_number_text_and_formula() {
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!("l123_engine_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "42")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 1, 0), "'hello")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "=A1+10")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(
            e2.get_cell(Address::new(SheetId::A, 0, 0)).unwrap().value,
            Value::Number(42.0)
        );
        assert_eq!(
            e2.get_cell(Address::new(SheetId::A, 1, 0)).unwrap().value,
            Value::Text("hello".into())
        );
        let a2 = e2.get_cell(Address::new(SheetId::A, 0, 1)).unwrap();
        assert_eq!(a2.value, Value::Number(52.0));
        assert!(a2.formula.is_some(), "formula preserved across round-trip");

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }
}
