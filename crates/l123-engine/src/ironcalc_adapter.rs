//! IronCalc implementation of the `Engine` trait.
//!
//! IronCalc uses 1-based (row, column) coordinates and Excel syntax
//! (`=SUM(A1:B2)`). This adapter bridges to L123's 0-based addressing.
//! L123's upper layers are responsible for translating 1-2-3 formula
//! syntax (`@SUM(A1..B2)`) to Excel shape *before* calling
//! `set_user_input` — that's the l123-parse crate's job.

use std::path::Path;

use ironcalc_xlsx::base::{
    expressions::utils::number_to_column,
    types::{
        Alignment as IcAlignment, Border as IcBorder, BorderItem as IcBorderItem,
        BorderStyle as IcBorderStyle, Cell, Comment as IcComment, HorizontalAlignment,
        SheetState as IcSheetState, Table as IcTable, TableColumn as IcTableColumn,
        TableStyleInfo as IcTableStyleInfo, VerticalAlignment,
    },
    Model,
};
use ironcalc_xlsx::export::save_to_xlsx;
use ironcalc_xlsx::import::load_from_xlsx;

#[cfg(feature = "wk3")]
use ironcalc_lotus::load_from_wk3_bytes;

use l123_core::{
    address::col_to_letters, Address, Alignment, Border, BorderEdge, BorderStyle, Comment, Fill,
    FillPattern, FontStyle, Format, HAlign, Merge, Range, RgbColor, SheetId, SheetState, Table,
    TableColumn, TableStyle, TextStyle, VAlign, Value,
};

use crate::engine::{CellView, Engine, EngineError, Result};
use crate::num_fmt;

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
            ironcalc_xlsx::base::cell::CellValue::None => Value::Empty,
            ironcalc_xlsx::base::cell::CellValue::String(s) => Value::Text(s),
            ironcalc_xlsx::base::cell::CellValue::Number(n) => Value::Number(n),
            ironcalc_xlsx::base::cell::CellValue::Boolean(b) => Value::Bool(b),
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
        // IronCalc moved the per-cell clear API onto the worksheet
        // type and dropped the `Model::cell_clear_contents` shim.
        // Reach through `worksheet_mut` to keep the L123 surface stable.
        let sheet = self.sheet_index(addr.sheet);
        self.model
            .workbook
            .worksheet_mut(sheet)
            .map_err(EngineError::Backend)?
            .cell_clear_contents(Self::row_1based(addr), Self::col_1based(addr))
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

    fn delete_sheet_at(&mut self, at: u16) -> Result<()> {
        let current = self.model.workbook.worksheets.len();
        if at as usize >= current {
            return Err(EngineError::BadAddress(format!(
                "delete_sheet_at: index {at} >= current count {current}"
            )));
        }
        self.model
            .delete_sheet(at as u32)
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

    #[cfg(feature = "wk3")]
    fn load_wk3(&mut self, path: &Path) -> Result<()> {
        // ironcalc_lotus's `load_from_wk3<'a>` ties `'a` across the path,
        // locale, tz, and language args, which collapses `Model<'a>` to
        // the borrowed-path lifetime — incompatible with the adapter's
        // `Model<'static>`. Read the bytes ourselves and feed
        // `load_from_wk3_bytes` a `'static` name so the result is
        // `Model<'static>`.
        let bytes = std::fs::read(path)?;
        let mut model =
            load_from_wk3_bytes(&bytes, "workbook", "en", "UTC", "en").map_err(|e| {
                tracing::error!(path = %path.display(), err = %e, "ironcalc load_from_wk3 failed");
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

    fn set_cell_text_style(&mut self, addr: Address, style: TextStyle) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        // Start from the cell's effective style so we preserve any
        // numeric format, fill, or borders the engine already knows
        // about. We only touch the three font bits that L123 models.
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        s.font.b = style.bold;
        s.font.i = style.italic;
        s.font.u = style.underline;
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_cell_format(&mut self, addr: Address, format: Format) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        s.num_fmt = num_fmt::to_num_fmt(format);
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_cell_alignment(&mut self, addr: Address, alignment: Alignment) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        s.alignment = if alignment.is_default() {
            None
        } else {
            Some(to_ic_alignment(alignment))
        };
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_cell_fill(&mut self, addr: Address, fill: Fill) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        s.fill = to_ic_fill(fill);
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_sheet_color(&mut self, sheet: SheetId, color: Option<RgbColor>) -> Result<()> {
        self.extend_sheets_to(sheet)?;
        // IronCalc's `set_sheet_color` accepts a `"#RRGGBB"` hex string
        // (with leading `#`) or an empty string to clear.
        let hex = color
            .map(|c| format!("#{}", c.to_rgb_hex()))
            .unwrap_or_default();
        self.model
            .set_sheet_color(self.sheet_index(sheet), &hex)
            .map_err(EngineError::Backend)
    }

    fn set_cell_font_style(&mut self, addr: Address, style: FontStyle) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        // Font size: IronCalc stores points as i32.  `None` means
        // "don't override" — fall back to the backend default (which
        // is whatever the style already had on it before this call).
        if let Some(sz) = style.size {
            s.font.sz = sz as i32;
        }
        // Font color: IronCalc expects a `#RRGGBB` hex string (no
        // alpha).  `None` clears the override by dropping back to
        // IronCalc's own default color string.
        s.font.color = style.color.map(|c| format!("#{}", c.to_rgb_hex()));
        s.font.strike = style.strike;
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_cell_border(&mut self, addr: Address, border: Border) -> Result<()> {
        self.extend_sheets_to(addr.sheet)?;
        let sheet = self.sheet_index(addr.sheet);
        let row = Self::row_1based(addr);
        let col = Self::col_1based(addr);
        let mut s = self
            .model
            .get_style_for_cell(sheet, row, col)
            .map_err(EngineError::Backend)?;
        // L123's Border doesn't model diagonals; clear them so a
        // previously-diagonal cell converges to the L123 view on save.
        s.border = IcBorder {
            diagonal_up: false,
            diagonal_down: false,
            left: border.left.map(to_ic_border_item),
            right: border.right.map(to_ic_border_item),
            top: border.top.map(to_ic_border_item),
            bottom: border.bottom.map(to_ic_border_item),
            diagonal: None,
        };
        self.model
            .set_cell_style(sheet, row, col, &s)
            .map_err(EngineError::Backend)
    }

    fn set_comment(&mut self, comment: Comment) -> Result<()> {
        self.extend_sheets_to(comment.addr.sheet)?;
        let sheet_idx = self.sheet_index(comment.addr.sheet);
        let cell_ref = format!(
            "{}{}",
            col_to_letters(comment.addr.col),
            comment.addr.row + 1
        );
        let ws = self
            .model
            .workbook
            .worksheet_mut(sheet_idx)
            .map_err(EngineError::Backend)?;
        // Remove any existing comment at the same address before
        // pushing the new one — set semantics, not append.
        ws.comments.retain(|c| c.cell_ref != cell_ref);
        ws.comments.push(IcComment {
            text: comment.text,
            author_name: comment.author,
            author_id: None,
            cell_ref,
        });
        Ok(())
    }

    fn delete_comment(&mut self, addr: Address) -> Result<()> {
        let sheet_idx = self.sheet_index(addr.sheet);
        let cell_ref = format!("{}{}", col_to_letters(addr.col), addr.row + 1);
        let Ok(ws) = self.model.workbook.worksheet_mut(sheet_idx) else {
            // Sheet doesn't exist → nothing to delete; treat as no-op.
            return Ok(());
        };
        ws.comments.retain(|c| c.cell_ref != cell_ref);
        Ok(())
    }

    fn set_merged_range(&mut self, merge: Merge) -> Result<()> {
        if merge.anchor == merge.end {
            // Single-cell "merge" — semantically empty, drop on the floor.
            return Ok(());
        }
        self.extend_sheets_to(merge.anchor.sheet)?;
        let s = merge_to_a1_range(merge);
        let ws = self
            .model
            .workbook
            .worksheet_mut(self.sheet_index(merge.anchor.sheet))
            .map_err(EngineError::Backend)?;
        // Idempotent: drop any existing identical entry before pushing.
        ws.merge_cells.retain(|r| r != &s);
        ws.merge_cells.push(s);
        Ok(())
    }

    fn unset_merged_range(&mut self, merge: Merge) -> Result<()> {
        let s = merge_to_a1_range(merge);
        let Ok(ws) = self
            .model
            .workbook
            .worksheet_mut(self.sheet_index(merge.anchor.sheet))
        else {
            return Ok(());
        };
        ws.merge_cells.retain(|r| r != &s);
        Ok(())
    }

    fn set_frozen_panes(&mut self, sheet: SheetId, rows: u32, cols: u16) -> Result<()> {
        self.extend_sheets_to(sheet)?;
        let idx = self.sheet_index(sheet);
        // IronCalc stores both as `i32`.  L123's `(rows, cols)` types
        // can't go negative; clamp on the way down.
        self.model
            .set_frozen_rows(idx, rows.try_into().unwrap_or(i32::MAX))
            .map_err(EngineError::Backend)?;
        self.model
            .set_frozen_columns(idx, cols.into())
            .map_err(EngineError::Backend)?;
        Ok(())
    }

    fn set_sheet_state(&mut self, sheet: SheetId, state: SheetState) -> Result<()> {
        self.extend_sheets_to(sheet)?;
        let ic = match state {
            SheetState::Visible => IcSheetState::Visible,
            SheetState::Hidden => IcSheetState::Hidden,
            SheetState::VeryHidden => IcSheetState::VeryHidden,
        };
        self.model
            .set_sheet_state(self.sheet_index(sheet), ic)
            .map_err(EngineError::Backend)
    }

    fn set_table(&mut self, sheet: SheetId, table: Table) -> Result<()> {
        self.extend_sheets_to(sheet)?;
        let sheet_name = self
            .sheet_name(sheet)
            .ok_or_else(|| EngineError::BadAddress(format!("unknown sheet {sheet:?}")))?;
        let ic = to_ic_table(&table, sheet_name);
        self.model.workbook.tables.insert(table.name.clone(), ic);
        Ok(())
    }

    fn unset_table(&mut self, name: &str) -> Result<()> {
        self.model.workbook.tables.remove(name);
        Ok(())
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

fn to_ic_alignment(a: Alignment) -> IcAlignment {
    IcAlignment {
        horizontal: match a.horizontal {
            HAlign::General => HorizontalAlignment::General,
            HAlign::Left => HorizontalAlignment::Left,
            HAlign::Center => HorizontalAlignment::Center,
            HAlign::Right => HorizontalAlignment::Right,
            HAlign::Fill => HorizontalAlignment::Fill,
            HAlign::Justify => HorizontalAlignment::Justify,
            HAlign::CenterAcross => HorizontalAlignment::CenterContinuous,
        },
        vertical: match a.vertical {
            VAlign::Top => VerticalAlignment::Top,
            VAlign::Center => VerticalAlignment::Center,
            VAlign::Bottom => VerticalAlignment::Bottom,
        },
        wrap_text: a.wrap_text,
    }
}

/// Translate L123's `Fill` into an IronCalc `Fill` ready to be
/// persisted on a cell style.  `Fill::DEFAULT` maps to
/// `pattern_type = "none"`; a solid fill goes on `bg_color` — IronCalc
/// (unlike the xlsx XML schema, where solid-fill color lives in
/// `<fgColor>`) treats `bg_color` as the solid fill's color, and its
/// save path serializes it there.  Keeping our reader and writer on
/// the same field avoids a save-then-read round trip losing the color.
fn to_ic_fill(f: Fill) -> ironcalc_xlsx::base::types::Fill {
    match f.pattern {
        FillPattern::None => ironcalc_xlsx::base::types::Fill {
            pattern_type: "none".to_string(),
            fg_color: None,
            bg_color: None,
        },
        FillPattern::Solid => ironcalc_xlsx::base::types::Fill {
            pattern_type: "solid".to_string(),
            fg_color: None,
            // IronCalc's xlsx exporter blindly writes `FF` + whatever
            // string we supply here, so we hand it a 6-char RGB (no
            // alpha) to produce a well-formed ARGB value on disk.
            bg_color: f.bg.map(|c| c.to_rgb_hex()),
        },
    }
}

/// Translate an IronCalc `Fill` back to L123's `Fill`.  Collapses every
/// non-`none` pattern to `Solid` (Excel's ~18 patterns can't render
/// in a terminal cell; v1 drops the hatch and keeps the color).
/// Reads `bg_color` first, falls back to `fg_color` so xlsx files
/// authored by other tools (which *do* follow the Excel spec and put
/// the solid color in `fg_color`) still render.  Unknown hex is
/// silently dropped so a corrupt xlsx doesn't panic the load path.
fn from_ic_fill(f: &ironcalc_xlsx::base::types::Fill) -> Fill {
    match f.pattern_type.as_str() {
        "none" | "" => Fill::DEFAULT,
        _ => {
            let color = f
                .bg_color
                .as_deref()
                .and_then(RgbColor::from_hex)
                .or_else(|| f.fg_color.as_deref().and_then(RgbColor::from_hex));
            Fill {
                pattern: FillPattern::Solid,
                bg: color,
            }
        }
    }
}

/// Translate L123's `Table` into an IronCalc `Table` ready to be
/// stored on the workbook.  Drops the dxf_id fields (round-trip
/// lossy for header / data / totals row formatting — see the
/// `l123-core::table` doc comment).
fn to_ic_table(t: &Table, sheet_name: String) -> IcTable {
    let reference = format!(
        "{}{}:{}{}",
        col_to_letters(t.range.start.col),
        t.range.start.row + 1,
        col_to_letters(t.range.end.col),
        t.range.end.row + 1,
    );
    IcTable {
        name: t.name.clone(),
        display_name: if t.display_name.is_empty() {
            t.name.clone()
        } else {
            t.display_name.clone()
        },
        sheet_name,
        reference,
        totals_row_count: if t.has_totals_row { 1 } else { 0 },
        header_row_count: if t.has_header_row { 1 } else { 0 },
        header_row_dxf_id: None,
        data_dxf_id: None,
        totals_row_dxf_id: None,
        columns: t
            .columns
            .iter()
            .map(|c| IcTableColumn {
                id: c.id,
                name: c.name.clone(),
                totals_row_label: c.totals_row_label.clone(),
                totals_row_function: c.totals_row_function.clone(),
                header_row_dxf_id: None,
                data_dxf_id: None,
                totals_row_dxf_id: None,
            })
            .collect(),
        style_info: IcTableStyleInfo {
            name: t.style.name.clone(),
            show_first_column: t.style.show_first_column,
            show_last_column: t.style.show_last_column,
            show_row_stripes: t.style.show_row_stripes,
            show_column_stripes: t.style.show_column_stripes,
        },
        has_filters: t.has_filters,
    }
}

/// Translate an IronCalc `Table` back to L123's `Table`.  Returns
/// `None` when the `reference` string doesn't parse as an A1 range
/// (cross-sheet refs, malformed input).
fn from_ic_table(ic: &IcTable, sheet: SheetId) -> Option<Table> {
    let range = parse_a1_range(&ic.reference, sheet)?;
    Some(Table {
        name: ic.name.clone(),
        display_name: ic.display_name.clone(),
        range: Range {
            start: range.anchor,
            end: range.end,
        },
        has_filters: ic.has_filters,
        has_header_row: ic.header_row_count > 0,
        has_totals_row: ic.totals_row_count > 0,
        columns: ic
            .columns
            .iter()
            .map(|c| TableColumn {
                id: c.id,
                name: c.name.clone(),
                totals_row_function: c.totals_row_function.clone(),
                totals_row_label: c.totals_row_label.clone(),
            })
            .collect(),
        style: TableStyle {
            name: ic.style_info.name.clone(),
            show_first_column: ic.style_info.show_first_column,
            show_last_column: ic.style_info.show_last_column,
            show_row_stripes: ic.style_info.show_row_stripes,
            show_column_stripes: ic.style_info.show_column_stripes,
        },
    })
}

/// Format a merge as the IronCalc `"A1:B2"` cell-range string.
fn merge_to_a1_range(m: Merge) -> String {
    format!(
        "{}{}:{}{}",
        col_to_letters(m.anchor.col),
        m.anchor.row + 1,
        col_to_letters(m.end.col),
        m.end.row + 1,
    )
}

/// Parse an A1-style range (`"A1:B2"`) into an L123 [`Merge`] pinned
/// to the given sheet.  Returns `None` for anything that doesn't fit
/// that exact shape (sheet-qualified refs, single cells, malformed).
fn parse_a1_range(s: &str, sheet: SheetId) -> Option<Merge> {
    let (a, b) = s.split_once(':')?;
    let anchor = parse_a1_cell_ref(a, sheet)?;
    let end = parse_a1_cell_ref(b, sheet)?;
    Merge::from_corners(anchor, end)
}

/// Parse an A1-style single-cell reference (e.g. `"B5"`) into an
/// L123 [`Address`] pinned to the given sheet.  Returns `None` for
/// anything more complex (ranges, sheet-qualified refs, malformed
/// hex).  Used to bridge IronCalc's string-typed `Comment.cell_ref`
/// back into typed addresses.
fn parse_a1_cell_ref(s: &str, sheet: SheetId) -> Option<Address> {
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == 0 || i == bytes.len() {
        return None;
    }
    let col_part = &s[..i];
    let row_part = &s[i..];
    if !row_part.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row_1b: u32 = row_part.parse().ok()?;
    if row_1b == 0 {
        return None;
    }
    let mut col_1b: u32 = 0;
    for ch in col_part.chars() {
        let v = ch.to_ascii_uppercase() as u32;
        if !(b'A' as u32..=b'Z' as u32).contains(&v) {
            return None;
        }
        col_1b = col_1b * 26 + (v - b'A' as u32 + 1);
    }
    let col_0b: u16 = (col_1b.checked_sub(1)?).try_into().ok()?;
    let row_0b: u32 = row_1b - 1;
    Some(Address::new(sheet, col_0b, row_0b))
}

/// Translate L123's `BorderEdge` into an IronCalc `BorderItem`.
/// Style collapse is deliberate: L123 only models 6 styles but
/// IronCalc has 9.  Writing out `Dashed` as `MediumDashed` is the
/// best lossless-seeming round-trip we can manage (Excel renderers
/// treat the medium-dash variants as "the basic dashed line", so
/// reopened files still look like L123 intended).
fn to_ic_border_item(edge: BorderEdge) -> IcBorderItem {
    let style = match edge.style {
        BorderStyle::Thin => IcBorderStyle::Thin,
        BorderStyle::Medium => IcBorderStyle::Medium,
        BorderStyle::Thick => IcBorderStyle::Thick,
        BorderStyle::Double => IcBorderStyle::Double,
        BorderStyle::Dashed => IcBorderStyle::MediumDashed,
        BorderStyle::Dotted => IcBorderStyle::Dotted,
    };
    IcBorderItem {
        style,
        color: edge.color.map(|c| format!("#{}", c.to_rgb_hex())),
    }
}

/// Translate an IronCalc `BorderItem` back to L123's `BorderEdge`.
/// Every dash-dot-variant collapses to `Dashed`; the rest map
/// one-to-one.  Unknown hex in `color` is dropped to `None` so a
/// corrupt xlsx doesn't panic the load path.
fn from_ic_border_item(item: &IcBorderItem) -> BorderEdge {
    let style = match item.style {
        IcBorderStyle::Thin => BorderStyle::Thin,
        IcBorderStyle::Medium => BorderStyle::Medium,
        IcBorderStyle::Thick => BorderStyle::Thick,
        IcBorderStyle::Double => BorderStyle::Double,
        IcBorderStyle::Dotted => BorderStyle::Dotted,
        IcBorderStyle::MediumDashed
        | IcBorderStyle::MediumDashDot
        | IcBorderStyle::MediumDashDotDot
        | IcBorderStyle::SlantDashDot => BorderStyle::Dashed,
    };
    BorderEdge {
        style,
        color: item.color.as_deref().and_then(RgbColor::from_hex),
    }
}

fn from_ic_border(b: &IcBorder) -> Border {
    Border {
        left: b.left.as_ref().map(from_ic_border_item),
        right: b.right.as_ref().map(from_ic_border_item),
        top: b.top.as_ref().map(from_ic_border_item),
        bottom: b.bottom.as_ref().map(from_ic_border_item),
    }
}

fn from_ic_alignment(a: &IcAlignment) -> Alignment {
    Alignment {
        horizontal: match a.horizontal {
            HorizontalAlignment::General => HAlign::General,
            HorizontalAlignment::Left => HAlign::Left,
            HorizontalAlignment::Center => HAlign::Center,
            HorizontalAlignment::Right => HAlign::Right,
            HorizontalAlignment::Fill => HAlign::Fill,
            HorizontalAlignment::Justify => HAlign::Justify,
            HorizontalAlignment::CenterContinuous => HAlign::CenterAcross,
            // Asian-language distributed collapses to Justify — same as
            // the HAlign::from_xlsx_str fallback, kept consistent so
            // there is one place to change this later.
            HorizontalAlignment::Distributed => HAlign::Justify,
        },
        vertical: match a.vertical {
            VerticalAlignment::Top => VAlign::Top,
            VerticalAlignment::Center => VAlign::Center,
            VerticalAlignment::Bottom => VAlign::Bottom,
            // terminal cells are single-line; collapse the less-common
            // verticals to Bottom.
            VerticalAlignment::Distributed | VerticalAlignment::Justify => VAlign::Bottom,
        },
        wrap_text: a.wrap_text,
    }
}

impl IronCalcEngine {
    /// All sheet names in workbook order, ready to index by SheetId.0.
    /// Used by the formula translator to expand sheet-qualified refs.
    pub fn all_sheet_names(&self) -> Vec<String> {
        self.model.workbook.get_worksheet_names()
    }

    /// Enumerate every Excel-format table across the workbook.
    /// Tables whose `sheet_name` doesn't match a worksheet, or whose
    /// `reference` doesn't parse as `A1:B2`, are silently dropped —
    /// keeps a corrupt xlsx from panicking the load path.  Order is
    /// not guaranteed (callers that want stable display order should
    /// sort by `(sheet, table.name)`).
    pub fn used_tables(&self) -> Vec<(SheetId, Table)> {
        // Pre-build a lookup from sheet name → SheetId so we can map
        // each table's `sheet_name` back to L123's typed identifier.
        let names = self.all_sheet_names();
        let mut name_to_id: std::collections::HashMap<String, SheetId> =
            std::collections::HashMap::with_capacity(names.len());
        for (i, n) in names.into_iter().enumerate() {
            name_to_id.insert(n, SheetId(i as u16));
        }
        let mut out = Vec::new();
        for ic in self.model.workbook.tables.values() {
            let Some(&sid) = name_to_id.get(&ic.sheet_name) else {
                continue;
            };
            let Some(t) = from_ic_table(ic, sid) else {
                continue;
            };
            out.push((sid, t));
        }
        out
    }

    /// Read a sheet's visibility state (`Visible`, `Hidden`, or
    /// `VeryHidden`).  Out-of-range sheet indices return `Visible` so
    /// callers don't have to special-case the "no such sheet" path.
    pub fn sheet_state(&self, sheet: SheetId) -> SheetState {
        let idx = self.sheet_index(sheet) as usize;
        let Some(ws) = self.model.workbook.worksheets.get(idx) else {
            return SheetState::Visible;
        };
        match ws.state {
            IcSheetState::Visible => SheetState::Visible,
            IcSheetState::Hidden => SheetState::Hidden,
            IcSheetState::VeryHidden => SheetState::VeryHidden,
        }
    }

    /// Read frozen-pane counts for a sheet as `(rows, cols)`.
    /// Returns `(0, 0)` for unfrozen sheets and for out-of-range
    /// sheet indices (so callers don't have to special-case the
    /// "no such sheet" path during render).
    pub fn frozen_panes(&self, sheet: SheetId) -> (u32, u16) {
        let idx = self.sheet_index(sheet) as usize;
        let Some(ws) = self.model.workbook.worksheets.get(idx) else {
            return (0, 0);
        };
        let rows = ws.frozen_rows.max(0) as u32;
        let cols = ws.frozen_columns.clamp(0, u16::MAX as i32) as u16;
        (rows, cols)
    }

    /// Read the sheet's tab color, if any.  Returns `None` when the
    /// sheet is out of range or carries no color override.  The hex
    /// string IronCalc stores (`"#RRGGBB"`) is parsed via
    /// [`RgbColor::from_hex`]; malformed hex degrades to `None`
    /// rather than panicking on load.
    pub fn sheet_color(&self, sheet: SheetId) -> Option<RgbColor> {
        let idx = self.sheet_index(sheet) as usize;
        let ws = self.model.workbook.worksheets.get(idx)?;
        let hex = ws.color.as_deref()?;
        RgbColor::from_hex(hex)
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

    /// Enumerate every cell that carries a non-plain text style
    /// (bold / italic / underline). Used after `load_xlsx` to
    /// repopulate the UI's `cell_text_styles` cache.  Cells whose
    /// style the engine considers "plain" are skipped so we don't
    /// flood the UI map with empty entries.
    pub fn used_cell_text_styles(&self) -> Vec<(Address, TextStyle)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    let ts = TextStyle {
                        bold: style.font.b,
                        italic: style.font.i,
                        underline: style.font.u,
                    };
                    if !ts.is_empty() {
                        out.push((addr, ts));
                    }
                }
            }
        }
        out
    }

    /// Enumerate every cell whose style carries a non-General number
    /// format. Used after `load_xlsx` to repopulate the UI's
    /// `cell_formats` map so xlsx files authored in Excel (and in l123
    /// itself across /FS → /FR) keep their Currency / Percent / etc.
    /// tags.
    pub fn used_cell_formats(&self) -> Vec<(Address, Format)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    if let Some(fmt) = num_fmt::parse(&style.num_fmt) {
                        out.push((addr, fmt));
                    }
                }
            }
        }
        out
    }

    /// Enumerate every cell whose style carries a non-default
    /// alignment. Used after `load_xlsx` to repopulate the UI's
    /// `cell_alignments` map so xlsx files authored in Excel survive
    /// /FS → /FR with their horizontal/vertical/wrap settings intact.
    pub fn used_cell_alignments(&self) -> Vec<(Address, Alignment)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    if let Some(ic) = &style.alignment {
                        let a = from_ic_alignment(ic);
                        if !a.is_default() {
                            out.push((addr, a));
                        }
                    }
                }
            }
        }
        out
    }

    /// Enumerate every merged range across every sheet, paired with
    /// the sheet it lives on.  Merges whose `cell_ref` doesn't parse
    /// as `A1:B2` (cross-sheet refs, malformed strings) are silently
    /// dropped — keeps a corrupt xlsx from panicking the load path.
    pub fn used_merged_cells(&self) -> Vec<(SheetId, Merge)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for s in &ws.merge_cells {
                if let Some(m) = parse_a1_range(s, sheet) {
                    out.push((sheet, m));
                }
            }
        }
        out
    }

    /// Enumerate every comment across every sheet.  Returns them in
    /// no particular order (callers that want stable display order
    /// should sort by `(sheet, row, col)`).  Comments whose
    /// `cell_ref` doesn't parse as a single-cell A1 reference are
    /// silently dropped — keeps a corrupt xlsx from panicking the
    /// load path.
    pub fn used_comments(&self) -> Vec<Comment> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for c in &ws.comments {
                let Some(addr) = parse_a1_cell_ref(&c.cell_ref, sheet) else {
                    continue;
                };
                out.push(Comment {
                    addr,
                    author: c.author_name.clone(),
                    text: c.text.clone(),
                });
            }
        }
        out
    }

    /// Enumerate every cell that carries any border override on any
    /// of its four edges.  Diagonals are silently skipped — they're
    /// not in L123's `Border` model.  Cells whose IronCalc border
    /// struct has every edge `None` and no diagonals are omitted
    /// (keeping the map small even on large styled sheets).
    pub fn used_cell_borders(&self) -> Vec<(Address, Border)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    let b = from_ic_border(&style.border);
                    if !b.is_default() {
                        out.push((addr, b));
                    }
                }
            }
        }
        out
    }

    /// Enumerate every cell whose font carries an override L123
    /// cares about (color, non-default size, or strikethrough).
    /// Bold / italic / underline are *not* reported here — they live
    /// on [`Self::used_cell_text_styles`] since they're the 1-2-3
    /// WYSIWYG triple.  `size` is reported whenever the stored point
    /// value differs from IronCalc's fresh-Model default (13pt) so
    /// we don't flood the UI map with every cell IronCalc ever
    /// touched.
    pub fn used_cell_font_styles(&self) -> Vec<(Address, FontStyle)> {
        // IronCalc 0.7 creates an empty Model with Font::default() ==
        // 13pt Calibri on every styled cell.  Treat `13` (or any
        // cell where font.sz matches the workbook-wide default) as
        // "no override" so the map stays small.
        const DEFAULT_SIZE: i32 = 13;
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    let color = style
                        .font
                        .color
                        .as_deref()
                        .and_then(RgbColor::from_hex)
                        // IronCalc fills in "#000000" as the default;
                        // treat pure black as "no override" so we don't
                        // tint cells the user never touched.  A user who
                        // genuinely wants black text can still set it,
                        // and the xlsx will carry it, just won't round
                        // back into the `used_cell_font_styles` view.
                        .filter(|c| *c != RgbColor::BLACK);
                    let size = if style.font.sz != DEFAULT_SIZE {
                        Some((style.font.sz as u8).max(1))
                    } else {
                        None
                    };
                    let fs = FontStyle {
                        color,
                        size,
                        strike: style.font.strike,
                    };
                    if !fs.is_default() {
                        out.push((addr, fs));
                    }
                }
            }
        }
        out
    }

    /// Enumerate every cell whose style carries a non-default fill.
    /// Used after `load_xlsx` to repopulate the UI's `cell_fills` map
    /// so xlsx files authored in Excel keep their background colors
    /// through /FS → /FR.
    pub fn used_cell_fills(&self) -> Vec<(Address, Fill)> {
        let mut out = Vec::new();
        for (sheet_idx, ws) in self.model.workbook.worksheets.iter().enumerate() {
            let sheet = SheetId(sheet_idx as u16);
            for (&row_1b, row_cells) in &ws.sheet_data {
                if row_1b < 1 {
                    continue;
                }
                let row_0b = (row_1b - 1) as u32;
                for &col_1b in row_cells.keys() {
                    if col_1b < 1 {
                        continue;
                    }
                    let col_0b = (col_1b - 1) as u16;
                    let addr = Address::new(sheet, col_0b, row_0b);
                    let Ok(style) = self
                        .model
                        .get_style_for_cell(sheet_idx as u32, row_1b, col_1b)
                    else {
                        continue;
                    };
                    let fill = from_ic_fill(&style.fill);
                    if !fill.is_default() {
                        out.push((addr, fill));
                    }
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
    fn delete_sheet_at_drops_and_shifts() {
        let mut e = IronCalcEngine::new().unwrap();
        e.insert_sheet_at(1).unwrap();
        e.set_user_input(Address::new(SheetId(1), 0, 0), "42")
            .unwrap();
        let to_keep = e.sheet_name(SheetId(1)).unwrap();
        // Delete sheet 0; the second sheet (and its data) shifts to 0.
        e.delete_sheet_at(0).unwrap();
        assert_eq!(e.sheet_count(), 1);
        assert_eq!(e.sheet_name(SheetId(0)).as_deref(), Some(to_keep.as_str()));
        assert_eq!(
            e.get_cell(Address::new(SheetId(0), 0, 0)).unwrap().value,
            Value::Number(42.0)
        );
    }

    #[test]
    fn delete_sheet_at_refuses_last_sheet() {
        let mut e = IronCalcEngine::new().unwrap();
        assert_eq!(e.sheet_count(), 1);
        assert!(e.delete_sheet_at(0).is_err());
        assert_eq!(e.sheet_count(), 1);
    }

    #[test]
    fn delete_sheet_at_out_of_range_fails() {
        let mut e = IronCalcEngine::new().unwrap();
        e.insert_sheet_at(1).unwrap();
        assert!(e.delete_sheet_at(5).is_err());
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
    fn cell_text_style_round_trips_with_values_set_first() {
        // Mirrors the UI's save order: values written via `set_user_input`
        // first, then text styles applied via `set_cell_text_style`.
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_style_rt_v_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("style_vrt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'hi")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "'bye")
            .unwrap();
        e.recalc();
        let bold_italic = TextStyle {
            bold: true,
            italic: true,
            underline: false,
        };
        e.set_cell_text_style(Address::new(SheetId::A, 0, 0), bold_italic)
            .unwrap();
        e.set_cell_text_style(Address::new(SheetId::A, 0, 1), TextStyle::UNDERLINE)
            .unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let styles: std::collections::HashMap<Address, TextStyle> =
            e2.used_cell_text_styles().into_iter().collect();
        assert_eq!(
            styles.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(bold_italic),
            "A1 should not pick up A2's underline"
        );
        assert_eq!(
            styles.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(TextStyle::UNDERLINE),
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_text_style_round_trips_through_xlsx() {
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_style_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("style_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        // Three cells with distinct style combinations.
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'bold")
            .unwrap();
        e.set_cell_text_style(Address::new(SheetId::A, 0, 0), TextStyle::BOLD)
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 1, 0), "'italic")
            .unwrap();
        e.set_cell_text_style(Address::new(SheetId::A, 1, 0), TextStyle::ITALIC)
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 2, 0), "'all")
            .unwrap();
        e.set_cell_text_style(
            Address::new(SheetId::A, 2, 0),
            TextStyle {
                bold: true,
                italic: true,
                underline: true,
            },
        )
        .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let styles: std::collections::HashMap<Address, TextStyle> =
            e2.used_cell_text_styles().into_iter().collect();
        assert_eq!(
            styles.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(TextStyle::BOLD)
        );
        assert_eq!(
            styles.get(&Address::new(SheetId::A, 1, 0)).copied(),
            Some(TextStyle::ITALIC)
        );
        assert_eq!(
            styles.get(&Address::new(SheetId::A, 2, 0)).copied(),
            Some(TextStyle {
                bold: true,
                italic: true,
                underline: true
            })
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn custom_sheet_names_round_trip_through_xlsx() {
        // Simulates opening an xlsx authored in Excel with meaningful
        // tab names. L123 still addresses sheets by letter (A/B/…), but
        // the underlying names must survive /FS → /FR so Excel reopens
        // them intact.
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_sheet_name_rt_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("sheet_names_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.model.rename_sheet_by_index(0, "Q1 Sales").unwrap();
        e.insert_sheet_at(1).unwrap();
        e.model.rename_sheet_by_index(1, "Q2 Budget").unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "100")
            .unwrap();
        e.set_user_input(Address::new(SheetId(1), 0, 0), "200")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(e2.sheet_name(SheetId::A).as_deref(), Some("Q1 Sales"));
        assert_eq!(e2.sheet_name(SheetId(1)).as_deref(), Some("Q2 Budget"));

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_format_round_trips_through_xlsx() {
        use l123_core::Format;
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_fmt_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("fmt_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "10000")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "0.125")
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 2), "42")
            .unwrap();
        e.recalc();
        e.set_cell_format(Address::new(SheetId::A, 0, 0), Format::currency(2))
            .unwrap();
        e.set_cell_format(Address::new(SheetId::A, 0, 1), Format::percent(1))
            .unwrap();
        // A3 left on General — should NOT appear in used_cell_formats.
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let fmts: std::collections::HashMap<Address, Format> =
            e2.used_cell_formats().into_iter().collect();
        assert_eq!(
            fmts.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(Format::currency(2))
        );
        assert_eq!(
            fmts.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(Format::percent(1))
        );
        assert!(
            !fmts.contains_key(&Address::new(SheetId::A, 0, 2)),
            "General-format cell should not surface in used_cell_formats"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[cfg(feature = "wk3")]
    #[test]
    fn load_wk3_reads_committed_fixture() {
        // Loads `tests/acceptance/fixtures/wk3/FILE0001.WK3` (a sparse
        // single-sheet WK3 with "Hello" at A1 and A6) and asserts the
        // engine surfaces both the cell value and the WK3 sheet name.
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let fixture = manifest
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-engine")
            .join("tests/acceptance/fixtures/wk3/FILE0001.WK3");
        if !fixture.exists() {
            eprintln!(
                "skipping load_wk3_reads_committed_fixture: {} missing",
                fixture.display()
            );
            return;
        }
        let mut e = IronCalcEngine::new().unwrap();
        e.load_wk3(&fixture).unwrap();
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 0)).unwrap().value,
            Value::Text("Hello".into()),
            "A1 should be Hello"
        );
        assert_eq!(
            e.get_cell(Address::new(SheetId::A, 0, 5)).unwrap().value,
            Value::Text("Hello".into()),
            "A6 should be Hello"
        );
        assert_eq!(
            e.sheet_name(SheetId::A).as_deref(),
            Some("A"),
            "WK3 sheet name preserved"
        );
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

    #[test]
    fn used_cell_alignments_reads_committed_fixture() {
        // Loads `tests/acceptance/fixtures/xlsx/alignment.xlsx` as
        // written by `cargo run -p l123-engine --example build_fixtures`
        // and asserts each A1..E1 cell's alignment.  This exercises
        // the external-fixture path that the acceptance transcript
        // relies on.
        use l123_core::{Alignment, HAlign, VAlign};
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let fixture = manifest
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-engine")
            .join("tests/acceptance/fixtures/xlsx/alignment.xlsx");
        if !fixture.exists() {
            // Fixture generation is manual (not part of the build).
            // Skip silently when missing so CI on a fresh checkout still
            // runs the rest of the suite; contributors regenerate via
            // the example binary.
            eprintln!(
                "skipping used_cell_alignments_reads_committed_fixture: {} missing",
                fixture.display()
            );
            return;
        }
        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&fixture).unwrap();
        let got: std::collections::HashMap<Address, Alignment> =
            e.used_cell_alignments().into_iter().collect();
        // A1..C1: horizontal-only overrides.
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Left,
                ..Default::default()
            }),
            "A1 should be left-aligned"
        );
        assert_eq!(
            got.get(&Address::new(SheetId::A, 1, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Center,
                ..Default::default()
            }),
            "B1 should be centered"
        );
        assert_eq!(
            got.get(&Address::new(SheetId::A, 2, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Right,
                ..Default::default()
            }),
            "C1 should be right-aligned"
        );
        // D1: left-aligned number.
        assert_eq!(
            got.get(&Address::new(SheetId::A, 3, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Left,
                ..Default::default()
            }),
            "D1 number should pick up left alignment"
        );
        // E1: top-vertical + wrap_text combo.
        assert_eq!(
            got.get(&Address::new(SheetId::A, 4, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Left,
                vertical: VAlign::Top,
                wrap_text: true,
            }),
            "E1 should carry top-vertical + wrap_text"
        );
    }

    #[test]
    fn comments_read_committed_fixture() {
        // The `comments.xlsx` fixture is built by
        // `cargo run -p l123-engine --example build_fixtures` —
        // IronCalc produces the base xlsx and we hand-inject
        // `xl/comments1.xml` + `xl/worksheets/_rels/sheet1.xml.rels`
        // (its exporter doesn't emit them).  IronCalc's importer
        // ignores the author block, so every loaded comment surfaces
        // with `author = ""`.
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let fixture = manifest
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-engine")
            .join("tests/acceptance/fixtures/xlsx/comments.xlsx");
        if !fixture.exists() {
            eprintln!(
                "skipping comments_read_committed_fixture: {} missing",
                fixture.display()
            );
            return;
        }
        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&fixture).unwrap();
        let mut got = e.used_comments();
        got.sort_by_key(|c| (c.addr.sheet.0, c.addr.row, c.addr.col));
        assert_eq!(got.len(), 2, "fixture has two commented cells");
        assert_eq!(got[0].addr, Address::new(SheetId::A, 0, 0));
        assert_eq!(got[0].text, "first note");
        assert_eq!(got[0].author, "", "IronCalc importer drops author");
        assert_eq!(got[1].addr, Address::new(SheetId::A, 1, 1));
        assert_eq!(got[1].text, "second note");
    }

    #[test]
    fn tables_read_committed_fixture() {
        // The `tables.xlsx` fixture is built by `cargo run -p
        // l123-engine --example build_fixtures` — IronCalc produces
        // the base xlsx and we hand-inject `xl/tables/table1.xml` +
        // `xl/worksheets/_rels/sheet1.xml.rels` (its exporter
        // doesn't emit them).
        use l123_core::Address;
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let fixture = manifest
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-engine")
            .join("tests/acceptance/fixtures/xlsx/tables.xlsx");
        if !fixture.exists() {
            eprintln!(
                "skipping tables_read_committed_fixture: {} missing",
                fixture.display()
            );
            return;
        }
        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&fixture).unwrap();
        let mut got = e.used_tables();
        got.sort_by_key(|(_, t)| t.name.clone());
        assert_eq!(got.len(), 1);
        let (sid, t) = &got[0];
        assert_eq!(*sid, SheetId::A);
        assert_eq!(t.name, "Table1");
        assert_eq!(t.display_name, "Table1");
        assert_eq!(t.range.start, Address::new(SheetId::A, 0, 0));
        assert_eq!(t.range.end, Address::new(SheetId::A, 3, 2));
        assert!(t.has_filters);
        assert!(t.has_header_row);
        assert!(!t.has_totals_row);
        assert_eq!(t.columns.len(), 4);
        let col_names: Vec<&str> = t.columns.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(col_names, vec!["Year", "Q1", "Q2", "Q3"]);
        // IronCalc 0.7's importer parses style info from a
        // mis-spelled `<tableInfo>` tag, not the actual
        // `<tableStyleInfo>` we wrote — so the style name silently
        // drops to None on read.  `show_row_stripes` defaults to
        // `true` in that path, which happens to match what we wrote.
        // Pin both behaviors so we notice when upstream fixes the
        // tag-name typo.
        assert_eq!(t.style.name, None);
        assert!(t.style.show_row_stripes);
    }

    #[test]
    fn tables_survive_in_memory_set_and_read() {
        // Like comments / sheet color, IronCalc 0.7's xlsx exporter
        // doesn't write tables — see the `*_upstream_gap` test below.
        // The setter writes into the workbook field; the reader pulls
        // back across all sheets.
        use l123_core::{Address, Range, Table, TableColumn};
        let mut e = IronCalcEngine::new().unwrap();
        let r = Range {
            start: Address::new(SheetId::A, 0, 0),
            end: Address::new(SheetId::A, 3, 5),
        };
        let t = Table {
            name: "Sales".into(),
            display_name: "Sales".into(),
            range: r,
            has_filters: true,
            has_header_row: true,
            has_totals_row: false,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Year".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Total".into(),
                    totals_row_function: Some("sum".into()),
                    ..Default::default()
                },
            ],
            ..Default::default()
        };
        e.set_table(SheetId::A, t.clone()).unwrap();
        let got = e.used_tables();
        assert_eq!(got.len(), 1);
        assert_eq!(got[0].0, SheetId::A);
        assert_eq!(got[0].1, t);

        e.unset_table("Sales").unwrap();
        assert!(e.used_tables().is_empty());
    }

    #[test]
    fn tables_are_dropped_on_xlsx_save_upstream_gap() {
        // IronCalc 0.7 doesn't serialize tables through its xlsx
        // exporter (no `xl/tables/*.xml` parts written).  Pin the
        // current behavior so we notice when upstream closes the gap.
        use l123_core::{Address, Range, Table, TableColumn};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_table_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("table.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'h")
            .unwrap();
        let r = Range {
            start: Address::new(SheetId::A, 0, 0),
            end: Address::new(SheetId::A, 1, 2),
        };
        e.set_table(
            SheetId::A,
            Table {
                name: "T1".into(),
                display_name: "T1".into(),
                range: r,
                has_header_row: true,
                columns: vec![TableColumn {
                    id: 1,
                    name: "A".into(),
                    ..Default::default()
                }],
                ..Default::default()
            },
        )
        .unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert!(
            e2.used_tables().is_empty(),
            "IronCalc xlsx save drops tables. If this starts returning \
             a non-empty Vec, upstream closed the gap; flip the assertion \
             and tell users /FS preserves tables."
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn sheet_state_round_trips_through_xlsx() {
        // IronCalc round-trips sheet state natively via the workbook
        // XML's `<sheet state="hidden|veryHidden"/>` attribute.
        use l123_core::SheetState;
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_sheet_state_rt_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("sheet_state_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.insert_sheet_at(1).unwrap();
        e.insert_sheet_at(2).unwrap();
        // A: visible (default), B: hidden, C: very-hidden.
        e.set_sheet_state(SheetId(1), SheetState::Hidden).unwrap();
        e.set_sheet_state(SheetId(2), SheetState::VeryHidden)
            .unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(e2.sheet_state(SheetId::A), SheetState::Visible);
        assert_eq!(e2.sheet_state(SheetId(1)), SheetState::Hidden);
        assert_eq!(e2.sheet_state(SheetId(2)), SheetState::VeryHidden);

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn frozen_panes_round_trip_through_xlsx() {
        // IronCalc 0.7 round-trips frozen rows + columns natively.
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_frozen_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("frozen_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.insert_sheet_at(1).unwrap();
        // Sheet A: 2 frozen rows, 1 frozen column.
        e.set_frozen_panes(SheetId::A, 2, 1).unwrap();
        // Sheet B: only frozen rows.
        e.set_frozen_panes(SheetId(1), 3, 0).unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(e2.frozen_panes(SheetId::A), (2, 1));
        assert_eq!(e2.frozen_panes(SheetId(1)), (3, 0));

        // Clearing both wipes the freeze.
        e2.set_frozen_panes(SheetId::A, 0, 0).unwrap();
        assert_eq!(e2.frozen_panes(SheetId::A), (0, 0));

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn merges_round_trip_through_xlsx() {
        // IronCalc 0.7's xlsx exporter DOES write `<mergeCells>` —
        // unlike comments / sheet color / dotted borders.  Full
        // round-trip is exercised here.
        use l123_core::Merge;
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_merge_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("merge_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        // A1:C1 (horizontal)
        let m1 = Merge::from_corners(
            Address::new(SheetId::A, 0, 0),
            Address::new(SheetId::A, 2, 0),
        )
        .unwrap();
        // B3:C5 (rectangular)
        let m2 = Merge::from_corners(
            Address::new(SheetId::A, 1, 2),
            Address::new(SheetId::A, 2, 4),
        )
        .unwrap();
        e.set_user_input(m1.anchor, "'header").unwrap();
        e.set_user_input(m2.anchor, "'box").unwrap();
        e.set_merged_range(m1).unwrap();
        e.set_merged_range(m2).unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let mut got: Vec<(SheetId, Merge)> = e2.used_merged_cells();
        got.sort_by_key(|(_, m)| (m.anchor.col, m.anchor.row));
        assert_eq!(got.len(), 2);
        assert_eq!(got[0], (SheetId::A, m1));
        assert_eq!(got[1], (SheetId::A, m2));

        // Unmerge one and verify it disappears.
        e2.unset_merged_range(m1).unwrap();
        let after = e2.used_merged_cells();
        assert_eq!(after.len(), 1);
        assert_eq!(after[0], (SheetId::A, m2));

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn comments_round_trip_via_in_memory_set_and_read() {
        // Like sheet_color, IronCalc 0.7's xlsx exporter does NOT
        // write comments out (see the *_upstream_gap test below).
        // The setter writes into the in-memory model field; the
        // reader pulls comments back across all sheets.
        use l123_core::Comment;
        let mut e = IronCalcEngine::new().unwrap();
        e.insert_sheet_at(1).unwrap();
        let c1 = Comment::new(Address::new(SheetId::A, 0, 0), "Alice", "looks high");
        let c2 = Comment::new(Address::new(SheetId::A, 1, 4), "Bob", "");
        let c3 = Comment::new(Address::new(SheetId(1), 0, 0), "Carol", "sheet 2 note");
        e.set_comment(c1.clone()).unwrap();
        e.set_comment(c2.clone()).unwrap();
        e.set_comment(c3.clone()).unwrap();
        let mut got = e.used_comments();
        // Reader doesn't promise an order; sort for deterministic compare.
        got.sort_by_key(|c| (c.addr.sheet.0, c.addr.row, c.addr.col));
        let mut want = vec![c1.clone(), c2.clone(), c3.clone()];
        want.sort_by_key(|c| (c.addr.sheet.0, c.addr.row, c.addr.col));
        assert_eq!(got, want);

        // Delete one and confirm it disappears.
        e.delete_comment(c2.addr).unwrap();
        let after_delete = e.used_comments();
        assert_eq!(after_delete.len(), 2);
        assert!(!after_delete.iter().any(|c| c.addr == c2.addr));
    }

    #[test]
    fn comments_are_dropped_on_xlsx_save_upstream_gap() {
        // IronCalc 0.7 doesn't serialize comments through its xlsx
        // exporter (no `xl/comments1.xml` part written).  Pin the
        // current behavior so we notice when upstream closes the gap.
        use l123_core::Comment;
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_comment_rt_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("comment.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "1")
            .unwrap();
        e.set_comment(Comment::new(
            Address::new(SheetId::A, 0, 0),
            "Alice",
            "this should not survive the save",
        ))
        .unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert!(
            e2.used_comments().is_empty(),
            "IronCalc xlsx save drops comments. If this starts returning a non-empty Vec, \
             upstream closed the gap; flip the assertion and tell users /FS preserves notes."
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_border_dotted_survives_in_memory_set_and_read() {
        // Dotted round-trip through xlsx fails (see the
        // `*_upstream_gap` test below).  This test proves the setter
        // and reader themselves handle Dotted correctly, so when
        // IronCalc closes the gap, no L123-side work is needed.
        use l123_core::{Border, BorderEdge, BorderStyle};
        let mut e = IronCalcEngine::new().unwrap();
        let b = Border {
            right: Some(BorderEdge {
                style: BorderStyle::Dotted,
                color: None,
            }),
            ..Default::default()
        };
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'x")
            .unwrap();
        e.set_cell_border(Address::new(SheetId::A, 0, 0), b)
            .unwrap();
        let got: std::collections::HashMap<Address, Border> =
            e.used_cell_borders().into_iter().collect();
        assert_eq!(got.get(&Address::new(SheetId::A, 0, 0)).copied(), Some(b));
    }

    #[test]
    fn dotted_border_degrades_on_xlsx_round_trip_upstream_gap() {
        // IronCalc 0.7's xlsx importer (xlsx/src/import/styles.rs
        // `get_border`) has no arm for `"dotted"` — any unknown style
        // falls back to `Thin`.  Pins the current behavior so we
        // notice when upstream adds the missing match arm.
        use l123_core::{Border, BorderEdge, BorderStyle};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_border_dotted_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("dotted.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'x")
            .unwrap();
        e.set_cell_border(
            Address::new(SheetId::A, 0, 0),
            Border {
                right: Some(BorderEdge {
                    style: BorderStyle::Dotted,
                    color: None,
                }),
                ..Default::default()
            },
        )
        .unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let got: std::collections::HashMap<Address, Border> =
            e2.used_cell_borders().into_iter().collect();
        // NOTE: expected to be `Dotted` once upstream fixes the gap.
        // Today it degrades to `Thin`.
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 0))
                .and_then(|b| b.right)
                .map(|e| e.style),
            Some(BorderStyle::Thin),
            "IronCalc's xlsx importer silently drops `dotted` → Thin. \
             If this starts returning Dotted, upstream closed the gap — \
             flip the assertion and tell users dotted borders round-trip."
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_border_round_trips_through_xlsx() {
        use l123_core::{Border, BorderEdge, BorderStyle, RgbColor};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_border_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("border_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor {
            r: 0xFF,
            g: 0x00,
            b: 0x00,
        };

        // A1: thin right border, no color.
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'thin")
            .unwrap();
        let a1_border = Border {
            right: Some(BorderEdge {
                style: BorderStyle::Thin,
                color: None,
            }),
            ..Default::default()
        };
        e.set_cell_border(Address::new(SheetId::A, 0, 0), a1_border)
            .unwrap();
        // A2: double top + thick bottom + red left.
        e.set_user_input(Address::new(SheetId::A, 0, 1), "'box")
            .unwrap();
        let a2_border = Border {
            left: Some(BorderEdge {
                style: BorderStyle::Medium,
                color: Some(red),
            }),
            top: Some(BorderEdge {
                style: BorderStyle::Double,
                color: None,
            }),
            bottom: Some(BorderEdge {
                style: BorderStyle::Thick,
                color: None,
            }),
            right: None,
        };
        e.set_cell_border(Address::new(SheetId::A, 0, 1), a2_border)
            .unwrap();
        // A3: dashed right border.  Dotted *does not round-trip*
        // through IronCalc 0.7's xlsx importer (see
        // `dotted_border_degrades_on_xlsx_round_trip_upstream_gap`
        // below for the pinned IronCalc bug); we only exercise the
        // round-trip with styles IronCalc's importer understands.
        e.set_user_input(Address::new(SheetId::A, 0, 2), "'dash")
            .unwrap();
        let a3_border = Border {
            right: Some(BorderEdge {
                style: BorderStyle::Dashed,
                color: None,
            }),
            ..Default::default()
        };
        e.set_cell_border(Address::new(SheetId::A, 0, 2), a3_border)
            .unwrap();
        // A4 plain — should NOT appear in used_cell_borders.
        e.set_user_input(Address::new(SheetId::A, 0, 3), "'plain")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let got: std::collections::HashMap<Address, Border> =
            e2.used_cell_borders().into_iter().collect();
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(a1_border),
            "A1: thin right border"
        );
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(a2_border),
            "A2: medium-left + double-top + thick-bottom"
        );
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 2)).copied(),
            Some(a3_border),
            "A3: dashed right"
        );
        assert!(
            !got.contains_key(&Address::new(SheetId::A, 0, 3)),
            "plain cell should not surface"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_font_style_round_trips_through_xlsx() {
        use l123_core::{FontStyle, RgbColor};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_fs_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("fs_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor {
            r: 0xFF,
            g: 0x00,
            b: 0x00,
        };
        let blue = RgbColor {
            r: 0x33,
            g: 0x66,
            b: 0xCC,
        };
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'red")
            .unwrap();
        e.set_cell_font_style(
            Address::new(SheetId::A, 0, 0),
            FontStyle {
                color: Some(red),
                ..Default::default()
            },
        )
        .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "'big blue struck")
            .unwrap();
        e.set_cell_font_style(
            Address::new(SheetId::A, 0, 1),
            FontStyle {
                color: Some(blue),
                size: Some(14),
                strike: true,
            },
        )
        .unwrap();
        // A3 — no font override; should NOT appear in used_cell_font_styles.
        e.set_user_input(Address::new(SheetId::A, 0, 2), "'plain")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let got: std::collections::HashMap<Address, FontStyle> =
            e2.used_cell_font_styles().into_iter().collect();
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(FontStyle {
                color: Some(red),
                ..Default::default()
            }),
            "A1 carries red only"
        );
        assert_eq!(
            got.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(FontStyle {
                color: Some(blue),
                size: Some(14),
                strike: true,
            }),
            "A2 carries blue + 14pt + strike"
        );
        assert!(
            !got.contains_key(&Address::new(SheetId::A, 0, 2)),
            "plain cell should not surface"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn sheet_color_reads_committed_fixture() {
        // The `sheet_color.xlsx` fixture is built by
        // `cargo run -p l123-engine --example build_fixtures` — IronCalc
        // produces the base xlsx and we hand-patch `<tabColor>` in
        // because IronCalc's exporter doesn't emit it. See the
        // `build_sheet_color` function in examples/build_fixtures.rs
        // for the patch shape.
        use l123_core::RgbColor;
        let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
        let fixture = manifest
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-engine")
            .join("tests/acceptance/fixtures/xlsx/sheet_color.xlsx");
        if !fixture.exists() {
            eprintln!(
                "skipping sheet_color_reads_committed_fixture: {} missing",
                fixture.display()
            );
            return;
        }
        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&fixture).unwrap();
        assert_eq!(e.sheet_color(SheetId::A), None, "Overview sheet untinted");
        assert_eq!(
            e.sheet_color(SheetId(1)),
            Some(RgbColor {
                r: 0xDB,
                g: 0xBE,
                b: 0x29
            }),
            "Q1 Red sheet should carry the patched tabColor"
        );
        assert_eq!(
            e.sheet_color(SheetId(2)),
            Some(RgbColor {
                r: 0x33,
                g: 0x66,
                b: 0xCC
            }),
            "Q2 Blue sheet should carry the patched tabColor"
        );
    }

    #[test]
    fn sheet_color_is_dropped_on_xlsx_save_upstream_gap() {
        // IronCalc 0.7's xlsx importer *reads* `<sheetPr><tabColor/>`
        // (see xlsx/src/import/worksheets.rs:197) but the exporter
        // doesn't write it.  That makes the round-trip asymmetric: we
        // can open an Excel-authored workbook with colored tabs and
        // see the colors in the engine, but saving through IronCalc
        // drops them.  This test pins that behavior so we notice (and
        // update the user-facing docs) when IronCalc closes the gap.
        use l123_core::RgbColor;
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_engine_sheet_color_rt_{}_{}",
            process::id(),
            nanos
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("sheet_color_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor {
            r: 0xDB,
            g: 0xBE,
            b: 0x29,
        };
        e.set_sheet_color(SheetId::A, Some(red)).unwrap();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        assert_eq!(
            e2.sheet_color(SheetId::A),
            None,
            "IronCalc xlsx save drops tab color — if this starts returning Some(...) \
             upstream has fixed the gap; flip this assertion and tell users /FS preserves tabs."
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn sheet_color_none_clears_override() {
        use l123_core::RgbColor;
        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor { r: 255, g: 0, b: 0 };
        e.set_sheet_color(SheetId::A, Some(red)).unwrap();
        assert_eq!(e.sheet_color(SheetId::A), Some(red));
        e.set_sheet_color(SheetId::A, None).unwrap();
        assert_eq!(e.sheet_color(SheetId::A), None);
    }

    #[test]
    fn cell_fill_survives_in_memory_set_and_read() {
        // Isolates whether the setter/reader pair work without an xlsx
        // round-trip.  If this passes and the xlsx round-trip doesn't,
        // the fault is in IronCalc's serializer (field choice).
        use l123_core::{Fill, FillPattern, RgbColor};
        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor {
            r: 0xFF,
            g: 0,
            b: 0,
        };
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'x")
            .unwrap();
        e.set_cell_fill(Address::new(SheetId::A, 0, 0), Fill::solid(red))
            .unwrap();
        let fills: std::collections::HashMap<Address, Fill> =
            e.used_cell_fills().into_iter().collect();
        assert_eq!(
            fills.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(Fill {
                pattern: FillPattern::Solid,
                bg: Some(red)
            })
        );
    }

    #[test]
    fn cell_fill_round_trips_through_xlsx() {
        use l123_core::{Fill, FillPattern, RgbColor};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_fill_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("fill_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        let red = RgbColor {
            r: 0xFF,
            g: 0,
            b: 0,
        };
        let blue = RgbColor {
            r: 0x33,
            g: 0x66,
            b: 0xCC,
        };
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'red bg")
            .unwrap();
        e.set_cell_fill(Address::new(SheetId::A, 0, 0), Fill::solid(red))
            .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "'blue bg")
            .unwrap();
        e.set_cell_fill(Address::new(SheetId::A, 0, 1), Fill::solid(blue))
            .unwrap();
        // A3 left on default — should NOT appear in used_cell_fills.
        e.set_user_input(Address::new(SheetId::A, 0, 2), "'plain")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let fills: std::collections::HashMap<Address, Fill> =
            e2.used_cell_fills().into_iter().collect();
        assert_eq!(
            fills.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(Fill {
                pattern: FillPattern::Solid,
                bg: Some(red)
            }),
            "A1 should round-trip the red fill"
        );
        assert_eq!(
            fills.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(Fill {
                pattern: FillPattern::Solid,
                bg: Some(blue)
            }),
            "A2 should round-trip the blue fill"
        );
        assert!(
            !fills.contains_key(&Address::new(SheetId::A, 0, 2)),
            "no-fill cell should not surface in used_cell_fills"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn cell_alignment_round_trips_through_xlsx() {
        use l123_core::{Alignment, HAlign, VAlign};
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_engine_align_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("align_rt.xlsx");

        let mut e = IronCalcEngine::new().unwrap();
        // Four cells with distinct horizontal/vertical/wrap combos.
        e.set_user_input(Address::new(SheetId::A, 0, 0), "'right")
            .unwrap();
        e.set_cell_alignment(
            Address::new(SheetId::A, 0, 0),
            Alignment {
                horizontal: HAlign::Right,
                ..Default::default()
            },
        )
        .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 1), "'center")
            .unwrap();
        e.set_cell_alignment(
            Address::new(SheetId::A, 0, 1),
            Alignment {
                horizontal: HAlign::Center,
                ..Default::default()
            },
        )
        .unwrap();
        e.set_user_input(Address::new(SheetId::A, 0, 2), "'top wrap")
            .unwrap();
        e.set_cell_alignment(
            Address::new(SheetId::A, 0, 2),
            Alignment {
                horizontal: HAlign::Left,
                vertical: VAlign::Top,
                wrap_text: true,
            },
        )
        .unwrap();
        // A4 left on default — should NOT appear in used_cell_alignments.
        e.set_user_input(Address::new(SheetId::A, 0, 3), "'plain")
            .unwrap();
        e.recalc();
        e.save_xlsx(&path).unwrap();

        let mut e2 = IronCalcEngine::new().unwrap();
        e2.load_xlsx(&path).unwrap();
        let aligns: std::collections::HashMap<Address, Alignment> =
            e2.used_cell_alignments().into_iter().collect();
        assert_eq!(
            aligns.get(&Address::new(SheetId::A, 0, 0)).copied(),
            Some(Alignment {
                horizontal: HAlign::Right,
                ..Default::default()
            })
        );
        assert_eq!(
            aligns.get(&Address::new(SheetId::A, 0, 1)).copied(),
            Some(Alignment {
                horizontal: HAlign::Center,
                ..Default::default()
            })
        );
        assert_eq!(
            aligns.get(&Address::new(SheetId::A, 0, 2)).copied(),
            Some(Alignment {
                horizontal: HAlign::Left,
                vertical: VAlign::Top,
                wrap_text: true,
            })
        );
        assert!(
            !aligns.contains_key(&Address::new(SheetId::A, 0, 3)),
            "default-aligned cell should not surface in used_cell_alignments"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }
}
