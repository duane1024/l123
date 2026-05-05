//! Mouse handling, icon-panel dispatch, and SmartIcon implementations.

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers, MouseButton, MouseEvent, MouseEventKind};
use l123_core::{
    Address, Border, BorderEdge, CellContents, Mode, Range, SheetId, TextStyle, Value,
};
use l123_engine::Engine;

use super::types::*;
use super::{App, MOUSE_SCROLL_STEP, ROW_GUTTER};

impl App {
    pub fn handle_mouse(&mut self, m: MouseEvent) {
        let icon_geom = self.icon_panel_area.get();
        let inside_icon = icon_geom.is_some_and(|g| {
            let r = g.rect;
            m.column >= r.x && m.column < r.x + r.width && m.row >= r.y && m.row < r.y + r.height
        });

        match m.kind {
            MouseEventKind::Moved => {
                // Slot 16 (pager) is excluded from hover tooltip —
                // its function is rendered on the slot itself.
                self.hovered_icon = match (inside_icon, icon_geom) {
                    (true, Some(g)) => match Self::hit_test_slot(&g, m.row) {
                        Some(slot) if slot < 16 => Some((self.current_panel, slot)),
                        _ => None,
                    },
                    _ => None,
                };
            }
            MouseEventKind::Down(MouseButton::Left) => {
                // Each fresh press resets the drag-anchor; the only
                // path that re-arms it is a grid-cell hit in
                // Ready/Point below.
                self.drag_anchor = None;
                // Icon panel wins over the grid when both would hit —
                // it sits on top visually.
                if let (true, Some(g)) = (inside_icon, icon_geom) {
                    self.dispatch_icon_click(&g, m.column, m.row);
                } else if let Some(addr) = self.cell_at_screen(m.column, m.row) {
                    match self.mode {
                        // In POINT, the anchor is untouched by
                        // move_pointer_to, so an anchored range
                        // extends (or shrinks) to the clicked cell
                        // naturally; unanchored POINT just moves the
                        // pointer.
                        Mode::Ready | Mode::Point => {
                            self.move_pointer_to(addr);
                            // Remember where the press landed so a
                            // follow-up Drag can promote READY into
                            // POINT anchored here. Cleared on Up.
                            self.drag_anchor = Some(addr);
                        }
                        // Mid-formula: splice the clicked cell's
                        // address into the entry buffer if the buffer
                        // is in a cell-ref-accepting position. The
                        // pointer itself stays put — the entry still
                        // belongs to the originally-selected cell.
                        Mode::Value | Mode::Edit => self.splice_cell_ref_into_entry(addr),
                        _ => {}
                    }
                }
            }
            MouseEventKind::Drag(MouseButton::Left) => {
                // Drag-to-select. Only acts if a prior Down landed on
                // a grid cell (`drag_anchor` set). Off-grid drag
                // motion freezes the pointer rather than snapping.
                let Some(anchor) = self.drag_anchor else {
                    return;
                };
                let Some(addr) = self.cell_at_screen(m.column, m.row) else {
                    return;
                };
                match self.mode {
                    Mode::Ready => {
                        // Promote into POINT anchored at the press
                        // cell, then move the free end to where the
                        // cursor is now.
                        self.menu = None;
                        self.point = Some(PointState {
                            anchor: Some(anchor),
                            pending: PendingCommand::MouseSelect,
                            typed: String::new(),
                        });
                        self.mode = Mode::Point;
                        self.move_pointer_to(addr);
                    }
                    Mode::Point => {
                        // Existing POINT (e.g. /RE) — its anchor is
                        // already set; live-extend by moving pointer.
                        self.move_pointer_to(addr);
                    }
                    _ => {}
                }
            }
            MouseEventKind::Up(MouseButton::Left) => {
                self.drag_anchor = None;
            }
            MouseEventKind::ScrollDown => {
                let wb = self.wb_mut();
                wb.viewport_row_offset = wb.viewport_row_offset.saturating_add(MOUSE_SCROLL_STEP);
            }
            MouseEventKind::ScrollUp => {
                let wb = self.wb_mut();
                wb.viewport_row_offset = wb.viewport_row_offset.saturating_sub(MOUSE_SCROLL_STEP);
            }
            _ => {}
        }
    }

    /// Jump the cell pointer to `addr`, scrolling the viewport only
    /// if needed. Used by mouse click-to-move; keyboard `/RG` (F5)
    /// has its own prompt-driven path.
    pub(super) fn move_pointer_to(&mut self, addr: Address) {
        self.wb_mut().pointer = addr;
        self.scroll_into_view();
    }

    /// Insert `addr` (short form, e.g. `C3`) at the entry buffer
    /// cursor — but only when the position before the cursor is one
    /// where a cell reference is grammatically valid (start of buffer,
    /// or after an operator/paren/comma/dot/comparison/logical token).
    /// Called from mouse click-in-grid during VALUE/EDIT, so the user
    /// can build formulas by clicking instead of typing references.
    pub(super) fn splice_cell_ref_into_entry(&mut self, addr: Address) {
        let Some(entry) = self.entry.as_mut() else {
            return;
        };
        if !matches!(entry.kind, EntryKind::Value | EntryKind::Edit) {
            return;
        }
        let accepts_ref = match entry.buffer[..entry.cursor].chars().last() {
            None => true,
            // `.` covers both the range separator `..` and an
            // in-progress `.`; `#` covers `#AND#`/`#OR#`/`#NOT#`.
            Some(c) => matches!(
                c,
                '+' | '-' | '*' | '/' | '^' | '(' | ',' | '.' | '=' | '<' | '>' | '#' | ' '
            ),
        };
        if !accepts_ref {
            return;
        }
        let s = addr.display_short();
        entry.buffer.insert_str(entry.cursor, &s);
        entry.cursor += s.len();
    }

    /// Map a screen coordinate to the cell address it sits on, or
    /// `None` for clicks on the column header, row-number gutter, or
    /// outside the grid. Requires that a grid has already rendered
    /// this session so `last_grid_area` is populated.
    pub(super) fn cell_at_screen(&self, col: u16, row: u16) -> Option<Address> {
        let area = self.last_grid_area.get()?;
        if area.width <= ROW_GUTTER || area.height < 2 {
            return None;
        }
        // Reject: outside rect, on column header row, on row gutter.
        if col < area.x + ROW_GUTTER
            || col >= area.x + area.width
            || row <= area.y
            || row >= area.y + area.height
        {
            return None;
        }

        let local_x = col - area.x - ROW_GUTTER;
        let local_y = row - area.y - 1;
        let content_width = area.width - ROW_GUTTER;

        let col_idx = self
            .visible_column_layout(content_width)
            .into_iter()
            .find(|(_, x_off, w)| local_x >= *x_off && local_x < *x_off + *w)
            .map(|(c, _, _)| c)?;

        let sheet = self.wb().pointer.sheet;
        let row_idx = self.wb().viewport_row_offset.saturating_add(local_y as u32);
        Some(Address::new(sheet, col_idx, row_idx))
    }

    /// Map a row within the icon panel to a slot index `0..=16`.
    /// Slot 16 is the panel navigator. Returns `None` for rows above
    /// or below the rendered cells, or when geometry is degenerate.
    ///
    /// Each icon spans a fractional cell (image is 1:17 aspect, fit
    /// into a typically taller area). Sampling at the cell midpoint
    /// pixel and mapping into the *true* rendered pixel height —
    /// not the ceiled cell height — pins each cell to the icon that
    /// covers most of it. The earlier cells-only formula compressed
    /// 17 icons across the post-ceiling cell count, which slowly
    /// drifted the bottom slots up by half an icon.
    pub(super) fn hit_test_slot(geom: &IconPanelGeom, row: u16) -> Option<usize> {
        if geom.rect.height == 0 || row < geom.rect.y {
            return None;
        }
        let offset = (row - geom.rect.y) as u32;
        if offset >= geom.rect.height as u32 {
            return None;
        }
        if geom.rendered_px_h == 0 || geom.font_px_h == 0 {
            return None;
        }
        let mid_px = offset * geom.font_px_h as u32 + geom.font_px_h as u32 / 2;
        // Bottom slack between the rendered image and the ceiled cell
        // rect — clicks here visually land on empty grey, but treat
        // them as the pager so the bottom row isn't an inert dead zone.
        if mid_px >= geom.rendered_px_h {
            return Some(16);
        }
        Some(((mid_px * 17) / geom.rendered_px_h).min(16) as usize)
    }

    /// Map a mouse click on the icon panel to a slot and fire that
    /// slot's action. Slot 16 is the panel navigator; clicks on its
    /// left half go to the previous panel, right half to the next.
    pub(super) fn dispatch_icon_click(&mut self, geom: &IconPanelGeom, column: u16, row: u16) {
        let Some(slot) = Self::hit_test_slot(geom, row) else {
            return;
        };

        if slot == 16 {
            // Pager: left half → previous panel, right half → next.
            let half = geom.rect.x + geom.rect.width / 2;
            self.current_panel = if column < half {
                self.current_panel.prev()
            } else {
                self.current_panel.next()
            };
            self.refresh_icon_panel();
            return;
        }

        let ids = self.current_panel.icon_ids();
        let id = ids[slot];
        match l123_graph::icon_action(id) {
            l123_graph::IconAction::MenuPath(path) => self.dispatch_menu_path(path),
            l123_graph::IconAction::WysiwygMenuPath(path) => self.dispatch_wysiwyg_menu_path(path),
            l123_graph::IconAction::TextStyleToggle { bits } => self.dispatch_icon_text_style(bits),
            l123_graph::IconAction::SysKey(act) => self.dispatch_sys_action(act),
            l123_graph::IconAction::PageNav => {} // reached only via slot 16
            l123_graph::IconAction::Noop => {}
        }
    }

    /// Open the slash menu and descend via the given accelerator
    /// letters — equivalent to the user typing "/" then each char.
    pub(super) fn dispatch_menu_path(&mut self, path: &str) {
        self.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        for c in path.chars() {
            self.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    /// Open the WYSIWYG colon menu and descend via the given accelerator
    /// letters — equivalent to the user typing `:` then each char.
    pub(super) fn dispatch_wysiwyg_menu_path(&mut self, path: &str) {
        self.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
        for c in path.chars() {
            self.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    /// SmartIcons toggle for Bold / Italic / Underline. Acts on the
    /// active POINT highlight if one is in progress, otherwise on the
    /// single cell at the pointer. If every cell in the target already
    /// carries `bits`, clear them; otherwise set on all. Always lands
    /// back in Ready.
    pub(super) fn dispatch_icon_text_style(&mut self, bits: TextStyle) {
        let range = if matches!(self.mode, Mode::Point) {
            self.highlight_range()
        } else {
            Range::single(self.wb().pointer)
        };
        self.point = None;
        self.menu = None;
        self.prompt = None;

        let r = range.normalized();
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let all_have_bits = sheets.iter().all(|sheet| {
            (r.start.row..=r.end.row).all(|row| {
                (r.start.col..=r.end.col).all(|col| {
                    let addr = Address::new(*sheet, col, row);
                    let cur = self
                        .wb()
                        .cell_text_styles
                        .get(&addr)
                        .copied()
                        .unwrap_or_default();
                    cur.merge(bits) == cur
                })
            })
        });
        self.execute_range_text_style(range, bits, !all_have_bits);
        self.mode = Mode::Ready;
    }

    pub(super) fn dispatch_sys_action(&mut self, action: l123_graph::SysAction) {
        use l123_graph::SysAction;
        let (code, mods) = match action {
            SysAction::GraphView => (KeyCode::F(10), KeyModifiers::NONE),
            SysAction::Undo => (KeyCode::F(4), KeyModifiers::ALT),
            SysAction::Home => (KeyCode::Home, KeyModifiers::NONE),
            SysAction::Recalc => (KeyCode::F(9), KeyModifiers::NONE),
            SysAction::Edit => (KeyCode::F(2), KeyModifiers::NONE),
            SysAction::Goto => (KeyCode::F(5), KeyModifiers::NONE),
            SysAction::NextSheet => (KeyCode::PageDown, KeyModifiers::CONTROL),
            SysAction::PrevSheet => (KeyCode::PageUp, KeyModifiers::CONTROL),
            SysAction::StepToggle => (KeyCode::F(2), KeyModifiers::ALT),
            SysAction::RunMacro => (KeyCode::F(3), KeyModifiers::ALT),
            // END+arrow / END+HOME have no native crossterm equivalent
            // — the End key is a one-shot prefix in 1-2-3, not a key
            // chord. Run the move directly.
            SysAction::BlockEndHome => {
                self.move_pointer_to(self.active_area_corner());
                return;
            }
            SysAction::BlockEndDown => return self.block_end_jump(0, 1),
            SysAction::BlockEndUp => return self.block_end_jump(0, -1),
            SysAction::BlockEndRight => return self.block_end_jump(1, 0),
            SysAction::BlockEndLeft => return self.block_end_jump(-1, 0),
            // Pure viewport scrolls — pointer untouched. The next
            // arrow press will pull the viewport back to the pointer
            // via `scroll_into_view`.
            SysAction::ScrollColumnLeft => return self.scroll_viewport(-1, 0),
            SysAction::ScrollColumnRight => return self.scroll_viewport(1, 0),
            SysAction::ScrollRowUp => return self.scroll_viewport(0, -1),
            SysAction::ScrollRowDown => return self.scroll_viewport(0, 1),
            SysAction::ScrollScreenLeft => {
                let n = self.visible_scrolling_cols() as i32;
                return self.scroll_viewport(-n, 0);
            }
            SysAction::ScrollScreenRight => {
                let n = self.visible_scrolling_cols() as i32;
                return self.scroll_viewport(n, 0);
            }
            SysAction::ScrollScreenUp => {
                let n = self.visible_scrolling_rows() as i32;
                return self.scroll_viewport(0, -n);
            }
            SysAction::ScrollScreenDown => {
                let n = self.visible_scrolling_rows() as i32;
                return self.scroll_viewport(0, n);
            }
            SysAction::SumRange => return self.dispatch_sum_smarticon(),
            SysAction::TodayDate => return self.dispatch_today_smarticon(),
            SysAction::OutlineRange => return self.dispatch_outline_smarticon(),
            SysAction::SortAscending => return self.dispatch_sort_smarticon(SortDir::Ascending),
            SysAction::SortDescending => return self.dispatch_sort_smarticon(SortDir::Descending),
        };
        self.handle_key(KeyEvent::new(code, mods));
    }

    /// Lower-right corner of the occupied rectangle on the current
    /// sheet. Empty sheet → A1. The corner cell itself need not be
    /// occupied — `(max_col, max_row)` are computed independently.
    pub(super) fn active_area_corner(&self) -> Address {
        let sheet = self.wb().pointer.sheet;
        let (max_col, max_row) = self
            .wb()
            .cells
            .keys()
            .filter(|a| a.sheet == sheet)
            .fold((0u16, 0u32), |(c, r), a| (c.max(a.col), r.max(a.row)));
        Address::new(sheet, max_col, max_row)
    }

    /// Shift the viewport offsets by the given deltas, clamped to
    /// `[0, MAX_COLS-1]` × `[0, MAX_ROWS-1]`. Pointer is untouched.
    pub(super) fn scroll_viewport(&mut self, d_col: i32, d_row: i32) {
        let max_col = (l123_core::address::MAX_COLS - 1) as i32;
        let max_row = (l123_core::address::MAX_ROWS - 1) as i32;
        let new_col = (self.wb().viewport_col_offset as i32 + d_col).clamp(0, max_col) as u16;
        let new_row = (self.wb().viewport_row_offset as i32 + d_row).clamp(0, max_row) as u32;
        self.wb_mut().viewport_col_offset = new_col;
        self.wb_mut().viewport_row_offset = new_row;
    }

    /// Number of scrolling rows currently visible in the body area.
    /// Falls back to 20 when no grid has rendered yet — matches the
    /// PgUp/PgDn convention.
    pub(super) fn visible_scrolling_rows(&self) -> u32 {
        let Some(area) = self.last_grid_area.get() else {
            return 20;
        };
        if area.height < 2 {
            return 1;
        }
        let visible = (area.height - 1) as u32;
        let sheet = self.wb().pointer.sheet;
        let frozen: u32 = self.wb().frozen.get(&sheet).map(|f| f.0).unwrap_or(0);
        visible.saturating_sub(frozen).max(1)
    }

    /// Number of scrolling columns currently visible. Falls back to 8
    /// columns at default width when no grid has rendered.
    pub(super) fn visible_scrolling_cols(&self) -> u16 {
        let Some(area) = self.last_grid_area.get() else {
            return 8;
        };
        if area.width <= ROW_GUTTER {
            return 1;
        }
        let content_width = area.width - ROW_GUTTER;
        let sheet = self.wb().pointer.sheet;
        let frozen: u16 = self.wb().frozen.get(&sheet).map(|f| f.1).unwrap_or(0);
        let layout = self.visible_column_layout(content_width);
        let scrolling = layout.iter().filter(|(c, _, _)| *c >= frozen).count();
        (scrolling as u16).max(1)
    }

    /// END+arrow scan rules per 1-2-3 R3.4: from a non-blank cell with
    /// a non-blank neighbour, jump to the last non-blank in that run;
    /// otherwise (cur is blank, or first step is blank) skip blanks
    /// to the first non-blank. If no non-blank is found, stop at the
    /// worksheet boundary.
    pub(super) fn block_end_jump(&mut self, d_col: i32, d_row: i32) {
        let start = self.wb().pointer;
        let cur_blank = !self.wb().cells.contains_key(&start);
        let next_nonblank = start
            .shifted(d_col, d_row)
            .map(|n| self.wb().cells.contains_key(&n));
        let stop_at_run_end = !cur_blank && matches!(next_nonblank, Some(true));
        let mut p = start;
        let target = loop {
            let Some(n) = p.shifted(d_col, d_row) else {
                break p;
            };
            if stop_at_run_end {
                if !self.wb().cells.contains_key(&n) {
                    break p;
                }
                p = n;
            } else {
                p = n;
                if self.wb().cells.contains_key(&p) {
                    break p;
                }
            }
        };
        self.move_pointer_to(target);
    }

    /// `@SUM` SmartIcon (icon 9). Detect the contiguous numeric run
    /// immediately above the cursor (preferred) or to its left, then
    /// write `@SUM(top..bottom)` at the cursor. Beep if neither
    /// neighbour is numeric, or if the cursor cell is protected.
    pub(super) fn dispatch_sum_smarticon(&mut self) {
        if self.is_cell_protected(self.wb().pointer) {
            self.request_beep();
            return;
        }
        let Some((from, to)) = self.detect_sum_source() else {
            self.request_beep();
            return;
        };
        let expr = format!("@SUM({}..{})", from.display_short(), to.display_short());
        let addr = self.wb().pointer;
        self.write_source_at(addr, &expr);
    }

    /// Outline SmartIcon (icon 20). Toggle a thin border on the
    /// perimeter of the active POINT highlight (or single cell at the
    /// pointer). If every perimeter edge is already set, clear them
    /// all; otherwise set them all. Journal the prior `Border` for
    /// each touched cell so undo can restore it.
    pub(super) fn dispatch_outline_smarticon(&mut self) {
        let range = if matches!(self.mode, Mode::Point) {
            self.highlight_range()
        } else {
            Range::single(self.wb().pointer)
        };
        let r = range.normalized();
        let sheet = r.start.sheet;
        let already_outlined = (r.start.row..=r.end.row).all(|row| {
            (r.start.col..=r.end.col).all(|col| {
                let addr = Address::new(sheet, col, row);
                let b = self
                    .wb()
                    .cell_borders
                    .get(&addr)
                    .copied()
                    .unwrap_or_default();
                let need_top = row == r.start.row;
                let need_bottom = row == r.end.row;
                let need_left = col == r.start.col;
                let need_right = col == r.end.col;
                (!need_top || b.top.is_some())
                    && (!need_bottom || b.bottom.is_some())
                    && (!need_left || b.left.is_some())
                    && (!need_right || b.right.is_some())
            })
        });
        let new_edge: Option<BorderEdge> = if already_outlined {
            None
        } else {
            Some(BorderEdge::default())
        };

        let mut prior: Vec<(Address, Option<Border>)> = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(sheet, col, row);
                let prev = self.wb().cell_borders.get(&addr).copied();
                let mut next = prev.unwrap_or_default();
                if row == r.start.row {
                    next.top = new_edge;
                }
                if row == r.end.row {
                    next.bottom = new_edge;
                }
                if col == r.start.col {
                    next.left = new_edge;
                }
                if col == r.end.col {
                    next.right = new_edge;
                }
                if next == prev.unwrap_or_default() {
                    continue;
                }
                prior.push((addr, prev));
                if next.is_default() {
                    self.wb_mut().cell_borders.remove(&addr);
                } else {
                    self.wb_mut().cell_borders.insert(addr, next);
                }
            }
        }
        let any_change = !prior.is_empty();
        if self.undo_enabled && any_change {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeBorder { entries: prior });
        }
        self.point = None;
        self.menu = None;
        self.mode = Mode::Ready;
        if any_change {
            self.wb_mut().dirty = true;
        }
    }

    /// `@NOW` SmartIcon (icon 45). Writes `@NOW` at the cursor; the
    /// engine handles formatting. Beeps if the cell is protected.
    pub(super) fn dispatch_today_smarticon(&mut self) {
        if self.is_cell_protected(self.wb().pointer) {
            self.request_beep();
            return;
        }
        let addr = self.wb().pointer;
        self.write_source_at(addr, "@NOW");
    }

    /// Find the @SUM source range for the SmartIcon: the contiguous
    /// run of numeric cells directly above the cursor (preferred) or
    /// directly to its left. Returns `(top_left, bottom_right)` of the
    /// range, or `None` if neither neighbour is numeric.
    pub(super) fn detect_sum_source(&self) -> Option<(Address, Address)> {
        let cursor = self.wb().pointer;
        if let Some(above) = cursor.shifted(0, -1) {
            if self.is_numeric_at(above) {
                let mut top = above;
                while let Some(prev) = top.shifted(0, -1) {
                    if !self.is_numeric_at(prev) {
                        break;
                    }
                    top = prev;
                }
                return Some((top, above));
            }
        }
        if let Some(left) = cursor.shifted(-1, 0) {
            if self.is_numeric_at(left) {
                let mut leftmost = left;
                while let Some(prev) = leftmost.shifted(-1, 0) {
                    if !self.is_numeric_at(prev) {
                        break;
                    }
                    leftmost = prev;
                }
                return Some((leftmost, left));
            }
        }
        None
    }

    /// True iff the cell at `addr` evaluates to a number — either a
    /// numeric constant or a formula whose cached value is a Number.
    /// Empty cells, labels, booleans, errors, and unevaluated formulas
    /// are all treated as non-numeric.
    pub(super) fn is_numeric_at(&self, addr: Address) -> bool {
        matches!(
            self.wb().cells.get(&addr),
            Some(CellContents::Constant(Value::Number(_)))
                | Some(CellContents::Formula {
                    cached_value: Some(Value::Number(_)),
                    ..
                })
        )
    }

    /// Auto-detect the rectangular block of contiguous non-empty cells
    /// surrounding the cursor. The block grows in two passes: first
    /// left/right along the cursor's row to fix the column band, then
    /// up/down across that band stopping at the first row where every
    /// column is blank. Returns `None` when the cursor itself is blank.
    /// The first row of the returned range is the database header.
    pub(super) fn auto_detect_database_range(&self, cursor: Address) -> Option<Range> {
        if !self.wb().cells.contains_key(&cursor) {
            return None;
        }
        let sheet = cursor.sheet;
        let mut col_lo = cursor.col;
        let mut col_hi = cursor.col;
        while col_lo > 0
            && self
                .wb()
                .cells
                .contains_key(&Address::new(sheet, col_lo - 1, cursor.row))
        {
            col_lo -= 1;
        }
        while col_hi + 1 < l123_core::address::MAX_COLS
            && self
                .wb()
                .cells
                .contains_key(&Address::new(sheet, col_hi + 1, cursor.row))
        {
            col_hi += 1;
        }
        let row_has_data = |row: u32| -> bool {
            (col_lo..=col_hi).any(|c| self.wb().cells.contains_key(&Address::new(sheet, c, row)))
        };
        let mut row_lo = cursor.row;
        let mut row_hi = cursor.row;
        while row_lo > 0 && row_has_data(row_lo - 1) {
            row_lo -= 1;
        }
        while row_hi + 1 < l123_core::address::MAX_ROWS && row_has_data(row_hi + 1) {
            row_hi += 1;
        }
        Some(Range {
            start: Address::new(sheet, col_lo, row_lo),
            end: Address::new(sheet, col_hi, row_hi),
        })
    }

    /// Sort SmartIcon (icons 31/32). Auto-detect the database around
    /// the cursor, treat the first row as a header, set the primary
    /// key to the cursor's column with `dir`, and run the sort. Beep +
    /// no-op when the cursor isn't inside a block of at least one
    /// header row plus one data row.
    pub(super) fn dispatch_sort_smarticon(&mut self, dir: SortDir) {
        let cursor = self.wb().pointer;
        let Some(block) = self.auto_detect_database_range(cursor) else {
            self.request_beep();
            return;
        };
        if block.end.row <= block.start.row {
            self.request_beep();
            return;
        }
        let data_range = Range {
            start: Address::new(block.start.sheet, block.start.col, block.start.row + 1),
            end: block.end,
        };
        self.data_sort.data_range = Some(data_range);
        self.data_sort.primary = Some((cursor.col, dir));
        self.data_sort.secondary = None;
        self.data_sort.extra = None;
        self.execute_data_sort();
    }
}
