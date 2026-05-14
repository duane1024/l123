//! Keystroke entry point and per-mode dispatch handlers.

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use l123_core::{Address, Mode};

use super::types::*;
use super::{
    adjust_file_list_view, adjust_name_list_view, adjust_sqlite_table_picker_view,
    is_retrievable_workbook, App,
};

impl App {
    pub fn handle_key(&mut self, k: KeyEvent) {
        // §4.7 / SPEC §7: Ctrl-Break aborts an in-flight long op,
        // dropping the queued work and returning to READY. This
        // takes precedence over splash/help/menus so a runaway load
        // is always escapable.
        if matches!(k.code, KeyCode::Pause)
            && k.modifiers.contains(KeyModifiers::CONTROL)
            && self.cancel_pending_async_op()
        {
            return;
        }
        // Startup splash consumes the first keystroke and drops to
        // READY without dispatching — matches the 1-2-3 R3.4a behavior
        // where any key clears the welcome screen.
        if self.splash.is_some() {
            self.splash = None;
            return;
        }
        // F1 HELP overlay sits on top of every other state.
        if self.help.is_some() {
            self.handle_key_help(k);
            return;
        }
        // F1 from any mode opens the help overlay. The current mode is
        // saved on `HelpState` so Esc returns the user to where they
        // left off (POINT, MENU, …) without losing their work.
        if matches!(k.code, KeyCode::F(1)) {
            self.open_help();
            return;
        }
        // F3 NAMES overlay takes precedence over everything else — it
        // sits on top of the underlying POINT / prompt state and owns
        // the keyboard in NAMES mode until Esc or Enter.
        if self.name_list.is_some() {
            self.handle_key_names(k);
            return;
        }
        // `/File Import Sqlite`'s table picker (v0.4 follow-up) shares
        // the NAMES-style overlay shape; it owns the keyboard until
        // Esc cancels or Enter loads the chosen table.
        if self.sqlite_table_picker.is_some() {
            self.handle_key_sqlite_table_picker(k);
            return;
        }
        // /File List overlay takes precedence when active — it owns the
        // keyboard in FILES mode.
        if self.file_list.is_some() {
            self.handle_key_file_list(k);
            return;
        }
        // Save-confirm submenu runs before the prompt/menu dispatcher:
        // it overrides line 2 with its own three-item picker and owns
        // the keyboard until the user commits or cancels.
        if self.save_confirm.is_some() {
            self.handle_key_save_confirm(k);
            return;
        }
        // `{MENUBRANCH}` / `{MENUCALL}` overlay sits above the
        // built-in menu dispatcher: it owns the keyboard until the
        // user picks an item or cancels.
        if self.custom_menu.is_some() {
            self.handle_key_custom_menu(k);
            return;
        }
        // STEP-mode pause owns Space and Esc directly so they don't
        // leak into the underlying mode (e.g. a Space in READY would
        // otherwise start a label entry). Other keys flow through
        // normally — the user can navigate / inspect the workbook
        // between steps.
        if !self.macro_pumping
            && matches!(
                self.macro_state.as_ref().and_then(|s| s.suspend.as_ref()),
                Some(MacroSuspend::StepPause)
            )
        {
            match k.code {
                KeyCode::Char(' ') => {
                    if let Some(s) = self.macro_state.as_mut() {
                        s.suspend = None;
                        s.step_advance = true;
                    }
                    self.pump_macro();
                    return;
                }
                KeyCode::Esc => {
                    self.macro_state = None;
                    return;
                }
                _ => {}
            }
        }
        // Erase-confirm submenu has the same precedence: shown after
        // the `/File Erase` filename prompt commits, owns the keyboard
        // until the user picks No or Yes.
        if self.erase_confirm.is_some() {
            self.handle_key_erase_confirm(k);
            return;
        }
        // A command argument prompt takes precedence over the mode-based
        // dispatcher — it intercepts keystrokes while the mode indicator
        // continues to reflect the underlying state (MENU/POINT/etc).
        if self.prompt.is_some() {
            self.handle_key_prompt(k);
            return;
        }
        match self.mode {
            Mode::Ready => self.handle_key_ready(k),
            Mode::Label | Mode::Value | Mode::Edit => self.handle_key_entry(k),
            Mode::Menu => self.handle_key_menu(k),
            Mode::Point => self.handle_key_point(k),
            Mode::Find => self.handle_key_find(k),
            Mode::Graph => self.handle_key_graph(k),
            Mode::Stat => self.handle_key_stat(k),
            Mode::Error => self.handle_key_error(k),
            _ => {}
        }
        // Record into the learn buffer if Alt-F5 is armed. Skip
        // synthetic keystrokes the macro pump generates (we don't
        // want a macro running under Learn to log itself), and skip
        // the Alt-F5 toggle itself so it doesn't end up in the
        // recorded source.
        if self.learn_recording
            && !self.macro_pumping
            && !(matches!(k.code, KeyCode::F(5)) && k.modifiers.contains(KeyModifiers::ALT))
        {
            self.record_keystroke(&k);
        }
        // Macro-pause resume hook: a `{?}` directive parks the
        // interpreter in `WaitEnter`. The user's next Enter is the
        // signal to resume — it still flows through the dispatcher
        // first (so an in-progress entry/prompt commits normally),
        // and afterwards we kick the pump.
        if !self.macro_pumping
            && matches!(k.code, KeyCode::Enter)
            && matches!(
                self.macro_state.as_ref().and_then(|s| s.suspend.as_ref()),
                Some(MacroSuspend::WaitEnter)
            )
        {
            if let Some(s) = self.macro_state.as_mut() {
                s.suspend = None;
            }
            self.pump_macro();
        }
    }
    fn handle_key_custom_menu(&mut self, k: KeyEvent) {
        let len = self
            .custom_menu
            .as_ref()
            .map(|m| m.items.len())
            .unwrap_or(0);
        if len == 0 {
            self.finish_custom_menu(None);
            return;
        }
        match k.code {
            KeyCode::Esc => self.finish_custom_menu(None),
            KeyCode::Enter => {
                let idx = self.custom_menu.as_ref().map(|m| m.highlight).unwrap_or(0);
                self.finish_custom_menu(Some(idx));
            }
            KeyCode::Left => {
                if let Some(m) = self.custom_menu.as_mut() {
                    m.highlight = m.highlight.checked_sub(1).unwrap_or(len - 1);
                }
            }
            KeyCode::Right => {
                if let Some(m) = self.custom_menu.as_mut() {
                    m.highlight = (m.highlight + 1) % len;
                }
            }
            KeyCode::Home => {
                if let Some(m) = self.custom_menu.as_mut() {
                    m.highlight = 0;
                }
            }
            KeyCode::End => {
                if let Some(m) = self.custom_menu.as_mut() {
                    m.highlight = len - 1;
                }
            }
            KeyCode::Char(c) => {
                let needle = c.to_ascii_lowercase();
                let pick = self.custom_menu.as_ref().and_then(|m| {
                    m.items.iter().position(|it| {
                        it.name
                            .chars()
                            .next()
                            .map(|f| f.to_ascii_lowercase() == needle)
                            .unwrap_or(false)
                    })
                });
                if let Some(idx) = pick {
                    self.finish_custom_menu(Some(idx));
                }
            }
            _ => {}
        }
    }
    fn handle_key_error(&mut self, k: KeyEvent) {
        if matches!(k.code, KeyCode::Esc | KeyCode::Enter) {
            self.error_message = None;
            self.mode = Mode::Ready;
        }
    }
    fn handle_key_stat(&mut self, k: KeyEvent) {
        // Any key dismisses the status panel, same shape as
        // handle_key_graph. 1-2-3 R3.4a used Esc specifically, but a
        // generic "any key" dismissal matches the splash screen's
        // behavior and is friendlier.
        let _ = k;
        self.mode = Mode::Ready;
    }
    fn handle_key_graph(&mut self, k: KeyEvent) {
        if matches!(
            k.code,
            KeyCode::Esc | KeyCode::Enter | KeyCode::Char(' ') | KeyCode::F(10)
        ) {
            self.graph_view = None;
            self.mode = Mode::Ready;
        }
    }
    fn handle_key_ready(&mut self, k: KeyEvent) {
        let ctrl = k.modifiers.contains(KeyModifiers::CONTROL);
        // Ctrl-End arms the FILE-navigation prefix. The next key
        // reads the prefix: Ctrl-PgUp/PgDn become file rotations.
        if ctrl && matches!(k.code, KeyCode::End) {
            self.file_nav_pending = true;
            return;
        }
        let file_nav = self.file_nav_pending;
        // Any keystroke other than the file-nav follow-through clears
        // the prefix — so a typo drops you back into normal nav.
        if file_nav {
            self.file_nav_pending = false;
        }
        match k.code {
            KeyCode::Up => self.move_pointer(0, -1),
            KeyCode::Down => self.move_pointer(0, 1),
            KeyCode::Left => self.move_pointer(-1, 0),
            KeyCode::Right => self.move_pointer(1, 0),
            KeyCode::Home => {
                let sheet = self.wb().pointer.sheet;
                let wb = self.wb_mut();
                wb.pointer = Address::new(sheet, 0, 0);
                wb.viewport_col_offset = 0;
                wb.viewport_row_offset = 0;
            }
            KeyCode::PageDown if ctrl && file_nav => self.rotate_files(1),
            KeyCode::PageUp if ctrl && file_nav => self.rotate_files(-1),
            KeyCode::PageDown if ctrl => self.move_sheet(1),
            KeyCode::PageUp if ctrl => self.move_sheet(-1),
            KeyCode::PageDown => self.move_pointer(0, 20),
            KeyCode::PageUp => self.move_pointer(0, -20),
            // Alt-F2 STEP: toggle single-step mode. Active macros
            // pause before each action; the user advances with
            // Space. Idle when no macro is running.
            KeyCode::F(2) if k.modifiers.contains(KeyModifiers::ALT) => {
                self.step_mode = !self.step_mode;
            }
            KeyCode::F(2) => {
                if self.is_cell_protected(self.wb().pointer) {
                    self.request_beep();
                } else {
                    self.begin_edit();
                }
            }
            // Alt-F3 RUN: pop up the NAMES picker; Enter on a name
            // runs the macro stored at that range. Same overlay as
            // F3, just with a "run on commit" intent.
            KeyCode::F(3) if k.modifiers.contains(KeyModifiers::ALT) => {
                self.open_name_list(NameListOrigin::RunMacro);
            }
            KeyCode::F(4) if k.modifiers.contains(KeyModifiers::ALT) => self.undo(),
            // Alt-F5 LEARN: toggle keystroke recording. Requires a
            // /Worksheet Learn Range to have been set first.
            KeyCode::F(5) if k.modifiers.contains(KeyModifiers::ALT) => {
                self.toggle_learn_recording();
            }
            // Alt+letter runs the macro stored at the named range
            // `\<letter>` (case-insensitive). Per SPEC §18 / PLAN.md
            // M9: the Lotus user-macro launch convention.
            KeyCode::Char(c)
                if k.modifiers.contains(KeyModifiers::ALT) && c.is_ascii_alphabetic() =>
            {
                let name = format!("\\{}", c.to_ascii_lowercase());
                self.run_named_macro(&name);
            }
            KeyCode::F(5) => self.begin_goto_prompt(),
            KeyCode::F(9) => self.do_recalc(),
            KeyCode::F(10) => self.enter_graph_view(),
            KeyCode::Char('/') => self.open_menu(),
            KeyCode::Char(':') => self.open_wysiwyg_menu(),
            KeyCode::Esc if self.input_range.is_some() => self.exit_input_mode(),
            KeyCode::Char(c) => {
                if self.is_cell_protected(self.wb().pointer) {
                    self.request_beep();
                } else {
                    self.begin_entry(c);
                }
            }
            _ => {}
        }
    }
    fn handle_key_entry(&mut self, k: KeyEvent) {
        let in_edit = matches!(self.entry.as_ref().map(|e| e.kind), Some(EntryKind::Edit));
        match k.code {
            KeyCode::Enter => self.commit_entry(),
            KeyCode::Esc => self.cancel_entry(),
            // Up/Down always commit-and-move (no in-buffer vertical
            // navigation in either initial entry or EDIT).
            KeyCode::Up => {
                self.commit_entry();
                self.move_pointer(0, -1);
            }
            KeyCode::Down => {
                self.commit_entry();
                self.move_pointer(0, 1);
            }
            // Left/Right: in LABEL/VALUE, commit-and-move (Lotus
            // tutorial §2.4 fast-entry idiom). In EDIT, move the cursor
            // within the buffer.
            KeyCode::Left if in_edit => {
                self.move_entry_cursor_left();
            }
            KeyCode::Left => {
                self.commit_entry();
                self.move_pointer(-1, 0);
            }
            KeyCode::Right if in_edit => {
                self.move_entry_cursor_right();
            }
            KeyCode::Right => {
                self.commit_entry();
                self.move_pointer(1, 0);
            }
            // Tab always commits-and-moves right; an in-buffer Tab has
            // no Lotus precedent.
            KeyCode::Tab => {
                self.commit_entry();
                self.move_pointer(1, 0);
            }
            // Home/End move the cursor in all three entry modes.
            KeyCode::Home => {
                if let Some(e) = self.entry.as_mut() {
                    e.cursor = 0;
                }
            }
            KeyCode::End => {
                if let Some(e) = self.entry.as_mut() {
                    e.cursor = e.buffer.len();
                }
            }
            KeyCode::Backspace => self.entry_backspace(),
            KeyCode::Delete => self.entry_delete(),
            // F2 mid-entry: promote LABEL/VALUE to EDIT preserving the
            // buffer; cursor stays where it is (typically at the end of
            // what the user just typed).
            KeyCode::F(2) => self.promote_entry_to_edit(),
            KeyCode::Char(c) => self.entry_insert_char(c),
            _ => {}
        }
    }
    fn handle_key_menu(&mut self, k: KeyEvent) {
        let Some(state) = self.menu.as_mut() else {
            self.mode = Mode::Ready;
            return;
        };
        match k.code {
            KeyCode::Esc => {
                if state.path.is_empty() {
                    self.close_menu();
                } else {
                    state.path.pop();
                    state.highlight = 0;
                    state.message = None;
                }
            }
            KeyCode::Left => {
                let len = state.level().len();
                if len > 0 {
                    state.highlight = (state.highlight + len - 1) % len;
                    state.message = None;
                }
            }
            KeyCode::Right | KeyCode::Tab => {
                let len = state.level().len();
                if len > 0 {
                    state.highlight = (state.highlight + 1) % len;
                    state.message = None;
                }
            }
            KeyCode::Home => {
                state.highlight = 0;
                state.message = None;
            }
            KeyCode::End => {
                let len = state.level().len();
                state.highlight = len.saturating_sub(1);
                state.message = None;
            }
            KeyCode::Enter => self.descend_highlighted(),
            KeyCode::Char(c) => self.descend_by_letter(c),
            _ => {}
        }
    }
    fn handle_key_file_list(&mut self, k: KeyEvent) {
        let Some(fl) = self.file_list.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => {
                self.file_list = None;
                self.mode = Mode::Ready;
            }
            // Vertical navigation is the primary axis for the overlay;
            // Left/Right are kept as aliases for muscle memory.
            KeyCode::Up | KeyCode::Left => {
                if fl.highlight > 0 {
                    fl.highlight -= 1;
                }
            }
            KeyCode::Down | KeyCode::Right => {
                if fl.highlight + 1 < fl.entries.len() {
                    fl.highlight += 1;
                }
            }
            KeyCode::PageUp => {
                fl.highlight = fl.highlight.saturating_sub(FILE_LIST_PAGE_SIZE);
            }
            KeyCode::PageDown => {
                if !fl.entries.is_empty() {
                    fl.highlight = (fl.highlight + FILE_LIST_PAGE_SIZE).min(fl.entries.len() - 1);
                }
            }
            KeyCode::Home => fl.highlight = 0,
            KeyCode::End => {
                if !fl.entries.is_empty() {
                    fl.highlight = fl.entries.len() - 1;
                }
            }
            KeyCode::Enter => {
                let Some(fl) = self.file_list.take() else {
                    return;
                };
                match fl.kind {
                    FileListKind::Worksheet => {
                        if let Some(path) = fl.entries.get(fl.highlight).cloned() {
                            self.load_workbook_from(path);
                        } else {
                            self.mode = Mode::Ready;
                        }
                    }
                    FileListKind::Active => {
                        // Already the active file — just dismiss.
                        self.mode = Mode::Ready;
                    }
                    FileListKind::Other => {
                        if let Some(path) = fl.entries.get(fl.highlight).cloned() {
                            if is_retrievable_workbook(&path) {
                                self.retrieve_by_extension(path);
                            } else {
                                self.mode = Mode::Ready;
                            }
                        } else {
                            self.mode = Mode::Ready;
                        }
                    }
                }
            }
            _ => return,
        }
        if let Some(fl) = self.file_list.as_mut() {
            adjust_file_list_view(fl);
        }
    }
    fn handle_key_help(&mut self, k: KeyEvent) {
        let Some(state) = self.help.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => self.close_help(),
            KeyCode::Up => state.focus_up(),
            KeyCode::Down => state.focus_down(),
            KeyCode::Left => state.focus_left(),
            KeyCode::Right => state.focus_right(),
            KeyCode::Enter => {
                if let Some(link) = state.page.links.get(state.focus) {
                    let target = link.target.clone();
                    state.follow(&target);
                }
            }
            KeyCode::Backspace => {
                state.pop();
            }
            _ => {}
        }
    }
    /// Keys while the `/File Import Sqlite` table picker owns the
    /// overlay. Mirrors `handle_key_names` but committing dispatches
    /// `queue_file_import_sqlite` with the highlighted table.
    fn handle_key_sqlite_table_picker(&mut self, k: KeyEvent) {
        let Some(p) = self.sqlite_table_picker.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => {
                self.sqlite_table_picker = None;
                self.mode = Mode::Ready;
                return;
            }
            KeyCode::Up | KeyCode::Left => {
                if p.highlight > 0 {
                    p.highlight -= 1;
                }
            }
            KeyCode::Down | KeyCode::Right => {
                if p.highlight + 1 < p.tables.len() {
                    p.highlight += 1;
                }
            }
            KeyCode::PageUp => {
                p.highlight = p.highlight.saturating_sub(SQLITE_TABLE_PICKER_PAGE_SIZE);
            }
            KeyCode::PageDown => {
                if !p.tables.is_empty() {
                    p.highlight = (p.highlight + SQLITE_TABLE_PICKER_PAGE_SIZE)
                        .min(p.tables.len() - 1);
                }
            }
            KeyCode::Home => p.highlight = 0,
            KeyCode::End => {
                if !p.tables.is_empty() {
                    p.highlight = p.tables.len() - 1;
                }
            }
            KeyCode::Enter => {
                let Some(state) = self.sqlite_table_picker.take() else {
                    return;
                };
                let Some(table) = state.tables.get(state.highlight).cloned() else {
                    self.mode = Mode::Ready;
                    return;
                };
                self.queue_file_import_sqlite(state.path, table);
                return;
            }
            _ => return,
        }
        if let Some(p) = self.sqlite_table_picker.as_mut() {
            adjust_sqlite_table_picker_view(p);
        }
    }

    fn handle_key_names(&mut self, k: KeyEvent) {
        let Some(nl) = self.name_list.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => self.dismiss_name_list(),
            KeyCode::Up | KeyCode::Left => {
                if nl.highlight > 0 {
                    nl.highlight -= 1;
                }
            }
            KeyCode::Down | KeyCode::Right => {
                if nl.highlight + 1 < nl.entries.len() {
                    nl.highlight += 1;
                }
            }
            KeyCode::PageUp => {
                nl.highlight = nl.highlight.saturating_sub(NAME_LIST_PAGE_SIZE);
            }
            KeyCode::PageDown => {
                if !nl.entries.is_empty() {
                    nl.highlight = (nl.highlight + NAME_LIST_PAGE_SIZE).min(nl.entries.len() - 1);
                }
            }
            KeyCode::Home => nl.highlight = 0,
            KeyCode::End => {
                if !nl.entries.is_empty() {
                    nl.highlight = nl.entries.len() - 1;
                }
            }
            KeyCode::Enter => self.commit_name_list(),
            _ => return,
        }
        if let Some(nl) = self.name_list.as_mut() {
            adjust_name_list_view(nl);
        }
    }
    fn handle_key_erase_confirm(&mut self, k: KeyEvent) {
        let Some(ec) = self.erase_confirm.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => {
                self.erase_confirm = None;
                self.mode = Mode::Ready;
            }
            KeyCode::Left if ec.highlight > 0 => ec.highlight -= 1,
            KeyCode::Right if ec.highlight + 1 < FILE_ERASE_CONFIRM_ITEMS.len() => {
                ec.highlight += 1;
            }
            KeyCode::Home => ec.highlight = 0,
            KeyCode::End => ec.highlight = FILE_ERASE_CONFIRM_ITEMS.len() - 1,
            KeyCode::Enter => {
                let choice = ec.highlight;
                self.commit_erase_confirm(choice);
            }
            KeyCode::Char(c) => {
                let upper = c.to_ascii_uppercase();
                if let Some(idx) = FILE_ERASE_CONFIRM_ITEMS
                    .iter()
                    .position(|(name, _)| name.starts_with(upper))
                {
                    self.commit_erase_confirm(idx);
                }
            }
            _ => {}
        }
    }
    fn handle_key_save_confirm(&mut self, k: KeyEvent) {
        let Some(sc) = self.save_confirm.as_mut() else {
            return;
        };
        match k.code {
            KeyCode::Esc => {
                self.save_confirm = None;
                self.mode = Mode::Ready;
            }
            KeyCode::Left if sc.highlight > 0 => sc.highlight -= 1,
            KeyCode::Right if sc.highlight + 1 < SAVE_CONFIRM_ITEMS.len() => {
                sc.highlight += 1;
            }
            KeyCode::Home => sc.highlight = 0,
            KeyCode::End => sc.highlight = SAVE_CONFIRM_ITEMS.len() - 1,
            KeyCode::Enter => {
                let choice = sc.highlight;
                self.commit_save_confirm(choice);
            }
            KeyCode::Char(c) => {
                // Letter accelerators: C(ancel), R(eplace), B(ackup).
                let upper = c.to_ascii_uppercase();
                if let Some(idx) = SAVE_CONFIRM_ITEMS
                    .iter()
                    .position(|(name, _)| name.starts_with(upper))
                {
                    self.commit_save_confirm(idx);
                }
            }
            _ => {}
        }
    }
    fn handle_key_point(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Up => self.move_pointer(0, -1),
            KeyCode::Down => self.move_pointer(0, 1),
            KeyCode::Left => self.move_pointer(-1, 0),
            KeyCode::Right => self.move_pointer(1, 0),
            KeyCode::Home => {
                self.wb_mut().pointer = Address::A1;
                self.scroll_into_view();
            }
            KeyCode::PageDown => self.move_pointer(0, 20),
            KeyCode::PageUp => self.move_pointer(0, -20),
            KeyCode::Enter => self.commit_point(),
            KeyCode::Esc => self.esc_in_point(),
            KeyCode::F(3) => self.open_name_list(NameListOrigin::Point),
            KeyCode::Backspace => {
                if let Some(ps) = self.point.as_mut() {
                    ps.typed.pop();
                }
            }
            // `.` is the anchor-cycle key when no typed range is in
            // progress, but once the user has started typing a range it
            // becomes the literal range separator (`A1..D5`).
            KeyCode::Char('.') => {
                let typing = self
                    .point
                    .as_ref()
                    .map(|p| !p.typed.is_empty())
                    .unwrap_or(false);
                if typing {
                    if let Some(ps) = self.point.as_mut() {
                        ps.typed.push('.');
                    }
                } else {
                    self.period_in_point();
                }
            }
            // Any other address-like char extends (or starts) the typed
            // range buffer. `:` only makes sense after a sheet letter,
            // so we ignore a leading `:`. `_` is accepted to support
            // typed range names (Lotus permits `_` inside names). `,`
            // is the multi-range separator (`A1..B2,C3..D4`).
            KeyCode::Char(c) if c.is_ascii_alphanumeric() || c == ':' || c == '_' || c == ',' => {
                if let Some(ps) = self.point.as_mut() {
                    if !((c == ':' || c == ',') && ps.typed.is_empty()) {
                        ps.typed.push(c);
                    }
                }
            }
            _ => {}
        }
    }
    fn handle_key_find(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Enter => {
                if let Some(s) = self.search.as_mut() {
                    if !s.matches.is_empty() {
                        s.cursor = (s.cursor + 1) % s.matches.len();
                        let next = s.matches[s.cursor];
                        self.wb_mut().pointer = next;
                        self.scroll_into_view();
                    }
                }
            }
            KeyCode::Esc => {
                self.search = None;
                self.mode = Mode::Ready;
            }
            _ => {}
        }
    }
    fn handle_key_prompt(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Enter => self.commit_prompt(),
            KeyCode::Esc => self.cancel_prompt(),
            KeyCode::F(3) => self.open_name_list_from_prompt(),
            KeyCode::Backspace => {
                if let Some(p) = self.prompt.as_mut() {
                    if p.fresh {
                        p.buffer.clear();
                        p.fresh = false;
                    } else {
                        p.buffer.pop();
                    }
                }
            }
            KeyCode::Char(c) => {
                if let Some(p) = self.prompt.as_mut() {
                    if !p.next.accepts_char(c) {
                        return;
                    }
                    if p.fresh {
                        p.buffer.clear();
                        p.fresh = false;
                    }
                    p.buffer.push(c);
                }
            }
            _ => {}
        }
    }
}
