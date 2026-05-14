//! Macro execution: lex + interpret 1-2-3 macro source, the call stack,
//! the parked-suspend points (`{?}`, `{GETLABEL}`, `{MENUBRANCH}`,
//! STEP-mode), `/Worksheet Learn` recording, and the post-key resume
//! hook driven by `handle_key`.

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use l123_core::{Address, CellContents, LabelPrefix, Mode, Range, Value};
use l123_engine::Engine;
use l123_macro::{lex as lex_macro, lex_actions as lex_macro_actions, MacroAction, MacroKey};

use super::types::*;
use super::App;

/// One row down from `pc`, or `None` if we hit the bottom of the
/// sheet. Macros run column-major, so this is the natural "next
/// line" rule.
fn next_macro_pc(pc: Address) -> Option<Address> {
    pc.row
        .checked_add(1)
        .filter(|r| *r < 8192)
        .map(|r| Address::new(pc.sheet, pc.col, r))
}

/// Reverse of [`macro_key_to_event`]: given a key the user just
/// pressed, return its representation in macro source form (so the
/// Learn recorder can replay it later). `None` for keys that have
/// no macro-source equivalent (modifier-only, unmapped function
/// keys with Alt, ...).
fn key_event_to_macro_source(k: &KeyEvent) -> Option<String> {
    let alt = k.modifiers.contains(KeyModifiers::ALT);
    let ctrl = k.modifiers.contains(KeyModifiers::CONTROL);
    // Alt-prefixed keys are macro launchers / system actions; they
    // shouldn't end up in a Learn recording as themselves.
    if alt {
        return None;
    }
    Some(match k.code {
        KeyCode::Enter => "~".to_string(),
        KeyCode::Up => "{UP}".to_string(),
        KeyCode::Down => "{DOWN}".to_string(),
        KeyCode::Left if ctrl => "{BIGLEFT}".to_string(),
        KeyCode::Right if ctrl => "{BIGRIGHT}".to_string(),
        KeyCode::Left => "{LEFT}".to_string(),
        KeyCode::Right => "{RIGHT}".to_string(),
        KeyCode::Home => "{HOME}".to_string(),
        KeyCode::End => "{END}".to_string(),
        KeyCode::PageUp => "{PGUP}".to_string(),
        KeyCode::PageDown => "{PGDN}".to_string(),
        KeyCode::Esc => "{ESC}".to_string(),
        KeyCode::Backspace => "{BS}".to_string(),
        KeyCode::Delete => "{DEL}".to_string(),
        KeyCode::Insert => "{INS}".to_string(),
        KeyCode::Tab => "{TAB}".to_string(),
        KeyCode::F(n) => match n {
            1 => "{HELP}".to_string(),
            2 => "{EDIT}".to_string(),
            3 => "{NAME}".to_string(),
            4 => "{ABS}".to_string(),
            5 => "{GOTO}".to_string(),
            6 => "{WINDOW}".to_string(),
            7 => "{QUERY}".to_string(),
            8 => "{TABLE}".to_string(),
            9 => "{CALC}".to_string(),
            10 => "{GRAPH}".to_string(),
            _ => return None,
        },
        KeyCode::Char(c) => match c {
            '~' => "{TILDE}".to_string(),
            '{' => "{LBRACE}".to_string(),
            '}' => "{RBRACE}".to_string(),
            other => other.to_string(),
        },
        _ => return None,
    })
}

/// Translate a [`MacroKey`] from `l123-macro` into the crossterm
/// [`KeyEvent`] the dispatcher already handles. The mapping is one-
/// to-one: a macro pressing `{DOWN}` should be exactly the same as
/// the user pressing the Down arrow.
fn macro_key_to_event(key: MacroKey) -> KeyEvent {
    let none = KeyModifiers::NONE;
    let ctrl = KeyModifiers::CONTROL;
    match key {
        MacroKey::Char(c) => KeyEvent::new(KeyCode::Char(c), none),
        MacroKey::Enter => KeyEvent::new(KeyCode::Enter, none),
        MacroKey::Up => KeyEvent::new(KeyCode::Up, none),
        MacroKey::Down => KeyEvent::new(KeyCode::Down, none),
        MacroKey::Left => KeyEvent::new(KeyCode::Left, none),
        MacroKey::Right => KeyEvent::new(KeyCode::Right, none),
        MacroKey::Home => KeyEvent::new(KeyCode::Home, none),
        MacroKey::End => KeyEvent::new(KeyCode::End, none),
        MacroKey::PageUp => KeyEvent::new(KeyCode::PageUp, none),
        MacroKey::PageDown => KeyEvent::new(KeyCode::PageDown, none),
        MacroKey::BigLeft => KeyEvent::new(KeyCode::Left, ctrl),
        MacroKey::BigRight => KeyEvent::new(KeyCode::Right, ctrl),
        MacroKey::Escape => KeyEvent::new(KeyCode::Esc, none),
        MacroKey::Backspace => KeyEvent::new(KeyCode::Backspace, none),
        MacroKey::Delete => KeyEvent::new(KeyCode::Delete, none),
        MacroKey::Insert => KeyEvent::new(KeyCode::Insert, none),
        MacroKey::Tab => KeyEvent::new(KeyCode::Tab, none),
        MacroKey::Function(n) => KeyEvent::new(KeyCode::F(n), none),
    }
}

impl App {
    /// Lex `text` as L123 macro source and feed each resulting key
    /// through [`Self::handle_key`] in order, exactly as if the user
    /// had typed the same sequence at the keyboard. This is the
    /// foundation that all later macro features build on (named-range
    /// invocation, `\A..\Z`, `{BRANCH}`, Learn replay, ...).
    ///
    /// On a malformed source string the macro halts at the bad token
    /// and the error surfaces in the standard ERROR-mode panel.
    pub fn run_macro_text(&mut self, text: &str) {
        let keys = match lex_macro(text) {
            Ok(k) => k,
            Err(e) => {
                self.set_error(format!("{e}"));
                return;
            }
        };
        for key in keys {
            self.handle_key(macro_key_to_event(key));
        }
    }

    /// Look up a named range (case-insensitive) and run the macro
    /// stored at its start cell. Returns `false` when the name is
    /// undefined — Lotus would beep here, we silently no-op.
    pub(super) fn run_named_macro(&mut self, name: &str) -> bool {
        let key = name.to_ascii_lowercase();
        let Some(range) = self.wb().named_ranges.get(&key).copied() else {
            return false;
        };
        self.run_macro_at(range.start);
        true
    }

    /// Read the source of one macro line at `addr`. Label cells
    /// return their `text`; constant/formula cells return
    /// `source_form()` so a stray number in the macro range types as
    /// digits rather than aborting the read. Empty / off-sheet cells
    /// return `None`, which the interpreter treats as end-of-frame.
    pub(super) fn read_macro_line(&self, addr: Address) -> Option<String> {
        match self.wb().cells.get(&addr)? {
            CellContents::Label { text, .. } => Some(text.clone()),
            other if !other.is_empty() => Some(other.source_form()),
            _ => None,
        }
    }

    /// Resolve a macro `loc` argument to an absolute [`Address`].
    /// Names take precedence over raw cell addresses (matches Lotus).
    /// Returns `None` if the loc is neither a known name nor a
    /// parseable address.
    pub(super) fn resolve_macro_loc(&self, loc: &str) -> Option<Address> {
        let trimmed = loc.trim();
        if trimmed.is_empty() {
            return None;
        }
        let key = trimmed.to_ascii_lowercase();
        if let Some(range) = self.wb().named_ranges.get(&key) {
            return Some(range.start);
        }
        Address::parse(trimmed).ok()
    }

    /// Evaluate `expr` as an `{IF}` condition. Returns `true` for
    /// non-zero numbers, `true` for booleans/text, `false` for
    /// zero/empty/error. Implementation: stash the formula in a
    /// scratch cell at the bottom-right of the current sheet, recalc,
    /// read the value, then clear the scratch cell.
    pub(super) fn eval_macro_condition(&mut self, expr: &str) -> bool {
        let sheet = self.wb().pointer.sheet;
        let scratch = Address::new(sheet, 255, 8190);
        let formula = format!("={}", expr.trim());
        let truthy = if self
            .wb_mut()
            .engine
            .set_user_input(scratch, &formula)
            .is_ok()
        {
            self.wb_mut().engine.recalc();
            match self.wb().engine.get_cell(scratch) {
                Ok(view) => match view.value {
                    Value::Number(n) => n != 0.0,
                    Value::Bool(b) => b,
                    Value::Text(_) => true,
                    _ => false,
                },
                Err(_) => false,
            }
        } else {
            false
        };
        let _ = self.wb_mut().engine.clear_cell(scratch);
        self.wb_mut().engine.recalc();
        truthy
    }

    /// Run a macro starting at `start`. Each cell is one logical
    /// line; the interpreter walks down the column, lexing and
    /// executing per-line actions. Sets up [`MacroState`] then
    /// calls [`pump_macro`] which runs to completion or to the
    /// first suspension point.
    pub(super) fn run_macro_at(&mut self, start: Address) {
        self.macro_state = Some(MacroState {
            frames: vec![MacroFrame::starting_at(start)],
            steps: 0,
            suspend: None,
            step_advance: false,
        });
        self.pump_macro();
    }

    /// Drive the active macro forward until it suspends or finishes.
    /// Synthetic keystrokes flow back through [`handle_key`]; the
    /// `macro_pumping` re-entrancy guard keeps that from re-entering
    /// the pump.
    pub(super) fn pump_macro(&mut self) {
        if self.macro_pumping {
            return;
        }
        self.macro_pumping = true;
        loop {
            // Suspended? Idle until handle_key clears the suspend
            // and re-pumps.
            let suspended = self
                .macro_state
                .as_ref()
                .map(|s| s.suspend.is_some())
                .unwrap_or(true);
            if suspended {
                break;
            }
            // Frame stack empty → macro done.
            let empty = self
                .macro_state
                .as_ref()
                .map(|s| s.frames.is_empty())
                .unwrap_or(true);
            if empty {
                self.macro_state = None;
                break;
            }
            if !self.step_macro() {
                break;
            }
        }
        self.macro_pumping = false;
    }

    /// Execute a single macro action. Returns `false` to stop the
    /// pump (e.g. on error / `{QUIT}` / suspend). Helper so the
    /// outer loop in [`pump_macro`] stays small.
    pub(super) fn step_macro(&mut self) -> bool {
        // Bump step counter and runaway guard.
        let steps = match self.macro_state.as_mut() {
            Some(s) => {
                s.steps += 1;
                s.steps
            }
            None => return false,
        };
        if steps > MAX_MACRO_STEPS {
            self.set_error("macro: step limit exceeded".to_string());
            self.macro_state = None;
            return false;
        }

        // Re-fill `remaining` from the next cell whenever empty.
        let needs_fill = self
            .macro_state
            .as_ref()
            .and_then(|s| s.frames.last())
            .map(|f| f.remaining.is_empty())
            .unwrap_or(false);
        if needs_fill {
            let pc = self.macro_state.as_ref().unwrap().frames.last().unwrap().pc;
            let line = self.read_macro_line(pc);
            let Some(line) = line else {
                self.macro_state.as_mut().unwrap().frames.pop();
                return true;
            };
            let actions = match lex_macro_actions(&line) {
                Ok(a) => a,
                Err(e) => {
                    self.set_error(format!("{e}"));
                    self.macro_state = None;
                    return false;
                }
            };
            let frame = self
                .macro_state
                .as_mut()
                .unwrap()
                .frames
                .last_mut()
                .unwrap();
            frame.remaining = actions.into_iter().collect();
            frame.pc = next_macro_pc(pc).unwrap_or(pc);
        }

        // STEP gate: pause before each action when single-step mode
        // is on, unless `step_advance` was set by the user pressing
        // Space (which fires exactly one action then re-pauses).
        if self.step_mode {
            let advance = self
                .macro_state
                .as_ref()
                .map(|s| s.step_advance)
                .unwrap_or(false);
            if !advance {
                if let Some(s) = self.macro_state.as_mut() {
                    s.suspend = Some(MacroSuspend::StepPause);
                }
                return false;
            }
            if let Some(s) = self.macro_state.as_mut() {
                s.step_advance = false;
            }
        }

        // Pop the next action from the top frame.
        let action = match self
            .macro_state
            .as_mut()
            .and_then(|s| s.frames.last_mut())
            .and_then(|f| f.remaining.pop_front())
        {
            Some(a) => a,
            None => return true,
        };

        match action {
            MacroAction::Key(k) => {
                self.handle_key(macro_key_to_event(k));
            }
            MacroAction::Branch(loc) => {
                let Some(addr) = self.resolve_macro_loc(&loc) else {
                    self.set_error(format!("macro: bad branch loc `{loc}`"));
                    self.macro_state = None;
                    return false;
                };
                if let Some(top) = self.macro_state.as_mut().and_then(|s| s.frames.last_mut()) {
                    top.pc = addr;
                    top.remaining.clear();
                }
            }
            MacroAction::Quit => {
                self.macro_state = None;
                return false;
            }
            MacroAction::Return => {
                if let Some(s) = self.macro_state.as_mut() {
                    s.frames.pop();
                }
            }
            MacroAction::If(expr) => {
                let truthy = self.eval_macro_condition(&expr);
                if !truthy {
                    if let Some(top) = self.macro_state.as_mut().and_then(|s| s.frames.last_mut()) {
                        top.remaining.clear();
                    }
                }
            }
            MacroAction::Subroutine { loc, args: _ } => {
                let Some(addr) = self.resolve_macro_loc(&loc) else {
                    self.set_error(format!("macro: bad subroutine loc `{loc}`"));
                    self.macro_state = None;
                    return false;
                };
                if let Some(s) = self.macro_state.as_mut() {
                    if s.frames.len() >= 64 {
                        self.set_error("macro: call stack overflow".to_string());
                        self.macro_state = None;
                        return false;
                    }
                    s.frames.push(MacroFrame::starting_at(addr));
                }
            }
            MacroAction::Define(_) => {
                // Positional-arg binding stub — recognized to avoid
                // an "unknown directive" error.
            }
            MacroAction::Let { loc, expr } => {
                self.execute_macro_let(&loc, &expr);
            }
            MacroAction::Blank(range_arg) => {
                self.execute_macro_blank(&range_arg);
            }
            MacroAction::Recalc(_) => {
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();
                self.recalc_pending = false;
            }
            MacroAction::QuestionPause => {
                if let Some(s) = self.macro_state.as_mut() {
                    s.suspend = Some(MacroSuspend::WaitEnter);
                }
            }
            MacroAction::GetLabel { prompt_text, loc } => {
                self.start_macro_get_input(prompt_text, loc, false);
            }
            MacroAction::GetNumber { prompt_text, loc } => {
                self.start_macro_get_input(prompt_text, loc, true);
            }
            MacroAction::MenuBranch(loc) => {
                self.open_custom_menu(&loc, false);
            }
            MacroAction::MenuCall(loc) => {
                self.open_custom_menu(&loc, true);
            }
            MacroAction::Beep => {
                self.beep_count = self.beep_count.saturating_add(1);
                self.beep_pending = true;
            }
            MacroAction::Wait(_)
            | MacroAction::BreakOff
            | MacroAction::BreakOn
            | MacroAction::OnError { .. } => {
                // Stubs: lexed so a macro source using them doesn't
                // halt with an unknown-directive error. Wall-clock
                // sleeps and Ctrl-Break interception are deferred
                // out of M9; ONERROR trap behavior needs set_error
                // to consult macro state.
            }
        }
        true
    }

    /// `{MENUBRANCH loc}` / `{MENUCALL loc}` — open a custom menu
    /// reading item names + descriptions out of the cells at `loc`.
    /// Items terminate at the first empty name cell (or 8 columns,
    /// whichever comes first — Lotus's hard cap).
    pub(super) fn open_custom_menu(&mut self, loc: &str, is_call: bool) {
        let Some(start) = self.resolve_macro_loc(loc) else {
            self.set_error(format!("macro: bad menu loc `{loc}`"));
            self.macro_state = None;
            return;
        };
        let mut items: Vec<CustomMenuItem> = Vec::new();
        for i in 0..8u16 {
            let col = start.col.saturating_add(i);
            let name_addr = Address::new(start.sheet, col, start.row);
            let desc_addr = Address::new(start.sheet, col, start.row.saturating_add(1));
            let Some(name) = self.read_macro_line(name_addr) else {
                break;
            };
            if name.is_empty() {
                break;
            }
            let description = self.read_macro_line(desc_addr).unwrap_or_default();
            items.push(CustomMenuItem { name, description });
        }
        if items.is_empty() {
            self.set_error("macro: empty custom menu".to_string());
            self.macro_state = None;
            return;
        }
        let action_row = Address::new(start.sheet, start.col, start.row.saturating_add(2));
        self.custom_menu = Some(CustomMenuState {
            items,
            action_row,
            is_call,
            highlight: 0,
        });
        if let Some(s) = self.macro_state.as_mut() {
            s.suspend = Some(MacroSuspend::MenuPick);
        }
        self.mode = Mode::Menu;
    }

    /// User picked item `idx` from the custom menu (or the menu
    /// was cancelled with `idx = None`). Resume the macro: BRANCH
    /// or CALL to the chosen action cell, or just continue past
    /// the `{MENUBRANCH}` if cancelled.
    pub(super) fn finish_custom_menu(&mut self, picked: Option<usize>) {
        let Some(menu) = self.custom_menu.take() else {
            return;
        };
        if let Some(idx) = picked {
            let action = Address::new(
                menu.action_row.sheet,
                menu.action_row.col.saturating_add(idx as u16),
                menu.action_row.row,
            );
            if let Some(s) = self.macro_state.as_mut() {
                if menu.is_call {
                    if s.frames.len() >= 64 {
                        self.set_error("macro: call stack overflow".to_string());
                        self.macro_state = None;
                        return;
                    }
                    s.frames.push(MacroFrame::starting_at(action));
                } else if let Some(top) = s.frames.last_mut() {
                    top.pc = action;
                    top.remaining.clear();
                }
            }
        }
        if let Some(s) = self.macro_state.as_mut() {
            s.suspend = None;
        }
        self.mode = Mode::Ready;
        self.pump_macro();
    }

    /// `/Worksheet Learn Cancel` — drop the learn range. Stops the
    /// recorder if it was on; the in-flight buffer is discarded.
    pub(super) fn cancel_learn(&mut self) {
        self.learn_range = None;
        self.learn_recording = false;
        self.learn_buffer.clear();
        self.close_menu();
    }

    /// `/Worksheet Learn Erase` — blank every cell in the learn
    /// range without dropping the range definition.
    pub(super) fn erase_learn_range(&mut self) {
        if let Some(r) = self.learn_range {
            self.execute_range_erase(r);
            self.wb_mut().dirty = true;
        }
        self.close_menu();
    }

    /// Alt-F5 LEARN toggle. Off→On arms recording; On→Off flushes
    /// the buffered macro source to cells of the learn range.
    /// v0.4: also opens / closes the `.l123log` sidecar writer
    /// (one JSON record per token) so the session is replayable
    /// outside the learn range.
    pub(super) fn toggle_learn_recording(&mut self) {
        if self.learn_range.is_none() {
            // No range set — Lotus would beep with "no learn range
            // defined". Silent no-op for now.
            return;
        }
        if self.learn_recording {
            self.learn_recording = false;
            self.flush_learn_buffer();
            self.close_learn_sidecar();
        } else {
            self.learn_buffer.clear();
            self.learn_recording = true;
            self.open_learn_sidecar();
        }
    }

    /// Open the `.l123log` sidecar writer for a fresh LEARN session.
    /// Path source order: explicit `learn_sidecar_path` (test hook) →
    /// `<active_path>.l123log` derived from the saved workbook. With
    /// neither, the sidecar stays closed and LEARN behaves as
    /// pre-v0.4 (in-range macro only).
    fn open_learn_sidecar(&mut self) {
        let path = match self.learn_sidecar_path.clone() {
            Some(p) => Some(p),
            None => self
                .wb()
                .active_path
                .as_ref()
                .map(|p| p.with_extension("l123log")),
        };
        let Some(path) = path else {
            return;
        };
        match std::fs::File::create(&path) {
            Ok(f) => {
                self.learn_sidecar_writer = Some(std::io::BufWriter::new(f));
            }
            Err(e) => {
                tracing::warn!(
                    "could not open learn sidecar {}: {e}",
                    path.display()
                );
            }
        }
    }

    /// Flush + drop the sidecar writer. Always called when LEARN
    /// turns off (even if `open_learn_sidecar` failed silently).
    fn close_learn_sidecar(&mut self) {
        if let Some(mut w) = self.learn_sidecar_writer.take() {
            use std::io::Write;
            let _ = w.flush();
        }
    }

    /// Append one JSON record `{"keys": "<token>"}` to the sidecar
    /// writer if one is open. Errors are swallowed (LEARN is the
    /// primary contract; the sidecar is additive).
    fn append_learn_sidecar(&mut self, token: &str) {
        let Some(w) = self.learn_sidecar_writer.as_mut() else {
            return;
        };
        use std::io::Write;
        let line = serde_json::json!({ "keys": token }).to_string();
        let _ = writeln!(w, "{line}");
    }

    /// Replay a `.l123log` sidecar onto the current App state. Each
    /// line is a JSON object with a `keys` string of macro tokens;
    /// the concatenation is dispatched through `run_macro_text` (which
    /// lexes the tokens and feeds them to `handle_key`). The
    /// regression-test harness invokes this via the `REPLAY` directive;
    /// `l123 --replay <path>` is the CLI counterpart.
    pub fn replay_sidecar(&mut self, path: &std::path::Path) -> Result<(), String> {
        let body =
            std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
        let mut src = String::new();
        for (i, line) in body.lines().enumerate() {
            if line.trim().is_empty() {
                continue;
            }
            let parsed: serde_json::Value = serde_json::from_str(line)
                .map_err(|e| format!("{}: line {}: {e}", path.display(), i + 1))?;
            let Some(tok) = parsed.get("keys").and_then(|v| v.as_str()) else {
                return Err(format!(
                    "{}: line {}: missing or non-string \"keys\" field",
                    path.display(),
                    i + 1
                ));
            };
            src.push_str(tok);
        }
        self.run_macro_text(&src);
        Ok(())
    }

    /// Test hook: pin the sidecar path used by the next LEARN session,
    /// bypassing the `active_path` derivation. Production code leaves
    /// `learn_sidecar_path` at `None`.
    pub fn test_set_learn_sidecar_path(&mut self, path: Option<std::path::PathBuf>) {
        self.learn_sidecar_path = path;
    }

    /// Write `learn_buffer` into the cells of `learn_range`,
    /// splitting at 240 chars (the 1-2-3 label-cell capacity) so
    /// the recording wraps to subsequent rows.
    pub(super) fn flush_learn_buffer(&mut self) {
        let Some(range) = self.learn_range else {
            return;
        };
        if self.learn_buffer.is_empty() {
            return;
        }
        let buf = std::mem::take(&mut self.learn_buffer);
        let mut row = range.start.row;
        let max_row = range.end.row;
        let chunk_size = 240;
        let mut chars = buf.chars().peekable();
        while chars.peek().is_some() && row <= max_row {
            let mut chunk = String::new();
            for _ in 0..chunk_size {
                match chars.next() {
                    Some(ch) => chunk.push(ch),
                    None => break,
                }
            }
            let addr = Address::new(range.start.sheet, range.start.col, row);
            let contents = CellContents::Label {
                prefix: LabelPrefix::Apostrophe,
                text: chunk,
            };
            if self.undo_enabled {
                let prev = self.wb().cells.get(&addr).cloned();
                let prev_format = self.wb().cell_formats.get(&addr).copied();
                self.wb_mut().journal.push(JournalEntry::CellEdit {
                    addr,
                    prev_contents: prev,
                    prev_format,
                });
            }
            self.push_to_engine_at(addr, &contents);
            self.wb_mut().cells.insert(addr, contents);
            self.wb_mut().dirty = true;
            row = match row.checked_add(1) {
                Some(r) => r,
                None => break,
            };
        }
    }

    /// Append the macro-source serialization of `k` to the learn
    /// buffer if recording is on. Skips Alt-F5 itself (the toggle)
    /// and modifier-only "dead" key events that don't represent
    /// user-typed input.
    pub(super) fn record_keystroke(&mut self, k: &KeyEvent) {
        if !self.learn_recording {
            return;
        }
        if let Some(token) = key_event_to_macro_source(k) {
            self.learn_buffer.push_str(&token);
            self.append_learn_sidecar(&token);
        }
    }

    /// Open a prompt in service of `{GETLABEL}` / `{GETNUMBER}` and
    /// park the macro until the user commits or cancels.
    pub(super) fn start_macro_get_input(
        &mut self,
        prompt_text: String,
        loc: String,
        numeric: bool,
    ) {
        if let Some(s) = self.macro_state.as_mut() {
            s.suspend = Some(MacroSuspend::GetInput);
        }
        let label = if prompt_text.trim().is_empty() {
            "Macro input:".to_string()
        } else {
            prompt_text.trim().to_string()
        };
        self.pending_macro_input_loc = Some(loc);
        self.prompt = Some(PromptState {
            label,
            buffer: String::new(),
            next: PromptNext::MacroGetInput { numeric },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `{LET loc, expr}` — write `expr`'s value to `loc` directly.
    /// Goes through the same source-form parser as a typed entry
    /// commit and journals the previous cell state for undo.
    pub(super) fn execute_macro_let(&mut self, loc: &str, expr: &str) {
        let Some(addr) = self.resolve_macro_loc(loc) else {
            self.set_error(format!("macro: {{LET}} bad loc `{loc}`"));
            return;
        };
        self.write_source_at(addr, expr);
    }

    /// Parse `source` (1-2-3 source form, e.g. `@SUM(A1..A5)` or
    /// `42`) and commit it to `addr`, journalling the previous cell
    /// for undo and recalcing the engine. Shared by `{LET}` and the
    /// SmartIcon writers (Sum, Today's Date).
    pub(super) fn write_source_at(&mut self, addr: Address, source: &str) {
        let intl = self.wb().international.clone();
        let (contents, format) =
            CellContents::from_source_with_format(source, self.default_label_prefix, &intl);
        if self.undo_enabled {
            let prev_contents = self.wb().cells.get(&addr).cloned();
            let prev_format = self.wb().cell_formats.get(&addr).copied();
            self.wb_mut().journal.push(JournalEntry::CellEdit {
                addr,
                prev_contents,
                prev_format,
            });
        }
        self.push_to_engine_at(addr, &contents);
        if contents.is_empty() {
            self.wb_mut().cells.remove(&addr);
        } else {
            self.wb_mut().cells.insert(addr, contents);
        }
        if let Some(fmt) = format {
            self.wb_mut().cell_formats.insert(addr, fmt);
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
    }

    /// `{BLANK range}` — erase every cell in `range`. Reuses the
    /// existing /Range Erase plumbing so the undo journal stays
    /// consistent.
    pub(super) fn execute_macro_blank(&mut self, range_arg: &str) {
        let Some(range) = self.parse_macro_range(range_arg) else {
            self.set_error(format!("macro: {{BLANK}} bad range `{range_arg}`"));
            return;
        };
        self.execute_range_erase(range);
        self.wb_mut().dirty = true;
    }

    /// Resolve a `range` argument to a [`Range`]. Accepts named
    /// ranges, single addresses (treated as 1×1 ranges), and
    /// `addr..addr` literal forms. Returns `None` if neither
    /// representation parses.
    pub(super) fn parse_macro_range(&self, arg: &str) -> Option<Range> {
        let trimmed = arg.trim();
        if trimmed.is_empty() {
            return None;
        }
        // Named range first.
        let key = trimmed.to_ascii_lowercase();
        if let Some(r) = self.wb().named_ranges.get(&key) {
            return Some(*r);
        }
        // `addr..addr` literal.
        if let Some((a, b)) = trimmed.split_once("..") {
            let start = Address::parse(a.trim()).ok()?;
            let end = Address::parse(b.trim()).ok()?;
            return Some(Range { start, end });
        }
        // Bare address → 1×1 range.
        let addr = Address::parse(trimmed).ok()?;
        Some(Range {
            start: addr,
            end: addr,
        })
    }
}
