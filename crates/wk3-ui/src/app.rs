//! The app loop: ratatui + crossterm, control panel + grid + status line.
//!
//! Scope as of M1 cycle 2:
//! - READY / LABEL / VALUE modes with first-character dispatch (LABEL only
//!   implemented this cycle; VALUE lands in cycle 3).
//! - `'` auto-prefixed labels. Enter commits; Ctrl-C quits.
//! - Three-line control panel, mode indicator, cell readout.

use std::collections::HashMap;
use std::io;
use std::time::Duration;

use crossterm::{
    event::{self, DisableMouseCapture, Event, KeyCode, KeyEvent, KeyEventKind, KeyModifiers},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::CrosstermBackend,
    buffer::Buffer,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Paragraph, Widget},
    Terminal,
};
use wk3_core::{
    address::col_to_letters, label::is_value_starter, Address, CellContents, LabelPrefix, Mode,
    SheetId, Value,
};

// Grid geometry — kept as consts so both render and cell-address-probe agree.
const ROW_GUTTER: u16 = 5;
const COL_WIDTH: u16 = 9;
const PANEL_HEIGHT: u16 = 4; // 3 content lines + 1 bottom border

#[derive(Debug, Clone, Copy, PartialEq)]
enum EntryKind {
    /// Label entry with an implicit or explicit prefix. Buffer holds the
    /// post-prefix text; the prefix is displayed only on commit / on line 1.
    Label(LabelPrefix),
    /// Value entry. Buffer is the literal characters typed.
    Value,
    /// F2-initiated edit of an existing cell. Buffer holds the full source
    /// form (including prefix for labels). Commit re-applies the first-char
    /// rule so the user may change the prefix or the type.
    Edit,
}

#[derive(Debug)]
struct Entry {
    kind: EntryKind,
    buffer: String,
}

pub struct App {
    pointer: Address,
    mode: Mode,
    running: bool,
    viewport_col_offset: u16,
    viewport_row_offset: u32,
    cells: HashMap<Address, CellContents>,
    entry: Option<Entry>,
    default_label_prefix: LabelPrefix,
}

impl App {
    pub fn new() -> Self {
        Self {
            pointer: Address::A1,
            mode: Mode::Ready,
            running: true,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            cells: HashMap::new(),
            entry: None,
            default_label_prefix: LabelPrefix::Apostrophe,
        }
    }

    pub fn run() -> anyhow::Result<()> {
        let mut stdout = io::stdout();
        enable_raw_mode()?;
        execute!(stdout, EnterAlternateScreen)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        let mut app = App::new();
        let result = app.event_loop(&mut terminal);

        disable_raw_mode()?;
        execute!(terminal.backend_mut(), LeaveAlternateScreen, DisableMouseCapture)?;
        terminal.show_cursor()?;

        result
    }

    fn event_loop<B: ratatui::backend::Backend>(
        &mut self,
        terminal: &mut Terminal<B>,
    ) -> anyhow::Result<()> {
        while self.running {
            terminal.draw(|f| self.render(f.area(), f.buffer_mut()))?;
            if event::poll(Duration::from_millis(100))? {
                if let Event::Key(k) = event::read()? {
                    if k.kind == KeyEventKind::Press {
                        self.handle_key(k);
                    }
                }
            }
        }
        Ok(())
    }

    // ---------------- test-surface accessors ----------------

    pub fn pointer(&self) -> Address {
        self.pointer
    }

    pub fn mode(&self) -> Mode {
        self.mode
    }

    pub fn is_running(&self) -> bool {
        self.running
    }

    pub fn render_to_buffer(&self, width: u16, height: u16) -> Buffer {
        let area = Rect::new(0, 0, width, height);
        let mut buf = Buffer::empty(area);
        self.render(area, &mut buf);
        buf
    }

    pub fn line_text(buf: &Buffer, y: u16) -> String {
        let mut s = String::new();
        for x in 0..buf.area.width {
            s.push_str(buf[(x, y)].symbol());
        }
        s.trim_end().to_string()
    }

    /// Read back the rendered text of a single grid cell by address
    /// (`"A:B5"` or `"B5"`). Returns None if the cell is outside the
    /// current viewport.
    pub fn cell_rendered_text(&self, buf: &Buffer, addr: &str) -> Option<String> {
        let a = Address::parse(addr).ok()?;
        if a.col < self.viewport_col_offset || a.row < self.viewport_row_offset {
            return None;
        }
        let dc = a.col - self.viewport_col_offset;
        let dr = (a.row - self.viewport_row_offset) as u16;
        let x0 = ROW_GUTTER + dc * COL_WIDTH;
        let y = PANEL_HEIGHT + 1 + dr; // +1 skips column header row
        if x0 + COL_WIDTH > buf.area.width || y >= buf.area.height {
            return None;
        }
        let mut s = String::with_capacity(COL_WIDTH as usize);
        for i in 0..COL_WIDTH {
            s.push_str(buf[(x0 + i, y)].symbol());
        }
        Some(s)
    }

    // ---------------- key handling ----------------

    pub fn handle_key(&mut self, k: KeyEvent) {
        // Ctrl-C (Ctrl-Break alias) always exits.
        if k.modifiers.contains(KeyModifiers::CONTROL) && k.code == KeyCode::Char('c') {
            self.running = false;
            return;
        }
        match self.mode {
            Mode::Ready => self.handle_key_ready(k),
            Mode::Label | Mode::Value | Mode::Edit => self.handle_key_entry(k),
            _ => {}
        }
    }

    fn handle_key_ready(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Up => self.move_pointer(0, -1),
            KeyCode::Down => self.move_pointer(0, 1),
            KeyCode::Left => self.move_pointer(-1, 0),
            KeyCode::Right => self.move_pointer(1, 0),
            KeyCode::Home => {
                self.pointer = Address::A1;
                self.viewport_col_offset = 0;
                self.viewport_row_offset = 0;
            }
            KeyCode::PageDown => self.move_pointer(0, 20),
            KeyCode::PageUp => self.move_pointer(0, -20),
            KeyCode::F(2) => self.begin_edit(),
            KeyCode::Char(c) => self.begin_entry(c),
            _ => {}
        }
    }

    fn begin_edit(&mut self) {
        let source = self
            .cells
            .get(&self.pointer)
            .map(|c| c.source_form())
            .unwrap_or_default();
        self.entry = Some(Entry { kind: EntryKind::Edit, buffer: source });
        self.mode = Mode::Edit;
    }

    fn handle_key_entry(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Enter => self.commit_entry(),
            KeyCode::Esc => self.cancel_entry(),
            // Arrow/Tab: commit then move. Pressing these during entry is
            // the canonical fast-entry idiom (see Tutorial §2.4).
            KeyCode::Up => {
                self.commit_entry();
                self.move_pointer(0, -1);
            }
            KeyCode::Down => {
                self.commit_entry();
                self.move_pointer(0, 1);
            }
            KeyCode::Left => {
                self.commit_entry();
                self.move_pointer(-1, 0);
            }
            KeyCode::Right | KeyCode::Tab => {
                self.commit_entry();
                self.move_pointer(1, 0);
            }
            KeyCode::Backspace => {
                if let Some(e) = self.entry.as_mut() {
                    e.buffer.pop();
                }
            }
            KeyCode::Char(c) => {
                if let Some(e) = self.entry.as_mut() {
                    e.buffer.push(c);
                }
            }
            _ => {}
        }
    }

    fn cancel_entry(&mut self) {
        self.entry = None;
        self.mode = Mode::Ready;
    }

    fn begin_entry(&mut self, c: char) {
        if is_value_starter(c) {
            self.entry = Some(Entry { kind: EntryKind::Value, buffer: c.to_string() });
            self.mode = Mode::Value;
        } else if matches!(c, '\'' | '"' | '^' | '\\') {
            // Explicit label prefix typed first: the char becomes the
            // LabelPrefix; the buffer starts empty.
            let prefix = LabelPrefix::from_char(c).expect("matched above");
            self.entry = Some(Entry { kind: EntryKind::Label(prefix), buffer: String::new() });
            self.mode = Mode::Label;
        } else {
            // Any other non-value-starter: default `'` prefix auto-inserted;
            // the typed char is the first char of the label text.
            self.entry = Some(Entry {
                kind: EntryKind::Label(self.default_label_prefix),
                buffer: c.to_string(),
            });
            self.mode = Mode::Label;
        }
    }

    fn commit_entry(&mut self) {
        let Some(entry) = self.entry.take() else {
            self.mode = Mode::Ready;
            return;
        };
        let contents = match entry.kind {
            EntryKind::Label(prefix) => CellContents::Label { prefix, text: entry.buffer },
            EntryKind::Value => match entry.buffer.parse::<f64>() {
                Ok(n) => CellContents::Constant(Value::Number(n)),
                Err(_) => CellContents::Label {
                    prefix: self.default_label_prefix,
                    text: entry.buffer,
                },
            },
            // EDIT commits re-parse the full source buffer so the user can
            // change prefix or type (label ↔ value) via the first-char rule.
            EntryKind::Edit => {
                CellContents::from_source(&entry.buffer, self.default_label_prefix)
            }
        };
        if contents.is_empty() {
            self.cells.remove(&self.pointer);
        } else {
            self.cells.insert(self.pointer, contents);
        }
        self.mode = Mode::Ready;
    }

    fn move_pointer(&mut self, d_col: i32, d_row: i32) {
        if let Some(next) = self.pointer.shifted(d_col, d_row) {
            self.pointer = next;
            self.scroll_into_view();
        }
    }

    fn scroll_into_view(&mut self) {
        if self.pointer.col < self.viewport_col_offset {
            self.viewport_col_offset = self.pointer.col;
        }
        if self.pointer.row < self.viewport_row_offset {
            self.viewport_row_offset = self.pointer.row;
        }
    }

    // ---------------- rendering ----------------

    fn render(&self, area: Rect, buf: &mut Buffer) {
        let chunks = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(PANEL_HEIGHT),
                Constraint::Min(1),
                Constraint::Length(1),
            ])
            .split(area);

        self.render_control_panel(chunks[0], buf);
        self.render_grid(chunks[1], buf);
        self.render_status(chunks[2], buf);
    }

    fn render_control_panel(&self, area: Rect, buf: &mut Buffer) {
        let block = Block::default().borders(Borders::BOTTOM);
        let inner = block.inner(area);
        block.render(area, buf);

        // Line 1: "<addr>: [(fmt)] <readout>" left; mode indicator right.
        let readout = self.cell_readout_for_line1();
        let left = format!(" {}: {}", self.pointer.display_full(), readout);
        let mode_str = self.mode.indicator();
        let pad = (area.width as usize)
            .saturating_sub(left.chars().count() + mode_str.len() + 1);
        let line1 = Line::from(vec![
            Span::raw(left),
            Span::raw(" ".repeat(pad)),
            Span::styled(mode_str, Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)),
            Span::raw(" "),
        ]);

        // Line 2: entry buffer during LABEL/VALUE/EDIT.
        let line2 = match self.entry.as_ref() {
            Some(e) => Line::from(format!(" {}", e.buffer)),
            None => Line::from(""),
        };

        // Line 3: reserved for menu preview / prompt (M3+).
        let line3 = Line::from("");

        Paragraph::new(vec![line1, line2, line3]).render(inner, buf);
    }

    fn cell_readout_for_line1(&self) -> String {
        self.cells
            .get(&self.pointer)
            .map(|c| c.control_panel_readout())
            .unwrap_or_default()
    }

    fn render_grid(&self, area: Rect, buf: &mut Buffer) {
        if area.width <= ROW_GUTTER || area.height < 2 {
            return;
        }

        let visible_cols = ((area.width - ROW_GUTTER) / COL_WIDTH).max(1);
        let visible_rows = area.height - 1;

        // Column header row
        for i in 0..visible_cols {
            let c = self.viewport_col_offset + i;
            let letters = col_to_letters(c);
            let x = area.x + ROW_GUTTER + i * COL_WIDTH;
            let y = area.y;
            let style = Style::default().add_modifier(Modifier::REVERSED);
            write_centered(buf, x, y, COL_WIDTH, &letters, style);
        }
        // Top-left gutter corner
        for k in 0..ROW_GUTTER {
            buf[(area.x + k, area.y)]
                .set_char(' ')
                .set_style(Style::default().add_modifier(Modifier::REVERSED));
        }

        // Body rows
        for r in 0..visible_rows {
            let row_idx = self.viewport_row_offset + r as u32;
            let y = area.y + 1 + r;
            // Row number gutter
            let label = format!("{:>width$}", row_idx + 1, width = (ROW_GUTTER - 1) as usize);
            let style = Style::default().add_modifier(Modifier::REVERSED);
            for (i, ch) in label.chars().enumerate() {
                buf[(area.x + i as u16, y)].set_char(ch).set_style(style);
            }
            buf[(area.x + ROW_GUTTER - 1, y)]
                .set_char(' ')
                .set_style(style);

            // Cells
            for c in 0..visible_cols {
                let col_idx = self.viewport_col_offset + c;
                let x = area.x + ROW_GUTTER + c * COL_WIDTH;
                let addr = Address::new(SheetId::A, col_idx, row_idx);
                let is_pointer = addr == self.pointer;
                let cell_style = if is_pointer {
                    Style::default().add_modifier(Modifier::REVERSED)
                } else {
                    Style::default()
                };
                // Blank background first
                for k in 0..COL_WIDTH {
                    buf[(x + k, y)].set_char(' ').set_style(cell_style);
                }
                // Content
                if let Some(contents) = self.cells.get(&addr) {
                    draw_cell_contents(buf, x, y, COL_WIDTH, contents, cell_style);
                }
            }
        }
    }

    fn render_status(&self, area: Rect, buf: &mut Buffer) {
        let left = " *untitled*";
        let hint = "Ctrl-C to quit";
        let pad = (area.width as usize).saturating_sub(left.len() + hint.len() + 1);
        let line = format!("{left}{}{hint} ", " ".repeat(pad));
        for (i, ch) in line.chars().enumerate().take(area.width as usize) {
            buf[(area.x + i as u16, area.y)]
                .set_char(ch)
                .set_style(Style::default().fg(Color::DarkGray));
        }
    }
}

impl Default for App {
    fn default() -> Self {
        Self::new()
    }
}

fn write_centered(buf: &mut Buffer, x: u16, y: u16, width: u16, text: &str, style: Style) {
    let w = width as usize;
    let t: String = if text.chars().count() >= w {
        text.chars().take(w).collect()
    } else {
        let pad_left = (w - text.chars().count()) / 2;
        let pad_right = w - text.chars().count() - pad_left;
        format!("{}{}{}", " ".repeat(pad_left), text, " ".repeat(pad_right))
    };
    for (i, ch) in t.chars().enumerate().take(w) {
        buf[(x + i as u16, y)].set_char(ch).set_style(style);
    }
}

/// Render a cell's contents into a fixed-width slot.
/// - Labels honor their prefix alignment (`'` left, `"` right, `^` center,
///   `\` repeats the text across cell width, `|` treated as left-align for
///   now — it's a print-only marker).
/// - Numbers render right-aligned in General format.
fn draw_cell_contents(
    buf: &mut Buffer,
    x: u16,
    y: u16,
    width: u16,
    contents: &CellContents,
    style: Style,
) {
    let w = width as usize;
    let rendered = match contents {
        CellContents::Empty => return,
        CellContents::Label { prefix, text } => render_label(*prefix, text, w),
        CellContents::Constant(Value::Number(n)) => {
            right_pad(&wk3_core::format_number_general(*n), w, /*right_align=*/ true)
        }
        CellContents::Constant(Value::Text(s)) => right_pad(s, w, /*right_align=*/ false),
        CellContents::Constant(_) => return,
    };
    for (i, ch) in rendered.chars().enumerate().take(w) {
        buf[(x + i as u16, y)].set_char(ch).set_style(style);
    }
}

fn render_label(prefix: LabelPrefix, text: &str, width: usize) -> String {
    if text.is_empty() {
        return " ".repeat(width);
    }
    match prefix {
        LabelPrefix::Apostrophe | LabelPrefix::Pipe => right_pad(text, width, false),
        LabelPrefix::Quote => right_pad(text, width, true),
        LabelPrefix::Caret => center_pad(text, width),
        LabelPrefix::Backslash => repeat_to_width(text, width),
    }
}

/// Left-pad (right_align=true) or right-pad with spaces to `width`, or
/// truncate to `width`.
fn right_pad(text: &str, width: usize, right_align: bool) -> String {
    let chars: Vec<char> = text.chars().collect();
    if chars.len() >= width {
        return chars.into_iter().take(width).collect();
    }
    let pad = width - chars.len();
    if right_align {
        format!("{}{text}", " ".repeat(pad))
    } else {
        format!("{text}{}", " ".repeat(pad))
    }
}

fn center_pad(text: &str, width: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    if chars.len() >= width {
        return chars.into_iter().take(width).collect();
    }
    let pad = width - chars.len();
    let left = pad / 2;
    let right = pad - left;
    format!("{}{text}{}", " ".repeat(left), " ".repeat(right))
}

fn repeat_to_width(pattern: &str, width: usize) -> String {
    if pattern.is_empty() {
        return " ".repeat(width);
    }
    let mut out = String::with_capacity(width);
    let pattern_chars: Vec<char> = pattern.chars().collect();
    while out.chars().count() < width {
        for ch in &pattern_chars {
            if out.chars().count() == width {
                break;
            }
            out.push(*ch);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn starts_at_a1() {
        let app = App::new();
        assert_eq!(app.pointer, Address::A1);
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.entry.is_none());
    }

    #[test]
    fn arrow_nav() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        assert_eq!(app.pointer.col, 1);
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        assert_eq!(app.pointer.row, 1);
        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
        assert_eq!(app.pointer.col, 0);
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        assert_eq!(app.pointer.row, 0);
    }

    #[test]
    fn left_from_a1_stays_at_a1() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        assert_eq!(app.pointer, Address::A1);
    }

    #[test]
    fn home_resets_pointer() {
        let mut app = App::new();
        app.pointer = Address::new(SheetId::A, 10, 10);
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        assert_eq!(app.pointer, Address::A1);
    }

    #[test]
    fn ctrl_c_quits() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL));
        assert!(!app.running);
    }

    #[test]
    fn pgdn_moves_twenty_rows() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::NONE));
        assert_eq!(app.pointer.row, 20);
    }

    #[test]
    fn letter_first_enters_label_mode() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('h'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Label);
        let e = app.entry.as_ref().unwrap();
        assert_eq!(e.buffer, "h");
        assert!(matches!(e.kind, EntryKind::Label(LabelPrefix::Apostrophe)));
    }

    #[test]
    fn digit_first_enters_value_mode() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('1'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Value);
        let e = app.entry.as_ref().unwrap();
        assert_eq!(e.buffer, "1");
        assert!(matches!(e.kind, EntryKind::Value));
    }

    #[test]
    fn f2_on_empty_cell_enters_edit_mode() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.entry.as_ref().unwrap().buffer, "");
    }

    #[test]
    fn f2_loads_label_source_into_buffer_with_prefix() {
        let mut app = App::new();
        app.cells.insert(
            Address::A1,
            CellContents::Label {
                prefix: LabelPrefix::Quote,
                text: "right".into(),
            },
        );
        app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.entry.as_ref().unwrap().buffer, "\"right");
    }

    #[test]
    fn f2_commit_reparses_via_first_char_rule() {
        let mut app = App::new();
        app.cells.insert(
            Address::A1,
            CellContents::Label {
                prefix: LabelPrefix::Apostrophe,
                text: "hello".into(),
            },
        );
        app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
        // Buffer is "'hello"; remove ' and " -prefix instead.
        app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
        for c in "ello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
            let _ = c;
        }
        // buffer = "h" now (we backspaced over 'ello but 'h' remains);
        // for a deterministic commit, clear fully.
        app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
        assert_eq!(app.entry.as_ref().unwrap().buffer, "");
        app.handle_key(KeyEvent::new(KeyCode::Char('"'), KeyModifiers::NONE));
        for c in "right".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        match app.cells.get(&Address::A1).unwrap() {
            CellContents::Label { prefix, text } => {
                assert_eq!(*prefix, LabelPrefix::Quote);
                assert_eq!(text, "right");
            }
            other => panic!("expected Label(Quote), got {other:?}"),
        }
    }

    #[test]
    fn esc_during_entry_cancels_and_leaves_cell_empty() {
        let mut app = App::new();
        for c in "hello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.entry.is_none());
        assert!(!app.cells.contains_key(&Address::A1));
    }

    #[test]
    fn arrow_commits_then_moves() {
        let mut app = App::new();
        for c in "hi".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(app.pointer, Address::new(SheetId::A, 0, 1));
        assert!(matches!(
            app.cells.get(&Address::A1),
            Some(CellContents::Label { .. })
        ));
    }

    #[test]
    fn backspace_edits_buffer() {
        let mut app = App::new();
        for c in "hello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
        assert_eq!(app.entry.as_ref().unwrap().buffer, "hell");
    }

    #[test]
    fn explicit_label_prefix_dispatch() {
        for (ch, want) in [
            ('\'', LabelPrefix::Apostrophe),
            ('"', LabelPrefix::Quote),
            ('^', LabelPrefix::Caret),
            ('\\', LabelPrefix::Backslash),
        ] {
            let mut app = App::new();
            app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::NONE));
            assert_eq!(app.mode, Mode::Label, "char {ch:?}");
            let e = app.entry.as_ref().unwrap();
            assert!(
                matches!(e.kind, EntryKind::Label(p) if p == want),
                "char {ch:?}: expected prefix {want:?}, got {:?}",
                e.kind
            );
            assert_eq!(e.buffer, "", "buffer should be empty after prefix char");
        }
    }

    #[test]
    fn render_label_alignments() {
        assert_eq!(render_label(LabelPrefix::Apostrophe, "hi", 9), "hi       ");
        assert_eq!(render_label(LabelPrefix::Quote, "hi", 9), "       hi");
        assert_eq!(render_label(LabelPrefix::Caret, "hi", 9), "   hi    ");
        assert_eq!(render_label(LabelPrefix::Backslash, "-", 9), "---------");
        assert_eq!(render_label(LabelPrefix::Backslash, "ab", 9), "ababababa");
        assert_eq!(render_label(LabelPrefix::Backslash, "abc", 9), "abcabcabc");
    }

    #[test]
    fn value_commit_stores_as_number() {
        let mut app = App::new();
        for c in "123".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        match app.cells.get(&Address::A1).unwrap() {
            CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 123.0),
            other => panic!("expected Number, got {other:?}"),
        }
    }

    #[test]
    fn value_commit_handles_decimal_and_negative() {
        let mut app = App::new();
        for c in "-1.25".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        match app.cells.get(&Address::A1).unwrap() {
            CellContents::Constant(Value::Number(n)) => {
                assert!((*n - (-1.25)).abs() < 1e-9, "got {n}");
            }
            other => panic!("expected Number, got {other:?}"),
        }
    }

    #[test]
    fn label_commit_stores_with_prefix_and_returns_to_ready() {
        let mut app = App::new();
        for c in "hello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.entry.is_none());
        let stored = app.cells.get(&Address::A1).unwrap();
        match stored {
            CellContents::Label { prefix, text } => {
                assert_eq!(*prefix, LabelPrefix::Apostrophe);
                assert_eq!(text, "hello");
            }
            other => panic!("expected Label, got {other:?}"),
        }
    }
}
