//! All rendering: top-level [`App::render`], per-region panels, overlays,
//! the visible-layout planners, and helper free fns. The test-surface
//! accessors that probe a rendered buffer also live here since they are
//! tightly coupled to the layout math the renderers use.

use l123_core::cell_render::{apply_halign_to_rendered, label_text_bounds};
use l123_core::{
    address::col_to_letters, plan_row_spill, render_label, render_value_in_cell, Address,
    CellContents, Format, HAlign, International, Mode, Range, RgbColor, SheetId, SpillSlot,
    TextStyle, Value,
};
use l123_engine::{Engine, RecalcMode};
use l123_menu::MenuBody;
use l123_print::WorkbookView;
use ratatui::{
    buffer::Buffer,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Paragraph, Widget},
};
use ratatui_image::{picker::ProtocolType, Image, Resize};

use crate::help::HelpState;

use super::types::*;
use super::{
    effective_label_prefix, format_file_list_row, format_size, letter_to_sheet_index, App,
    COMMENT_MARKER, PANEL_HEIGHT, ROW_GUTTER, STATUS_SHEET_NAME_MAX,
};

fn display_mode_default_style(mode: DisplayMode) -> Style {
    match mode {
        DisplayMode::BW => Style::default(),
        DisplayMode::Color => Style::default()
            .bg(Color::Rgb(0xFF, 0xFF, 0xFF))
            .fg(Color::Rgb(0x00, 0x00, 0x00)),
        DisplayMode::Reverse => Style::default()
            .bg(Color::Rgb(0x00, 0x00, 0x00))
            .fg(Color::Rgb(0xFF, 0xFF, 0xFF)),
    }
}
/// One body row of the help overlay: a slice of plain text plus zero
/// or more link spans expressed as `(byte_range, link_index)` pairs.
/// `link_index` references the page's `links` slice so the renderer
/// can decide whether to paint that link as focused.
struct HelpRow<'a> {
    text: &'a str,
    /// `(start, end, link_index)` byte ranges into `text`.
    links: Vec<(usize, usize, usize)>,
}

/// Split a help page's body into rows for rendering. Each newline
/// becomes a row; every link is attached to the row(s) it covers
/// (links never cross newlines in the corpus, but we clip defensively).
fn build_help_rows(state: &HelpState) -> Vec<HelpRow<'_>> {
    let body = state.page.body.as_str();
    let mut rows: Vec<HelpRow<'_>> = Vec::new();
    let mut row_start = 0;
    let mut row_idx_starts: Vec<usize> = vec![0];
    for (i, ch) in body.char_indices() {
        if ch == '\n' {
            rows.push(HelpRow {
                text: &body[row_start..i],
                links: Vec::new(),
            });
            row_start = i + 1;
            row_idx_starts.push(row_start);
        }
    }
    rows.push(HelpRow {
        text: &body[row_start..],
        links: Vec::new(),
    });

    for (li, link) in state.page.links.iter().enumerate() {
        // Find which row contains link.start.
        let row = match row_idx_starts.binary_search(&link.start) {
            Ok(i) => i,
            Err(i) => i.saturating_sub(1),
        };
        if row >= rows.len() {
            continue;
        }
        let r_start = row_idx_starts[row];
        let r_end = if row + 1 < row_idx_starts.len() {
            row_idx_starts[row + 1] - 1 // strip the '\n'
        } else {
            body.len()
        };
        let s = link.start.saturating_sub(r_start);
        let e = link.end.min(r_end).saturating_sub(r_start);
        if s < e {
            rows[row].links.push((s, e, li));
        }
    }
    rows
}

/// Visible row index of the link at `focus`, if any.
fn link_row_for_focus(rows: &[HelpRow<'_>], focus: usize) -> Option<usize> {
    rows.iter()
        .position(|r| r.links.iter().any(|(_, _, li)| *li == focus))
}

/// Paint a help body row into `buf`. The row is left-padded by one
/// space (matches the existing overlay) and clipped to `width`.
fn render_help_row(buf: &mut Buffer, x: u16, y: u16, width: u16, row: &HelpRow<'_>, focus: usize) {
    let normal = Style::default();
    let link_style = Style::default().fg(Color::Green);
    let focus_style = Style::default()
        .fg(Color::Black)
        .bg(Color::Cyan)
        .add_modifier(Modifier::BOLD);

    let buf_left = buf.area.x;
    let buf_right = buf.area.x + buf.area.width;
    let mut col = x;
    if col < buf_right && y < buf.area.y + buf.area.height {
        buf[(col, y)].set_char(' ').set_style(normal);
        col += 1;
    }
    let avail = width.saturating_sub(1);
    let end_col = x + 1 + avail;

    let text_bytes = row.text.as_bytes();
    let mut i = 0;
    while i < text_bytes.len() && col < end_col && col < buf_right {
        let in_link = row
            .links
            .iter()
            .find(|(s, e, _)| i >= *s && i < *e)
            .map(|(_, _, li)| *li);
        let style = match in_link {
            Some(li) if li == focus => focus_style,
            Some(_) => link_style,
            None => normal,
        };
        let ch_end = next_char_boundary_str(row.text, i);
        let ch = row.text[i..ch_end].chars().next().unwrap_or(' ');
        if col >= buf_left {
            buf[(col, y)].set_char(ch).set_style(style);
        }
        col += 1;
        i = ch_end;
    }
    // Pad remainder of the row with spaces so a previous frame doesn't
    // bleed through.
    while col < end_col && col < buf_right {
        buf[(col, y)].set_char(' ').set_style(normal);
        col += 1;
    }
}

fn next_char_boundary_str(s: &str, mut i: usize) -> usize {
    i += 1;
    while i < s.len() && !s.is_char_boundary(i) {
        i += 1;
    }
    i
}

/// Write `text` at `(x, y)` into `buf`, padded or truncated to exactly
/// `width` cells, with `style` applied over the whole span.
fn set_line(buf: &mut Buffer, x: u16, y: u16, text: &str, width: u16, style: Style) {
    let mut chars = text.chars();
    for i in 0..width {
        let cx = x + i;
        if cx >= buf.area.x + buf.area.width || y >= buf.area.y + buf.area.height {
            break;
        }
        let ch = chars.next().unwrap_or(' ');
        buf[(cx, y)].set_char(ch).set_style(style);
    }
}

/// Build the line-2 control panel rendering of an active entry,
/// showing a reverse-video cursor cell at the buffer cursor position.
/// When the cursor sits at the end of the buffer, the highlighted cell
/// is a single space (the canonical "I-beam at end" appearance).
fn render_entry_l2(e: &Entry) -> Line<'static> {
    let cursor = e.cursor.min(e.buffer.len());
    let before = &e.buffer[..cursor];
    let (cursor_cell, after) = match e.buffer[cursor..].chars().next() {
        Some(c) => {
            let len = c.len_utf8();
            (
                e.buffer[cursor..cursor + len].to_string(),
                e.buffer[cursor + len..].to_string(),
            )
        }
        None => (" ".to_string(), String::new()),
    };
    let cursor_style = Style::default().add_modifier(Modifier::REVERSED);
    Line::from(vec![
        Span::raw(" "),
        Span::raw(before.to_string()),
        Span::styled(cursor_cell, cursor_style),
        Span::raw(after),
    ])
}
/// Map a per-cell [`TextStyle`] into the ratatui [`Modifier`] bits that
/// render it: bold → `BOLD`, italic → `ITALIC`, underline → `UNDERLINED`.
/// Empty style yields `Modifier::empty()` and adds no visible attributes.
pub(super) fn text_style_modifier(style: TextStyle) -> Modifier {
    let mut m = Modifier::empty();
    if style.bold {
        m |= Modifier::BOLD;
    }
    if style.italic {
        m |= Modifier::ITALIC;
    }
    if style.underline {
        m |= Modifier::UNDERLINED;
    }
    m
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

/// Render non-label cell contents into exactly `width` chars. An
/// `Empty` / `Value::Empty` / unevaluated formula produces blanks so
/// the result can slot directly into [`SpillSlot::Rendered`].
fn render_own_width(
    contents: &CellContents,
    width: usize,
    format: Format,
    intl: &International,
    excel_override: Option<&str>,
) -> String {
    // `(T)` Text format on a formula cell shows the formula source
    // (left-aligned, truncated to width) instead of the cached value
    // — per SPEC §12: "Text — show formula, not value." Non-formula
    // cells fall through to the normal numeric path; Text on a label
    // is unreachable here (labels go through SpillSlot::Label).
    if matches!(format.kind, l123_core::FormatKind::Text) {
        if let CellContents::Formula { expr, .. } = contents {
            return l123_core::cell_render::right_pad(expr, width, false);
        }
    }
    match contents {
        CellContents::Empty => " ".repeat(width),
        CellContents::Label { .. } => {
            // Labels are routed through SpillSlot::Label; we should
            // never hit this branch from the planner caller.
            " ".repeat(width)
        }
        CellContents::Constant(v) => render_value_in_cell(v, width, format, intl, excel_override)
            .unwrap_or_else(|| " ".repeat(width)),
        CellContents::Formula {
            cached_value: Some(v),
            ..
        } => render_value_in_cell(v, width, format, intl, excel_override)
            .unwrap_or_else(|| " ".repeat(width)),
        CellContents::Formula {
            cached_value: None, ..
        } => " ".repeat(width),
    }
}

impl App {
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

    /// Concatenate the symbols of one buffer column top-to-bottom.
    /// Mirror of `line_text` for column-major content like vertical
    /// Y-Axis titles.
    pub fn column_text(buf: &Buffer, x: u16) -> String {
        let mut s = String::new();
        for y in 0..buf.area.height {
            s.push_str(buf[(x, y)].symbol());
        }
        s.trim().to_string()
    }

    /// Find the buffer y coordinate for a given grid row, honoring
    /// frozen rows + the current row scroll.  Returns `None` when the
    /// row is outside the visible body region.
    pub(super) fn cell_y_in_buffer(&self, buf: &Buffer, row: u32) -> Option<u16> {
        // Body height: total - panel - column-header (1) - status line (1).
        let body_rows = buf.area.height.saturating_sub(PANEL_HEIGHT + 2);
        let layout = self.visible_row_layout(body_rows);
        let (_, y_off) = *layout.iter().find(|(r, _)| *r == row)?;
        Some(PANEL_HEIGHT + 1 + y_off)
    }

    /// Read back the rendered text of a single grid cell by address
    /// (`"A:B5"` or `"B5"`). Returns None if the cell is outside the
    /// current viewport.
    pub fn cell_rendered_text(&self, buf: &Buffer, addr: &str) -> Option<String> {
        let a = Address::parse(addr).ok()?;
        let y = self.cell_y_in_buffer(buf, a.row)?;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = self.visible_column_layout(content_width);
        let (_, x_off, w) = *layout.iter().find(|(c, _, _)| *c == a.col)?;
        let x0 = ROW_GUTTER + x_off;
        let mut s = String::with_capacity(w as usize);
        for i in 0..w {
            s.push_str(buf[(x0 + i, y)].symbol());
        }
        Some(s)
    }

    /// Read back the text-style override for a cell by address
    /// (`"A:B5"` or `"B5"`).  Returns `None` if the cell has no
    /// override (i.e. plain).  Used by the acceptance harness's
    /// `ASSERT_CELL_STYLE` directive.
    pub fn cell_text_style(&self, addr: &str) -> Option<TextStyle> {
        let a = Address::parse(addr).ok()?;
        self.wb().cell_text_styles.get(&a).copied()
    }

    /// Comma-joined table names on the given sheet, in the order
    /// L123 stores them after load.  Empty string when the sheet has
    /// no tables.  Used by the acceptance harness's `ASSERT_TABLES`
    /// directive to verify xlsx-imported tables survive load.
    pub fn table_names(&self, sheet_letter: char) -> String {
        let Some(idx) = letter_to_sheet_index(sheet_letter) else {
            return String::new();
        };
        let sid = SheetId(idx);
        match self.wb().tables.get(&sid) {
            Some(tables) => tables
                .iter()
                .map(|t| t.name.as_str())
                .collect::<Vec<_>>()
                .join(","),
            None => String::new(),
        }
    }

    /// Read back the rendered fg color of the sheet-letter cell on
    /// the status line, as `(r, g, b)`.  Returns `None` when the
    /// workbook has only one sheet (so no letter is shown), or when
    /// the letter renders in the default `DarkGray` (i.e. the sheet
    /// has no tab color).  Used by the acceptance harness's
    /// `ASSERT_STATUS_SHEET_FG` directive.
    pub fn status_sheet_letter_fg(&self, buf: &Buffer) -> Option<(u8, u8, u8)> {
        if self.wb().engine.sheet_count() <= 1 {
            return None;
        }
        // The status line is the last row of the rendered buffer.
        let y = buf.area.height.saturating_sub(1);
        // Locate the `[` that opens the sheet-indicator and read the
        // letter immediately after it — that's the char we tint.
        for x in 0..buf.area.width {
            if buf[(x, y)].symbol() == "[" && x + 1 < buf.area.width {
                match buf[(x + 1, y)].fg {
                    Color::Rgb(r, g, b) => return Some((r, g, b)),
                    _ => return None,
                }
            }
        }
        None
    }

    /// Read the rendered character at the rightmost column of a
    /// cell's slot.  Used by the acceptance harness to verify that
    /// xlsx-imported right borders paint the expected box-drawing
    /// glyph.  Returns `None` when the cell is outside the viewport.
    pub fn cell_right_edge_char(&self, buf: &Buffer, addr: &str) -> Option<String> {
        let a = Address::parse(addr).ok()?;
        let y = self.cell_y_in_buffer(buf, a.row)?;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = self.visible_column_layout(content_width);
        let (_, x_off, w) = *layout.iter().find(|(c, _, _)| *c == a.col)?;
        if w == 0 {
            return None;
        }
        let bx = ROW_GUTTER + x_off + w - 1;
        Some(buf[(bx, y)].symbol().to_string())
    }

    /// Read back the rendered foreground color at a cell's left-edge
    /// buffer position, as `(r, g, b)`.  Semantics mirror
    /// [`Self::cell_bg_rendered`]: `None` when the cell is off-screen
    /// or the fg renders with a non-RGB color (including the default
    /// terminal fg).  Used by the acceptance harness's
    /// `ASSERT_CELL_FG` directive.
    pub fn cell_fg_rendered(&self, buf: &Buffer, addr: &str) -> Option<(u8, u8, u8)> {
        let a = Address::parse(addr).ok()?;
        let y = self.cell_y_in_buffer(buf, a.row)?;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = self.visible_column_layout(content_width);
        let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == a.col)?;
        let x0 = ROW_GUTTER + x_off;
        match buf[(x0, y)].fg {
            Color::Rgb(r, g, b) => Some((r, g, b)),
            _ => None,
        }
    }

    /// Report whether the cell's left-edge character is rendered with
    /// the `CROSSED_OUT` modifier (xlsx strikethrough).  Returns
    /// `false` when the cell is outside the viewport.  Used by the
    /// acceptance harness's `ASSERT_CELL_STRIKE` directive.
    pub fn cell_strike_rendered(&self, buf: &Buffer, addr: &str) -> bool {
        let Ok(a) = Address::parse(addr) else {
            return false;
        };
        let Some(y) = self.cell_y_in_buffer(buf, a.row) else {
            return false;
        };
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = self.visible_column_layout(content_width);
        let Some((_, x_off, _)) = layout.iter().find(|(c, _, _)| *c == a.col).copied() else {
            return false;
        };
        let x0 = ROW_GUTTER + x_off;
        buf[(x0, y)].modifier.contains(Modifier::CROSSED_OUT)
    }

    /// Read back the rendered background color at a cell's left-edge
    /// buffer position, as `(r, g, b)`.  Returns `None` when the cell
    /// is outside the viewport, or when the rendered buffer has no
    /// explicit RGB background set (the terminal default shows through).
    /// Used by the acceptance harness's `ASSERT_CELL_BG` directive to
    /// verify that an xlsx-imported fill survives both the load and
    /// the grid-render pipeline.
    pub fn cell_bg_rendered(&self, buf: &Buffer, addr: &str) -> Option<(u8, u8, u8)> {
        let a = Address::parse(addr).ok()?;
        let y = self.cell_y_in_buffer(buf, a.row)?;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = self.visible_column_layout(content_width);
        let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == a.col)?;
        let x0 = ROW_GUTTER + x_off;
        match buf[(x0, y)].bg {
            Color::Rgb(r, g, b) => Some((r, g, b)),
            _ => None,
        }
    }
    pub(super) fn ideal_left_for_rightmost(&self, target_col: u16, content_width: u16) -> u16 {
        let sheet = self.wb().pointer.sheet;
        let mut total: u16 = 0;
        let mut col = target_col;
        loop {
            if !self.wb().hidden_cols.contains(&(sheet, col)) {
                let w = self.col_width_of(sheet, col) as u16;
                if w > 0 {
                    let new_total = total.saturating_add(w);
                    if new_total > content_width && col != target_col {
                        return col + 1;
                    }
                    total = new_total;
                }
            }
            if col == 0 {
                return 0;
            }
            col -= 1;
        }
    }

    // ---------------- rendering ----------------

    pub(super) fn render(&self, area: Rect, buf: &mut Buffer) {
        // Startup splash is full-screen — no control panel, no grid,
        // no status line. Draws until the first keystroke dismisses it.
        if let Some(info) = self.splash.as_ref() {
            self.render_splash(area, buf, info);
            return;
        }

        // GRAPH mode is full-screen too — 1-2-3 R3.4a hands the entire
        // display to the graph, no cell-pointer panel and no status
        // line. Esc dismisses it back to the framed layout.
        if self.mode == Mode::Graph {
            self.icon_panel_area.set(None);
            self.last_grid_area.set(None);
            self.render_graph_overlay(area, buf);
            return;
        }

        let chunks = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(PANEL_HEIGHT),
                Constraint::Min(1),
                Constraint::Length(1),
            ])
            .split(area);

        // Clear last frame's stashed icon-panel rect. render_icon_panel
        // re-sets it iff it actually draws the panel this frame.
        self.icon_panel_area.set(None);
        // Same pattern for the grid rect — render_grid re-sets it when
        // a grid is actually drawn (not during overlays).
        self.last_grid_area.set(None);

        self.render_control_panel(chunks[0], buf);
        // Split the middle area horizontally when the v3.1 WYSIWYG icon
        // panel is live. File-list and graph-view overlays take the full
        // width; everything else (grid, menu, point) keeps room for the
        // panel on the right.
        let (main_area, icon_area) = self.split_for_icon_panel(chunks[1]);
        if self.help.is_some() {
            self.render_help_overlay(chunks[1], buf);
        } else if self.sqlite_table_picker.is_some() {
            self.render_sqlite_table_picker_overlay(chunks[1], buf);
        } else if self.external_list.is_some() {
            self.render_external_list_overlay(chunks[1], buf);
        } else if self.name_list.is_some() {
            self.render_name_list_overlay(chunks[1], buf);
        } else if self.file_list.is_some() {
            self.render_file_list_overlay(chunks[1], buf);
        } else if self.mode == Mode::Stat {
            self.render_stat_overlay(chunks[1], buf);
        } else if self.is_in_graph_menu() {
            self.render_graph_settings_overlay(chunks[1], buf);
        } else {
            self.render_grid(main_area, buf);
            if let Some(area) = icon_area {
                self.render_icon_panel(area, buf);
            }
        }
        self.render_status(chunks[2], buf);
    }

    pub(super) fn render_splash(&self, area: Rect, buf: &mut Buffer, info: &SplashInfo) {
        // Classic DOS VGA "cyan" (palette index 3) is #00AAAA — the
        // shade the 1-2-3 R3.4a welcome screen fills its field with.
        // Using explicit RGB triples keeps the colors stable across
        // terminals that remap their ANSI slots.
        const TEAL: Color = Color::Rgb(0, 170, 170);
        const BLACK: Color = Color::Rgb(0, 0, 0);
        let teal_bg = Style::default().bg(TEAL);
        for y in 0..area.height {
            for x in 0..area.width {
                buf[(area.x + x, area.y + y)].set_style(teal_bg);
            }
        }

        let title_style = Style::default()
            .bg(BLACK)
            .fg(Color::White)
            .add_modifier(Modifier::BOLD);
        let banner = [
            Line::from(""),
            Line::from(Span::styled("l123", title_style)),
            Line::from(Span::styled(
                format!("Release {}", env!("CARGO_PKG_VERSION")),
                title_style,
            )),
            Line::from(""),
            Line::from("A terminal spreadsheet in the 1-2-3 tradition."),
            Line::from(""),
            Line::from("Copyright 2026 Duane Moore"),
            Line::from("All Rights Reserved."),
            Line::from(""),
        ];

        let banner_w = 60.min(area.width.saturating_sub(4));
        let banner_h = banner.len() as u16 + 2;
        if area.width < banner_w + 2 || area.height < banner_h + 6 {
            return;
        }

        let banner_x = area.x + (area.width - banner_w) / 2;
        let banner_y = area.y + 2;
        let banner_rect = Rect::new(banner_x, banner_y, banner_w, banner_h);
        let body_style = Style::default().bg(BLACK).fg(TEAL);
        let banner_block = Block::default().borders(Borders::ALL).style(body_style);
        let banner_inner = banner_block.inner(banner_rect);
        banner_block.render(banner_rect, buf);
        Paragraph::new(banner.to_vec())
            .alignment(ratatui::layout::Alignment::Center)
            .style(body_style)
            .render(banner_inner, buf);

        let licensing_y = banner_y + banner_h + 2;
        if licensing_y + 4 >= area.y + area.height {
            return;
        }

        const USER_LABEL: &str = "User name:     ";
        const ORG_LABEL: &str = "Organization:  ";
        const FOOTER: [&str; 3] = [
            "Use, duplication, or sale of this product, except as described",
            "in the project's license agreement, is strictly prohibited.",
            "Violators may be prosecuted.",
        ];
        const HEADING: &str = "LICENSING INFORMATION:";

        // Width of the license block is the longest line it contains:
        // rows and footer anchor the left edge so the labels and legal
        // text read as a single centered column.
        let user_w = USER_LABEL.len() + info.user.chars().count();
        let org_w = ORG_LABEL.len() + info.organization.chars().count();
        let footer_w = FOOTER.iter().map(|s| s.chars().count()).max().unwrap_or(0);
        let content_w = [HEADING.chars().count(), user_w, org_w, footer_w]
            .into_iter()
            .max()
            .unwrap_or(0) as u16;
        let block_w = content_w.min(area.width.saturating_sub(4));
        let block_x = area.x + area.width.saturating_sub(block_w) / 2;

        let heading = Paragraph::new(Line::from(Span::styled(
            HEADING,
            Style::default()
                .bg(TEAL)
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD),
        )))
        .alignment(ratatui::layout::Alignment::Center);
        heading.render(Rect::new(block_x, licensing_y, block_w, 1), buf);

        let label_style = Style::default().bg(TEAL).fg(BLACK);
        let value_style = Style::default()
            .bg(TEAL)
            .fg(Color::Yellow)
            .add_modifier(Modifier::BOLD);
        let rows = [
            Line::from(vec![
                Span::styled(USER_LABEL, label_style),
                Span::styled(info.user.clone(), value_style),
            ]),
            Line::from(vec![
                Span::styled(ORG_LABEL, label_style),
                Span::styled(info.organization.clone(), value_style),
            ]),
        ];
        Paragraph::new(rows.to_vec()).render(Rect::new(block_x, licensing_y + 2, block_w, 2), buf);

        // Bottom legal notice, in black on teal to match the 1-2-3
        // R3.4a welcome screen. Rendered only when there is room
        // beneath the licensing block.
        let footer_y = licensing_y + 5;
        if footer_y + 3 >= area.y + area.height {
            return;
        }
        let footer = Paragraph::new(FOOTER.iter().map(|s| Line::from(*s)).collect::<Vec<_>>())
            .style(Style::default().bg(TEAL).fg(BLACK));
        footer.render(Rect::new(block_x, footer_y, block_w, 3), buf);
    }

    /// The v3.1 manual shows the icon panel occupying the right edge
    /// of the worksheet area. Three columns is enough for readable
    /// icons at terminal resolution; less than 20 columns of grid
    /// would be awkward, so very narrow terminals hide the panel.
    const ICON_PANEL_COLS: u16 = 3;
    const ICON_PANEL_MIN_GRID_COLS: u16 = 20;

    pub(super) fn split_for_icon_panel(&self, area: Rect) -> (Rect, Option<Rect>) {
        if self.icon_panel.is_none()
            || self.mode == Mode::Graph
            || self.file_list.is_some()
            || self.name_list.is_some()
            || self.help.is_some()
            || area.width < Self::ICON_PANEL_MIN_GRID_COLS + Self::ICON_PANEL_COLS
        {
            return (area, None);
        }
        let main_width = area.width - Self::ICON_PANEL_COLS;
        let main = Rect::new(area.x, area.y, main_width, area.height);
        let icons = Rect::new(
            area.x + main_width,
            area.y,
            Self::ICON_PANEL_COLS,
            area.height,
        );
        (main, Some(icons))
    }

    pub(super) fn render_icon_panel(&self, area: Rect, buf: &mut Buffer) {
        let (Some(picker), Some(img)) = (self.image_picker.as_ref(), self.icon_panel.as_ref())
        else {
            return;
        };
        if picker.protocol_type() == ProtocolType::Halfblocks {
            return;
        }
        if let Ok(protocol) = picker.new_protocol(img.clone(), area, Resize::Fit(None)) {
            // `Resize::Fit` ceilings the rendered cell rect, but the
            // image itself only fills the pre-ceiling pixel size. Hit-
            // testing must use the true rendered pixel height so each
            // mouse cell maps to the icon containing the majority of
            // its pixels — otherwise the half-cell of bottom slack
            // accumulates as a downward slot drift, sliding the
            // tooltips ahead of the icons the user sees.
            let font = picker.font_size();
            let rendered_px_h = Self::fit_rendered_px_h(area, font);
            let rendered = protocol.area();
            Image::new(&protocol).render(area, buf);
            self.icon_panel_area.set(Some(IconPanelGeom {
                rect: Rect::new(
                    area.x,
                    area.y,
                    rendered.width.min(area.width),
                    rendered.height.min(area.height),
                ),
                rendered_px_h,
                font_px_h: font.1,
            }));
        }
    }

    /// Reproduce ratatui-image's `Resize::Fit` math to recover the
    /// exact pixel height of the rendered icon-panel image. The image
    /// has a 1:17 aspect; whichever of width or height is the binding
    /// constraint determines the scale, and the height pixels is the
    /// PNG height times that scale.
    pub(super) fn fit_rendered_px_h(area: Rect, font: (u16, u16)) -> u32 {
        let img_w = l123_graph::icons::ICON_PANEL_WIDTH_PX as u64;
        let img_h = l123_graph::icons::ICON_PANEL_HEIGHT_PX as u64;
        let avail_w = area.width as u64 * font.0 as u64;
        let avail_h = area.height as u64 * font.1 as u64;
        let nw = avail_w.min(img_w);
        let nh = avail_h.min(img_h);
        // ratio_w = nw/img_w, ratio_h = nh/img_h; the smaller one wins.
        // Compare cross-products to avoid float arithmetic.
        if nw * img_h <= nh * img_w {
            (nw * img_h / img_w) as u32
        } else {
            nh as u32
        }
    }
    pub(super) fn render_graph_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(overlay) = self.graph_view.as_ref() else {
            self.render_grid(area, buf);
            return;
        };
        // Graphical path: render the chart at the area's pixel size
        // so the image and area aspect ratios match — `Resize::Fit`
        // then scales 1:1 instead of leaving letterbox bars on a
        // terminal whose aspect doesn't match plotters' default 4:3.
        // Cached by pixel dims via `overlay.img_cache` so static
        // graph view doesn't re-rasterize on every 100ms event-loop
        // tick. Protocol creation can fail (encoding error, terminal
        // query hiccup); on any failure we fall through to the
        // unicode path so the user still sees something.
        if let Some(picker) = self.image_picker.as_ref() {
            if picker.protocol_type() != ProtocolType::Halfblocks && !overlay.values.is_empty() {
                let (cell_w, cell_h) = picker.font_size();
                let target_w = (area.width as u32) * (cell_w as u32);
                let target_h = (area.height as u32) * (cell_h as u32);
                if target_w > 0 && target_h > 0 {
                    let img = self.graph_image_for_dims(overlay, target_w, target_h);
                    if let Some(img) = img {
                        if let Ok(protocol) = picker.new_protocol(img, area, Resize::Fit(None)) {
                            Image::new(&protocol).render(area, buf);
                            return;
                        }
                    }
                }
            }
        }
        l123_graph::render_unicode(&self.wb().current_graph, &overlay.values, area, buf);
    }

    /// Pull a graph raster from `overlay.img_cache` if its dims match
    /// `(w, h)`; otherwise rasterize fresh and update the cache.
    /// Returns the image to hand to the picker. The clone is cheap —
    /// `DynamicImage` is reference-counted internally for the buffer.
    fn graph_image_for_dims(
        &self,
        overlay: &GraphOverlay,
        w: u32,
        h: u32,
    ) -> Option<image::DynamicImage> {
        if let Some((cw, ch, img)) = overlay.img_cache.borrow().as_ref() {
            if *cw == w && *ch == h {
                return Some(img.clone());
            }
        }
        let img = l123_graph::render_dynamic_image_sized(
            &self.wb().current_graph,
            &overlay.values,
            w,
            h,
        )?;
        *overlay.img_cache.borrow_mut() = Some((w, h, img.clone()));
        Some(img)
    }

    /// True when the user has descended into the top-level `/Graph`
    /// menu and is still navigating it. Drives the Graph Settings
    /// overlay; cleared when the menu is dismissed (Esc / Quit) or
    /// when a leaf action takes the app back to READY.
    pub(super) fn is_in_graph_menu(&self) -> bool {
        if self.mode != Mode::Menu {
            return false;
        }
        let Some(state) = self.menu.as_ref() else {
            return false;
        };
        state.override_root.is_none() && state.path.first() == Some(&'G')
    }

    /// `/Graph` settings sheet (R3.1 Reference p. 2-230). Four panels:
    /// Graph Type (top-left), Data Ranges (bottom-left), Graph Type
    /// Features (center), Options (right). Pure projection of the
    /// current graph; menu commands mutate the graph and the next
    /// frame redraws.
    pub(super) fn render_graph_settings_overlay(&self, area: Rect, buf: &mut Buffer) {
        const GREEN: Color = Color::Rgb(0, 170, 85);
        const BRIGHT: Color = Color::Rgb(120, 255, 120);
        const BLACK: Color = Color::Rgb(0, 0, 0);
        let text_style = Style::default().bg(BLACK).fg(GREEN);
        let on_style = Style::default().bg(BLACK).fg(BRIGHT);

        for y in 0..area.height {
            for x in 0..area.width {
                buf[(area.x + x, area.y + y)].set_style(Style::default().bg(BLACK));
            }
        }

        let outer = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Graph Settings")
            .title_alignment(ratatui::layout::Alignment::Center)
            .style(text_style);
        let inner = outer.inner(area);
        outer.render(area, buf);

        let cols = Layout::default()
            .direction(Direction::Horizontal)
            .constraints([
                Constraint::Length(22),
                Constraint::Length(28),
                Constraint::Min(20),
            ])
            .split(inner);

        let left = Layout::default()
            .direction(Direction::Vertical)
            .constraints([Constraint::Length(7), Constraint::Min(10)])
            .split(cols[0]);

        self.render_graph_panel_type(left[0], buf, text_style, on_style);
        self.render_graph_panel_ranges(left[1], buf, text_style);
        self.render_graph_panel_features(cols[1], buf, text_style, on_style);
        self.render_graph_panel_options(cols[2], buf, text_style, on_style);
    }

    fn render_graph_panel_type(
        &self,
        area: Rect,
        buf: &mut Buffer,
        text: Style,
        on: Style,
    ) {
        let block = Block::default()
            .borders(Borders::ALL)
            .border_style(text)
            .title("Graph Type")
            .style(text);
        let inner = block.inner(area);
        block.render(area, buf);

        let g = &self.wb().current_graph;
        let mark = |selected: bool, label: &str| -> Span<'static> {
            if selected {
                Span::styled(format!("x {label}"), on)
            } else {
                Span::styled(format!("  {label}"), text)
            }
        };
        let t = g.graph_type;
        use l123_graph::GraphType as GT;
        let lines = vec![
            Line::from(vec![
                mark(t == GT::Line, "Line       "),
                mark(t == GT::Pie, "Pie"),
            ]),
            Line::from(vec![
                mark(t == GT::Bar, "Bar        "),
                mark(t == GT::HLCO, "HLCO"),
            ]),
            Line::from(vec![
                mark(t == GT::XY, "XY         "),
                mark(t == GT::Mixed, "Mixed"),
            ]),
            Line::from(mark(t == GT::Stack, "Stacked Bar")),
        ];
        Paragraph::new(lines).style(text).render(inner, buf);
    }

    fn render_graph_panel_ranges(&self, area: Rect, buf: &mut Buffer, text: Style) {
        let block = Block::default()
            .borders(Borders::ALL)
            .border_style(text)
            .title("Data Ranges")
            .style(text);
        let inner = block.inner(area);
        block.render(area, buf);

        let g = &self.wb().current_graph;
        let row = |label: &str, r: Option<l123_core::Range>| -> Line<'static> {
            let val = match r {
                Some(rr) => format!("{}..{}", rr.start.display_full(), rr.end.display_full()),
                None => String::new(),
            };
            Line::from(format!("{label}: {val}"))
        };
        let lines = vec![
            row("X", g.x),
            row("A", g.data[0]),
            row("B", g.data[1]),
            row("C", g.data[2]),
            row("D", g.data[3]),
            row("E", g.data[4]),
            row("F", g.data[5]),
        ];
        Paragraph::new(lines).style(text).render(inner, buf);
    }

    fn render_graph_panel_features(
        &self,
        area: Rect,
        buf: &mut Buffer,
        text: Style,
        on: Style,
    ) {
        let block = Block::default()
            .borders(Borders::ALL)
            .border_style(text)
            .title("Graph Type Features")
            .style(text);
        let inner = block.inner(area);
        block.render(area, buf);

        let f = &self.wb().current_graph.features;
        let mark = |selected: bool, label: &str| -> Span<'static> {
            if selected {
                Span::styled(format!("x {label}"), on)
            } else {
                Span::styled(format!("  {label}"), text)
            }
        };
        let yax = |i: usize| match f.y_axis[i] {
            l123_graph::YAxis::First => "Y",
            l123_graph::YAxis::Second => "2Y",
        };
        // Two-column layout inside the panel: a fixed-width left column
        // for the Y/2Y per-series indicators, plus a right column for
        // orientation and frame controls. `slot` formats the left
        // column at a stable width so the right column lines up.
        let slot = |letter: char, axis: &str| format!("  {letter}: {axis:<5}");

        let lines = vec![
            Line::from(vec![
                Span::raw("Y/2Y      "),
                mark(f.orientation == l123_graph::Orientation::Vertical, "Vertical"),
            ]),
            Line::from(vec![
                Span::raw(slot('A', yax(0))),
                mark(
                    f.orientation == l123_graph::Orientation::Horizontal,
                    "Horizontal",
                ),
            ]),
            Line::from(slot('B', yax(1))),
            Line::from(vec![Span::raw(slot('C', yax(2))), Span::raw("Frame")]),
            Line::from(vec![Span::raw(slot('D', yax(3))), mark(f.frame.left, "Left")]),
            Line::from(vec![Span::raw(slot('E', yax(4))), mark(f.frame.right, "Right")]),
            Line::from(vec![Span::raw(slot('F', yax(5))), mark(f.frame.top, "Top")]),
            Line::from(vec![Span::raw("          "), mark(f.frame.bottom, "Bottom")]),
            Line::from(vec![Span::raw("          "), mark(f.frame.y_axis, "Y-axis")]),
            Line::from(""),
            Line::from(mark(f.stacked, "Stack data ranges")),
            Line::from(mark(f.percent, "Percentage")),
            Line::from(mark(f.drop_shadow, "Drop-shadow")),
            Line::from(mark(f.three_d, "3-D")),
            Line::from(mark(f.table, "Table")),
        ];
        Paragraph::new(lines).style(text).render(inner, buf);
    }

    fn render_graph_panel_options(
        &self,
        area: Rect,
        buf: &mut Buffer,
        text: Style,
        on: Style,
    ) {
        let block = Block::default()
            .borders(Borders::ALL)
            .border_style(text)
            .title("Options")
            .style(text);
        let inner = block.inner(area);
        block.render(area, buf);

        let o = &self.wb().current_graph.options;
        let mark = |selected: bool, label: &str| -> Span<'static> {
            if selected {
                Span::styled(format!("x {label}"), on)
            } else {
                Span::styled(format!("  {label}"), text)
            }
        };

        let lines = vec![
            Line::from(mark(o.color, "Colors on")),
            Line::from(""),
            Line::from(Span::raw("Grid Lines")),
            Line::from(mark(o.grid.horizontal, "Horizontal")),
            Line::from(mark(o.grid.vertical, "Vertical")),
            Line::from(mark(
                matches!(
                    o.grid.y_axis,
                    Some(l123_graph::GridYAxisOrigin::First | l123_graph::GridYAxisOrigin::Both),
                ),
                "Y-Axis",
            )),
            Line::from(mark(
                matches!(
                    o.grid.y_axis,
                    Some(l123_graph::GridYAxisOrigin::Second | l123_graph::GridYAxisOrigin::Both),
                ),
                "2Y-Axis",
            )),
        ];
        Paragraph::new(lines).style(text).render(inner, buf);
    }

    pub(super) fn render_stat_overlay(&self, area: Rect, buf: &mut Buffer) {
        if self.stat_view == StatView::Defaults {
            self.render_defaults_overlay(area, buf);
            return;
        }
        // Monochrome CRT look: green-on-black, like the R3.4a status
        // page. Explicit RGB so terminals that remap their ANSI green
        // slot don't lose the effect.
        const GREEN: Color = Color::Rgb(0, 170, 85);
        const BLACK: Color = Color::Rgb(0, 0, 0);
        let text_style = Style::default().bg(BLACK).fg(GREEN);

        for y in 0..area.height {
            for x in 0..area.width {
                buf[(area.x + x, area.y + y)].set_style(Style::default().bg(BLACK));
            }
        }

        let outer = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Worksheet Status")
            .title_alignment(ratatui::layout::Alignment::Center)
            .style(text_style);
        let outer_inner = outer.inner(area);
        outer.render(area, buf);

        // Upper band: two side-by-side sub-boxes (Recalculation + Cell
        // display). Lower band: the environment readout — the same
        // flat list the R3.4a status page used (memory, processor,
        // protection, circular reference). International settings
        // belong on `/Worksheet Global Default Other International`,
        // not here; the original DOS status page never showed them.
        let band = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(6),
                Constraint::Length(1),
                Constraint::Min(1),
            ])
            .split(outer_inner);
        let split = Layout::default()
            .direction(Direction::Horizontal)
            .constraints([Constraint::Percentage(45), Constraint::Percentage(55)])
            .split(band[0]);

        let recalc = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Recalculation")
            .style(text_style);
        let recalc_inner = recalc.inner(split[0]);
        recalc.render(split[0], buf);
        let method = match self.recalc_mode {
            RecalcMode::Automatic => "Automatic",
            RecalcMode::Manual => "Manual",
        };
        let order = self.recalc_order.label();
        let iterations = self.recalc_iterations;
        Paragraph::new(vec![
            Line::from(format!("Method:     {method}")),
            Line::from(format!("Order:      {order}")),
            Line::from(format!("Iterations: {iterations}")),
        ])
        .style(text_style)
        .render(recalc_inner, buf);

        let cell = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Cell display")
            .style(text_style);
        let cell_inner = cell.inner(split[1]);
        cell.render(split[1], buf);
        let prefix = self.default_label_prefix.char();
        let col_width = self.wb().default_col_width;
        let zero = self.zero_display.label();
        let global_format = self.wb().global_format;
        Paragraph::new(vec![
            Line::from(format!("Format:       {global_format}")),
            Line::from(format!("Label prefix: {prefix}")),
            Line::from(format!("Column width: {col_width}")),
            Line::from(format!("Zero setting: {zero}")),
        ])
        .style(text_style)
        .render(cell_inner, buf);

        let info = crate::sysinfo::SysInfo::probe();
        let mem_free = info
            .memory_free
            .map(crate::sysinfo::format_bytes)
            .unwrap_or_else(|| "—".to_string());
        let mem_total = info
            .memory_total
            .map(crate::sysinfo::format_bytes)
            .unwrap_or_else(|| "—".to_string());
        let pad = 20;
        Paragraph::new(vec![
            Line::from(""),
            Line::from(format!(
                "{label:<pad$}{mem_free} bytes",
                label = "Available memory:"
            )),
            Line::from(format!(
                "{label:<pad$}{mem_total} bytes",
                label = "        out of:"
            )),
            Line::from(""),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Processor:",
                value = info.processor
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Math coprocessor:",
                value = info.coprocessor
            )),
            Line::from(""),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Global protection:",
                value = if self.global_protection { "On" } else { "Off" }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Circular reference:",
                value = match self.first_circular_reference() {
                    Some(addr) => addr.display_full(),
                    None => "(None)".to_string(),
                }
            )),
        ])
        .style(text_style)
        .render(band[2], buf);
    }

    pub(super) fn render_defaults_overlay(&self, area: Rect, buf: &mut Buffer) {
        const GREEN: Color = Color::Rgb(0, 170, 85);
        const BLACK: Color = Color::Rgb(0, 0, 0);
        let text_style = Style::default().bg(BLACK).fg(GREEN);

        for y in 0..area.height {
            for x in 0..area.width {
                buf[(area.x + x, area.y + y)].set_style(Style::default().bg(BLACK));
            }
        }

        let outer = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Global Default Settings")
            .title_alignment(ratatui::layout::Alignment::Center)
            .style(text_style);
        let outer_inner = outer.inner(area);
        outer.render(area, buf);

        let band = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(13),
                Constraint::Length(1),
                Constraint::Min(1),
            ])
            .split(outer_inner);

        let split = Layout::default()
            .direction(Direction::Horizontal)
            .constraints([Constraint::Percentage(50), Constraint::Percentage(50)])
            .split(band[0]);

        let printer = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Printer")
            .style(text_style);
        let printer_inner = printer.inner(split[0]);
        printer.render(split[0], buf);
        let d = &self.defaults;
        let on_off = |b: bool| if b { "Yes" } else { "No" };
        let pad = 11;
        Paragraph::new(vec![
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Interface:",
                value = d.printer_interface
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "AutoLf:",
                value = on_off(d.printer_autolf)
            )),
            Line::from(format!(
                "{label:<pad$}L{} R{} T{} B{}",
                d.printer_left,
                d.printer_right,
                d.printer_top,
                d.printer_bottom,
                label = "Margins:",
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Pg-Length:",
                value = d.printer_pg_length
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Wait:",
                value = on_off(d.printer_wait)
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Setup:",
                value = if d.printer_setup.is_empty() {
                    "(none)"
                } else {
                    &d.printer_setup
                }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Name:",
                value = if d.printer_name.is_empty() {
                    "(default)"
                } else {
                    &d.printer_name
                }
            )),
        ])
        .style(text_style)
        .render(printer_inner, buf);

        let other = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("Files & Graph")
            .style(text_style);
        let other_inner = other.inner(split[1]);
        other.render(split[1], buf);
        Paragraph::new(vec![
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Dir:",
                value = if d.default_dir.is_empty() {
                    "(unset)"
                } else {
                    &d.default_dir
                }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Temp:",
                value = if d.temp_dir.is_empty() {
                    "(unset)"
                } else {
                    &d.temp_dir
                }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Ext Save:",
                value = if d.ext_save.is_empty() {
                    "(unset)"
                } else {
                    &d.ext_save
                }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Ext List:",
                value = if d.ext_list.is_empty() {
                    "(any)"
                } else {
                    &d.ext_list
                }
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Autoexec:",
                value = on_off(d.autoexec)
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Graph Grp:",
                value = d.graph_group.label()
            )),
            Line::from(format!(
                "{label:<pad$}{value}",
                label = "Graph Save:",
                value = d.graph_save.label()
            )),
        ])
        .style(text_style)
        .render(other_inner, buf);

        let cnf_path = crate::config::default_config_path()
            .map(|p| p.display().to_string())
            .unwrap_or_else(|| "<$HOME not set>".into());
        Paragraph::new(vec![
            Line::from(""),
            Line::from(format!("Update writes: {cnf_path}")),
            Line::from(""),
            Line::from("Press any key to return to READY."),
        ])
        .style(text_style)
        .render(band[2], buf);
    }

    pub(super) fn render_control_panel(&self, area: Rect, buf: &mut Buffer) {
        let block = Block::default().borders(Borders::BOTTOM);
        let inner = block.inner(area);
        block.render(area, buf);

        // Line 1: "<addr>: [(fmt) [Wn] {Style}] <readout>" left; mode
        // indicator right.  The `{Style}` marker comes after numeric
        // format/width tags so readers scan parens → brackets → braces
        // in a consistent order.
        let readout = self.cell_readout_for_line1();
        let format_tag = self.format_tag_for_line1();
        let width_tag = self.width_tag_for_line1();
        let style_marker = self.text_style_marker_for_line1();
        let mut tags: Vec<&str> = Vec::new();
        if !format_tag.is_empty() {
            tags.push(&format_tag);
        }
        if !width_tag.is_empty() {
            tags.push(&width_tag);
        }
        if !style_marker.is_empty() {
            tags.push(&style_marker);
        }
        let left = if readout.is_empty() && tags.is_empty() {
            format!(" {}: ", self.wb().pointer.display_full())
        } else if tags.is_empty() {
            format!(" {}: {}", self.wb().pointer.display_full(), readout)
        } else if readout.is_empty() {
            format!(" {}: {}", self.wb().pointer.display_full(), tags.join(" "))
        } else {
            format!(
                " {}: {} {}",
                self.wb().pointer.display_full(),
                tags.join(" "),
                readout
            )
        };
        let mode_str = self.mode.indicator();
        let pad = (area.width as usize).saturating_sub(left.chars().count() + mode_str.len() + 1);
        let line1 = Line::from(vec![
            Span::raw(left),
            Span::raw(" ".repeat(pad)),
            Span::styled(
                mode_str,
                Style::default()
                    .fg(Color::Yellow)
                    .add_modifier(Modifier::BOLD),
            ),
            Span::raw(" "),
        ]);

        // Line 2 & 3 depend on mode. F3 NAMES, file-list and
        // save-confirm take absolute precedence (they own the
        // keyboard); then a command-argument prompt; then
        // mode-specific rendering.
        let (line2, line3) = if self.sqlite_table_picker.is_some() {
            self.render_sqlite_table_picker_lines()
        } else if self.external_list.is_some() {
            self.render_external_list_lines()
        } else if self.name_list.is_some() {
            self.render_name_list_lines()
        } else if self.file_list.is_some() {
            self.render_file_list_lines()
        } else if self.save_confirm.is_some() {
            self.render_save_confirm_lines()
        } else if self.erase_confirm.is_some() {
            self.render_erase_confirm_lines()
        } else if self.custom_menu.is_some() {
            self.render_custom_menu_lines()
        } else if let Some(msg) = self.error_message.as_ref() {
            (
                Line::from(format!(" {msg}")),
                Line::from(" Press ESC or ENTER to clear"),
            )
        } else if let Some(p) = self.prompt.as_ref() {
            (
                Line::from(format!(" {}", p.buffer)),
                Line::from(format!(" {}", p.label)),
            )
        } else {
            match self.mode {
                Mode::Menu => self.render_menu_lines(),
                Mode::Point => self.render_point_lines(),
                Mode::Wait => {
                    let l3 = match self.pending_async_op.as_ref() {
                        Some(op) => Line::from(op.render_line3()),
                        None => Line::from(""),
                    };
                    (Line::from(""), l3)
                }
                _ => {
                    let l2 = match self.entry.as_ref() {
                        Some(e) => render_entry_l2(e),
                        None => Line::from(""),
                    };
                    // Icon-hover description — gated on entry being
                    // idle so an in-progress value/label/edit never
                    // has its canonical line-3 space (formula preview)
                    // clobbered by a stray mouse-over.
                    //
                    // Cell-comment readout — fallback when the pointer
                    // lands on a cell with an xlsx-imported comment
                    // and no other claimant for line 3.  Format is
                    // `<author>: <text>`, truncated to fit the panel.
                    let l3 = match (self.entry.as_ref(), self.hovered_icon) {
                        (None, Some((panel, slot))) => {
                            Line::from(format!(" {}", l123_graph::slot_description(panel, slot)))
                        }
                        (None, None) => match self.wb().comments.get(&self.wb().pointer) {
                            Some(c) => Line::from(format!(" {}", c.summary())),
                            None => Line::from(""),
                        },
                        _ => Line::from(""),
                    };
                    (l2, l3)
                }
            }
        };

        Paragraph::new(vec![line1, line2, line3]).render(inner, buf);
    }

    pub(super) fn render_file_list_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(fl) = self.file_list.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        // Panel lines are just the header + highlighted path / count.
        // The full picker lives in the overlay below.
        let header = match fl.kind {
            FileListKind::Worksheet => " File List — Worksheet",
            FileListKind::Active => " File List — Active",
            FileListKind::Other => " File List — Other",
        };
        let tail = if fl.entries.is_empty() {
            match fl.kind {
                FileListKind::Worksheet => " (no worksheet files in directory)".to_string(),
                FileListKind::Active => " (no active file)".to_string(),
                FileListKind::Other => " (no files in directory)".to_string(),
            }
        } else {
            format!(
                " {}   [{}/{}]   Enter: retrieve  Esc: cancel",
                fl.entries
                    .get(fl.highlight)
                    .map(|p| p.display().to_string())
                    .unwrap_or_default(),
                fl.highlight + 1,
                fl.entries.len(),
            )
        };
        (Line::from(header), Line::from(tail))
    }

    /// Draw the scrollable file picker in `area`. Each row shows the
    /// file name and, when available, its size (in bytes). Highlighted
    /// row is reverse-video.
    pub(super) fn render_file_list_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(fl) = self.file_list.as_ref() else {
            return;
        };
        let width = area.width as usize;
        let rows = area.height as usize;
        if rows == 0 || width == 0 {
            return;
        }

        let size_col_width: usize = 10;
        let name_col_width = width.saturating_sub(size_col_width + 3);

        // Header row.
        let header = format_file_list_row("NAME", "SIZE", name_col_width, size_col_width, width);
        set_line(buf, area.x, area.y, &header, area.width, Style::default());

        if fl.entries.is_empty() {
            let empty_msg = match fl.kind {
                FileListKind::Worksheet => "(no worksheet files in directory)",
                FileListKind::Active => "(no active file)",
                FileListKind::Other => "(no files in directory)",
            };
            set_line(
                buf,
                area.x,
                area.y + 1,
                empty_msg,
                area.width,
                Style::default(),
            );
            return;
        }

        let visible_rows = rows.saturating_sub(1);
        let start = fl.view_offset.min(fl.entries.len());
        let end = (start + visible_rows).min(fl.entries.len());
        for (i, path) in fl.entries[start..end].iter().enumerate() {
            let idx = start + i;
            let name = path
                .file_name()
                .map(|n| n.to_string_lossy().into_owned())
                .unwrap_or_default();
            let size = std::fs::metadata(path)
                .map(|m| format_size(m.len()))
                .unwrap_or_default();
            let row = format_file_list_row(&name, &size, name_col_width, size_col_width, width);
            let style = if idx == fl.highlight {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            set_line(buf, area.x, area.y + 1 + i as u16, &row, area.width, style);
        }
    }

    pub(super) fn render_name_list_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(nl) = self.name_list.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let header = " Name List";
        let tail = if nl.entries.is_empty() {
            " (no defined names)".to_string()
        } else {
            let (name, range) = &nl.entries[nl.highlight];
            format!(
                " {}  {}..{}   [{}/{}]   Enter: select  Esc: cancel",
                name,
                range.start.display_full(),
                range.end.display_full(),
                nl.highlight + 1,
                nl.entries.len(),
            )
        };
        (Line::from(header), Line::from(tail))
    }

    /// Draw the scrollable name picker in `area`. Each row shows the
    /// range name and the range it points at. Highlighted row is
    /// reverse-video.
    pub(super) fn render_name_list_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(nl) = self.name_list.as_ref() else {
            return;
        };
        let width = area.width as usize;
        let rows = area.height as usize;
        if rows == 0 || width == 0 {
            return;
        }

        let range_col_width: usize = 24;
        let name_col_width = width.saturating_sub(range_col_width + 3);

        let header = format_file_list_row("NAME", "RANGE", name_col_width, range_col_width, width);
        set_line(buf, area.x, area.y, &header, area.width, Style::default());

        if nl.entries.is_empty() {
            set_line(
                buf,
                area.x,
                area.y + 1,
                "(no defined names)",
                area.width,
                Style::default(),
            );
            return;
        }

        let visible_rows = rows.saturating_sub(1);
        let start = nl.view_offset.min(nl.entries.len());
        let end = (start + visible_rows).min(nl.entries.len());
        for (i, (name, range)) in nl.entries[start..end].iter().enumerate() {
            let idx = start + i;
            let range_str = format!(
                "{}..{}",
                range.start.display_full(),
                range.end.display_full()
            );
            let row =
                format_file_list_row(name, &range_str, name_col_width, range_col_width, width);
            let style = if idx == nl.highlight {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            set_line(buf, area.x, area.y + 1 + i as u16, &row, area.width, style);
        }
    }

    /// Panel lines 2 / 3 while the `/File Import Sqlite` table picker
    /// is on screen. Line 2 names the picker; line 3 echoes the
    /// highlighted table + key hints.
    pub(super) fn render_sqlite_table_picker_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(p) = self.sqlite_table_picker.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let basename = p
            .path
            .file_name()
            .map(|s| s.to_string_lossy().into_owned())
            .unwrap_or_else(|| p.path.display().to_string());
        let header = format!(" Pick table from {basename}");
        let tail = if p.tables.is_empty() {
            " (no user tables)".to_string()
        } else {
            format!(
                " {}   [{}/{}]   Enter: load  Esc: cancel",
                p.tables[p.highlight],
                p.highlight + 1,
                p.tables.len(),
            )
        };
        (Line::from(header), Line::from(tail))
    }

    /// Draw the sqlite table picker — one row per table, highlight in
    /// reverse video. Mirrors `render_name_list_overlay` but without
    /// the secondary RANGE column.
    pub(super) fn render_sqlite_table_picker_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(p) = self.sqlite_table_picker.as_ref() else {
            return;
        };
        let width = area.width as usize;
        let rows = area.height as usize;
        if rows == 0 || width == 0 {
            return;
        }
        set_line(buf, area.x, area.y, "TABLE", area.width, Style::default());
        if p.tables.is_empty() {
            set_line(
                buf,
                area.x,
                area.y + 1,
                "(no user tables)",
                area.width,
                Style::default(),
            );
            return;
        }
        let visible_rows = rows.saturating_sub(1);
        let start = p.view_offset.min(p.tables.len());
        let end = (start + visible_rows).min(p.tables.len());
        for (i, name) in p.tables[start..end].iter().enumerate() {
            let idx = start + i;
            let style = if idx == p.highlight {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            set_line(buf, area.x, area.y + 1 + i as u16, name, area.width, style);
        }
    }

    /// Panel lines 2 / 3 while the `/Data External List` overlay is
    /// on screen (M12 v0.4 slice 2). Read-only — Enter is a no-op
    /// today; line 3 just shows source count + Esc hint.
    pub(super) fn render_external_list_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(el) = self.external_list.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let header = " External sources";
        let tail = if el.entries.is_empty() {
            " (no sources connected)".to_string()
        } else {
            let (name, _conn, _ts) = &el.entries[el.highlight];
            format!(
                " {}   [{}/{}]   Esc: close",
                name,
                el.highlight + 1,
                el.entries.len(),
            )
        };
        (Line::from(header), Line::from(tail))
    }

    /// Draw the `/Data External List` overlay: NAME / CONNECTION /
    /// REFRESHED columns, one row per registered source. Read-only;
    /// no row highlight selection in slice 2.
    pub(super) fn render_external_list_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(el) = self.external_list.as_ref() else {
            return;
        };
        let width = area.width as usize;
        let rows = area.height as usize;
        if rows == 0 || width == 0 {
            return;
        }
        let name_col = 16usize;
        let refreshed_col = 12usize;
        let conn_col = width.saturating_sub(name_col + refreshed_col + 4);
        let header = format!(
            " {name:<name_col$}  {conn:<conn_col$}  {ts:<refreshed_col$}",
            name = "NAME",
            conn = "CONNECTION",
            ts = "REFRESHED",
            name_col = name_col,
            conn_col = conn_col,
            refreshed_col = refreshed_col,
        );
        set_line(buf, area.x, area.y, &header, area.width, Style::default());
        if el.entries.is_empty() {
            set_line(
                buf,
                area.x,
                area.y + 1,
                "(no sources connected)",
                area.width,
                Style::default(),
            );
            return;
        }
        let visible_rows = rows.saturating_sub(1);
        let start = el.view_offset.min(el.entries.len());
        let end = (start + visible_rows).min(el.entries.len());
        for (i, (name, conn, ts)) in el.entries[start..end].iter().enumerate() {
            let idx = start + i;
            let ts_str = ts
                .map(|t| t.to_string())
                .unwrap_or_else(|| "(never)".to_string());
            let row = format!(
                " {name:<name_col$}  {conn:<conn_col$}  {ts:<refreshed_col$}",
                name = name,
                conn = conn,
                ts = ts_str,
                name_col = name_col,
                conn_col = conn_col,
                refreshed_col = refreshed_col,
            );
            let style = if idx == el.highlight {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            set_line(buf, area.x, area.y + 1 + i as u16, &row, area.width, style);
        }
    }

    /// Draw the F1 help overlay: header bar with the page title, body
    /// (with hyperlinks colorized — focused link reversed), footer with
    /// key hints. Body rows are clipped to the available height; the
    /// renderer auto-scrolls so the focused link stays visible.
    pub(super) fn render_help_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(state) = self.help.as_ref() else {
            return;
        };
        if area.height == 0 || area.width == 0 {
            return;
        }
        let header_style = Style::default()
            .fg(Color::Black)
            .bg(Color::Cyan)
            .add_modifier(Modifier::BOLD);
        let footer_style = Style::default().fg(Color::Black).bg(Color::Cyan);

        // Top header: " <Title>                                       HELP "
        let title = state.page.title.as_str();
        let header_left = format!(" {} ", title);
        set_line(buf, area.x, area.y, &header_left, area.width, header_style);
        // Right-aligned "HELP" tag.
        let tag = " HELP ";
        if (tag.len() as u16) <= area.width {
            let tag_x = area.x + area.width - tag.len() as u16;
            set_line(buf, tag_x, area.y, tag, tag.len() as u16, header_style);
        }

        // Reserve one row for the footer; the rest is the body window.
        let footer_y = area.y + area.height.saturating_sub(1);
        let body_top = area.y + 1;
        let body_height = footer_y.saturating_sub(body_top);
        let body_width = area.width;

        // Split body into (line_text, link_spans) per row, with link
        // spans expressed as byte ranges into the line. We then auto-
        // scroll so the focused link's row is on screen.
        let rows = build_help_rows(state);
        let focus_row = link_row_for_focus(&rows, state.focus);
        let scroll = match focus_row {
            Some(fr) => {
                let bh = body_height as usize;
                fr.saturating_sub(bh.saturating_sub(1))
            }
            None => 0,
        };

        let visible = rows
            .iter()
            .skip(scroll)
            .take(body_height as usize)
            .enumerate();
        for (i, row) in visible {
            render_help_row(
                buf,
                area.x,
                body_top + i as u16,
                body_width,
                row,
                state.focus,
            );
        }
        // Blank out remaining body rows.
        let drawn = rows.len().saturating_sub(scroll).min(body_height as usize);
        for i in drawn..body_height as usize {
            set_line(
                buf,
                area.x,
                body_top + i as u16,
                "",
                area.width,
                Style::default(),
            );
        }

        // Footer.
        let footer = " ↑/↓: next/prev link   ENTER: follow   BACKSPACE: back   ESC: close ";
        set_line(buf, area.x, footer_y, footer, area.width, footer_style);
    }

    pub(super) fn render_erase_confirm_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(ec) = self.erase_confirm.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let mut spans: Vec<Span<'_>> = Vec::with_capacity(FILE_ERASE_CONFIRM_ITEMS.len() * 2 + 1);
        spans.push(Span::raw(" "));
        for (i, (name, _)) in FILE_ERASE_CONFIRM_ITEMS.iter().enumerate() {
            if i > 0 {
                spans.push(Span::raw("  "));
            }
            if i == ec.highlight {
                spans.push(Span::styled(
                    *name,
                    Style::default().add_modifier(Modifier::REVERSED),
                ));
            } else {
                spans.push(Span::raw(*name));
            }
        }
        let line2 = Line::from(spans);
        let help = FILE_ERASE_CONFIRM_ITEMS
            .get(ec.highlight)
            .map(|(_, h)| *h)
            .unwrap_or("");
        let line3 = Line::from(format!(" {} {}", ec.path.display(), help));
        (line2, line3)
    }

    /// Render `{MENUBRANCH}` / `{MENUCALL}` overlay onto lines 2/3.
    /// Same shape as the static menu: items horizontally on line 2
    /// with the highlight reverse-video, description on line 3.
    pub(super) fn render_custom_menu_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(menu) = self.custom_menu.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let mut spans: Vec<Span<'_>> = Vec::with_capacity(menu.items.len() * 2 + 1);
        spans.push(Span::raw(" "));
        for (i, item) in menu.items.iter().enumerate() {
            if i > 0 {
                spans.push(Span::raw("  "));
            }
            if i == menu.highlight {
                spans.push(Span::styled(
                    item.name.clone(),
                    Style::default().add_modifier(Modifier::REVERSED),
                ));
            } else {
                spans.push(Span::raw(item.name.clone()));
            }
        }
        let desc = menu
            .items
            .get(menu.highlight)
            .map(|i| i.description.as_str())
            .unwrap_or("");
        (Line::from(spans), Line::from(format!(" {desc}")))
    }

    pub(super) fn render_save_confirm_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(sc) = self.save_confirm.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let mut spans: Vec<Span<'_>> = Vec::with_capacity(SAVE_CONFIRM_ITEMS.len() * 2 + 1);
        spans.push(Span::raw(" "));
        for (i, (name, _)) in SAVE_CONFIRM_ITEMS.iter().enumerate() {
            if i > 0 {
                spans.push(Span::raw("  "));
            }
            if i == sc.highlight {
                spans.push(Span::styled(
                    *name,
                    Style::default().add_modifier(Modifier::REVERSED),
                ));
            } else {
                spans.push(Span::raw(*name));
            }
        }
        let line2 = Line::from(spans);
        let help = SAVE_CONFIRM_ITEMS
            .get(sc.highlight)
            .map(|(_, h)| *h)
            .unwrap_or("");
        let line3 = Line::from(format!(" {help}"));
        (line2, line3)
    }

    pub(super) fn render_point_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(ps) = self.point.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        // If the user is typing a literal range, show the buffer
        // verbatim — it replaces the auto-derived highlight string.
        let range_str = if ps.typed.is_empty() {
            let range = self.highlight_range();
            format!(
                "{}..{}",
                range.start.display_full(),
                range.end.display_full()
            )
        } else {
            ps.typed.clone()
        };
        let line3_text = format!(" {} {}", ps.pending.prompt(), range_str);
        (Line::from(""), Line::from(line3_text))
    }

    pub(super) fn render_menu_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(state) = self.menu.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let level = state.level();

        // Line 2: items joined by two spaces, with the highlighted item
        // in reverse video.
        let mut spans: Vec<Span<'_>> = Vec::with_capacity(level.len() * 2 + 1);
        spans.push(Span::raw(" "));
        for (i, item) in level.iter().enumerate() {
            if i > 0 {
                spans.push(Span::raw("  "));
            }
            if i == state.highlight {
                spans.push(Span::styled(
                    item.name,
                    Style::default().add_modifier(Modifier::REVERSED),
                ));
            } else {
                spans.push(Span::raw(item.name));
            }
        }
        let line2 = Line::from(spans);

        // Line 3: if the highlighted item is a parent, preview its
        // children's names; else show the item's help text (and any
        // NotImplemented message).
        let line3_text = if let Some(item) = state.highlighted() {
            if let Some(msg) = state.message {
                format!(" Not yet implemented: {msg}")
            } else {
                match item.body {
                    MenuBody::Submenu(children) => {
                        let names: Vec<&str> = children.iter().map(|m| m.name).collect();
                        format!(" {}", names.join(" "))
                    }
                    _ => format!(" {}", item.help),
                }
            }
        } else {
            String::new()
        };
        let line3 = Line::from(line3_text);
        (line2, line3)
    }

    pub(super) fn cell_readout_for_line1(&self) -> String {
        self.wb()
            .cells
            .get(&self.wb().pointer)
            .map(|c| c.control_panel_readout())
            .unwrap_or_default()
    }

    /// Parenthesized format tag (e.g. `(G)`, `(C2)`, `(F3)`) for the current
    /// cell. Empty for labels, empty cells, or cells using `Reset` format.
    pub(super) fn format_tag_for_line1(&self) -> String {
        match self.wb().cells.get(&self.wb().pointer) {
            Some(CellContents::Constant(Value::Number(_))) | Some(CellContents::Formula { .. }) => {
                match self.format_for_cell(self.wb().pointer).tag() {
                    Some(s) => format!("({s})"),
                    None => String::new(),
                }
            }
            _ => String::new(),
        }
    }

    /// Brace-wrapped WYSIWYG attribute marker (e.g. `{Bold}`,
    /// `{Bold Italic}`) for the current cell.  Empty when no text-style
    /// override is set — which is common, so callers should treat the
    /// empty string as "omit the marker from line 1".
    pub(super) fn text_style_marker_for_line1(&self) -> String {
        match self.wb().cell_text_styles.get(&self.wb().pointer) {
            Some(style) => style.to_string(),
            None => String::new(),
        }
    }

    /// Resolve the format for a given cell — the per-cell override if
    /// set, else the workbook's global default ([`Workbook::global_format`]).
    pub(super) fn format_for_cell(&self, addr: Address) -> Format {
        self.wb().format_for_cell(addr)
    }

    /// Negative-value foreground override for the cell at `addr`.
    /// Returns `Some(rgb)` only when (a) the format carries a
    /// `negative_color` and (b) the cell's value is a negative number.
    /// Anything else (no override, non-numeric value, zero, positive)
    /// yields `None` so callers fall back to the cell's own font color.
    pub(super) fn negative_color_override(&self, addr: Address) -> Option<RgbColor> {
        let format = self.wb().format_for_cell(addr);
        let color = format.negative_color?;
        let cell = self.wb().cells.get(&addr)?;
        let n = match cell {
            CellContents::Constant(Value::Number(n)) => *n,
            CellContents::Formula {
                cached_value: Some(Value::Number(n)),
                ..
            } => *n,
            _ => return None,
        };
        if n < 0.0 {
            Some(color)
        } else {
            None
        }
    }

    /// `[Wn]` tag when the current column's width differs from the
    /// workbook's global default.
    pub(super) fn width_tag_for_line1(&self) -> String {
        let w = self.col_width_of(self.wb().pointer.sheet, self.wb().pointer.col);
        if w == self.wb().default_col_width {
            String::new()
        } else {
            format!("[W{w}]")
        }
    }

    /// Lay out the visible columns starting at `viewport_col_offset`,
    /// honoring per-column width overrides from `col_widths` and the
    /// `hidden_cols` set. Returns `(col_0b, x_offset, drawn_width)` for
    /// each column that has any on-screen footprint. Hidden columns are
    /// skipped entirely — the next visible column takes the slot.
    /// `x_offset` is measured from the start of the content area
    /// (after `ROW_GUTTER`). The last entry may be truncated to fit
    /// `content_width`.
    pub(super) fn visible_column_layout(&self, content_width: u16) -> Vec<(u16, u16, u16)> {
        let mut out = Vec::new();
        if content_width == 0 {
            return out;
        }
        let sheet = self.wb().pointer.sheet;
        let frozen_cols = self.wb().frozen.get(&sheet).map(|f| f.1).unwrap_or(0);
        let mut x_off: u16 = 0;
        // Emit the frozen columns first at fixed positions starting
        // from x_off = 0, regardless of viewport_col_offset.
        for col in 0..frozen_cols {
            let hidden = self.wb().hidden_cols.contains(&(sheet, col));
            let w = self.col_width_of(sheet, col) as u16;
            if hidden || w == 0 {
                continue;
            }
            let remaining = content_width.saturating_sub(x_off);
            if remaining == 0 {
                return out;
            }
            let drawn = w.min(remaining);
            out.push((col, x_off, drawn));
            x_off = x_off.saturating_add(drawn);
            if x_off >= content_width {
                return out;
            }
        }
        // Then emit the scrolling columns, starting at the greater of
        // the user's scroll offset and `frozen_cols` (so a user who
        // scrolled into the frozen range conceptually clamps back).
        let mut col = self.wb().viewport_col_offset.max(frozen_cols);
        loop {
            let hidden = self.wb().hidden_cols.contains(&(sheet, col));
            let w = self.col_width_of(sheet, col) as u16;
            if hidden || w == 0 {
                col = col.saturating_add(1);
                if col == u16::MAX {
                    break;
                }
                continue;
            }
            let remaining = content_width - x_off;
            let drawn = w.min(remaining);
            out.push((col, x_off, drawn));
            x_off = x_off.saturating_add(drawn);
            if x_off >= content_width {
                break;
            }
            if col == u16::MAX {
                break;
            }
            col += 1;
        }
        out
    }

    /// Visible body rows as `(row_idx, y_off_from_first_data_row)`.
    /// Frozen rows come first at fixed positions; scrolling rows
    /// follow, starting at `max(viewport_row_offset, frozen_rows)`.
    /// Each entry consumes one row in the buffer (cells are 1 line
    /// tall in this TUI), so `y_off` doubles as the row index within
    /// the body region.
    pub(super) fn visible_row_layout(&self, body_rows: u16) -> Vec<(u32, u16)> {
        let mut out = Vec::new();
        if body_rows == 0 {
            return out;
        }
        let sheet = self.wb().pointer.sheet;
        let frozen_rows: u32 = self.wb().frozen.get(&sheet).map(|f| f.0).unwrap_or(0);
        let mut y_off: u16 = 0;
        for row in 0..frozen_rows {
            if y_off >= body_rows {
                return out;
            }
            out.push((row, y_off));
            y_off += 1;
        }
        let mut row = (self.wb().viewport_row_offset as u64).max(frozen_rows as u64) as u32;
        while y_off < body_rows {
            out.push((row, y_off));
            y_off += 1;
            row = row.saturating_add(1);
            if row == u32::MAX {
                break;
            }
        }
        out
    }

    pub(super) fn render_grid(&self, area: Rect, buf: &mut Buffer) {
        self.last_grid_area.set(Some(area));
        if area.width <= ROW_GUTTER || area.height < 2 {
            return;
        }

        let content_width = area.width - ROW_GUTTER;
        let layout = self.visible_column_layout(content_width);
        let visible_rows = area.height - 1;

        // Column header row. The column / row containing the active
        // pointer wears `header_active_style`; everything else wears
        // the resting `header_style`. In the default theme both are
        // identical so today's uniform look is preserved.
        let header_style = self.theme.header_style();
        let header_active = self.theme.header_active_style();
        let active_col = self.wb().pointer.col;
        let active_row = self.wb().pointer.row;
        for &(col_idx, x_off, w) in &layout {
            let letters = col_to_letters(col_idx);
            let x = area.x + ROW_GUTTER + x_off;
            let style = if col_idx == active_col {
                header_active
            } else {
                header_style
            };
            write_centered(buf, x, area.y, w, &letters, style);
        }
        // Top-left gutter corner — sheet-identity area; wears the
        // active style so DOS paints the corner CGA blue alongside
        // the active column header and active row gutter.  The
        // active sheet's letter (A..IV) is centered in the gutter so
        // the user can see at a glance which sheet they're on, the
        // way 1-2-3 R3 stamps the worksheet identifier on the frame.
        let sheet_label = self.wb().pointer.sheet.letter();
        write_centered(buf, area.x, area.y, ROW_GUTTER, &sheet_label, header_active);

        // Body rows: frozen rows pinned at the top, then scrolling
        // rows starting at the viewport offset (clamped to skip past
        // the frozen prefix).
        let row_layout = self.visible_row_layout(visible_rows);
        for &(row_idx, r) in &row_layout {
            let y = area.y + 1 + r;
            // Row number gutter
            let label = format!("{:>width$}", row_idx + 1, width = (ROW_GUTTER - 1) as usize);
            let style = if row_idx == active_row {
                header_active
            } else {
                header_style
            };
            for (i, ch) in label.chars().enumerate() {
                buf[(area.x + i as u16, y)].set_char(ch).set_style(style);
            }
            buf[(area.x + ROW_GUTTER - 1, y)]
                .set_char(' ')
                .set_style(style);

            // During POINT, highlight the whole range; otherwise just the pointer.
            let highlight = if self.mode == Mode::Point {
                self.highlight_range()
            } else {
                Range::single(self.wb().pointer)
            };

            // Classify each visible column for the spill planner. Labels
            // are passed through as-is so they can overflow into empty
            // neighbors; everything else is pre-rendered in its own
            // column width and becomes an opaque spill blocker.
            let sheet = self.wb().pointer.sheet;
            let widths: Vec<usize> = layout.iter().map(|(_, _, w)| *w as usize).collect();
            let row_inputs: Vec<RowInput> = layout
                .iter()
                .map(|&(col_idx, _, w)| {
                    let addr = Address::new(sheet, col_idx, row_idx);
                    // Non-anchor cells of a merge render as a blank
                    // Rendered slot — blocks label-spill from
                    // neighbors and keeps the spill planner from
                    // double-painting over the merge area.  The
                    // anchor's content is repainted across the merge
                    // span by a dedicated pass after this loop.
                    if let Some(m) = self.wb().merge_at(addr) {
                        if !m.is_anchor(addr) {
                            return RowInput::Rendered(" ".repeat(w as usize));
                        }
                    }
                    // xlsx-set horizontal alignment overrides the 1-2-3
                    // default (label-prefix for labels, right for
                    // numbers / booleans / errors). HAlign::General
                    // leaves the default in place.
                    let halign = self
                        .wb()
                        .cell_alignments
                        .get(&addr)
                        .map(|a| a.horizontal)
                        .unwrap_or(HAlign::General);
                    match self.wb().cells.get(&addr) {
                        None | Some(CellContents::Empty) => RowInput::Empty,
                        Some(CellContents::Label { prefix, text }) => {
                            let eff_prefix = effective_label_prefix(*prefix, halign);
                            RowInput::Label {
                                prefix: eff_prefix,
                                text: text.clone(),
                            }
                        }
                        Some(other) => {
                            let fmt = self.format_for_cell(addr);
                            let ovr = self.wb().format_override_for_cell(addr);
                            let s = render_own_width(
                                other,
                                w as usize,
                                fmt,
                                &self.wb().international,
                                ovr,
                            );
                            RowInput::Rendered(apply_halign_to_rendered(&s, halign, w as usize))
                        }
                    }
                })
                .collect();
            let slots: Vec<SpillSlot<'_>> = row_inputs
                .iter()
                .map(|inp| match inp {
                    RowInput::Empty => SpillSlot::Empty,
                    RowInput::Label { prefix, text } => SpillSlot::Label {
                        prefix: *prefix,
                        text: text.as_str(),
                    },
                    RowInput::Rendered(s) => SpillSlot::Rendered(s.clone()),
                })
                .collect();
            let painted = plan_row_spill(&slots, &widths);

            for (&(col_idx, x_off, w), slot) in layout.iter().zip(painted.iter()) {
                let x = area.x + ROW_GUTTER + x_off;
                let addr = Address::new(sheet, col_idx, row_idx);
                let highlighted = highlight.contains(addr);
                // Pointer highlight is per-physical-cell; WYSIWYG text
                // style follows the owning cell so label spillover
                // into empty neighbors carries the owner's bold /
                // italic / underline.
                let owner_col = layout[slot.owner].0;
                let style_addr = Address::new(sheet, owner_col, row_idx);
                let mut cell_style = if highlighted {
                    self.theme.cell_highlight_style()
                } else {
                    display_mode_default_style(self.display_mode)
                };
                // xlsx-imported background fill paints behind the cell
                // contents.  Skip on the pointer highlight so the
                // inverted selection stays visually loud; the fill
                // returns the moment the pointer leaves.
                if !highlighted {
                    let fill_bg = self
                        .wb()
                        .cell_fills
                        .get(&style_addr)
                        .and_then(|fill| fill.bg);
                    if let Some(rgb) = fill_bg {
                        cell_style = cell_style.bg(Color::Rgb(rgb.r, rgb.g, rgb.b));
                    }
                    // xlsx-imported font color tints the text.  Same
                    // pointer-suppression rule as fill.  Order of
                    // precedence: an explicit font color from the
                    // workbook always wins; otherwise the cell defers
                    // to the terminal's own foreground so unfilled
                    // and filled cells share one aesthetic — except
                    // when the fill is light enough to wash out the
                    // typical light terminal fg.  That last case
                    // mirrors Excel's "automatic" font color flipping
                    // to black on light fills in dark themes.
                    let explicit_fg = self
                        .wb()
                        .cell_font_styles
                        .get(&style_addr)
                        .and_then(|fs| fs.color);
                    // `/Range Format Other Color Negative` wins over an
                    // xlsx-imported font color when the cell's value is
                    // negative — that's the whole point of the override.
                    let neg_override = self.negative_color_override(addr);
                    let resolved_fg = neg_override
                        .or(explicit_fg)
                        .or_else(|| fill_bg.and_then(|bg| bg.auto_contrast_for_dark_terminal()));
                    if let Some(rgb) = resolved_fg {
                        cell_style = cell_style.fg(Color::Rgb(rgb.r, rgb.g, rgb.b));
                    }
                }
                // Strikethrough applies whether highlighted or not —
                // it's an attribute of the glyph, not of the
                // selection state.
                if let Some(fs) = self.wb().cell_font_styles.get(&style_addr) {
                    if fs.strike {
                        cell_style = cell_style.add_modifier(Modifier::CROSSED_OUT);
                    }
                }
                if let Some(style) = self.wb().cell_text_styles.get(&style_addr).copied() {
                    cell_style = cell_style.add_modifier(text_style_modifier(style));
                }
                // Underline applies to glyphs, not to padding spaces
                // before/after them.  The spill planner reports the
                // exact text range per slot — including internal
                // whitespace at cell seams of a spilled label, which
                // a per-slot trim heuristic would mistakenly clip.
                let pad_style = cell_style.remove_modifier(Modifier::UNDERLINED);
                let mut printed = 0u16;
                for (idx, ch) in slot.text.chars().take(w as usize).enumerate() {
                    let in_text = idx >= slot.text_start && idx < slot.text_end;
                    let style = if in_text { cell_style } else { pad_style };
                    buf[(x + printed, y)].set_char(ch).set_style(style);
                    printed += 1;
                }
                // Pad any shortfall with blank cells so highlight still
                // fills the whole column.
                while printed < w {
                    buf[(x + printed, y)].set_char(' ').set_style(pad_style);
                    printed += 1;
                }
                // `:Display Options Grid Yes` — paint a dim dashed glyph at the
                // cell's rightmost column when that position would
                // otherwise be a space. Real R3.4a draws magenta lines
                // in the inter-glyph pixels; we don't have sub-character
                // precision and we're rendering text not graphics, so a
                // plain DarkGray reads as a subtle separator on every
                // terminal theme without competing with cell content.
                // Skip on highlighted cells so the REVERSED selection
                // stays loud, and skip when the cell content reached
                // the edge (don't overwrite data).
                if self.show_gridlines && !highlighted && w > 0 {
                    let gx = x + w - 1;
                    if buf[(gx, y)].symbol() == " " {
                        let mut g_style = pad_style;
                        g_style = g_style.fg(Color::DarkGray);
                        buf[(gx, y)].set_char('┊').set_style(g_style);
                    }
                }
            }

            // Second pass: overlay vertical (right-edge) borders on
            // each cell's rightmost column.  When two adjacent cells
            // both set a border on the seam between them, pick the
            // heavier via `merge_heavier`.  Pointer-highlighted cells
            // skip the overlay so the REVERSED selection stays loud.
            for (i, &(col_idx, x_off, w)) in layout.iter().enumerate() {
                if w == 0 {
                    continue;
                }
                let addr = Address::new(sheet, col_idx, row_idx);
                if highlight.contains(addr) {
                    continue;
                }
                let own_right = self.wb().cell_borders.get(&addr).and_then(|b| b.right);
                let neighbor_left = if let Some(&(next_col, _, _)) = layout.get(i + 1) {
                    let next_addr = Address::new(sheet, next_col, row_idx);
                    // Don't borrow from a highlighted neighbor — its
                    // border was suppressed too.
                    if highlight.contains(next_addr) {
                        None
                    } else {
                        self.wb().cell_borders.get(&next_addr).and_then(|b| b.left)
                    }
                } else {
                    None
                };
                let edge = match (own_right, neighbor_left) {
                    (Some(a), Some(b)) => Some(a.merge_heavier(b)),
                    (Some(a), None) => Some(a),
                    (None, Some(b)) => Some(b),
                    (None, None) => None,
                };
                let Some(edge) = edge else { continue };
                let glyph = edge.style.vertical_glyph();
                let bx = area.x + ROW_GUTTER + x_off + w - 1;
                // Preserve whatever bg the cell-paint pass put down
                // (fill, blanks, etc.); only the fg switches to the
                // border color so the glyph reads against the cell's
                // existing surface.
                let bg = buf[(bx, y)].bg;
                let mut bstyle = Style::default().bg(bg);
                if let Some(rgb) = edge.color {
                    bstyle = bstyle.fg(Color::Rgb(rgb.r, rgb.g, rgb.b));
                }
                buf[(bx, y)].set_char(glyph).set_style(bstyle);
            }

            // Third pass: comment corner markers.  A small `'` in red
            // sits on the rightmost column of cells that carry a
            // comment, evoking Excel's red-triangle indicator.  This
            // overrides any right-border glyph painted above (a
            // commented cell's "look here" cue is more actionable
            // than the visual seam).  Highlighted cells suppress
            // their marker so the REVERSED selection stays loud.
            for &(col_idx, x_off, w) in &layout {
                if w == 0 {
                    continue;
                }
                let addr = Address::new(sheet, col_idx, row_idx);
                if highlight.contains(addr) {
                    continue;
                }
                if !self.wb().comments.contains_key(&addr) {
                    continue;
                }
                let bx = area.x + ROW_GUTTER + x_off + w - 1;
                let bg = buf[(bx, y)].bg;
                buf[(bx, y)]
                    .set_char(COMMENT_MARKER)
                    .set_style(Style::default().bg(bg).fg(Color::Red));
            }

            // Fourth pass: merge anchor expansion.  For each merge
            // whose top-left anchor sits on this row and is visible
            // in the current layout, repaint the anchor's content
            // across the merge's column span — overwriting the blank
            // non-anchor slots prepared by the row-input builder.
            // Multi-row merges only expand on the anchor's row;
            // subsequent rows of the merge stay blank (top-aligned,
            // matching Excel's default vertical alignment).
            if let Some(merge_list) = self.wb().merges.get(&sheet) {
                for m in merge_list {
                    if m.anchor.row != row_idx {
                        continue;
                    }
                    // Look up the anchor's slot in the current layout;
                    // bail if it's not visible.
                    let Some(anchor_idx) = layout.iter().position(|(c, _, _)| *c == m.anchor.col)
                    else {
                        continue;
                    };
                    let (_, anchor_x_off, _) = layout[anchor_idx];
                    // Sum widths of every visible column from anchor
                    // through the merge's end (clamped to viewport).
                    let span_w: u16 = layout
                        .iter()
                        .filter(|(c, _, _)| *c >= m.anchor.col && *c <= m.end.col)
                        .map(|(_, _, w)| *w)
                        .sum();
                    if span_w == 0 {
                        continue;
                    }
                    // Render the anchor's content at the wider width.
                    let halign = self
                        .wb()
                        .cell_alignments
                        .get(&m.anchor)
                        .map(|a| a.horizontal)
                        .unwrap_or(HAlign::General);
                    let (painted_text, anchor_text_start, anchor_text_end) =
                        match self.wb().cells.get(&m.anchor) {
                            None | Some(CellContents::Empty) => {
                                (" ".repeat(span_w as usize), 0usize, 0usize)
                            }
                            Some(CellContents::Label { prefix, text }) => {
                                let eff_prefix = effective_label_prefix(*prefix, halign);
                                let painted = render_label(eff_prefix, text, span_w as usize);
                                let (s, e) = label_text_bounds(
                                    eff_prefix,
                                    text.chars().count(),
                                    span_w as usize,
                                );
                                (painted, s, e)
                            }
                            Some(other) => {
                                let fmt = self.format_for_cell(m.anchor);
                                let ovr = self.wb().format_override_for_cell(m.anchor);
                                let s = render_own_width(
                                    other,
                                    span_w as usize,
                                    fmt,
                                    &self.wb().international,
                                    ovr,
                                );
                                let painted = apply_halign_to_rendered(&s, halign, span_w as usize);
                                // Rendered values (numbers, formulas) have
                                // no internal whitespace runs, so trimming
                                // captures the text region exactly.
                                let chars: Vec<char> = painted.chars().collect();
                                let first = chars.iter().position(|c| *c != ' ');
                                let last = chars.iter().rposition(|c| *c != ' ');
                                let (ts, te) = match (first, last) {
                                    (Some(a), Some(b)) => (a, b + 1),
                                    _ => (0, 0),
                                };
                                (painted, ts, te)
                            }
                        };
                    // Build the anchor's full visual style — same
                    // layering as the cell-paint loop: pointer
                    // suppresses fill/font; text-style modifiers
                    // always apply.  When the pointer sits on the
                    // anchor, the merge's whole span wears the
                    // theme's cell-highlight style.
                    let anchor_highlighted = highlight.contains(m.anchor);
                    let mut astyle = if anchor_highlighted {
                        self.theme.cell_highlight_style()
                    } else {
                        Style::default()
                    };
                    if !anchor_highlighted {
                        if let Some(fill) = self.wb().cell_fills.get(&m.anchor) {
                            if let Some(rgb) = fill.bg {
                                astyle = astyle.bg(Color::Rgb(rgb.r, rgb.g, rgb.b));
                            }
                        }
                        if let Some(fs) = self.wb().cell_font_styles.get(&m.anchor) {
                            if let Some(rgb) = fs.color {
                                astyle = astyle.fg(Color::Rgb(rgb.r, rgb.g, rgb.b));
                            }
                        }
                    }
                    if let Some(fs) = self.wb().cell_font_styles.get(&m.anchor) {
                        if fs.strike {
                            astyle = astyle.add_modifier(Modifier::CROSSED_OUT);
                        }
                    }
                    if let Some(ts) = self.wb().cell_text_styles.get(&m.anchor).copied() {
                        astyle = astyle.add_modifier(text_style_modifier(ts));
                    }
                    let pad_astyle = astyle.remove_modifier(Modifier::UNDERLINED);
                    let x = area.x + ROW_GUTTER + anchor_x_off;
                    let mut printed = 0u16;
                    for (idx, ch) in painted_text.chars().take(span_w as usize).enumerate() {
                        let in_text = idx >= anchor_text_start && idx < anchor_text_end;
                        let style = if in_text { astyle } else { pad_astyle };
                        buf[(x + printed, y)].set_char(ch).set_style(style);
                        printed += 1;
                    }
                    while printed < span_w {
                        buf[(x + printed, y)].set_char(' ').set_style(pad_astyle);
                        printed += 1;
                    }
                }
            }
        }
    }

    pub(super) fn render_status(&self, area: Rect, buf: &mut Buffer) {
        // Left slot: filename or clock, per `/Worksheet Global Default
        // Other Clock`. Filename mode falls back to the International
        // clock when no file is loaded so the slot isn't blank in a
        // fresh session.
        let filename = || {
            self.wb().active_path.as_ref().map(|p| {
                p.file_name()
                    .and_then(|n| n.to_str())
                    .map(|s| s.to_string())
                    .unwrap_or_else(|| p.display().to_string())
            })
        };
        let left_text = match self.clock_display {
            ClockDisplay::Standard => {
                crate::clock::format_ddmmmyy_hhmm_ampm(crate::clock::local_now())
            }
            ClockDisplay::International => {
                crate::clock::format_ddmmmyyyy_hhmm(crate::clock::local_now())
            }
            ClockDisplay::None => String::new(),
            ClockDisplay::Filename => filename()
                .unwrap_or_else(|| crate::clock::format_ddmmmyyyy_hhmm(crate::clock::local_now())),
        };
        // Multi-sheet workbook → append "[<Letter>: <name>]" so users
        // can see which Excel tab their letter-addressed pointer is on.
        // Single-sheet workbooks skip this since there's nothing to
        // disambiguate and authentic 1-2-3 never showed a sheet name.
        let (sheet_suffix, letter_abs_col) = if self.wb().engine.sheet_count() > 1 {
            let sid = self.wb().pointer.sheet;
            let letter = sid.letter();
            let name = self
                .wb()
                .engine
                .sheet_name(sid)
                .unwrap_or_else(|| "?".to_string());
            let trimmed: String = name.chars().take(STATUS_SHEET_NAME_MAX).collect();
            let suffix = format!("  [{letter}: {trimmed}]");
            // Position of the letter char within `left`: 1 leading
            // space + left_text + 3 prefix chars ("  [") = left_text.len() + 4.
            let letter_pos = 1 + left_text.chars().count() + 3;
            (suffix, Some(letter_pos))
        } else {
            (String::new(), None)
        };
        // SPEC §4: prefix the left slot with `*` when the workbook has
        // unsaved changes; otherwise a leading space keeps the column
        // alignment.
        let prefix = if self.is_dirty() { '*' } else { ' ' };
        let left = format!("{prefix}{left_text}{sheet_suffix}");
        // Active status indicators, in the order 1-2-3 displays them.
        // For M2 we emit CALC only; the others arrive with their features.
        let mut indicators = Vec::new();
        if self.file_nav_pending {
            indicators.push("FILE");
        }
        if self.group_mode {
            indicators.push("GROUP");
        }
        if self.undo_enabled {
            indicators.push("UNDO");
        }
        if self.recalc_pending {
            indicators.push("CALC");
        }
        // M12 v0.4 slice 6 — light PROT when the pointer is over a
        // cell inside any `/Data External Use` binding. Direct edit
        // on those cells is refused; the user has to /DER or /DED
        // to mutate them.
        if self.addr_is_externally_bound(self.wb().pointer) {
            indicators.push("PROT");
        }
        if self.learn_recording {
            indicators.push("LEARN");
        }
        if self.step_mode {
            indicators.push("STEP");
        }
        // SST = "single-step suspended" — currently parked at a
        // STEP-mode pause waiting for the user to advance.
        if matches!(
            self.macro_state.as_ref().and_then(|s| s.suspend.as_ref()),
            Some(MacroSuspend::StepPause)
        ) {
            indicators.push("SST");
        }
        let right_chunk = indicators.join(" ");
        let pad = (area.width as usize).saturating_sub(left.len() + right_chunk.len() + 1);
        let line = format!("{left}{}{right_chunk} ", " ".repeat(pad));
        // Tint just the sheet letter when the active sheet carries a
        // tab color.  The rest of the status line stays DarkGray so the
        // letter stands out and the status text remains legible.
        let active_sheet = self.wb().pointer.sheet;
        let letter_color = letter_abs_col
            .and_then(|_| self.wb().sheet_colors.get(&active_sheet).copied())
            .map(|c| Color::Rgb(c.r, c.g, c.b));
        for (i, ch) in line.chars().enumerate().take(area.width as usize) {
            let fg = match (letter_abs_col, letter_color) {
                (Some(pos), Some(rgb)) if i == pos => rgb,
                _ => Color::DarkGray,
            };
            buf[(area.x + i as u16, area.y)]
                .set_char(ch)
                .set_style(Style::default().fg(fg));
        }
    }
}
