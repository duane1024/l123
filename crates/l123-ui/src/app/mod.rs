//! Top-level `App` state machine for the l123 TUI.
//!
//! The implementation is split across submodules — see the
//! "app/ module layout" section in `CLAUDE.md` for the full table and
//! conventions. This file owns the `App` struct itself plus the
//! remaining command impls (file / print / range / data / worksheet /
//! global) that have not yet been split into per-domain submodules.
//! Anything mode/render/macro/mouse/async/run-loop related lives in a
//! sibling file; new code should land there too.

use std::cell::Cell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io;
use std::path::{Path, PathBuf};

use tokio::runtime::{Builder as TokioBuilder, Runtime};

#[cfg(test)]
use crossterm::event::{KeyCode, KeyEvent, KeyModifiers, MouseButton, MouseEvent, MouseEventKind};
use l123_core::cell_render::halign_to_label_prefix;
use l123_core::{
    label::is_value_starter, Address, Alignment, Border, CellContents, Comment, CurrencyPosition,
    DateIntl, ErrKind, Fill, FontStyle, Format, FormatKind, HAlign, International, LabelPrefix,
    Merge, Mode, NegativeStyle, Punctuation, Range, RangeInput, RgbColor, SheetId, SheetState,
    Table, TextStyle, TimeIntl, Value,
};
use l123_engine::{CellView, Engine, IronCalcEngine, RecalcMode};
use l123_graph::{GraphDef, GraphType, Series};
use l123_menu::{self as menu, Action, MenuBody, MenuItem};
use l123_print::{PrintContentMode, PrintFormatMode, PrintSettings};
use ratatui::layout::Rect;
#[cfg(test)]
use ratatui::{
    buffer::Buffer,
    style::{Color, Modifier},
};
use ratatui_image::{picker::Picker, picker::ProtocolType};

use crate::help::HelpState;

mod async_ops;
mod keys;
mod macros;
mod mouse;
mod render;
mod run;
mod types;

#[cfg(test)]
mod tests;
#[cfg(test)]
use render::text_style_modifier;
pub use types::{
    ClockDisplay, DisplayMode, GlobalDefaults, GraphGroupOrientation, GraphSaveFormat, RecalcOrder,
    SplashInfo, ZeroDisplay,
};
use types::{
    ColorTarget, CustomMenuState, DataParseState, DataQueryState, DataRegressionState,
    DataSortState, Entry, EntryKind, EraseConfirmState, FormatField, FormatSnapshot, GraphOverlay,
    IconPanelGeom, JournalEntry, LabelDirection, MacroState, MenuState, PendingAsyncOp,
    PendingCommand, PointState, PrintDestination, PrintSession, PromptNext, PromptState, QueuedOp,
    SaveConfirmState, SearchScope, SearchSession, SortDir, SortKeySlot, StatView, Workbook,
    FILE_LIST_PAGE_SIZE, NAME_LIST_PAGE_SIZE,
};
pub(crate) use types::{
    CombineKind, FileListKind, FileListState, NameListOrigin, NameListState, TitlesKind, XtractKind,
};

// Grid geometry — kept as consts so both render and cell-address-probe agree.
const ROW_GUTTER: u16 = 5;
const PANEL_HEIGHT: u16 = 4; // 3 content lines + 1 bottom border

/// Rows shifted per scroll-wheel tick. Conventional 3-row step; small
/// enough that the user can land on the row they want without
/// overshooting, large enough that wheeling through a long sheet
/// doesn't feel sluggish.
const MOUSE_SCROLL_STEP: u32 = 3;

/// Glyph painted at the top-right of a commented cell (mimics
/// Excel's red-triangle "this cell has a note" indicator).  Sits on
/// the rightmost column of the cell's slot in red; suppressed on the
/// pointer-highlighted cell so the REVERSED selection stays loud.
const COMMENT_MARKER: char = '\'';

/// Cap for the per-sheet name shown in the status line after the
/// filename. Longer xlsx tab names get truncated to keep the right-
/// side indicator zone (`FILE GROUP UNDO CALC CIRC MEM NUM …`) from
/// getting shoved off-screen on an 80-column terminal.
const STATUS_SHEET_NAME_MAX: usize = 20;

/// 8-color palette for `:Format Color`. Matches the classic 1-2-3 R3
/// WYSIWYG palette plus xlsx_fill.tsv's green (00C800) for parity with
/// the existing fixture.
const PALETTE_BLACK: RgbColor = RgbColor { r: 0, g: 0, b: 0 };
const PALETTE_WHITE: RgbColor = RgbColor {
    r: 255,
    g: 255,
    b: 255,
};
const PALETTE_RED: RgbColor = RgbColor { r: 255, g: 0, b: 0 };
const PALETTE_GREEN: RgbColor = RgbColor { r: 0, g: 200, b: 0 };
const PALETTE_BLUE: RgbColor = RgbColor { r: 0, g: 0, b: 255 };
const PALETTE_YELLOW: RgbColor = RgbColor {
    r: 255,
    g: 255,
    b: 0,
};
const PALETTE_CYAN: RgbColor = RgbColor {
    r: 0,
    g: 255,
    b: 255,
};
const PALETTE_MAGENTA: RgbColor = RgbColor {
    r: 255,
    g: 0,
    b: 255,
};

/// Build a `ParseConfig` from the workbook's current `International`
/// for handing to `l123_parse::to_engine_source_with_config`. Argument
/// separator and decimal point come from the punctuation table.
fn parse_config_from(intl: &International) -> l123_parse::ParseConfig {
    l123_parse::ParseConfig {
        argument_sep: intl.punctuation.argument_sep(),
        decimal_point: intl.punctuation.decimal_char(),
    }
}

/// Emit a single BEL character to stdout and flush. The terminal's
/// own preferences decide whether that rings, flashes, or is ignored —
/// matching the soft, user-configurable behavior the user asked for.
fn emit_bell() {
    use std::io::Write;
    let mut out = io::stdout();
    let _ = out.write_all(b"\x07");
    let _ = out.flush();
}

/// Threshold for routing `F9` recalc through WAIT mode (PLAN §4.7).
/// Below this cell count the recalc stays synchronous so day-to-day
/// edits don't pay a tokio round-trip.
const RECALC_WAIT_CELL_THRESHOLD: usize = 50_000;

/// Pick the label prefix to render with, given the cell's stored
/// prefix and any xlsx-imported `HAlign` override. The stored
/// `Backslash` (repeat-fill) prefix is a Lotus directive with no
/// faithful Excel equivalent — Excel saves such cells with
/// `HAlign::Left` (its text default), which we must NOT let clobber
/// the fill semantics on re-import. For every other stored prefix,
/// an explicit halign override wins.
fn effective_label_prefix(stored: LabelPrefix, halign: HAlign) -> LabelPrefix {
    if stored == LabelPrefix::Backslash {
        return LabelPrefix::Backslash;
    }
    halign_to_label_prefix(halign).unwrap_or(stored)
}

pub struct App {
    mode: Mode,
    running: bool,
    entry: Option<Entry>,
    default_label_prefix: LabelPrefix,
    recalc_mode: RecalcMode,
    /// `/Worksheet Global Recalc` direction setting. IronCalc always
    /// evaluates in natural (dependency) order, so this is stored for
    /// the status panel but doesn't change calculation today. Slotted
    /// for a real effect once the engine grows explicit ordering.
    recalc_order: RecalcOrder,
    /// `/Worksheet Global Recalc Iteration` count (1..=50). Stored
    /// for the status panel; iterative solving isn't wired yet.
    recalc_iterations: u16,
    recalc_pending: bool,
    /// `/Worksheet Global Zero` — hide numeric zeros in cell
    /// rendering. Stored for the status panel; cell_render doesn't
    /// honor it yet.
    zero_display: ZeroDisplay,
    /// `/Worksheet Global Protection` — when On, edits to cells not
    /// listed in `Workbook::cell_unprotected` are refused (the input
    /// is dropped and an error beep fires).
    global_protection: bool,
    /// Live `/Range Input` constraint. While `Some(range)`, pointer
    /// movement is restricted to unprotected cells inside `range`.
    /// Esc clears it.
    input_range: Option<Range>,
    /// §4.7 — pending long-running file op. Queued when the user
    /// commits a `/File Retrieve` filename; mode flips to Wait while
    /// it's Some. Drained by `tick()`, which the production event
    /// loop calls each iteration. Ctrl-Break clears it without
    /// running.
    pending_async_op: Option<PendingAsyncOp>,
    /// §4.7 acceptance hook. While true, `tick()` skips the drain so
    /// transcripts can observe mid-flight WAIT mode. Cleared by
    /// `test_resume_async_op` or by Ctrl-Break.
    block_next_async_op: bool,
    /// 1-2-3 GROUP mode: when true, format and row/col operations
    /// propagate across all sheets of the active file. Toggled by
    /// `/Worksheet Global Group Enable|Disable`. Lights the GROUP
    /// indicator on the status line.
    group_mode: bool,
    /// True when `/Worksheet Global Default Other Undo` is enabled.
    /// While true, mutating commands push reverse entries onto the
    /// journal; Alt-F4 pops and applies. L123 defaults this to ON.
    undo_enabled: bool,
    /// `/Worksheet Global Default Other Clock` — picks what the
    /// status line's clock slot shows.
    clock_display: ClockDisplay,
    menu: Option<MenuState>,
    point: Option<PointState>,
    prompt: Option<PromptState>,
    /// Message displayed on control-panel line 2 while `Mode::Error` is
    /// active. Cleared by Esc/Enter, which also returns to `Mode::Ready`.
    error_message: Option<String>,
    /// Transient slot for the two-step /Range Name Create flow — the
    /// typed name is stashed here after the prompt step and consumed by
    /// commit_point.
    pending_name: Option<String>,
    /// After committing a filename that already exists on disk, this
    /// carries the chosen path through the Cancel/Replace/Backup
    /// submenu. Mode stays MENU while present.
    save_confirm: Option<SaveConfirmState>,
    /// After the `/File Erase` filename prompt commits, this carries the
    /// chosen path through the No/Yes confirm submenu.  Mode stays MENU
    /// while present.
    erase_confirm: Option<EraseConfirmState>,
    /// Transient slot for the two-step /File Xtract flow — the typed
    /// filename is stashed here after the prompt step and consumed by
    /// commit_point.
    pending_xtract_path: Option<PathBuf>,
    /// Transient slot for the two-step `/File Combine …
    /// Named/Specified-Range` flow — after the filename prompt commits,
    /// the path is stashed here while the user types the source range.
    pending_combine_path: Option<PathBuf>,
    /// Overlay state for /File List. When present, the mode is Files
    /// and the grid is obscured by a horizontal picker on lines 2/3.
    file_list: Option<FileListState>,
    /// Overlay state for F3 NAMES. When present, the mode is Names and
    /// the grid is obscured by a vertical name picker. Underlying
    /// POINT / prompt state is preserved so dismissal returns to it.
    name_list: Option<NameListState>,
    /// Overlay state for F1 HELP. Some while the help overlay is open;
    /// underlying mode is restored on Esc.
    help: Option<HelpState>,
    /// Active files in session order. A single-file session is a Vec
    /// of length 1; `/File Open` appends or inserts. Ctrl-End +
    /// Ctrl-PgUp/PgDn rotates `current` through the Vec.
    active_files: Vec<Workbook>,
    /// Index of the foreground file within `active_files`.
    current: usize,
    /// True after Ctrl-End until the next key. While true, the FILE
    /// indicator lights and Ctrl-PgUp/PgDn cycle between active files
    /// instead of between sheets.
    file_nav_pending: bool,
    /// In-flight `/Print File` session. Set when the filename prompt
    /// commits; cleared on Go-and-done or explicit Quit.
    print: Option<PrintSession>,
    /// In-flight `/Range Search` session between the search string
    /// commit and the Find/Replace leaf.
    search: Option<SearchSession>,
    /// Pre-resolved values for the currently-displayed full-screen
    /// graph. Some while in [`Mode::Graph`]; None otherwise. Snapshotting
    /// at F10 time keeps the renderer free of an engine dependency and
    /// means mid-view edits don't redraw until the user re-enters.
    graph_view: Option<GraphOverlay>,
    /// Terminal graphics-protocol picker, populated once at startup by
    /// [`App::probe_image_picker`]. `None` in headless tests (and any
    /// live session where the query fails) — the renderer falls back
    /// to the unicode path in that case.
    image_picker: Option<Picker>,
    /// Pre-decoded icon panel for `current_panel`, populated at startup
    /// iff the picker is graphical and re-rendered when the user pages
    /// through panels via the slot-16 navigator. On halfblocks /
    /// headless this stays `None` and the panel isn't drawn.
    icon_panel: Option<image::DynamicImage>,
    /// Which of the seven icon panels is currently displayed. The
    /// pager at slot 16 cycles through these.
    current_panel: l123_graph::Panel,
    /// Geometry of the last-rendered icon panel, stashed so mouse
    /// hover/click can hit-test against it without recomputing the
    /// layout. Cleared at the top of each frame; re-set by
    /// `render_icon_panel`. Stores both the cell rect and the actual
    /// rendered image pixel height so hit-tests can map mouse cells
    /// to icons even when each icon spans a fractional cell.
    icon_panel_area: Cell<Option<IconPanelGeom>>,
    /// Icon slot the mouse is currently over, if any. Drives the
    /// hover description in control-panel line 3 during READY. Slot 16
    /// (the pager) is intentionally excluded — its function is obvious
    /// from its rendered label.
    hovered_icon: Option<(l123_graph::Panel, usize)>,
    /// Rect the spreadsheet grid last occupied on screen. Cursor moves
    /// happen between renders, so `scroll_into_view` reads this stale
    /// rect to decide whether the new pointer fits below/right of the
    /// visible window. Cleared at the top of each frame; re-set by
    /// `render_grid`.
    last_grid_area: Cell<Option<Rect>>,
    /// Cell where the user pressed the left mouse button inside the
    /// grid, set on the Down event and cleared on Up. While `Some`, a
    /// subsequent Drag promotes Ready into POINT anchored here, or
    /// extends an existing POINT. `None` ⇒ Drag events are ignored
    /// (e.g. press landed off the grid, or no press at all).
    drag_anchor: Option<Address>,
    /// Startup welcome screen. `Some` while the splash is up; any
    /// keypress consumes the state and drops to READY without
    /// dispatching. Always `None` for `App::new()` so existing
    /// transcripts aren't blocked on a dismiss keystroke.
    splash: Option<SplashInfo>,
    /// Active chrome theme. Today this only colors the column-letter
    /// row and the row-number gutter; cell content, `:Format Color`,
    /// and xlsx fills are all unaffected. Seeded at startup from
    /// `--theme` > `L123_THEME` > `L123.CNF` `theme=` > built-in
    /// default.
    theme: crate::Theme,
    /// When true, the pointer-edge collision path fires a soft
    /// terminal bell (BEL, `\x07`). Toggled at runtime by
    /// `/Worksheet Global Default Other Beep Enable|Disable`; the
    /// startup value is seeded from [`crate::Config::error_beep_enabled`].
    beep_enabled: bool,
    /// Monotonic count of beep requests observed so far. Driven by
    /// [`App::request_beep`] and exposed for acceptance-transcript
    /// assertions — the TUI itself never reads it.
    beep_count: u64,
    /// Set by [`App::request_beep`] and consumed by
    /// [`App::take_pending_beep`] once per event-loop iteration so
    /// the terminal bell is emitted at most once per frame no matter
    /// how many times it was requested.
    beep_pending: bool,
    /// Persisted defaults set by `/Worksheet Global Default …` and
    /// written back to `L123.CNF` by `/Worksheet Global Default Update`.
    defaults: GlobalDefaults,
    /// `:Display Mode` — empty-cell color fallback. Cells with an
    /// xlsx-imported fill or font color paint that color regardless.
    display_mode: DisplayMode,
    /// `:Display Options Grid` — when true, paint a dim dashed glyph at each
    /// cell's rightmost column whenever that position would otherwise
    /// be a space. Best-effort vertical gridlines only — horizontals
    /// would cost a whole terminal row per cell row, which halves the
    /// visible row count. Defaults to off so every existing acceptance
    /// transcript that snapshots cell content sees the same byte stream.
    show_gridlines: bool,
    /// Which payload `Mode::Stat` is currently rendering — the standard
    /// `/Worksheet Status` panel or the `/Worksheet Global Default
    /// Status` defaults panel.
    stat_view: StatView,
    /// Set by `Action::System` and consumed by the event loop on the
    /// next iteration: leaves the alt-screen, drops raw mode, spawns
    /// `$SHELL`, and on its exit restores the TUI. Lives on `App`
    /// instead of being executed inline because the dispatcher doesn't
    /// own the `Terminal`; the event loop does.
    pending_system_suspend: bool,
    /// Active macro execution state. `Some` while a macro is
    /// running (possibly suspended for user input); `None` when
    /// idle. Constructed by [`run_macro_at`] / [`run_named_macro`]
    /// and torn down when the frame stack empties or `{QUIT}` fires.
    macro_state: Option<MacroState>,
    /// Re-entrancy guard for the macro pump. Synthetic key events
    /// from the macro flow back through [`handle_key`]; without
    /// this flag they would recursively pump and overflow.
    macro_pumping: bool,
    /// Destination cell for the active `{GETLABEL}`/`{GETNUMBER}`
    /// prompt. Side-cursor because [`PromptNext`] is `Copy` and
    /// can't carry an owned `String`.
    pending_macro_input_loc: Option<String>,
    /// Active `{MENUBRANCH}`/`{MENUCALL}` overlay. While `Some`,
    /// keystrokes are intercepted for menu navigation (similar to
    /// `save_confirm` / `name_list`).
    custom_menu: Option<CustomMenuState>,
    /// Destination range for Alt-F5 LEARN recordings. Set via
    /// `/Worksheet Learn Range`; cleared by `/WLC`.
    learn_range: Option<Range>,
    /// True while Alt-F5 has armed recording. Each user keystroke
    /// flowing through `handle_key` appends a macro-source token to
    /// `learn_buffer`.
    learn_recording: bool,
    /// Buffered macro source for the current learn session. Flushed
    /// to `learn_range` cells when the user toggles recording off.
    learn_buffer: String,
    /// Macro STEP mode (Alt-F2). When true, every macro action
    /// pauses for the user to advance with Space. Lights the STEP
    /// indicator on the status line; the running macro additionally
    /// shows SST while parked at a step.
    step_mode: bool,
    /// `/Data Sort` settings — sticky across Sort-menu visits and
    /// even across separate `/DS` sessions until cleared by Reset.
    data_sort: DataSortState,
    /// Which sort-key slot the in-flight Asc/Desc submenu writes
    /// into, set when `Primary-Key` / `Secondary-Key` POINT commits.
    pending_sort_key_slot: Option<SortKeySlot>,
    /// Column of the in-flight sort key, captured from the POINT
    /// cell that fired the Asc/Desc submenu.
    pending_sort_key_col: Option<u16>,
    /// `/Data Regression` settings — sticky across Regression-menu
    /// visits and across separate `/DR` sessions until cleared.
    data_regression: DataRegressionState,
    /// `/Data Parse` settings — sticky across Parse-menu visits and
    /// across separate `/DP` sessions until cleared by Reset.
    data_parse: DataParseState,
    /// `/Data Query` settings — sticky across Query-menu visits and
    /// across separate `/DQ` sessions until cleared by Reset.
    data_query: DataQueryState,
    /// §4.7 — current-thread tokio runtime that backs every long-
    /// running op via `spawn_blocking`. Owning it on `App` keeps
    /// the runtime's lifetime tied to the UI's; tests use the same
    /// handle to `block_on` a parked op to completion.
    runtime: Runtime,
    /// PLAN §4.7 cell-count threshold above which F9 recalc routes
    /// through the WAIT path. Initialized to
    /// `RECALC_WAIT_CELL_THRESHOLD`; transcripts can lower it via
    /// `test_set_recalc_wait_threshold` so a 5-cell sheet exercises
    /// the same code path without seeding 50k cells.
    recalc_wait_cell_threshold: usize,
}

fn parse_margin(buffer: &str, prev: u16) -> u16 {
    buffer.parse::<u16>().unwrap_or(prev).min(1000)
}

/// `/Data Query` criterion-vs-input comparator. Numbers compare
/// numerically; labels and `Value::Text` constants compare
/// case-insensitively; an empty input cell never matches a
/// non-empty criterion. Anything else (formula criterion, date,
/// error) returns `false` — the MVP slice doesn't evaluate
/// criterion expressions.
fn cell_values_equal_for_query(crit: &CellContents, input: Option<&CellContents>) -> bool {
    fn as_text_lower(c: &CellContents) -> Option<String> {
        match c {
            CellContents::Label { text, .. } => Some(text.to_ascii_lowercase()),
            CellContents::Constant(Value::Text(s)) => Some(s.to_ascii_lowercase()),
            _ => None,
        }
    }
    fn as_number(c: &CellContents) -> Option<f64> {
        match c {
            CellContents::Constant(Value::Number(n)) => Some(*n),
            CellContents::Formula {
                cached_value: Some(Value::Number(n)),
                ..
            } => Some(*n),
            _ => None,
        }
    }
    match (crit, input) {
        (_, None) => false,
        (CellContents::Constant(Value::Number(cn)), Some(in_c)) => as_number(in_c)
            .map(|n| (n - cn).abs() < 1e-12)
            .unwrap_or(false),
        (c, Some(in_c)) => {
            if let (Some(a), Some(b)) = (as_text_lower(c), as_text_lower(in_c)) {
                a == b
            } else {
                false
            }
        }
    }
}

/// Auto-generate a `/Data Parse` format line from a sample data
/// label. Each char is classified (digits + sign + dot → `V`,
/// whitespace → space gap, anything else → `L`); the first char
/// of each run emits the marker, subsequent chars in the same run
/// emit `>`. Whitespace gaps are preserved verbatim so the field
/// boundaries align character-for-character with the source label.
fn build_format_line(label: &str) -> String {
    let mut out = String::from("|");
    let mut prev_class: Option<char> = None;
    for c in label.chars() {
        let class = if c.is_ascii_digit() || matches!(c, '.' | '-' | '+') {
            'V'
        } else if c.is_whitespace() {
            ' '
        } else {
            'L'
        };
        if Some(class) != prev_class {
            out.push(class);
        } else if class == ' ' {
            out.push(' ');
        } else {
            out.push('>');
        }
        prev_class = Some(class);
    }
    out
}

/// Parse a Lotus-style format line into a list of fields. The
/// leading `|` is consumed; each marker char (L/V/D/T/S) opens a
/// new field at the current character position, and each `>`
/// extends the current field. Returns `(start_char, end_char,
/// kind)` — half-open byte-position-as-char-index ranges.
fn parse_format_line(fl: &str) -> Vec<(usize, usize, FormatField)> {
    let body: Vec<char> = if let Some(rest) = fl.strip_prefix('|') {
        rest.chars().collect()
    } else {
        fl.chars().collect()
    };
    let mut fields = Vec::new();
    let mut current: Option<(usize, FormatField)> = None;
    for (i, &c) in body.iter().enumerate() {
        let kind = match c {
            'L' | 'l' => Some(FormatField::Label),
            'V' | 'v' => Some(FormatField::Value),
            'D' | 'd' => Some(FormatField::Date),
            'T' | 't' => Some(FormatField::Time),
            'S' | 's' => Some(FormatField::Skip),
            _ => None,
        };
        if let Some(k) = kind {
            if let Some((start, ty)) = current.take() {
                fields.push((start, i, ty));
            }
            current = Some((i, k));
        }
        // `>` and any other char (space, etc.) just extend the
        // current field; if no field is open, they're ignored.
    }
    if let Some((start, ty)) = current.take() {
        fields.push((start, body.len(), ty));
    }
    fields
}

/// Round `n` to `sig` significant decimal digits — used by the
/// matrix kernels to suppress IEEE-754 noise (`0.6000000000000001`
/// → `0.6`) before storing into cells where General-format display
/// would otherwise leak the trailing junk.
fn round_to_significant(n: f64, sig: i32) -> f64 {
    if n == 0.0 || !n.is_finite() {
        return n;
    }
    let magnitude = n.abs().log10().floor() as i32;
    let factor = 10f64.powi(sig - 1 - magnitude);
    (n * factor).round() / factor
}

/// In-place Gauss-Jordan elimination with partial pivoting on the
/// augmented matrix `[mat | I]`. Returns the inverse, or `None`
/// when `mat` is singular (no usable pivot found in a column).
fn gauss_jordan_invert(mut mat: Vec<Vec<f64>>) -> Option<Vec<Vec<f64>>> {
    let n = mat.len();
    if n == 0 || mat.iter().any(|r| r.len() != n) {
        return None;
    }
    let mut inv: Vec<Vec<f64>> = (0..n)
        .map(|i| {
            let mut row = vec![0.0_f64; n];
            row[i] = 1.0;
            row
        })
        .collect();
    for col in 0..n {
        let mut pivot = col;
        for r in col + 1..n {
            if mat[r][col].abs() > mat[pivot][col].abs() {
                pivot = r;
            }
        }
        if mat[pivot][col].abs() < 1e-12 {
            return None;
        }
        if pivot != col {
            mat.swap(col, pivot);
            inv.swap(col, pivot);
        }
        let p = mat[col][col];
        for c in 0..n {
            mat[col][c] /= p;
            inv[col][c] /= p;
        }
        for r in 0..n {
            if r == col {
                continue;
            }
            let factor = mat[r][col];
            if factor == 0.0 {
                continue;
            }
            for c in 0..n {
                mat[r][c] -= factor * mat[col][c];
                inv[r][c] -= factor * inv[col][c];
            }
        }
    }
    Some(inv)
}

/// Build an apostrophe-prefixed label cell. Convenience for code
/// paths (e.g. `/Data Regression` output) that need to write
/// header strings into the grid.
fn label_cell(text: &str) -> CellContents {
    CellContents::Label {
        prefix: LabelPrefix::Apostrophe,
        text: text.into(),
    }
}

/// Comparator used by `/Data Sort` for two key cells. Numbers
/// compare numerically, labels lexicographically, and a missing
/// (empty) cell sorts after any populated cell — matching 1-2-3's
/// "blanks last in ascending sort" rule. Mixed numeric/label keys
/// put numbers before labels.
fn compare_cell_contents(a: Option<&CellContents>, b: Option<&CellContents>) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    fn rank(c: Option<&CellContents>) -> u8 {
        match c {
            Some(CellContents::Constant(Value::Number(_))) => 0,
            Some(CellContents::Formula { cached_value, .. }) => match cached_value {
                Some(Value::Number(_)) => 0,
                Some(Value::Text(_)) => 1,
                _ => 2,
            },
            Some(CellContents::Label { .. }) | Some(CellContents::Constant(Value::Text(_))) => 1,
            _ => 2,
        }
    }
    fn number_of(c: Option<&CellContents>) -> Option<f64> {
        match c {
            Some(CellContents::Constant(Value::Number(n))) => Some(*n),
            Some(CellContents::Formula {
                cached_value: Some(Value::Number(n)),
                ..
            }) => Some(*n),
            _ => None,
        }
    }
    fn text_of(c: Option<&CellContents>) -> Option<String> {
        match c {
            Some(CellContents::Label { text, .. }) => Some(text.clone()),
            Some(CellContents::Constant(Value::Text(s))) => Some(s.clone()),
            Some(CellContents::Formula {
                cached_value: Some(Value::Text(s)),
                ..
            }) => Some(s.clone()),
            _ => None,
        }
    }
    let ra = rank(a);
    let rb = rank(b);
    if ra != rb {
        return ra.cmp(&rb);
    }
    if ra == 0 {
        let na = number_of(a).unwrap_or(0.0);
        let nb = number_of(b).unwrap_or(0.0);
        return na.partial_cmp(&nb).unwrap_or(Ordering::Equal);
    }
    if ra == 1 {
        let ta = text_of(a).unwrap_or_default();
        let tb = text_of(b).unwrap_or_default();
        return ta.cmp(&tb);
    }
    Ordering::Equal
}

/// Resolve which destination anchors a `/Copy` should paste into,
/// given source and destination ranges. Implements the Lotus tutorial
/// dimension matrix:
/// - source 1×1 → replicate at every (col, row) on every dest sheet
/// - dest 1×1 OR same dims as source → single anchor at dest top-left
///   on every dest sheet (3D destination paste once per sheet)
/// - both multi-cell with mismatched dims → error string for the
///   caller to surface
fn copy_paste_anchors(src: Range, dest: Range) -> Result<Vec<Address>, &'static str> {
    let src_cols = u32::from(src.end.col - src.start.col + 1);
    let src_rows = src.end.row - src.start.row + 1;
    let dst_cols = u32::from(dest.end.col - dest.start.col + 1);
    let dst_rows = dest.end.row - dest.start.row + 1;
    let single_src = src_cols == 1 && src_rows == 1;
    let same_size = src_cols == dst_cols && src_rows == dst_rows;
    let single_dest = dst_cols == 1 && dst_rows == 1;
    if !single_src && !same_size && !single_dest {
        return Err("Copy: source and destination ranges have different sizes");
    }
    let mut anchors = Vec::new();
    for sheet_idx in dest.start.sheet.0..=dest.end.sheet.0 {
        let sheet = SheetId(sheet_idx);
        if single_src && !single_dest {
            // Replicate the single source at every (col, row) of dest.
            for col in dest.start.col..=dest.end.col {
                for row in dest.start.row..=dest.end.row {
                    anchors.push(Address::new(sheet, col, row));
                }
            }
        } else {
            anchors.push(Address::new(sheet, dest.start.col, dest.start.row));
        }
    }
    Ok(anchors)
}

/// Resolve a user-typed filename into a save-target path. If the input
/// has no extension, default to `.xlsx` (L123's modern save format).
fn resolve_save_path(input: &str) -> PathBuf {
    let mut p = PathBuf::from(input);
    if p.extension().is_none() {
        p.set_extension("xlsx");
    }
    p
}

/// 1-2-3 R3.4a range-name rules: 1..=15 chars, first char is a
/// letter, no embedded whitespace or special characters that would
/// look like operators or sheet refs (`+ - * / ^ ( ) , ; : . #
/// & < > = !`).
fn is_valid_range_name(s: &str) -> bool {
    let len = s.chars().count();
    if !(1..=15).contains(&len) {
        return false;
    }
    let mut chars = s.chars();
    let first = chars.next().unwrap();
    if !first.is_ascii_alphabetic() && first != '_' {
        return false;
    }
    s.chars()
        .all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
}

/// Render a `Range` in 1-2-3 source form. Same-sheet ranges get a
/// single sheet prefix on the start address (`A:A1..A1`); cross-sheet
/// ranges get prefixes on both ends. The compact same-sheet form
/// matches what /RNT writes in 1-2-3 R3.4a.
fn range_to_lotus_form(r: Range) -> String {
    let r = r.normalized();
    if r.start.sheet == r.end.sheet {
        format!("{}..{}", r.start.display_full(), r.end.display_short())
    } else {
        format!("{}..{}", r.start.display_full(), r.end.display_full())
    }
}

/// True if `expr` (a 1-2-3-shape formula source) references the
/// range name `name` (compared case-insensitively, key already
/// lowercased) as a whole word — adjacent chars must be non-name
/// characters so `tax` does not match `taxes` or `tax_rate`.
fn formula_uses_name(expr: &str, name: &str) -> bool {
    let lower = expr.to_ascii_lowercase();
    let bytes = lower.as_bytes();
    let target = name.as_bytes();
    if target.is_empty() {
        return false;
    }
    let mut i = 0;
    while i + target.len() <= bytes.len() {
        if &bytes[i..i + target.len()] == target {
            let before = if i == 0 { None } else { Some(bytes[i - 1]) };
            let after = bytes.get(i + target.len()).copied();
            let is_word = |b: u8| b.is_ascii_alphanumeric() || b == b'_' || b == b'.';
            if before.is_none_or(|b| !is_word(b)) && after.is_none_or(|b| !is_word(b)) {
                return true;
            }
        }
        i += 1;
    }
    false
}

/// Replace any formula cell with its cached value (used by /Range
/// Value and /Range Trans). Empty cells stay empty; non-formula
/// cells are passed through unchanged.
fn freeze_to_value(c: Option<CellContents>) -> CellContents {
    match c {
        Some(CellContents::Formula {
            cached_value: Some(v),
            ..
        }) => CellContents::Constant(v),
        Some(CellContents::Formula {
            cached_value: None, ..
        })
        | None => CellContents::Empty,
        Some(other) => other,
    }
}

/// Greedy word-wrap of `text` into chunks no wider than `width`
/// columns. Words longer than `width` are emitted on their own line
/// and may exceed the limit (1-2-3 R3.4a same-cell behavior — long
/// tokens spill rather than break mid-word).
fn wrap_text_to_width(text: &str, width: usize) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let mut line = String::new();
    for word in text.split_whitespace() {
        if line.is_empty() {
            line.push_str(word);
            continue;
        }
        if line.chars().count() + 1 + word.chars().count() <= width {
            line.push(' ');
            line.push_str(word);
        } else {
            out.push(std::mem::take(&mut line));
            line.push_str(word);
        }
    }
    if !line.is_empty() {
        out.push(line);
    }
    out
}

/// Replace whole-word occurrences of `name` (case-insensitive,
/// already lowercase) in `expr` with `replacement`. Mirrors the
/// matching rules of `formula_uses_name`.
fn replace_name_in_formula(expr: &str, name: &str, replacement: &str) -> String {
    if name.is_empty() {
        return expr.to_string();
    }
    let lower = expr.to_ascii_lowercase();
    let lower_bytes = lower.as_bytes();
    let target = name.as_bytes();
    let is_word = |b: u8| b.is_ascii_alphanumeric() || b == b'_' || b == b'.';
    let mut out = String::with_capacity(expr.len());
    let src = expr.as_bytes();
    let mut i = 0;
    while i < src.len() {
        let matches_here = i + target.len() <= src.len()
            && lower_bytes[i..i + target.len()] == *target
            && (i == 0 || !is_word(lower_bytes[i - 1]))
            && lower_bytes
                .get(i + target.len())
                .copied()
                .is_none_or(|b| !is_word(b));
        if matches_here {
            out.push_str(replacement);
            i += target.len();
        } else {
            out.push(src[i] as char);
            i += 1;
        }
    }
    out
}

/// Format one row of the /File List overlay: name left-padded into
/// `name_w` chars, size right-aligned into `size_w`, separated by a
/// gap. Truncated to `total_w` so the caller can write a full line.
fn format_file_list_row(
    name: &str,
    size: &str,
    name_w: usize,
    size_w: usize,
    total_w: usize,
) -> String {
    let name_trunc = truncate_to(name, name_w);
    let size_trunc = truncate_to(size, size_w);
    let mut out = String::with_capacity(total_w);
    out.push(' ');
    out.push_str(&name_trunc);
    // Pad name to name_w.
    let pad = name_w.saturating_sub(name_trunc.chars().count());
    out.extend(std::iter::repeat_n(' ', pad));
    out.push(' ');
    // Right-align size into size_w.
    let size_pad = size_w.saturating_sub(size_trunc.chars().count());
    out.extend(std::iter::repeat_n(' ', size_pad));
    out.push_str(&size_trunc);
    // Final trim to total_w.
    let chars: Vec<char> = out.chars().collect();
    if chars.len() > total_w {
        chars.into_iter().take(total_w).collect()
    } else {
        let mut s: String = chars.into_iter().collect();
        s.extend(std::iter::repeat_n(' ', total_w - s.chars().count()));
        s
    }
}

fn truncate_to(s: &str, max: usize) -> String {
    s.chars().take(max).collect()
}

/// Human-readable byte size: B / K / M / G with one decimal place for
/// the larger units.
fn format_size(bytes: u64) -> String {
    const KB: u64 = 1024;
    const MB: u64 = KB * 1024;
    const GB: u64 = MB * 1024;
    if bytes < KB {
        format!("{bytes}B")
    } else if bytes < MB {
        format!("{:.1}K", bytes as f64 / KB as f64)
    } else if bytes < GB {
        format!("{:.1}M", bytes as f64 / MB as f64)
    } else {
        format!("{:.1}G", bytes as f64 / GB as f64)
    }
}

/// Keep the /File List view window centered around the highlighted
/// row: `view_offset <= highlight < view_offset + FILE_LIST_PAGE_SIZE`.
fn adjust_file_list_view(fl: &mut FileListState) {
    if fl.highlight < fl.view_offset {
        fl.view_offset = fl.highlight;
    } else if fl.highlight >= fl.view_offset + FILE_LIST_PAGE_SIZE {
        fl.view_offset = fl.highlight + 1 - FILE_LIST_PAGE_SIZE;
    }
}

fn adjust_name_list_view(nl: &mut NameListState) {
    if nl.highlight < nl.view_offset {
        nl.view_offset = nl.highlight;
    } else if nl.highlight >= nl.view_offset + NAME_LIST_PAGE_SIZE {
        nl.view_offset = nl.highlight + 1 - NAME_LIST_PAGE_SIZE;
    }
}

/// List every worksheet file (`.xlsx`, plus `.WK3` when built with
/// the `wk3` feature) in `dir`, sorted by filename. Hidden files and
/// non-file entries are skipped.
fn list_worksheet_files_in(dir: &Path) -> Vec<PathBuf> {
    let Ok(read) = std::fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut entries: Vec<PathBuf> = read
        .filter_map(|e| e.ok())
        .filter(|e| e.file_type().map(|t| t.is_file()).unwrap_or(false))
        .map(|e| e.path())
        .filter(|p| {
            p.extension()
                .map(|x| {
                    if x.eq_ignore_ascii_case("xlsx") {
                        return true;
                    }
                    #[cfg(feature = "wk3")]
                    if x.eq_ignore_ascii_case("wk3") {
                        return true;
                    }
                    false
                })
                .unwrap_or(false)
        })
        .collect();
    entries.sort();
    entries
}

/// List every regular file in `dir`, sorted by filename. Hidden files
/// (leading dot) are skipped to match `/File List Worksheet`'s convention.
/// Backs `/File List Other`.
fn list_all_files_in(dir: &Path) -> Vec<PathBuf> {
    let Ok(read) = std::fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut entries: Vec<PathBuf> = read
        .filter_map(|e| e.ok())
        .filter(|e| e.file_type().map(|t| t.is_file()).unwrap_or(false))
        .map(|e| e.path())
        .filter(|p| {
            p.file_name()
                .and_then(|n| n.to_str())
                .map(|n| !n.starts_with('.'))
                .unwrap_or(false)
        })
        .collect();
    entries.sort();
    entries
}

/// True if `path`'s extension is one l123 knows how to retrieve as a
/// workbook — driver for `/File List Other`'s Enter behavior.
fn is_retrievable_workbook(path: &Path) -> bool {
    let Some(ext) = path.extension().and_then(|e| e.to_str()) else {
        return false;
    };
    if ext.eq_ignore_ascii_case("xlsx") || ext.eq_ignore_ascii_case("csv") {
        return true;
    }
    #[cfg(feature = "wk3")]
    if ext.eq_ignore_ascii_case("wk3") {
        return true;
    }
    false
}

/// Render a source-engine [`CellView`] into the `set_user_input`
/// string shape appropriate for `/File Xtract`'s kind. Formulas keeps
/// the formula string; Values flattens it to the cached scalar.
fn xtract_cell_input(cv: &CellView, kind: XtractKind) -> String {
    if let Some(f) = &cv.formula {
        if matches!(kind, XtractKind::Formulas) {
            return f.clone();
        }
    }
    match &cv.value {
        Value::Number(n) => l123_core::format_number_general(*n),
        Value::Text(s) => format!("'{s}"),
        Value::Bool(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        _ => String::new(),
    }
}

/// Best-effort reconstruction of [`CellContents`] from a
/// freshly-loaded engine cell. Formulas are stored with the leading `=`
/// stripped so that `to_engine_source` will re-prepend it cleanly on
/// save (the reverse Excel→1-2-3 translation is a later milestone;
/// edits of loaded formulas will see the Excel-shape expression).
/// If the workbook's pointer sits on a non-visible sheet (Hidden /
/// VeryHidden), advance it to the first sheet that *is* visible so
/// the user lands somewhere they can interact with.  No-op when the
/// pointer is already on a visible sheet, or when no visible sheets
/// exist at all.  Resets viewport offsets on redirect so the new
/// sheet shows from A1.
/// Translate a single sheet letter (`A`, `B`, …, `Z`) into its
/// `SheetId` index.  `'A'` → 0, `'B'` → 1, `'Z'` → 25.  Lowercase
/// works too.  Returns `None` for non-letters.  Used by harness
/// assertion directives that take `<letter> <…>` arguments.
fn letter_to_sheet_index(c: char) -> Option<u16> {
    let upper = c.to_ascii_uppercase() as u32;
    if (b'A' as u32..=b'Z' as u32).contains(&upper) {
        Some((upper - b'A' as u32) as u16)
    } else {
        None
    }
}

fn redirect_pointer_off_hidden(wb: &mut Workbook) {
    let active = wb.pointer.sheet;
    let active_visible = wb
        .sheet_states
        .get(&active)
        .copied()
        .unwrap_or(SheetState::Visible)
        .is_visible();
    if active_visible {
        return;
    }
    let count = wb.engine.sheet_count();
    for i in 0..count {
        let sid = SheetId(i);
        let visible = wb
            .sheet_states
            .get(&sid)
            .copied()
            .unwrap_or(SheetState::Visible)
            .is_visible();
        if visible {
            wb.pointer = Address::new(sid, 0, 0);
            wb.viewport_col_offset = 0;
            wb.viewport_row_offset = 0;
            return;
        }
    }
}

fn cell_view_to_contents(cv: &CellView, sheets: &[&str]) -> Option<CellContents> {
    if let Some(f) = &cv.formula {
        let body = f.strip_prefix('=').unwrap_or(f);
        // Reverse the engine's Excel form back to a 1-2-3 source so
        // the panel and the cell cache stay authentic across save +
        // reload. Forward and reverse round-trip cleanly for the
        // supported subset (renames, niladic parens, `:`/`..`,
        // sheet refs, INDIRECT, `#VALUE!`); arg-fix and emulated
        // functions display in their decomposed Excel form.
        let expr = l123_parse::to_lotus_source(body, sheets);
        return Some(CellContents::Formula {
            expr,
            cached_value: Some(cv.value.clone()),
        });
    }
    match &cv.value {
        Value::Empty => None,
        Value::Text(s) => Some(CellContents::Label {
            prefix: LabelPrefix::Apostrophe,
            text: s.clone(),
        }),
        other => Some(CellContents::Constant(other.clone())),
    }
}

impl App {
    pub fn new() -> Self {
        Self {
            mode: Mode::Ready,
            running: true,
            entry: None,
            default_label_prefix: LabelPrefix::Apostrophe,
            recalc_mode: RecalcMode::Automatic,
            recalc_order: RecalcOrder::Natural,
            recalc_iterations: 1,
            recalc_pending: false,
            zero_display: ZeroDisplay::No,
            global_protection: false,
            input_range: None,
            pending_async_op: None,
            block_next_async_op: false,
            group_mode: false,
            undo_enabled: true,
            clock_display: ClockDisplay::default(),
            menu: None,
            point: None,
            prompt: None,
            error_message: None,
            pending_name: None,
            save_confirm: None,
            erase_confirm: None,
            pending_xtract_path: None,
            pending_combine_path: None,
            file_list: None,
            name_list: None,
            help: None,
            active_files: vec![Workbook::new()],
            current: 0,
            file_nav_pending: false,
            print: None,
            search: None,
            graph_view: None,
            image_picker: None,
            icon_panel: None,
            current_panel: l123_graph::Panel::One,
            icon_panel_area: Cell::new(None),
            hovered_icon: None,
            last_grid_area: Cell::new(None),
            drag_anchor: None,
            splash: None,
            theme: crate::Theme::default(),
            beep_enabled: true,
            beep_count: 0,
            beep_pending: false,
            defaults: GlobalDefaults::default(),
            display_mode: DisplayMode::default(),
            show_gridlines: false,
            stat_view: StatView::Worksheet,
            pending_system_suspend: false,
            macro_state: None,
            macro_pumping: false,
            pending_macro_input_loc: None,
            custom_menu: None,
            learn_range: None,
            learn_recording: false,
            learn_buffer: String::new(),
            step_mode: false,
            data_sort: DataSortState::default(),
            pending_sort_key_slot: None,
            pending_sort_key_col: None,
            data_regression: DataRegressionState::default(),
            data_parse: DataParseState::default(),
            data_query: DataQueryState::default(),
            runtime: TokioBuilder::new_current_thread()
                .enable_time()
                .build()
                .expect("tokio current-thread runtime init"),
            recalc_wait_cell_threshold: RECALC_WAIT_CELL_THRESHOLD,
        }
    }

    /// Construct an app with the startup splash active. Normal
    /// [`App::run`] uses this; tests and [`App::new_with_file`] stay
    /// splashless so they can get straight to work.
    pub fn new_with_splash(user: String, organization: String) -> Self {
        let mut app = Self::new();
        app.splash = Some(SplashInfo { user, organization });
        app
    }

    /// Construct an app pre-loaded from `path`, skipping the splash —
    /// mirrors the `l123 file.xlsx` CLI invocation where the user has
    /// already told us which file they want.
    pub fn new_with_file(path: PathBuf) -> Self {
        let mut app = Self::new();
        app.retrieve_by_extension(path);
        app
    }

    /// Dispatch a retrieve-style load by file extension. `.csv` goes
    /// through the CSV path; everything else through the xlsx path.
    /// Shared by the CLI entry point and `/File Retrieve`.
    fn retrieve_by_extension(&mut self, path: PathBuf) {
        match path
            .extension()
            .and_then(|e| e.to_str())
            .map(|s| s.to_ascii_lowercase())
            .as_deref()
        {
            Some("csv") => self.load_csv_workbook_from(path),
            _ => self.load_workbook_from(path),
        }
        self.try_autoexec();
    }

    /// Hook fired after a successful /File Retrieve. When the
    /// loaded workbook defines `\0` and `/WGD Default Other Autoexec`
    /// is enabled (the default), the macro at `\0` runs once before
    /// control returns to READY.
    fn try_autoexec(&mut self) {
        if !self.defaults.autoexec {
            return;
        }
        // Skip when the load itself put us in ERROR mode — the user
        // needs to see and dismiss the error first.
        if matches!(self.mode, Mode::Error) {
            return;
        }
        self.run_named_macro("\\0");
    }

    /// Flip the startup splash on with the given identity strings.
    /// Acceptance transcripts use this via the `SPLASH` directive so
    /// they don't have to re-create the app mid-run.
    pub fn show_splash(&mut self, user: String, organization: String) {
        self.splash = Some(SplashInfo { user, organization });
    }

    /// Pin the hover state for the acceptance harness. Production
    /// code drives this via mouse-move events in `handle_mouse`; the
    /// harness renders into a headless buffer where no real mouse
    /// coordinates map to the (unrendered) icon panel, so transcripts
    /// set this directly to exercise the render contract.
    pub fn set_hovered_icon(&mut self, panel: l123_graph::Panel, slot: usize) {
        self.hovered_icon = Some((panel, slot));
    }

    /// Companion to [`Self::set_hovered_icon`].
    pub fn clear_hovered_icon(&mut self) {
        self.hovered_icon = None;
    }

    /// Dispatch the icon at `(panel, slot)` as if the user had clicked
    /// it. Switches `current_panel` to `panel` so SmartIcon dispatch
    /// resolves through the same code path as a real click. Slot 16
    /// (the pager) cycles `current_panel` forward — direction-aware
    /// pagination needs a real x-coordinate, which the harness can't
    /// supply.
    pub fn dispatch_icon_for_test(&mut self, panel: l123_graph::Panel, slot: usize) {
        if slot >= 17 {
            return;
        }
        self.current_panel = panel;
        if slot == 16 {
            self.current_panel = panel.next();
            self.refresh_icon_panel();
            return;
        }
        let id = panel.icon_ids()[slot];
        match l123_graph::icon_action(id) {
            l123_graph::IconAction::MenuPath(p) => self.dispatch_menu_path(p),
            l123_graph::IconAction::WysiwygMenuPath(p) => self.dispatch_wysiwyg_menu_path(p),
            l123_graph::IconAction::TextStyleToggle { bits } => self.dispatch_icon_text_style(bits),
            l123_graph::IconAction::SysKey(act) => self.dispatch_sys_action(act),
            l123_graph::IconAction::PageNav => {}
            l123_graph::IconAction::Noop => {}
        }
    }

    /// True while the startup splash is up.
    pub fn splash_active(&self) -> bool {
        self.splash.is_some()
    }

    fn wb(&self) -> &Workbook {
        &self.active_files[self.current]
    }

    fn wb_mut(&mut self) -> &mut Workbook {
        &mut self.active_files[self.current]
    }

    pub fn recalc_mode(&self) -> RecalcMode {
        self.recalc_mode
    }

    pub fn set_recalc_mode(&mut self, mode: RecalcMode) {
        self.recalc_mode = mode;
    }

    pub fn recalc_pending(&self) -> bool {
        self.recalc_pending
    }

    /// True if the active workbook has unsaved changes. Drives the
    /// `/QY` warn-on-quit second confirm.
    pub fn is_dirty(&self) -> bool {
        self.wb().dirty
    }

    // ---------------- test-surface accessors ----------------

    pub fn pointer(&self) -> Address {
        self.wb().pointer
    }

    pub fn mode(&self) -> Mode {
        self.mode
    }

    pub fn is_running(&self) -> bool {
        self.running
    }

    /// Byte index of the entry buffer cursor, or `None` if no entry is
    /// active. Always lands on a UTF-8 char boundary.
    pub fn entry_cursor(&self) -> Option<usize> {
        self.entry.as_ref().map(|e| e.cursor)
    }

    /// Address of the first cell whose cached formula value is a
    /// circular-reference error, searched in address order. Returns
    /// `None` when no cell currently reports a cycle.
    ///
    /// Reads the UI cell cache rather than re-interrogating the
    /// engine. IronCalc doesn't surface `#CIRC!` through our current
    /// adapter yet, so this typically returns `None` even on workbooks
    /// that cycle; the mechanism is in place for when the adapter
    /// maps the error through.
    fn first_circular_reference(&self) -> Option<Address> {
        let mut entries: Vec<(&Address, &CellContents)> = self.wb().cells.iter().collect();
        entries.sort_by_key(|(a, _)| (a.sheet, a.row, a.col));
        entries.into_iter().find_map(|(addr, cc)| match cc {
            CellContents::Formula {
                cached_value: Some(Value::Error(ErrKind::Circular)),
                ..
            } => Some(*addr),
            _ => None,
        })
    }

    /// Current graph's type as an ASCII all-caps token, for use in
    /// `ASSERT_GRAPH_TYPE` transcript directives.
    pub fn graph_type_str(&self) -> &'static str {
        match self.wb().current_graph.graph_type {
            GraphType::Line => "LINE",
            GraphType::Bar => "BAR",
            GraphType::XY => "XY",
            GraphType::Stack => "STACK",
            GraphType::Pie => "PIE",
            GraphType::HLCO => "HLCO",
            GraphType::Mixed => "MIXED",
        }
    }

    /// Current graph's range for a given series slot, formatted like
    /// `A:A1..A:A3`. Empty string when the slot is unset. `slot` is
    /// one of `X A B C D E F` (case-insensitive).
    pub fn graph_series_str(&self, slot: char) -> String {
        let s = match slot.to_ascii_uppercase() {
            'X' => Series::X,
            'A' => Series::A,
            'B' => Series::B,
            'C' => Series::C,
            'D' => Series::D,
            'E' => Series::E,
            'F' => Series::F,
            _ => return String::new(),
        };
        match self.wb().current_graph.get(s) {
            None => String::new(),
            Some(r) => format!("{}..{}", r.start.display_full(), r.end.display_full()),
        }
    }

    // ---------------- key handling ----------------

    fn set_error(&mut self, msg: impl Into<String>) {
        let msg = msg.into();
        tracing::error!(error = %msg, "user-visible error");
        self.error_message = Some(msg);
        self.mode = Mode::Error;
    }

    fn begin_goto_prompt(&mut self) {
        self.start_name_prompt("Enter address to go to:", PromptNext::Goto);
    }

    fn begin_edit(&mut self) {
        let pointer = self.wb().pointer;
        let source = self
            .wb()
            .cells
            .get(&pointer)
            .map(|c| c.source_form())
            .unwrap_or_default();
        let cursor = source.len();
        self.entry = Some(Entry {
            kind: EntryKind::Edit,
            buffer: source,
            cursor,
        });
        self.mode = Mode::Edit;
    }

    /// F2 mid-entry: promote LABEL/VALUE to EDIT preserving the buffer
    /// and cursor. No-op when already in EDIT.
    fn promote_entry_to_edit(&mut self) {
        if let Some(e) = self.entry.as_mut() {
            if !matches!(e.kind, EntryKind::Edit) {
                e.kind = EntryKind::Edit;
            }
        }
        if self.entry.is_some() {
            self.mode = Mode::Edit;
        }
    }

    /// Step `cursor` left by one char (UTF-8 safe).
    fn move_entry_cursor_left(&mut self) {
        let Some(e) = self.entry.as_mut() else {
            return;
        };
        if e.cursor == 0 {
            return;
        }
        let new_cursor = e.buffer[..e.cursor]
            .char_indices()
            .next_back()
            .map(|(i, _)| i)
            .unwrap_or(0);
        e.cursor = new_cursor;
    }

    /// Step `cursor` right by one char (UTF-8 safe).
    fn move_entry_cursor_right(&mut self) {
        let Some(e) = self.entry.as_mut() else {
            return;
        };
        if e.cursor >= e.buffer.len() {
            return;
        }
        let next = e.buffer[e.cursor..]
            .char_indices()
            .nth(1)
            .map(|(i, _)| e.cursor + i)
            .unwrap_or(e.buffer.len());
        e.cursor = next;
    }

    /// Delete the char before the cursor; cursor moves to the deleted
    /// char's start byte. No-op at cursor=0.
    fn entry_backspace(&mut self) {
        let Some(e) = self.entry.as_mut() else {
            return;
        };
        if e.cursor == 0 {
            return;
        }
        let prev = e.buffer[..e.cursor]
            .char_indices()
            .next_back()
            .map(|(i, _)| i)
            .unwrap_or(0);
        e.buffer.replace_range(prev..e.cursor, "");
        e.cursor = prev;
    }

    /// Delete the char at the cursor; cursor stays put. No-op at end.
    fn entry_delete(&mut self) {
        let Some(e) = self.entry.as_mut() else {
            return;
        };
        if e.cursor >= e.buffer.len() {
            return;
        }
        let next = e.buffer[e.cursor..]
            .char_indices()
            .nth(1)
            .map(|(i, _)| e.cursor + i)
            .unwrap_or(e.buffer.len());
        e.buffer.replace_range(e.cursor..next, "");
    }

    /// Insert `c` at the cursor; cursor advances past the new char.
    fn entry_insert_char(&mut self, c: char) {
        let Some(e) = self.entry.as_mut() else {
            return;
        };
        e.buffer.insert(e.cursor, c);
        e.cursor += c.len_utf8();
    }

    fn cancel_entry(&mut self) {
        self.entry = None;
        self.mode = Mode::Ready;
    }

    fn begin_entry(&mut self, c: char) {
        // `(L)` Label-only on the target cell forces every entry into
        // LABEL mode, including ones starting with a digit / minus /
        // operator. The default label prefix is auto-inserted; the
        // typed char is the first char of the label text. Explicit
        // label-prefix chars (`'`/`"`/`^`/`\`/`|`) still pick their
        // own prefix — the user is being explicit, so honor them.
        let pointer = self.wb().pointer;
        let label_only = matches!(
            self.wb().cell_formats.get(&pointer).copied(),
            Some(f) if matches!(f.kind, FormatKind::LabelOnly)
        ) || (!self.wb().cell_formats.contains_key(&pointer)
            && matches!(self.wb().global_format.kind, FormatKind::LabelOnly));
        if label_only && !matches!(c, '\'' | '"' | '^' | '\\' | '|') {
            let buffer = c.to_string();
            let cursor = buffer.len();
            self.entry = Some(Entry {
                kind: EntryKind::Label(self.default_label_prefix),
                buffer,
                cursor,
            });
            self.mode = Mode::Label;
            return;
        }
        if is_value_starter(c) {
            let buffer = c.to_string();
            let cursor = buffer.len();
            self.entry = Some(Entry {
                kind: EntryKind::Value,
                buffer,
                cursor,
            });
            self.mode = Mode::Value;
        } else if matches!(c, '\'' | '"' | '^' | '\\' | '|') {
            // Explicit label prefix typed first: the char becomes the
            // LabelPrefix; the buffer starts empty.
            let prefix = LabelPrefix::from_char(c).expect("matched above");
            self.entry = Some(Entry {
                kind: EntryKind::Label(prefix),
                buffer: String::new(),
                cursor: 0,
            });
            self.mode = Mode::Label;
        } else {
            // Any other non-value-starter: default `'` prefix auto-inserted;
            // the typed char is the first char of the label text.
            let buffer = c.to_string();
            let cursor = buffer.len();
            self.entry = Some(Entry {
                kind: EntryKind::Label(self.default_label_prefix),
                buffer,
                cursor,
            });
            self.mode = Mode::Label;
        }
    }

    fn commit_entry(&mut self) {
        let Some(entry) = self.entry.take() else {
            self.mode = Mode::Ready;
            return;
        };
        // Capture the prior state at the pointer before committing so
        // Alt-F4 can revert.
        if self.undo_enabled {
            let addr = self.wb().pointer;
            let prev_contents = self.wb().cells.get(&addr).cloned();
            let prev_format = self.wb().cell_formats.get(&addr).copied();
            self.wb_mut().journal.push(JournalEntry::CellEdit {
                addr,
                prev_contents,
                prev_format,
            });
        }
        let (mut contents, inferred_format) = match entry.kind {
            EntryKind::Label(prefix) => (
                CellContents::Label {
                    prefix,
                    text: entry.buffer,
                },
                None,
            ),
            EntryKind::Value => {
                let intl = self.wb().international.clone();
                match l123_core::parse_typed_value(&entry.buffer, &intl) {
                    Some(iv) => (CellContents::Constant(Value::Number(iv.number)), iv.format),
                    None => (
                        CellContents::Formula {
                            expr: entry.buffer,
                            cached_value: None,
                        },
                        None,
                    ),
                }
            }
            // EDIT commits re-parse the full source buffer so the user can
            // change prefix or type (label ↔ value) via the first-char rule.
            EntryKind::Edit => {
                let intl = self.wb().international.clone();
                CellContents::from_source_with_format(
                    &entry.buffer,
                    self.default_label_prefix,
                    &intl,
                )
            }
        };
        self.push_to_engine(&contents);
        match self.recalc_mode {
            RecalcMode::Automatic => {
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();
                self.recalc_pending = false;
                // Pick up the just-computed value for the committed cell.
                if let CellContents::Formula { expr, .. } = &contents {
                    let p = self.wb().pointer;
                    let view = self.wb_mut().engine.get_cell(p).ok();
                    let cached = view.map(|v| v.value);
                    contents = CellContents::Formula {
                        expr: expr.clone(),
                        cached_value: cached,
                    };
                }
            }
            RecalcMode::Manual => {
                self.recalc_pending = true;
            }
        }
        let p = self.wb().pointer;
        if contents.is_empty() {
            self.wb_mut().cells.remove(&p);
        } else {
            self.wb_mut().cells.insert(p, contents);
        }
        // Apply the inferred display format (Currency / Percent / Comma)
        // when the typed value carried Lotus-style markers. A plain
        // numeric commit leaves any pre-existing format alone — re-typing
        // `100` over a `(C2)` cell keeps the C2 format, matching 1-2-3.
        if let Some(fmt) = inferred_format {
            self.wb_mut().set_cell_format(p, fmt);
        }
        self.wb_mut().dirty = true;
        self.mode = Mode::Ready;
    }

    /// Record a batch of inverse entries. Empty batch and disabled
    /// undo are both no-ops. Single-entry batches are unwrapped to
    /// keep Batch usage to true multi-entry cases.
    fn push_journal_batch(&mut self, batch: Vec<JournalEntry>) {
        if !self.undo_enabled || batch.is_empty() {
            return;
        }
        let entry = if batch.len() == 1 {
            batch.into_iter().next().unwrap()
        } else {
            JournalEntry::Batch(batch)
        };
        self.wb_mut().journal.push(entry);
    }

    /// Pop the most recent journal entry and replay its inverse.
    /// No-op when the journal is empty.
    fn undo(&mut self) {
        let Some(entry) = self.wb_mut().journal.pop() else {
            return;
        };
        self.apply_undo(entry);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
    }

    fn apply_undo(&mut self, entry: JournalEntry) {
        match entry {
            JournalEntry::CellEdit {
                addr,
                prev_contents,
                prev_format,
            } => {
                match prev_contents {
                    Some(c) => {
                        self.push_to_engine_at(addr, &c);
                        self.wb_mut().cells.insert(addr, c);
                    }
                    None => {
                        self.wb_mut().cells.remove(&addr);
                        let _ = self.wb_mut().engine.clear_cell(addr);
                    }
                }
                match prev_format {
                    Some(f) => {
                        self.wb_mut().cell_formats.insert(addr, f);
                    }
                    None => {
                        self.wb_mut().cell_formats.remove(&addr);
                    }
                }
            }
            JournalEntry::RowDelete {
                sheet,
                at,
                cells,
                formats,
                text_styles,
            } => {
                if self.wb_mut().engine.insert_rows(sheet, at, 1).is_ok() {
                    shift_cells_rows(&mut self.wb_mut().cells, sheet, at, 1);
                    for (addr, contents) in cells {
                        self.push_to_engine_at(addr, &contents);
                        self.wb_mut().cells.insert(addr, contents);
                    }
                    for (addr, fmt) in formats {
                        self.wb_mut().cell_formats.insert(addr, fmt);
                    }
                    for (addr, style) in text_styles {
                        self.wb_mut().cell_text_styles.insert(addr, style);
                    }
                }
            }
            JournalEntry::RowInsert { sheet, at } => {
                if self.wb_mut().engine.delete_rows(sheet, at, 1).is_ok() {
                    self.wb_mut()
                        .cells
                        .retain(|a, _| !(a.sheet == sheet && a.row == at));
                    self.wb_mut()
                        .cell_formats
                        .retain(|a, _| !(a.sheet == sheet && a.row == at));
                    self.wb_mut()
                        .cell_text_styles
                        .retain(|a, _| !(a.sheet == sheet && a.row == at));
                    shift_cells_rows(&mut self.wb_mut().cells, sheet, at + 1, -1);
                }
            }
            JournalEntry::ColDelete {
                sheet,
                at,
                cells,
                formats,
                text_styles,
            } => {
                if self.wb_mut().engine.insert_cols(sheet, at, 1).is_ok() {
                    shift_cells_cols(&mut self.wb_mut().cells, sheet, at, 1);
                    for (addr, contents) in cells {
                        self.push_to_engine_at(addr, &contents);
                        self.wb_mut().cells.insert(addr, contents);
                    }
                    for (addr, fmt) in formats {
                        self.wb_mut().cell_formats.insert(addr, fmt);
                    }
                    for (addr, style) in text_styles {
                        self.wb_mut().cell_text_styles.insert(addr, style);
                    }
                }
            }
            JournalEntry::ColInsert { sheet, at } => {
                if self.wb_mut().engine.delete_cols(sheet, at, 1).is_ok() {
                    self.wb_mut()
                        .cells
                        .retain(|a, _| !(a.sheet == sheet && a.col == at));
                    self.wb_mut()
                        .cell_formats
                        .retain(|a, _| !(a.sheet == sheet && a.col == at));
                    self.wb_mut()
                        .cell_text_styles
                        .retain(|a, _| !(a.sheet == sheet && a.col == at));
                    shift_cells_cols(&mut self.wb_mut().cells, sheet, at + 1, -1);
                }
            }
            JournalEntry::RangeRestore {
                cells,
                formats,
                text_styles,
            } => {
                for (addr, contents) in cells {
                    self.push_to_engine_at(addr, &contents);
                    self.wb_mut().cells.insert(addr, contents);
                }
                for (addr, fmt) in formats {
                    self.wb_mut().cell_formats.insert(addr, fmt);
                }
                for (addr, style) in text_styles {
                    self.wb_mut().cell_text_styles.insert(addr, style);
                }
            }
            JournalEntry::RangeFormat { entries } => {
                for (addr, prev) in entries {
                    match prev {
                        Some(f) => {
                            self.wb_mut().cell_formats.insert(addr, f);
                        }
                        None => {
                            self.wb_mut().cell_formats.remove(&addr);
                        }
                    }
                }
            }
            JournalEntry::RangeTextStyle { entries } => {
                for (addr, prev) in entries {
                    match prev {
                        Some(s) => {
                            self.wb_mut().cell_text_styles.insert(addr, s);
                        }
                        None => {
                            self.wb_mut().cell_text_styles.remove(&addr);
                        }
                    }
                }
            }
            JournalEntry::RangeAlignment { entries } => {
                for (addr, prev) in entries {
                    match prev {
                        Some(a) => {
                            self.wb_mut().cell_alignments.insert(addr, a);
                        }
                        None => {
                            self.wb_mut().cell_alignments.remove(&addr);
                        }
                    }
                }
            }
            JournalEntry::RangeColor { entries } => {
                for (addr, prev_fill, prev_font) in entries {
                    match prev_fill {
                        Some(f) => {
                            self.wb_mut().cell_fills.insert(addr, f);
                        }
                        None => {
                            self.wb_mut().cell_fills.remove(&addr);
                        }
                    }
                    match prev_font {
                        Some(fs) => {
                            self.wb_mut().cell_font_styles.insert(addr, fs);
                        }
                        None => {
                            self.wb_mut().cell_font_styles.remove(&addr);
                        }
                    }
                }
            }
            JournalEntry::RangeBorder { entries } => {
                for (addr, prev) in entries {
                    match prev {
                        Some(b) => {
                            self.wb_mut().cell_borders.insert(addr, b);
                        }
                        None => {
                            self.wb_mut().cell_borders.remove(&addr);
                        }
                    }
                }
            }
            JournalEntry::ColWidth {
                sheet,
                col,
                prev_width,
            } => {
                let key = (sheet, col);
                match prev_width {
                    Some(w) => {
                        self.wb_mut().col_widths.insert(key, w);
                    }
                    None => {
                        self.wb_mut().col_widths.remove(&key);
                    }
                }
            }
            JournalEntry::ColHidden {
                sheet,
                col,
                prev_hidden,
            } => {
                let key = (sheet, col);
                if prev_hidden {
                    self.wb_mut().hidden_cols.insert(key);
                } else {
                    self.wb_mut().hidden_cols.remove(&key);
                }
            }
            JournalEntry::GlobalColWidth { prev } => {
                self.wb_mut().default_col_width = prev;
            }
            JournalEntry::GlobalFormat { prev } => {
                self.wb_mut().global_format = prev;
            }
            JournalEntry::GlobalInternational { prev } => {
                self.wb_mut().international = prev;
            }
            JournalEntry::DefaultLabelPrefix { prev } => {
                self.default_label_prefix = prev;
            }
            JournalEntry::Frozen { sheet, prev } => match prev {
                Some(f) => {
                    self.wb_mut().frozen.insert(sheet, f);
                }
                None => {
                    self.wb_mut().frozen.remove(&sheet);
                }
            },
            JournalEntry::SheetVisibility { sheet, prev } => {
                if prev == SheetState::Visible {
                    self.wb_mut().sheet_states.remove(&sheet);
                } else {
                    self.wb_mut().sheet_states.insert(sheet, prev);
                }
            }
            JournalEntry::RangeNameReset { prev } => {
                for (name, range) in prev {
                    let _ = self.wb_mut().engine.define_name(&name, range);
                    self.wb_mut()
                        .named_ranges
                        .insert(name.to_ascii_lowercase(), range);
                }
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();
            }
            JournalEntry::RangeNameLabels {
                created,
                overwritten,
            } => {
                for name in created {
                    let _ = self.wb_mut().engine.delete_name(&name);
                    self.wb_mut().named_ranges.remove(&name);
                }
                for (name, range) in overwritten {
                    let _ = self.wb_mut().engine.define_name(&name, range);
                    self.wb_mut().named_ranges.insert(name, range);
                }
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();
            }
            JournalEntry::RangeNameUndefine {
                name,
                range,
                note,
                cell_writes,
            } => {
                for (addr, prev) in cell_writes {
                    self.restore_cell_contents(addr, prev);
                }
                let _ = self.wb_mut().engine.define_name(&name, range);
                self.wb_mut()
                    .named_ranges
                    .insert(name.to_ascii_lowercase(), range);
                if let Some(text) = note {
                    self.wb_mut()
                        .name_notes
                        .insert(name.to_ascii_lowercase(), text);
                }
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();
            }
            JournalEntry::RangeNameNote { name, prev } => {
                let key = name.to_ascii_lowercase();
                match prev {
                    Some(text) => {
                        self.wb_mut().name_notes.insert(key, text);
                    }
                    None => {
                        self.wb_mut().name_notes.remove(&key);
                    }
                }
            }
            JournalEntry::RangeNameNoteReset { prev } => {
                for (name, text) in prev {
                    self.wb_mut().name_notes.insert(name, text);
                }
            }
            JournalEntry::RangeProtection { entries } => {
                for (addr, was_unprotected) in entries {
                    if was_unprotected {
                        self.wb_mut().cell_unprotected.insert(addr);
                    } else {
                        self.wb_mut().cell_unprotected.remove(&addr);
                    }
                }
            }
            JournalEntry::Batch(entries) => {
                // Apply in reverse order so the "outer" state restores
                // after the "inner" details.
                for e in entries.into_iter().rev() {
                    self.apply_undo(e);
                }
            }
        }
    }

    /// Explicit recalculation — invoked by F9 in READY mode. Safe to call
    /// repeatedly; no-op in terms of values but always clears the pending
    /// flag. PLAN §4.7: workbooks above `RECALC_WAIT_CELL_THRESHOLD`
    /// route through the async WAIT path so the UI doesn't freeze on
    /// big sheets; smaller workbooks stay synchronous to avoid a
    /// tokio round-trip on every F9.
    fn do_recalc(&mut self) {
        if self.wb().cells.len() > self.recalc_wait_cell_threshold {
            let placeholder = IronCalcEngine::new().expect("IronCalc placeholder engine init");
            let engine = std::mem::replace(&mut self.wb_mut().engine, placeholder);
            self.queue_async_op("Recalculating", String::new(), QueuedOp::Recalc { engine });
        } else {
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();
            self.recalc_pending = false;
        }
    }

    // ---------------- MENU mode ----------------

    fn open_menu(&mut self) {
        self.menu = Some(MenuState::fresh());
        self.mode = Mode::Menu;
    }

    /// `:` in READY opens the WYSIWYG colon-menu.  Uses a secondary root
    /// so the existing menu navigation, help, and descent machinery
    /// applies unchanged.
    fn open_wysiwyg_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::WYSIWYG_ROOT));
        self.mode = Mode::Menu;
    }

    fn close_menu(&mut self) {
        self.menu = None;
        self.mode = Mode::Ready;
    }

    fn set_graph_type(&mut self, t: GraphType) {
        self.wb_mut().current_graph.graph_type = t;
        self.close_menu();
    }

    /// F10 / `/Graph View`. Snapshot the current graph's series values
    /// and transition to [`Mode::Graph`]. An empty graph shows a "define
    /// ranges first" placeholder rather than silently no-op'ing so the
    /// user sees something happened.
    fn enter_graph_view(&mut self) {
        let def = self.wb().current_graph.clone();
        let values = self.collect_graph_values(&def);
        let img = if self.picker_is_graphical() && !values.is_empty() {
            let png = l123_graph::render_png(&def, &values);
            image::load_from_memory(&png).ok()
        } else {
            None
        };
        self.graph_view = Some(GraphOverlay { values, img });
        self.menu = None;
        self.mode = Mode::Graph;
    }

    fn enter_worksheet_status(&mut self) {
        self.menu = None;
        self.stat_view = StatView::Worksheet;
        self.mode = Mode::Stat;
    }

    fn enter_defaults_status(&mut self) {
        self.menu = None;
        self.stat_view = StatView::Defaults;
        self.mode = Mode::Stat;
    }

    fn start_wgd_path_prompt(&mut self, next: PromptNext, label: &str, current: String) {
        self.menu = None;
        let fresh = !current.is_empty();
        self.prompt = Some(PromptState {
            label: label.into(),
            buffer: current,
            next,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn start_wgd_numeric_prompt(&mut self, next: PromptNext, label: &str, current: u32) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: label.into(),
            buffer: current.to_string(),
            next,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Global Default Update` — write the current defaults
    /// back to the L123.CNF config file. Failure is silent so it doesn't
    /// hijack the menu flow; the user can re-run after fixing perms.
    fn execute_wgd_update(&mut self) {
        if let Some(path) = crate::config::default_config_path() {
            let _ = self.defaults.write_to_path(&path);
        }
        self.close_menu();
    }

    fn picker_is_graphical(&self) -> bool {
        match self.image_picker.as_ref() {
            Some(p) => p.protocol_type() != ProtocolType::Halfblocks,
            None => false,
        }
    }

    /// Called once by [`App::run`] after raw mode is enabled. Queries
    /// the terminal for its graphics-protocol capability; if the query
    /// fails (tmux, legacy terminals, redirected stdio) the picker is
    /// left as `None` and F10 uses the unicode renderer. When the
    /// picker reports a non-halfblocks protocol, also pre-decode the
    /// v3.1 WYSIWYG icon panel PNG so it's ready at first draw.
    pub fn probe_image_picker(&mut self) {
        let mut picker = Picker::from_query_stdio().ok();

        // iTerm2-family hosts need the OSC 1337 "Iterm2" protocol, but
        // `Picker::from_query_stdio` can steer us away from it in two
        // ways: (1) when font-size detection fails, the library drops
        // to Halfblocks and discards its own iTerm2 env hint; (2)
        // iTerm2 3.5+ advertises partial Kitty graphics support, so
        // Kitty wins the stdio probe — but iTerm2 doesn't implement
        // the Unicode-placeholder variant ratatui-image renders with,
        // so nothing actually draws. Mirror the WezTerm/Konsole
        // treatment already in upstream and force Iterm2 in both
        // cases. Sixel is left alone: when iTerm2 users turn it on,
        // it genuinely works.
        let needs_iterm2_override = picker.as_ref().is_none_or(|p| {
            matches!(
                p.protocol_type(),
                ProtocolType::Halfblocks | ProtocolType::Kitty,
            )
        });
        if needs_iterm2_override {
            let term_program = std::env::var("TERM_PROGRAM").ok();
            let lc_terminal = std::env::var("LC_TERMINAL").ok();
            if is_iterm2_compatible_env(term_program.as_deref(), lc_terminal.as_deref()) {
                let mut p = picker.take().unwrap_or_else(Picker::halfblocks);
                p.set_protocol_type(ProtocolType::Iterm2);
                picker = Some(p);
            }
        }

        self.image_picker = picker;
        if self.picker_is_graphical() {
            self.refresh_icon_panel();
        }
    }

    /// Re-rasterize the current panel into a [`DynamicImage`] ready
    /// for ratatui-image. Called at startup and whenever the user
    /// pages to a different panel via the slot-16 navigator.
    fn refresh_icon_panel(&mut self) {
        let bytes = l123_graph::render_panel_png(self.current_panel);
        self.icon_panel = image::load_from_memory(&bytes).ok();
    }

    fn start_graph_save_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter graph file name:".into(),
            buffer: String::new(),
            next: PromptNext::GraphSaveFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn commit_graph_save(&mut self, buffer: &str) {
        let trimmed = buffer.trim();
        if trimmed.is_empty() {
            self.mode = Mode::Ready;
            return;
        }
        // Match extension to format: svg is the default, .cgm is
        // preserved if typed but the bytes are still SVG. Anything
        // else also gets SVG bytes — the user got what they asked for.
        let path = if std::path::Path::new(trimmed).extension().is_some() {
            PathBuf::from(trimmed)
        } else {
            PathBuf::from(format!("{trimmed}.svg"))
        };
        let def = self.wb().current_graph.clone();
        let values = self.collect_graph_values(&def);
        let svg = l123_graph::render_svg(&def, &values);
        let _ = std::fs::write(&path, svg);
        self.mode = Mode::Ready;
    }

    fn collect_graph_values(&self, def: &GraphDef) -> l123_graph::GraphValues {
        let mut out = l123_graph::GraphValues::default();
        if let Some(r) = def.x {
            out.x = Some(self.read_series_values(r));
        }
        for (i, slot) in def.data.iter().enumerate() {
            if let Some(r) = *slot {
                out.data[i] = Some(self.read_series_values(r));
            }
        }
        out
    }

    /// Flatten a range to a sequence of numeric values. Non-numeric or
    /// error cells become `NaN` so positional alignment is preserved.
    /// Column-major: for a single-column range this is just the column
    /// read top to bottom, which is the common 1-2-3 idiom.
    fn read_series_values(&self, r: Range) -> Vec<f64> {
        let n = r.normalized();
        let mut out = Vec::new();
        for col in n.start.col..=n.end.col {
            for row in n.start.row..=n.end.row {
                let addr = Address {
                    sheet: n.start.sheet,
                    col,
                    row,
                };
                let v = match self.wb().engine.get_cell(addr) {
                    Ok(cv) => match cv.value {
                        Value::Number(f) => f,
                        _ => f64::NAN,
                    },
                    Err(_) => f64::NAN,
                };
                out.push(v);
            }
        }
        out
    }

    fn descend_highlighted(&mut self) {
        let Some(state) = self.menu.as_ref() else {
            return;
        };
        let Some(item) = state.highlighted() else {
            return;
        };
        self.descend_into(item);
    }

    fn descend_by_letter(&mut self, c: char) {
        let Some(state) = self.menu.as_ref() else {
            return;
        };
        let level = state.level();
        let item = level
            .iter()
            .find(|m| m.letter.eq_ignore_ascii_case(&c))
            .copied();
        if let Some(item) = item {
            self.descend_into(&item);
        }
    }

    fn execute_action(&mut self, action: Action) {
        match action {
            Action::Cancel => self.close_menu(),
            Action::QuitConfirm => {
                if self.is_dirty() {
                    self.menu = Some(MenuState::rooted_at(menu::QUIT_DIRTY_MENU));
                } else {
                    self.running = false;
                    self.close_menu();
                }
            }
            Action::Quit => {
                self.running = false;
                self.close_menu();
            }
            Action::WorksheetInsertRow => self.insert_row_at_pointer(1),
            Action::WorksheetInsertColumn => self.insert_col_at_pointer(1),
            Action::WorksheetInsertSheetBefore => self.insert_sheet_before_current(),
            Action::WorksheetInsertSheetAfter => self.insert_sheet_after_current(),
            Action::WorksheetDeleteRow => self.delete_row_at_pointer(1),
            Action::WorksheetDeleteColumn => self.delete_col_at_pointer(1),
            Action::WorksheetDeleteSheet => self.delete_sheet_at_pointer(),
            Action::WorksheetDeleteFile => self.delete_current_file(),
            Action::WorksheetGlobalRecalcAutomatic => {
                self.recalc_mode = RecalcMode::Automatic;
                // Switching into Automatic catches up on any pending work.
                self.do_recalc();
                self.close_menu();
            }
            Action::WorksheetGlobalRecalcManual => {
                self.recalc_mode = RecalcMode::Manual;
                self.close_menu();
            }
            Action::WorksheetGlobalRecalcNatural => {
                self.recalc_order = RecalcOrder::Natural;
                self.close_menu();
            }
            Action::WorksheetGlobalRecalcColumnwise => {
                self.recalc_order = RecalcOrder::Columnwise;
                self.close_menu();
            }
            Action::WorksheetGlobalRecalcRowwise => {
                self.recalc_order = RecalcOrder::Rowwise;
                self.close_menu();
            }
            Action::WorksheetGlobalRecalcIteration => {
                self.start_recalc_iteration_prompt();
            }
            Action::WorksheetGlobalZeroNo => {
                self.zero_display = ZeroDisplay::No;
                self.close_menu();
            }
            Action::WorksheetGlobalZeroYes => {
                self.zero_display = ZeroDisplay::Yes;
                self.close_menu();
            }
            Action::WorksheetGlobalProtectionEnable => {
                self.global_protection = true;
                self.close_menu();
            }
            Action::WorksheetGlobalProtectionDisable => {
                self.global_protection = false;
                self.close_menu();
            }
            Action::WorksheetGlobalGroupEnable => {
                self.group_mode = true;
                self.close_menu();
            }
            Action::WorksheetGlobalGroupDisable => {
                self.group_mode = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherUndoEnable => {
                self.undo_enabled = true;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherUndoDisable => {
                // Clear the existing journal so Alt-F4 can't pop a
                // pre-disable entry after the user has explicitly
                // turned undo off.
                self.wb_mut().journal.clear();
                self.undo_enabled = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherBeepEnable => {
                self.beep_enabled = true;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherBeepDisable => {
                // Drop any pending beep so the transition is clean —
                // the user just told us to be quiet.
                self.beep_pending = false;
                self.beep_enabled = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationA => {
                self.set_punctuation(Punctuation::A)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationB => {
                self.set_punctuation(Punctuation::B)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationC => {
                self.set_punctuation(Punctuation::C)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationD => {
                self.set_punctuation(Punctuation::D)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationE => {
                self.set_punctuation(Punctuation::E)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationF => {
                self.set_punctuation(Punctuation::F)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationG => {
                self.set_punctuation(Punctuation::G)
            }
            Action::WorksheetGlobalDefaultOtherIntlPunctuationH => {
                self.set_punctuation(Punctuation::H)
            }
            Action::WorksheetGlobalDefaultOtherIntlCurrencyPrefix => {
                self.start_currency_symbol_prompt(CurrencyPosition::Prefix)
            }
            Action::WorksheetGlobalDefaultOtherIntlCurrencySuffix => {
                self.start_currency_symbol_prompt(CurrencyPosition::Suffix)
            }
            Action::WorksheetGlobalDefaultOtherIntlDateA => self.set_date_intl(DateIntl::A),
            Action::WorksheetGlobalDefaultOtherIntlDateB => self.set_date_intl(DateIntl::B),
            Action::WorksheetGlobalDefaultOtherIntlDateC => self.set_date_intl(DateIntl::C),
            Action::WorksheetGlobalDefaultOtherIntlDateD => self.set_date_intl(DateIntl::D),
            Action::WorksheetGlobalDefaultOtherIntlTimeA => self.set_time_intl(TimeIntl::A),
            Action::WorksheetGlobalDefaultOtherIntlTimeB => self.set_time_intl(TimeIntl::B),
            Action::WorksheetGlobalDefaultOtherIntlTimeC => self.set_time_intl(TimeIntl::C),
            Action::WorksheetGlobalDefaultOtherIntlTimeD => self.set_time_intl(TimeIntl::D),
            Action::WorksheetGlobalDefaultOtherIntlNegativeParens => {
                self.set_negative_style(NegativeStyle::Parens)
            }
            Action::WorksheetGlobalDefaultOtherIntlNegativeSign => {
                self.set_negative_style(NegativeStyle::Sign)
            }
            Action::WorksheetTitlesBoth => self.set_titles(TitlesKind::Both),
            Action::WorksheetTitlesHorizontal => self.set_titles(TitlesKind::Horizontal),
            Action::WorksheetTitlesVertical => self.set_titles(TitlesKind::Vertical),
            Action::WorksheetTitlesClear => self.clear_titles(),
            Action::WorksheetPageRow => self.insert_page_break_row_at_pointer(),
            Action::WorksheetPageColumn => self.insert_page_break_column_at_pointer(),
            Action::WorksheetHideEnable => self.hide_current_sheet(),
            Action::WorksheetHideDisable => self.unhide_all_sheets(),
            Action::WorksheetLearnRange => {
                self.begin_point(PendingCommand::WorksheetLearnRange);
            }
            Action::WorksheetLearnCancel => self.cancel_learn(),
            Action::WorksheetLearnErase => self.erase_learn_range(),
            Action::WorksheetGlobalDefaultOtherClockStandard => {
                self.clock_display = ClockDisplay::Standard;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherClockInternational => {
                self.clock_display = ClockDisplay::International;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherClockNone => {
                self.clock_display = ClockDisplay::None;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultOtherClockFilename => {
                self.clock_display = ClockDisplay::Filename;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultStatus => self.enter_defaults_status(),
            Action::WorksheetGlobalDefaultUpdate => self.execute_wgd_update(),
            Action::WorksheetGlobalDefaultDir => self.start_wgd_path_prompt(
                PromptNext::WgdDir,
                "Enter default directory:",
                self.defaults.default_dir.clone(),
            ),
            Action::WorksheetGlobalDefaultTemp => self.start_wgd_path_prompt(
                PromptNext::WgdTemp,
                "Enter temporary file directory:",
                self.defaults.temp_dir.clone(),
            ),
            Action::WorksheetGlobalDefaultAutoexecYes => {
                self.defaults.autoexec = true;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultAutoexecNo => {
                self.defaults.autoexec = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultExtSave => self.start_wgd_path_prompt(
                PromptNext::WgdExtSave,
                "Enter default save extension:",
                self.defaults.ext_save.clone(),
            ),
            Action::WorksheetGlobalDefaultExtList => self.start_wgd_path_prompt(
                PromptNext::WgdExtList,
                "Enter default file-list extension:",
                self.defaults.ext_list.clone(),
            ),
            Action::WorksheetGlobalDefaultPrinterInterface => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterInterface,
                "Enter printer interface (1..9):",
                u32::from(self.defaults.printer_interface),
            ),
            Action::WorksheetGlobalDefaultPrinterAutoLfYes => {
                self.defaults.printer_autolf = true;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultPrinterAutoLfNo => {
                self.defaults.printer_autolf = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultPrinterMarginLeft => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterMarginLeft,
                "Enter default left margin (0..1000):",
                u32::from(self.defaults.printer_left),
            ),
            Action::WorksheetGlobalDefaultPrinterMarginRight => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterMarginRight,
                "Enter default right margin (0..1000):",
                u32::from(self.defaults.printer_right),
            ),
            Action::WorksheetGlobalDefaultPrinterMarginTop => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterMarginTop,
                "Enter default top margin (0..1000):",
                u32::from(self.defaults.printer_top),
            ),
            Action::WorksheetGlobalDefaultPrinterMarginBottom => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterMarginBottom,
                "Enter default bottom margin (0..1000):",
                u32::from(self.defaults.printer_bottom),
            ),
            Action::WorksheetGlobalDefaultPrinterPgLength => self.start_wgd_numeric_prompt(
                PromptNext::WgdPrinterPgLength,
                "Enter default page length (1..1000):",
                u32::from(self.defaults.printer_pg_length),
            ),
            Action::WorksheetGlobalDefaultPrinterWaitYes => {
                self.defaults.printer_wait = true;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultPrinterWaitNo => {
                self.defaults.printer_wait = false;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultPrinterSetup => self.start_wgd_path_prompt(
                PromptNext::WgdPrinterSetup,
                "Enter default printer setup string:",
                self.defaults.printer_setup.clone(),
            ),
            Action::WorksheetGlobalDefaultPrinterName => self.start_wgd_path_prompt(
                PromptNext::WgdPrinterName,
                "Enter default printer name:",
                self.defaults.printer_name.clone(),
            ),
            Action::WorksheetGlobalDefaultPrinterQuit => self.close_menu(),
            Action::WorksheetGlobalDefaultGraphGroupColumnwise => {
                self.defaults.graph_group = GraphGroupOrientation::Columnwise;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultGraphGroupRowwise => {
                self.defaults.graph_group = GraphGroupOrientation::Rowwise;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultGraphSaveCgm => {
                self.defaults.graph_save = GraphSaveFormat::Cgm;
                self.close_menu();
            }
            Action::WorksheetGlobalDefaultGraphSavePic => {
                self.defaults.graph_save = GraphSaveFormat::Pic;
                self.close_menu();
            }
            Action::WorksheetEraseConfirm => self.execute_worksheet_erase(),
            Action::WorksheetColumnSetWidth => self.start_col_width_prompt(),
            Action::WorksheetColumnResetWidth => self.execute_col_reset_width(),
            Action::WorksheetColumnRangeSetWidth => self.start_col_range_width_prompt(),
            Action::WorksheetColumnRangeResetWidth => {
                self.begin_point(PendingCommand::ColumnRangeResetWidth)
            }
            Action::WorksheetColumnHide => self.begin_point(PendingCommand::ColumnHide),
            Action::WorksheetColumnDisplay => self.begin_point(PendingCommand::ColumnDisplay),
            Action::WorksheetStatus => self.enter_worksheet_status(),
            Action::WorksheetGlobalColWidth => self.start_global_col_width_prompt(),
            Action::WorksheetGlobalLabelLeft => {
                self.set_default_label_prefix(LabelPrefix::Apostrophe)
            }
            Action::WorksheetGlobalLabelRight => self.set_default_label_prefix(LabelPrefix::Quote),
            Action::WorksheetGlobalLabelCenter => self.set_default_label_prefix(LabelPrefix::Caret),
            Action::RangeNameCreate => {
                self.start_name_prompt("Enter name:", PromptNext::RangeNameCreate)
            }
            Action::RangeNameDelete => {
                self.start_name_prompt("Enter name to delete:", PromptNext::RangeNameDelete)
            }
            Action::RangeNameReset => self.range_name_reset(),
            Action::RangeNameLabelsRight => self.begin_point(PendingCommand::RangeNameLabels {
                direction: LabelDirection::Right,
            }),
            Action::RangeNameLabelsDown => self.begin_point(PendingCommand::RangeNameLabels {
                direction: LabelDirection::Down,
            }),
            Action::RangeNameLabelsLeft => self.begin_point(PendingCommand::RangeNameLabels {
                direction: LabelDirection::Left,
            }),
            Action::RangeNameLabelsUp => self.begin_point(PendingCommand::RangeNameLabels {
                direction: LabelDirection::Up,
            }),
            Action::RangeNameTable => self.begin_point(PendingCommand::RangeNameTable),
            Action::RangeNameUndefine => {
                self.start_name_prompt("Enter name to undefine:", PromptNext::RangeNameUndefine)
            }
            Action::RangeNameNoteCreate => {
                self.start_name_prompt("Enter name to annotate:", PromptNext::RangeNameNoteCreate)
            }
            Action::RangeNameNoteDelete => self.start_name_prompt(
                "Enter name whose note to delete:",
                PromptNext::RangeNameNoteDelete,
            ),
            Action::RangeNameNoteReset => self.range_name_note_reset(),
            Action::RangeNameNoteTable => self.begin_point(PendingCommand::RangeNameNoteTable),
            Action::RangeProtect => {
                self.begin_point(PendingCommand::RangeProtect { unprotected: false })
            }
            Action::RangeUnprotect => {
                self.begin_point(PendingCommand::RangeProtect { unprotected: true })
            }
            Action::RangeInput => self.begin_point(PendingCommand::RangeInput),
            Action::RangeValue => self.begin_point(PendingCommand::RangeValueFrom),
            Action::RangeTrans => self.begin_point(PendingCommand::RangeTransFrom),
            Action::RangeJustify => self.begin_point(PendingCommand::RangeJustify),
            Action::RangeErase => self.begin_point(PendingCommand::RangeErase),
            Action::Copy => self.begin_point(PendingCommand::CopyFrom),
            Action::Move => self.begin_point(PendingCommand::MoveFrom),
            Action::RangeLabelLeft => self.begin_point(PendingCommand::RangeLabel {
                new_prefix: LabelPrefix::Apostrophe,
            }),
            Action::RangeLabelRight => self.begin_point(PendingCommand::RangeLabel {
                new_prefix: LabelPrefix::Quote,
            }),
            Action::RangeLabelCenter => self.begin_point(PendingCommand::RangeLabel {
                new_prefix: LabelPrefix::Caret,
            }),
            Action::RangeFormatFixed => self.start_decimals_prompt(FormatKind::Fixed),
            Action::RangeFormatScientific => self.start_decimals_prompt(FormatKind::Scientific),
            Action::RangeFormatCurrency => self.start_decimals_prompt(FormatKind::Currency),
            Action::RangeFormatComma => self.start_decimals_prompt(FormatKind::Comma),
            Action::RangeFormatPercent => self.start_decimals_prompt(FormatKind::Percent),
            Action::RangeFormatGeneral => self.begin_point(PendingCommand::RangeFormat {
                format: Format::GENERAL,
            }),
            Action::RangeFormatPlusMinus => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::PlusMinus),
            }),
            Action::RangeFormatAutomatic => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::Automatic),
            }),
            Action::RangeFormatLabelOnly => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::LabelOnly),
            }),
            Action::RangeFormatParensYes => {
                self.begin_point(PendingCommand::RangeParens { value: true })
            }
            Action::RangeFormatParensNo => {
                self.begin_point(PendingCommand::RangeParens { value: false })
            }
            Action::RangeFormatNegColorBlack => self.begin_neg_color(Some(PALETTE_BLACK)),
            Action::RangeFormatNegColorWhite => self.begin_neg_color(Some(PALETTE_WHITE)),
            Action::RangeFormatNegColorRed => self.begin_neg_color(Some(PALETTE_RED)),
            Action::RangeFormatNegColorGreen => self.begin_neg_color(Some(PALETTE_GREEN)),
            Action::RangeFormatNegColorBlue => self.begin_neg_color(Some(PALETTE_BLUE)),
            Action::RangeFormatNegColorYellow => self.begin_neg_color(Some(PALETTE_YELLOW)),
            Action::RangeFormatNegColorCyan => self.begin_neg_color(Some(PALETTE_CYAN)),
            Action::RangeFormatNegColorMagenta => self.begin_neg_color(Some(PALETTE_MAGENTA)),
            Action::RangeFormatNegColorReset => self.begin_neg_color(None),
            Action::RangeFormatReset => self.begin_point(PendingCommand::RangeFormat {
                format: Format::RESET,
            }),
            Action::FormatBoldSet => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::BOLD,
                set: true,
            }),
            Action::FormatBoldClear => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::BOLD,
                set: false,
            }),
            Action::FormatItalicSet => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::ITALIC,
                set: true,
            }),
            Action::FormatItalicClear => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::ITALIC,
                set: false,
            }),
            Action::FormatUnderlineSet => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::UNDERLINE,
                set: true,
            }),
            Action::FormatUnderlineClear => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle::UNDERLINE,
                set: false,
            }),
            Action::FormatReset => self.begin_point(PendingCommand::RangeTextStyle {
                bits: TextStyle {
                    bold: true,
                    italic: true,
                    underline: true,
                },
                set: false,
            }),
            Action::FormatAlignmentLeft => self.begin_point(PendingCommand::RangeAlignment {
                halign: HAlign::Left,
            }),
            Action::FormatAlignmentRight => self.begin_point(PendingCommand::RangeAlignment {
                halign: HAlign::Right,
            }),
            Action::FormatAlignmentCenter => self.begin_point(PendingCommand::RangeAlignment {
                halign: HAlign::Center,
            }),
            Action::FormatAlignmentGeneral => self.begin_point(PendingCommand::RangeAlignment {
                halign: HAlign::General,
            }),
            Action::FormatColorBgBlack => self.begin_color(ColorTarget::Background, PALETTE_BLACK),
            Action::FormatColorBgWhite => self.begin_color(ColorTarget::Background, PALETTE_WHITE),
            Action::FormatColorBgRed => self.begin_color(ColorTarget::Background, PALETTE_RED),
            Action::FormatColorBgGreen => self.begin_color(ColorTarget::Background, PALETTE_GREEN),
            Action::FormatColorBgBlue => self.begin_color(ColorTarget::Background, PALETTE_BLUE),
            Action::FormatColorBgYellow => {
                self.begin_color(ColorTarget::Background, PALETTE_YELLOW)
            }
            Action::FormatColorBgCyan => self.begin_color(ColorTarget::Background, PALETTE_CYAN),
            Action::FormatColorBgMagenta => {
                self.begin_color(ColorTarget::Background, PALETTE_MAGENTA)
            }
            Action::FormatColorTextBlack => self.begin_color(ColorTarget::Text, PALETTE_BLACK),
            Action::FormatColorTextWhite => self.begin_color(ColorTarget::Text, PALETTE_WHITE),
            Action::FormatColorTextRed => self.begin_color(ColorTarget::Text, PALETTE_RED),
            Action::FormatColorTextGreen => self.begin_color(ColorTarget::Text, PALETTE_GREEN),
            Action::FormatColorTextBlue => self.begin_color(ColorTarget::Text, PALETTE_BLUE),
            Action::FormatColorTextYellow => self.begin_color(ColorTarget::Text, PALETTE_YELLOW),
            Action::FormatColorTextCyan => self.begin_color(ColorTarget::Text, PALETTE_CYAN),
            Action::FormatColorTextMagenta => self.begin_color(ColorTarget::Text, PALETTE_MAGENTA),
            Action::FormatColorReset => self.begin_point(PendingCommand::RangeColor {
                target: ColorTarget::Both,
                color: None,
            }),
            Action::DisplayModeColor => {
                self.display_mode = DisplayMode::Color;
                self.close_menu();
            }
            Action::DisplayModeBW => {
                self.display_mode = DisplayMode::BW;
                self.close_menu();
            }
            Action::DisplayModeReverse => {
                self.display_mode = DisplayMode::Reverse;
                self.close_menu();
            }
            Action::DisplayOptionsGridYes => {
                self.show_gridlines = true;
                self.close_menu();
            }
            Action::DisplayOptionsGridNo => {
                self.show_gridlines = false;
                self.close_menu();
            }
            Action::SpecialCopy => self.begin_point(PendingCommand::SpecialCopyFrom),
            Action::SpecialMove => self.begin_point(PendingCommand::SpecialMoveFrom),
            Action::RangeFormatText => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::Text),
            }),
            Action::RangeFormatDateDmy => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::DateDmy),
            }),
            Action::RangeFormatDateDm => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::DateDm),
            }),
            Action::RangeFormatDateMy => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::DateMy),
            }),
            Action::RangeFormatDateLongIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::DateLongIntl),
            }),
            Action::RangeFormatDateShortIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::DateShortIntl),
            }),
            Action::RangeFormatHidden => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::Hidden),
            }),
            Action::RangeFormatTimeHmsAmPm => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::TimeHmsAmPm),
            }),
            Action::RangeFormatTimeHmAmPm => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::TimeHmAmPm),
            }),
            Action::RangeFormatTimeLongIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::TimeLongIntl),
            }),
            Action::RangeFormatTimeShortIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format::from_kind(FormatKind::TimeShortIntl),
            }),
            Action::WorksheetGlobalFormatFixed => {
                self.start_global_decimals_prompt(FormatKind::Fixed)
            }
            Action::WorksheetGlobalFormatScientific => {
                self.start_global_decimals_prompt(FormatKind::Scientific)
            }
            Action::WorksheetGlobalFormatCurrency => {
                self.start_global_decimals_prompt(FormatKind::Currency)
            }
            Action::WorksheetGlobalFormatComma => {
                self.start_global_decimals_prompt(FormatKind::Comma)
            }
            Action::WorksheetGlobalFormatPercent => {
                self.start_global_decimals_prompt(FormatKind::Percent)
            }
            Action::WorksheetGlobalFormatGeneral => self.set_global_format(Format::GENERAL),
            Action::WorksheetGlobalFormatPlusMinus => {
                self.set_global_format(Format::from_kind(FormatKind::PlusMinus))
            }
            Action::WorksheetGlobalFormatAutomatic => {
                self.set_global_format(Format::from_kind(FormatKind::Automatic))
            }
            Action::WorksheetGlobalFormatLabelOnly => {
                self.set_global_format(Format::from_kind(FormatKind::LabelOnly))
            }
            Action::WorksheetGlobalFormatParensYes => self.set_global_parens(true),
            Action::WorksheetGlobalFormatParensNo => self.set_global_parens(false),
            Action::WorksheetGlobalFormatNegColorBlack => {
                self.set_global_neg_color(Some(PALETTE_BLACK))
            }
            Action::WorksheetGlobalFormatNegColorWhite => {
                self.set_global_neg_color(Some(PALETTE_WHITE))
            }
            Action::WorksheetGlobalFormatNegColorRed => {
                self.set_global_neg_color(Some(PALETTE_RED))
            }
            Action::WorksheetGlobalFormatNegColorGreen => {
                self.set_global_neg_color(Some(PALETTE_GREEN))
            }
            Action::WorksheetGlobalFormatNegColorBlue => {
                self.set_global_neg_color(Some(PALETTE_BLUE))
            }
            Action::WorksheetGlobalFormatNegColorYellow => {
                self.set_global_neg_color(Some(PALETTE_YELLOW))
            }
            Action::WorksheetGlobalFormatNegColorCyan => {
                self.set_global_neg_color(Some(PALETTE_CYAN))
            }
            Action::WorksheetGlobalFormatNegColorMagenta => {
                self.set_global_neg_color(Some(PALETTE_MAGENTA))
            }
            Action::WorksheetGlobalFormatNegColorReset => self.set_global_neg_color(None),
            Action::WorksheetGlobalFormatReset => self.set_global_format(Format::GENERAL),
            Action::WorksheetGlobalFormatText => {
                self.set_global_format(Format::from_kind(FormatKind::Text))
            }
            Action::WorksheetGlobalFormatDateDmy => {
                self.set_global_format(Format::from_kind(FormatKind::DateDmy))
            }
            Action::WorksheetGlobalFormatDateDm => {
                self.set_global_format(Format::from_kind(FormatKind::DateDm))
            }
            Action::WorksheetGlobalFormatDateMy => {
                self.set_global_format(Format::from_kind(FormatKind::DateMy))
            }
            Action::WorksheetGlobalFormatDateLongIntl => {
                self.set_global_format(Format::from_kind(FormatKind::DateLongIntl))
            }
            Action::WorksheetGlobalFormatDateShortIntl => {
                self.set_global_format(Format::from_kind(FormatKind::DateShortIntl))
            }
            Action::FileSave => self.start_file_save_prompt(),
            Action::FileRetrieve => self.start_file_retrieve_prompt(),
            Action::FileXtractFormulas => self.start_file_xtract_prompt(XtractKind::Formulas),
            Action::FileXtractValues => self.start_file_xtract_prompt(XtractKind::Values),
            Action::FileImportNumbers => self.start_file_import_numbers_prompt(),
            Action::FileImportText => self.start_file_import_text_prompt(),
            Action::FileNew => self.execute_file_new(),
            Action::FileOpenBefore => self.start_file_open_prompt(true),
            Action::FileOpenAfter => self.start_file_open_prompt(false),
            Action::PrintFile => self.start_print_file_prompt(),
            Action::PrintPrinter => self.start_print_printer(),
            Action::PrintEncoded => self.start_print_encoded_prompt(),
            Action::PrintCancel => self.finish_print_session(),
            Action::PrintSessionRange => self.begin_point(PendingCommand::PrintFileRange),
            Action::PrintSessionGo => self.execute_print_go(),
            Action::PrintSessionQuit => self.finish_print_session(),
            Action::PrintSessionAlign => {
                if let Some(s) = self.print.as_mut() {
                    s.next_page = 1;
                }
                self.enter_print_file_menu();
            }
            Action::PrintSessionClear => {
                if let Some(s) = self.print.as_mut() {
                    s.clear_all();
                }
                self.enter_print_file_menu();
            }
            Action::PrintSessionOptionsHeader => self.start_print_header_prompt(),
            Action::PrintSessionOptionsFooter => self.start_print_footer_prompt(),
            Action::PrintSessionOptionsSetup => self.start_print_setup_prompt(),
            Action::PrintSessionOptionsQuit => self.enter_print_file_menu(),
            Action::PrintSessionOptionsOtherAsDisplayed => {
                self.set_print_content_mode(PrintContentMode::AsDisplayed)
            }
            Action::PrintSessionOptionsOtherCellFormulas => {
                self.set_print_content_mode(PrintContentMode::CellFormulas)
            }
            Action::PrintSessionOptionsOtherFormatted => {
                self.set_print_format_mode(PrintFormatMode::Formatted)
            }
            Action::PrintSessionOptionsOtherUnformatted => {
                self.set_print_format_mode(PrintFormatMode::Unformatted)
            }
            Action::PrintSessionOptionsMarginLeft => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginLeft, "left")
            }
            Action::PrintSessionOptionsMarginRight => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginRight, "right")
            }
            Action::PrintSessionOptionsMarginTop => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginTop, "top")
            }
            Action::PrintSessionOptionsMarginBottom => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginBottom, "bottom")
            }
            Action::PrintSessionOptionsMarginsQuit => self.enter_print_options_menu(),
            Action::PrintSessionOptionsPgLength => {
                self.start_print_pg_length_prompt();
            }
            Action::PrintSessionOptionsAdvancedDevice => self.start_print_advanced_device_prompt(),
            Action::PrintSessionOptionsAdvancedQuit => self.enter_print_options_menu(),
            Action::RangeSearchFormulas => self.begin_point(PendingCommand::RangeSearchRange {
                scope: SearchScope::Formulas,
            }),
            Action::RangeSearchLabels => self.begin_point(PendingCommand::RangeSearchRange {
                scope: SearchScope::Labels,
            }),
            Action::RangeSearchBoth => self.begin_point(PendingCommand::RangeSearchRange {
                scope: SearchScope::Both,
            }),
            Action::RangeSearchFind => self.execute_range_search_find(),
            Action::RangeSearchReplace => self.start_range_search_replace_prompt(),
            Action::FileDir => self.start_file_dir_prompt(),
            Action::FileListWorksheet => self.open_file_list(FileListKind::Worksheet),
            Action::FileListActive => self.open_file_list(FileListKind::Active),
            Action::FileListOther => self.open_file_list(FileListKind::Other),
            Action::System => {
                self.menu = None;
                self.pending_system_suspend = true;
                self.mode = Mode::Ready;
            }
            Action::GraphTypeLine => self.set_graph_type(GraphType::Line),
            Action::GraphTypeBar => self.set_graph_type(GraphType::Bar),
            Action::GraphTypeXY => self.set_graph_type(GraphType::XY),
            Action::GraphTypeStack => self.set_graph_type(GraphType::Stack),
            Action::GraphTypePie => self.set_graph_type(GraphType::Pie),
            Action::GraphTypeHLCO => self.set_graph_type(GraphType::HLCO),
            Action::GraphTypeMixed => self.set_graph_type(GraphType::Mixed),
            Action::GraphX => self.begin_point(PendingCommand::GraphSeries { series: Series::X }),
            Action::GraphA => self.begin_point(PendingCommand::GraphSeries { series: Series::A }),
            Action::GraphB => self.begin_point(PendingCommand::GraphSeries { series: Series::B }),
            Action::GraphC => self.begin_point(PendingCommand::GraphSeries { series: Series::C }),
            Action::GraphD => self.begin_point(PendingCommand::GraphSeries { series: Series::D }),
            Action::GraphE => self.begin_point(PendingCommand::GraphSeries { series: Series::E }),
            Action::GraphF => self.begin_point(PendingCommand::GraphSeries { series: Series::F }),
            Action::GraphResetGraph => {
                self.wb_mut().current_graph.reset();
                self.close_menu();
            }
            Action::GraphView => self.enter_graph_view(),
            Action::GraphSave => self.start_graph_save_prompt(),
            Action::GraphQuit => self.close_menu(),
            // Forward-declared in the menu enum but not yet implemented.
            // Hitting these from the menu currently is a no-op back to
            // READY; flesh out behavior when the feature lands.
            Action::FileEraseWorksheet
            | Action::FileErasePrint
            | Action::FileEraseGraph
            | Action::FileEraseOther => self.start_file_erase_prompt(),
            // `/File Admin` leaves are wired as named actions so future
            // implementation can hang behavior on them without further
            // menu surgery, but today they all just close the menu.
            Action::FileAdminReservationGet
            | Action::FileAdminReservationRelease
            | Action::FileAdminSealFile
            | Action::FileAdminSealReservationSetting
            | Action::FileAdminSealDisable
            | Action::FileAdminTableWorksheet
            | Action::FileAdminTablePrint
            | Action::FileAdminTableGraph
            | Action::FileAdminTableOther
            | Action::FileAdminTableActive
            | Action::FileAdminTableLinked
            | Action::FileAdminLinkRefresh => self.close_menu(),
            Action::FileCombineCopyEntire => {
                self.start_file_combine_prompt(CombineKind::Copy, true)
            }
            Action::FileCombineCopyNamed => {
                self.start_file_combine_prompt(CombineKind::Copy, false)
            }
            Action::FileCombineAddEntire => self.start_file_combine_prompt(CombineKind::Add, true),
            Action::FileCombineAddNamed => self.start_file_combine_prompt(CombineKind::Add, false),
            Action::FileCombineSubtractEntire => {
                self.start_file_combine_prompt(CombineKind::Subtract, true)
            }
            Action::FileCombineSubtractNamed => {
                self.start_file_combine_prompt(CombineKind::Subtract, false)
            }
            Action::DataFill => self.begin_point(PendingCommand::DataFillRange),
            Action::DataSortDataRange => self.begin_point(PendingCommand::DataSortDataRange),
            Action::DataSortPrimaryKey => {
                self.pending_sort_key_slot = Some(SortKeySlot::Primary);
                self.begin_point(PendingCommand::DataSortKey);
            }
            Action::DataSortSecondaryKey => {
                self.pending_sort_key_slot = Some(SortKeySlot::Secondary);
                self.begin_point(PendingCommand::DataSortKey);
            }
            Action::DataSortExtraKey => {
                self.pending_sort_key_slot = Some(SortKeySlot::Extra);
                self.begin_point(PendingCommand::DataSortKey);
            }
            Action::DataSortReset => {
                self.data_sort = DataSortState::default();
                self.pending_sort_key_slot = None;
                self.enter_data_sort_menu();
            }
            Action::DataSortGo => self.execute_data_sort(),
            Action::DataSortQuit => {
                self.menu = None;
                self.mode = Mode::Ready;
            }
            Action::DataSortAscending => self.bind_data_sort_dir(SortDir::Ascending),
            Action::DataSortDescending => self.bind_data_sort_dir(SortDir::Descending),
            Action::DataDistribution => self.begin_point(PendingCommand::DataDistributionValues),
            Action::DataRegressionXRange => self.begin_point(PendingCommand::DataRegressionXRange),
            Action::DataRegressionYRange => self.begin_point(PendingCommand::DataRegressionYRange),
            Action::DataRegressionOutputRange => {
                self.begin_point(PendingCommand::DataRegressionOutputRange)
            }
            Action::DataRegressionInterceptCompute => {
                self.data_regression.intercept_zero = false;
                self.enter_data_regression_menu();
            }
            Action::DataRegressionInterceptZero => {
                self.data_regression.intercept_zero = true;
                self.enter_data_regression_menu();
            }
            Action::DataRegressionReset => {
                self.data_regression = DataRegressionState::default();
                self.enter_data_regression_menu();
            }
            Action::DataRegressionGo => self.execute_data_regression(),
            Action::DataRegressionQuit => {
                self.menu = None;
                self.mode = Mode::Ready;
            }
            Action::DataMatrixInvert => self.begin_point(PendingCommand::DataMatrixInvertInput),
            Action::DataMatrixMultiply => self.begin_point(PendingCommand::DataMatrixMultiplyA),
            Action::DataParseInputColumn => self.begin_point(PendingCommand::DataParseInputColumn),
            Action::DataParseOutputRange => self.begin_point(PendingCommand::DataParseOutputRange),
            Action::DataParseReset => {
                self.data_parse = DataParseState::default();
                self.enter_data_parse_menu();
            }
            Action::DataParseGo => self.execute_data_parse(),
            Action::DataParseQuit => {
                self.menu = None;
                self.mode = Mode::Ready;
            }
            Action::DataTable1 => self.begin_point(PendingCommand::DataTable1Range),
            Action::DataTable2 => self.begin_point(PendingCommand::DataTable2Range),
            Action::DataTableReset => {
                self.menu = None;
                self.mode = Mode::Ready;
            }
            Action::DataParseFormatLineCreate => self.execute_parse_format_line_create(),
            Action::DataParseFormatLineEdit => self.execute_parse_format_line_edit(),
            Action::DataQueryInput => self.begin_point(PendingCommand::DataQueryInput),
            Action::DataQueryCriteria => self.begin_point(PendingCommand::DataQueryCriteria),
            Action::DataQueryOutput => self.begin_point(PendingCommand::DataQueryOutput),
            Action::DataQueryFind => self.execute_data_query_find(),
            Action::DataQueryExtract => self.execute_data_query_extract(false),
            Action::DataQueryUnique => self.execute_data_query_extract(true),
            Action::DataQueryDel => self.execute_data_query_del(),
            Action::DataQueryReset => {
                self.data_query = DataQueryState::default();
                self.enter_data_query_menu();
            }
            Action::DataQueryQuit => {
                self.menu = None;
                self.mode = Mode::Ready;
            }
            Action::DataTable3Stub => {
                self.set_error("Data Table 3 (3D table): not yet implemented in L123")
            }
            Action::DataTableLabeledStub => {
                self.set_error("Data Table Labeled: not yet implemented in L123")
            }
            Action::DataQueryModifyStub => {
                self.set_error("Data Query Modify: not yet implemented in L123")
            }
            Action::DataExternalStub => {
                self.set_error("Data External: no external-database driver configured")
            }
        }
    }

    fn start_file_save_prompt(&mut self) {
        self.menu = None;
        let (buffer, fresh) = match &self.wb_mut().active_path {
            Some(p) => (p.to_string_lossy().into_owned(), true),
            None => (String::new(), false),
        };
        self.prompt = Some(PromptState {
            label: "Enter save file name:".into(),
            buffer,
            next: PromptNext::FileSaveFilename,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    /// /FL — populate `file_list` with the requested set of files and
    /// enter FILES mode.
    fn open_file_list(&mut self, kind: FileListKind) {
        self.menu = None;
        let entries = match kind {
            FileListKind::Worksheet => {
                let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
                list_worksheet_files_in(&cwd)
            }
            FileListKind::Active => self.wb_mut().active_path.iter().cloned().collect(),
            FileListKind::Other => {
                let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
                list_all_files_in(&cwd)
            }
        };
        self.file_list = Some(FileListState {
            kind,
            entries,
            highlight: 0,
            view_offset: 0,
        });
        self.mode = Mode::Files;
    }

    fn open_help(&mut self) {
        // Don't double-open if already in HELP (defensive — the
        // dispatcher gate above should have routed F1 to
        // handle_key_help instead).
        if self.help.is_some() {
            return;
        }
        let return_mode = self.mode;
        let Some(state) = HelpState::open(return_mode) else {
            return;
        };
        self.help = Some(state);
        self.mode = Mode::Help;
    }

    fn close_help(&mut self) {
        let Some(state) = self.help.take() else {
            return;
        };
        self.mode = state.return_mode;
    }

    /// Esc from the name list: clear the overlay and restore the
    /// underlying mode (POINT or the prompt's MENU mode).
    fn dismiss_name_list(&mut self) {
        let Some(nl) = self.name_list.take() else {
            return;
        };
        self.mode = match nl.origin {
            NameListOrigin::Point => Mode::Point,
            NameListOrigin::Goto | NameListOrigin::PromptName => Mode::Menu,
            NameListOrigin::RunMacro => Mode::Ready,
        };
    }

    /// Enter on the name list: dispatch per origin.
    fn commit_name_list(&mut self) {
        let Some(nl) = self.name_list.take() else {
            return;
        };
        // Empty list — nothing to commit; fall back to dismiss.
        let Some((name, range)) = nl.entries.get(nl.highlight).cloned() else {
            self.mode = match nl.origin {
                NameListOrigin::Point => Mode::Point,
                NameListOrigin::Goto | NameListOrigin::PromptName => Mode::Menu,
                NameListOrigin::RunMacro => Mode::Ready,
            };
            return;
        };
        match nl.origin {
            NameListOrigin::Point => {
                let Some(ps) = self.point.take() else {
                    self.mode = Mode::Ready;
                    return;
                };
                self.apply_pending_with_ranges(ps.pending, &[range]);
            }
            NameListOrigin::Goto => {
                self.prompt = None;
                self.move_pointer_to(range.start);
                self.mode = Mode::Ready;
            }
            NameListOrigin::PromptName => {
                if let Some(p) = self.prompt.as_mut() {
                    p.buffer = name;
                    p.fresh = false;
                }
                self.mode = Mode::Menu;
            }
            NameListOrigin::RunMacro => {
                self.mode = Mode::Ready;
                self.run_named_macro(&name);
            }
        }
    }

    /// /FN — wipe the current workbook back to a blank slate. Both the
    /// `/Worksheet Erase Yes` — drop every active file and replace the
    /// workspace with a single blank workbook. Session-level prompts,
    /// menus, and modal overlays are also cleared so the user lands in
    /// a predictable READY state on A:A1.
    fn set_titles(&mut self, kind: TitlesKind) {
        let sheet = self.wb().pointer.sheet;
        let row = self.wb().pointer.row;
        let col = self.wb().pointer.col;
        let new = match kind {
            TitlesKind::Both => (row, col),
            TitlesKind::Horizontal => (row, 0),
            TitlesKind::Vertical => (0, col),
        };
        let prev = self.wb().frozen.get(&sheet).copied();
        self.wb_mut().frozen.insert(sheet, new);
        self.push_journal_batch(vec![JournalEntry::Frozen { sheet, prev }]);
        self.wb_mut().dirty = true;
        self.close_menu();
    }

    fn clear_titles(&mut self) {
        let sheet = self.wb().pointer.sheet;
        let prev = self.wb().frozen.get(&sheet).copied();
        if prev.is_some() {
            self.wb_mut().frozen.remove(&sheet);
            self.push_journal_batch(vec![JournalEntry::Frozen { sheet, prev }]);
            self.wb_mut().dirty = true;
        }
        self.close_menu();
    }

    fn insert_page_break_row_at_pointer(&mut self) {
        let sheet = self.wb().pointer.sheet;
        let at = self.wb().pointer.row;
        self.menu = None;
        let mut batch: Vec<JournalEntry> = Vec::new();
        if self.wb_mut().engine.insert_rows(sheet, at, 1).is_ok() {
            shift_cells_rows(&mut self.wb_mut().cells, sheet, at, 1);
            batch.push(JournalEntry::RowInsert { sheet, at });
        }
        let marker = Address::new(sheet, 0, at);
        let prev_contents = self.wb_mut().cells.remove(&marker);
        let prev_format = self.wb_mut().cell_formats.remove(&marker);
        let label = CellContents::Label {
            prefix: LabelPrefix::Pipe,
            text: "::".into(),
        };
        self.push_to_engine_at(marker, &label);
        self.wb_mut().cells.insert(marker, label);
        batch.push(JournalEntry::CellEdit {
            addr: marker,
            prev_contents,
            prev_format,
        });
        self.push_journal_batch(batch);
        self.mode = Mode::Ready;
    }

    fn insert_page_break_column_at_pointer(&mut self) {
        let sheet = self.wb().pointer.sheet;
        let at = self.wb().pointer.col;
        self.menu = None;
        let mut batch: Vec<JournalEntry> = Vec::new();
        if self.wb_mut().engine.insert_cols(sheet, at, 1).is_ok() {
            shift_cells_cols(&mut self.wb_mut().cells, sheet, at, 1);
            batch.push(JournalEntry::ColInsert { sheet, at });
        }
        let marker = Address::new(sheet, at, 0);
        let prev_contents = self.wb_mut().cells.remove(&marker);
        let prev_format = self.wb_mut().cell_formats.remove(&marker);
        let label = CellContents::Label {
            prefix: LabelPrefix::Pipe,
            text: "::".into(),
        };
        self.push_to_engine_at(marker, &label);
        self.wb_mut().cells.insert(marker, label);
        batch.push(JournalEntry::CellEdit {
            addr: marker,
            prev_contents,
            prev_format,
        });
        self.push_journal_batch(batch);
        self.mode = Mode::Ready;
    }

    fn hide_current_sheet(&mut self) {
        let sheet = self.wb().pointer.sheet;
        let count = self.wb().engine.sheet_count();
        let visible_other = (0..count).any(|i| {
            let sid = SheetId(i);
            sid != sheet
                && self
                    .wb()
                    .sheet_states
                    .get(&sid)
                    .copied()
                    .unwrap_or(SheetState::Visible)
                    .is_visible()
        });
        if !visible_other {
            self.menu = None;
            self.set_error("Cannot hide the only visible sheet");
            return;
        }
        let prev = self
            .wb()
            .sheet_states
            .get(&sheet)
            .copied()
            .unwrap_or(SheetState::Visible);
        self.wb_mut().sheet_states.insert(sheet, SheetState::Hidden);
        self.push_journal_batch(vec![JournalEntry::SheetVisibility { sheet, prev }]);
        redirect_pointer_off_hidden(self.wb_mut());
        self.close_menu();
    }

    fn unhide_all_sheets(&mut self) {
        let count = self.wb().engine.sheet_count();
        let mut batch: Vec<JournalEntry> = Vec::new();
        for i in 0..count {
            let sid = SheetId(i);
            let prev = self
                .wb()
                .sheet_states
                .get(&sid)
                .copied()
                .unwrap_or(SheetState::Visible);
            if prev != SheetState::Visible {
                self.wb_mut().sheet_states.remove(&sid);
                batch.push(JournalEntry::SheetVisibility { sheet: sid, prev });
            }
        }
        if !batch.is_empty() {
            self.push_journal_batch(batch);
        }
        self.close_menu();
    }

    fn execute_worksheet_erase(&mut self) {
        self.entry = None;
        self.menu = None;
        self.prompt = None;
        self.point = None;
        self.save_confirm = None;
        self.erase_confirm = None;
        self.pending_name = None;
        self.pending_xtract_path = None;
        self.pending_combine_path = None;
        self.file_list = None;
        self.active_files = vec![Workbook::new()];
        self.current = 0;
        self.recalc_pending = false;
        self.mode = Mode::Ready;
    }

    /// Before and After branches collapse to this same reset for now;
    /// true multi-file insertion is M5.
    fn execute_file_new(&mut self) {
        self.wb_mut().cells.clear();
        self.wb_mut().clear_all_cell_formats();
        self.wb_mut().cell_text_styles.clear();
        self.wb_mut().cell_alignments.clear();
        self.wb_mut().cell_fills.clear();
        self.wb_mut().cell_font_styles.clear();
        self.wb_mut().cell_borders.clear();
        self.wb_mut().comments.clear();
        self.wb_mut().merges.clear();
        self.wb_mut().frozen.clear();
        self.wb_mut().sheet_states.clear();
        self.wb_mut().tables.clear();
        self.wb_mut().sheet_colors.clear();
        self.wb_mut().col_widths.clear();
        self.wb_mut().default_col_width = 9;
        self.wb_mut().hidden_cols.clear();
        self.wb_mut().named_ranges.clear();
        self.wb_mut().name_notes.clear();
        self.entry = None;
        self.menu = None;
        self.prompt = None;
        self.point = None;
        self.save_confirm = None;
        self.erase_confirm = None;
        self.pending_name = None;
        self.pending_xtract_path = None;
        self.pending_combine_path = None;
        self.wb_mut().active_path = None;
        self.wb_mut().pointer = Address::A1;
        self.wb_mut().viewport_col_offset = 0;
        self.wb_mut().viewport_row_offset = 0;
        self.recalc_pending = false;
        if let Ok(engine) = IronCalcEngine::new() {
            self.wb_mut().engine = engine;
        }
        self.mode = Mode::Ready;
    }

    fn start_print_file_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter print file name:".into(),
            buffer: String::new(),
            next: PromptNext::PrintFileFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// `/Print Printer`: open a printer session straight away — no path
    /// prompt. Shares the submenu with `/Print File`; the Go branch
    /// sends output through CUPS `lp` instead of writing a file.
    fn start_print_printer(&mut self) {
        self.print = Some(PrintSession::new_printer());
        self.enter_print_file_menu();
    }

    /// `/Print Encoded`: prompt for the destination path. On commit a
    /// session is opened with [`PrintDestination::Encoded`] and the
    /// shared `/PF` submenu is entered.
    fn start_print_encoded_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter encoded file name:".into(),
            buffer: String::new(),
            next: PromptNext::PrintEncodedFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn enter_print_file_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::PRINT_FILE_MENU));
        self.mode = Mode::Menu;
    }

    /// Re-enter the Options sub-sub-menu at path=['O'] under the
    /// PRINT_FILE_MENU root. Used after an Options-level prompt
    /// commits so the user stays in Options for further tweaks.
    fn enter_print_options_menu(&mut self) {
        self.menu = Some(MenuState {
            path: vec!['O'],
            highlight: 0,
            message: None,
            override_root: Some(menu::PRINT_FILE_MENU),
        });
        self.mode = Mode::Menu;
    }

    fn set_print_content_mode(&mut self, mode: PrintContentMode) {
        if let Some(s) = self.print.as_mut() {
            s.content_mode = mode;
        }
        // After the setting commits, return to the Options submenu so
        // the user can pick another option or Quit.
        self.enter_print_options_menu();
    }

    fn set_print_format_mode(&mut self, mode: PrintFormatMode) {
        if let Some(s) = self.print.as_mut() {
            s.format_mode = mode;
        }
        self.enter_print_options_menu();
    }

    /// Re-enter the Margins sub-sub-menu at path=['O', 'M'] under
    /// the /PF root, matching the flow where each margin prompt
    /// committing drops the user back into Margins.
    fn enter_print_margins_menu(&mut self) {
        self.menu = Some(MenuState {
            path: vec!['O', 'M'],
            highlight: 0,
            message: None,
            override_root: Some(menu::PRINT_FILE_MENU),
        });
        self.mode = Mode::Menu;
    }

    /// Re-enter the Advanced sub-sub-menu at path=['O', 'A'] under
    /// the /PF root so each Advanced leaf returns to its sibling list
    /// (mirrors `enter_print_margins_menu`).
    fn enter_print_advanced_menu(&mut self) {
        self.menu = Some(MenuState {
            path: vec!['O', 'A'],
            highlight: 0,
            message: None,
            override_root: Some(menu::PRINT_FILE_MENU),
        });
        self.mode = Mode::Menu;
    }

    fn start_print_advanced_device_prompt(&mut self) {
        self.menu = None;
        let buffer = self
            .print
            .as_ref()
            .map(|s| s.lp_destination.clone())
            .unwrap_or_default();
        let fresh = !buffer.is_empty();
        self.prompt = Some(PromptState {
            label: "Enter printer name:".into(),
            buffer,
            next: PromptNext::PrintSessionOptionsAdvancedDevice,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn start_print_margin_prompt(&mut self, next: PromptNext, which: &str) {
        self.menu = None;
        let current: u16 = match (next, self.print.as_ref()) {
            (PromptNext::PrintFileMarginLeft, Some(s)) => s.margin_left,
            (PromptNext::PrintFileMarginRight, Some(s)) => s.margin_right,
            (PromptNext::PrintFileMarginTop, Some(s)) => s.margin_top,
            (PromptNext::PrintFileMarginBottom, Some(s)) => s.margin_bottom,
            _ => 0,
        };
        self.prompt = Some(PromptState {
            label: format!("Enter {which} margin (0..1000):"),
            buffer: current.to_string(),
            next,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    fn start_print_pg_length_prompt(&mut self) {
        self.menu = None;
        let current: u16 = self.print.as_ref().map(|s| s.pg_length).unwrap_or(0);
        self.prompt = Some(PromptState {
            label: "Enter page length (0 = no pagination, 1..1000):".into(),
            buffer: current.to_string(),
            next: PromptNext::PrintFilePgLength,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    fn start_print_header_prompt(&mut self) {
        self.menu = None;
        let buffer = self
            .print
            .as_ref()
            .map(|s| s.header.clone())
            .unwrap_or_default();
        let fresh = !buffer.is_empty();
        self.prompt = Some(PromptState {
            label: "Enter print header (L|C|R):".into(),
            buffer,
            next: PromptNext::PrintFileHeader,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn start_print_footer_prompt(&mut self) {
        self.menu = None;
        let buffer = self
            .print
            .as_ref()
            .map(|s| s.footer.clone())
            .unwrap_or_default();
        let fresh = !buffer.is_empty();
        self.prompt = Some(PromptState {
            label: "Enter print footer (L|C|R):".into(),
            buffer,
            next: PromptNext::PrintFileFooter,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn start_print_setup_prompt(&mut self) {
        self.menu = None;
        let buffer = self
            .print
            .as_ref()
            .map(|s| s.setup_string.clone())
            .unwrap_or_default();
        let fresh = !buffer.is_empty();
        self.prompt = Some(PromptState {
            label: "Enter setup string:".into(),
            buffer,
            next: PromptNext::PrintFileSetup,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn finish_print_session(&mut self) {
        self.print = None;
        self.close_menu();
    }

    fn execute_print_go(&mut self) {
        let Some(session) = self.print.as_ref() else {
            self.close_menu();
            return;
        };
        if session.ranges.is_empty() {
            // No range selected — bounce back to the menu without
            // writing anything. Matches 1-2-3's "Go with no range =
            // no-op".
            self.enter_print_file_menu();
            return;
        }
        // Render each range with running page numbers so multi-range
        // jobs (`A1..B2,C3..D4` typed in POINT) stay paginated as one
        // logical document. Each part contributes its own pages to the
        // merged grid; `next_page` advances by the total page count.
        let base_settings = PrintSettings {
            header: session.header.clone(),
            footer: session.footer.clone(),
            content_mode: session.content_mode,
            format_mode: session.format_mode,
            margin_left: session.margin_left,
            margin_right: session.margin_right,
            margin_top: session.margin_top,
            margin_bottom: session.margin_bottom,
            pg_length: session.pg_length,
            start_page: session.next_page,
        };
        let mut grid_pages: Vec<l123_print::grid::Page> = Vec::new();
        let mut page_width: u16 = 0;
        let mut start_page = session.next_page;
        for r in &session.ranges {
            let settings = PrintSettings {
                start_page,
                ..base_settings.clone()
            };
            let g = l123_print::render(self.wb(), *r, &settings);
            page_width = page_width.max(g.page_width);
            start_page += g.pages.len() as u32;
            grid_pages.extend(g.pages);
        }
        let grid = l123_print::grid::PageGrid {
            pages: grid_pages,
            page_width: page_width.max(1),
        };
        let pages = grid.pages.len() as u32;
        match &session.destination {
            PrintDestination::File(path) => {
                if let Some(parent) = path.parent() {
                    if !parent.as_os_str().is_empty() {
                        let _ = std::fs::create_dir_all(parent);
                    }
                }
                // `.pdf` extension (case-insensitive) → PDF encoding.
                // Any other extension — or no extension — gets the
                // classic .prn ASCII stream.
                let is_pdf = path
                    .extension()
                    .and_then(|e| e.to_str())
                    .is_some_and(|e| e.eq_ignore_ascii_case("pdf"));
                let bytes: Vec<u8> = if is_pdf {
                    l123_print::encode::pdf::to_pdf(
                        &grid,
                        &l123_print::encode::pdf::PdfOptions::default(),
                    )
                } else {
                    let mut out = session.setup_string.as_bytes().to_vec();
                    out.extend_from_slice(l123_print::to_ascii(&grid).as_bytes());
                    out
                };
                let _ = std::fs::write(path, bytes);
            }
            PrintDestination::Printer(lp_opts) => {
                #[cfg(unix)]
                {
                    let mut effective = lp_opts.clone();
                    if !session.setup_string.is_empty() {
                        effective.setup_string = Some(session.setup_string.clone());
                    }
                    if !session.lp_destination.is_empty() {
                        effective.destination = Some(session.lp_destination.clone());
                    }
                    let _ = l123_print::encode::lp::to_lp(&grid, &effective);
                }
                #[cfg(not(unix))]
                {
                    let _ = lp_opts; // hold field live on non-unix
                }
            }
            PrintDestination::Encoded(path) => {
                if let Some(parent) = path.parent() {
                    if !parent.as_os_str().is_empty() {
                        let _ = std::fs::create_dir_all(parent);
                    }
                }
                let mut out = session.setup_string.as_bytes().to_vec();
                out.extend_from_slice(l123_print::to_ascii(&grid).as_bytes());
                let _ = std::fs::write(path, out);
            }
        }
        // Session stays alive so the user can issue further commands
        // (Options, another Go, Align, Clear, …). Quit is the way
        // out. Advance the page counter for the next Go.
        if let Some(s) = self.print.as_mut() {
            s.next_page = s.next_page.saturating_add(pages);
        }
        self.enter_print_file_menu();
    }

    fn start_file_open_prompt(&mut self, before: bool) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter file to open:".into(),
            buffer: String::new(),
            next: PromptNext::FileOpenFilename { before },
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// Load an xlsx from `path` as an additional active file.
    /// `before = true` inserts it immediately before the current slot
    /// and makes it the foreground file. `before = false` appends it
    /// after the current slot without disturbing the current view.
    fn open_file_alongside(&mut self, path: PathBuf, before: bool) {
        let Ok(mut engine) = IronCalcEngine::new() else {
            self.mode = Mode::Ready;
            return;
        };
        if engine.load_xlsx(&path).is_err() {
            self.mode = Mode::Ready;
            return;
        }
        // Pre-populate the new file's cells cache from the engine.
        let mut cells = HashMap::new();
        let sheet_names = engine.all_sheet_names();
        let sheet_refs: Vec<&str> = sheet_names.iter().map(String::as_str).collect();
        for (addr, cv) in engine.used_cells() {
            if let Some(contents) = cell_view_to_contents(&cv, &sheet_refs) {
                cells.insert(addr, contents);
            }
        }
        // Apply the formula-source sidecar if present, overriding
        // the cosmetic reverse-translated `expr` for any cell that
        // has a stored Lotus source.
        if let Ok(sources) = l123_io::formula_sources::read_from_xlsx(&path) {
            for (addr, src) in sources {
                if let Some(CellContents::Formula { expr, .. }) = cells.get_mut(&addr) {
                    *expr = src;
                }
            }
        }
        let mut col_widths: HashMap<(SheetId, u16), u8> = HashMap::new();
        for (addr, w) in engine.used_column_widths() {
            col_widths.insert((addr.sheet, addr.col), w);
        }
        let mut cell_text_styles: HashMap<Address, TextStyle> = HashMap::new();
        for (addr, style) in engine.used_cell_text_styles() {
            cell_text_styles.insert(addr, style);
        }
        let mut cell_formats: HashMap<Address, Format> = HashMap::new();
        for (addr, fmt) in engine.used_cell_formats() {
            cell_formats.insert(addr, fmt);
        }
        let mut cell_format_overrides: HashMap<Address, String> = HashMap::new();
        for (addr, raw) in engine.used_cell_format_strings() {
            cell_format_overrides.insert(addr, raw);
        }
        // Layer the cell-format-extras sidecar (kind override for
        // non-Excel kinds, parens flag, negative-color) on top of
        // whatever the engine pulled from `num_fmt`. See
        // `l123_io::cell_formats` for the on-disk shape.
        let format_extras = l123_io::cell_formats::read_from_xlsx(&path).unwrap_or_default();
        for (addr, fe) in &format_extras.cells {
            let base = cell_formats.get(addr).copied().unwrap_or(Format::GENERAL);
            cell_formats.insert(*addr, fe.apply_to(base));
        }
        let mut cell_alignments: HashMap<Address, Alignment> = HashMap::new();
        for (addr, a) in engine.used_cell_alignments() {
            cell_alignments.insert(addr, a);
        }
        let mut cell_fills: HashMap<Address, Fill> = HashMap::new();
        for (addr, f) in engine.used_cell_fills() {
            cell_fills.insert(addr, f);
        }
        let mut cell_font_styles: HashMap<Address, FontStyle> = HashMap::new();
        for (addr, fs) in engine.used_cell_font_styles() {
            cell_font_styles.insert(addr, fs);
        }
        let mut cell_borders: HashMap<Address, Border> = HashMap::new();
        for (addr, b) in engine.used_cell_borders() {
            cell_borders.insert(addr, b);
        }
        let mut comments: HashMap<Address, Comment> = HashMap::new();
        for c in engine.used_comments() {
            comments.insert(c.addr, c);
        }
        let mut merges: HashMap<SheetId, Vec<Merge>> = HashMap::new();
        for (sheet, m) in engine.used_merged_cells() {
            merges.entry(sheet).or_default().push(m);
        }
        let mut frozen: HashMap<SheetId, (u32, u16)> = HashMap::new();
        for sheet_idx in 0..engine.sheet_count() {
            let sid = SheetId(sheet_idx);
            let f = engine.frozen_panes(sid);
            if f != (0, 0) {
                frozen.insert(sid, f);
            }
        }
        let mut sheet_states: HashMap<SheetId, SheetState> = HashMap::new();
        for sheet_idx in 0..engine.sheet_count() {
            let sid = SheetId(sheet_idx);
            let st = engine.sheet_state(sid);
            if st != SheetState::Visible {
                sheet_states.insert(sid, st);
            }
        }
        let mut tables: HashMap<SheetId, Vec<Table>> = HashMap::new();
        for (sheet, t) in engine.used_tables() {
            tables.entry(sheet).or_default().push(t);
        }
        let mut sheet_colors: HashMap<SheetId, RgbColor> = HashMap::new();
        for sheet_idx in 0..engine.sheet_count() {
            let sid = SheetId(sheet_idx);
            if let Some(c) = engine.sheet_color(sid) {
                sheet_colors.insert(sid, c);
            }
        }
        let new_file = Workbook {
            engine,
            cells,
            cell_formats,
            cell_format_overrides,
            global_format: format_extras
                .global
                .map(|fe| fe.apply_to(Format::GENERAL))
                .unwrap_or(Format::GENERAL),
            international: International::default(),
            cell_text_styles,
            cell_alignments,
            cell_fills,
            cell_font_styles,
            cell_borders,
            comments,
            merges,
            frozen,
            sheet_states,
            tables,
            sheet_colors,
            col_widths,
            default_col_width: 9,
            hidden_cols: HashSet::new(),
            active_path: Some(path),
            dirty: false,
            pointer: Address::A1,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            journal: Vec::new(),
            current_graph: GraphDef::default(),
            graphs: BTreeMap::new(),
            named_ranges: HashMap::new(),
            name_notes: HashMap::new(),
            cell_unprotected: HashSet::new(),
        };
        // If the active sheet is hidden / very-hidden, redirect to the
        // first visible sheet so the user lands somewhere they can
        // interact with.  Works on the freshly-built `new_file`
        // before it's inserted into the active-files list.
        let mut new_file = new_file;
        redirect_pointer_off_hidden(&mut new_file);
        if before {
            self.active_files.insert(self.current, new_file);
            // `current` still points to the old (now shifted) file;
            // Before convention is that the new file takes focus, so
            // move focus to the just-inserted slot.
            // After insert at `current`, old file is now at current+1
            // and new file is at current. Keep current as-is.
        } else {
            self.active_files.insert(self.current + 1, new_file);
        }
        self.mode = Mode::Ready;
    }

    /// Rotate the foreground file by `delta` slots. +1 = Ctrl-PgDn
    /// (next file); -1 = Ctrl-PgUp (prev file). No-op if only one
    /// file is active. Clears the Ctrl-End prefix.
    fn rotate_files(&mut self, delta: i32) {
        self.file_nav_pending = false;
        let n = self.active_files.len();
        if n <= 1 || delta == 0 {
            return;
        }
        let len = n as i32;
        let mut next = self.current as i32 + delta;
        next = ((next % len) + len) % len;
        self.current = next as usize;
    }

    fn start_file_dir_prompt(&mut self) {
        self.menu = None;
        let buffer = std::env::current_dir()
            .ok()
            .map(|p| p.to_string_lossy().into_owned())
            .unwrap_or_default();
        let fresh = !buffer.is_empty();
        self.prompt = Some(PromptState {
            label: "Enter new session directory:".into(),
            buffer,
            next: PromptNext::FileDirPath,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    fn start_file_import_numbers_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter import file name:".into(),
            buffer: String::new(),
            next: PromptNext::FileImportNumbersFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn start_file_import_text_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter import file name:".into(),
            buffer: String::new(),
            next: PromptNext::FileImportTextFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// `/File Combine` — open the source-filename prompt.  `entire`
    /// distinguishes the Entire-File branch (commit applies the merge)
    /// from Named-Or-Specified-Range (commit chains into a second
    /// prompt for the source range).
    fn start_file_combine_prompt(&mut self, kind: CombineKind, entire: bool) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter source file name:".into(),
            buffer: String::new(),
            next: PromptNext::FileCombineFilename { kind, entire },
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// Second step of `/File Combine … Named/Specified-Range`. The
    /// filename was stashed in `pending_combine_path` by the prior
    /// prompt commit; this one collects the source range string.
    fn start_file_combine_range_prompt(&mut self, kind: CombineKind) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter source range:".into(),
            buffer: String::new(),
            next: PromptNext::FileCombineRange { kind },
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// Read the source xlsx at `path` into a temporary engine, then
    /// merge non-empty source cells into the active workbook starting
    /// at the pointer.  When `source_range` is `None`, every non-empty
    /// cell on the source's first sheet is considered; otherwise only
    /// cells inside the range.  Failures (bad path, bad range) surface
    /// on line 3 via the standard error path.
    fn combine_from(&mut self, path: PathBuf, kind: CombineKind, source_range: Option<Range>) {
        let mut src = match IronCalcEngine::new() {
            Ok(e) => e,
            Err(e) => {
                self.set_error(format!("Combine: engine init failed: {e}"));
                return;
            }
        };
        if let Err(e) = src.load_xlsx(&path) {
            self.set_error(format!("Cannot open {}: {e}", path.display()));
            return;
        }
        src.recalc();
        let origin = self.wb_mut().pointer;
        let target_sheet = origin.sheet;
        let source_sheet = SheetId(0);
        // Determine the cells to scan on the source.  Without a typed
        // range we scan a generous window — the active first sheet up
        // to (256, 8192) — matching 1-2-3 R3's sheet bounds well enough
        // that any realistic Combine source fits.
        let (rmin, rmax, cmin, cmax) = match source_range {
            Some(r) => {
                let n = r.normalized();
                (n.start.row, n.end.row, n.start.col, n.end.col)
            }
            None => (0u32, 8191u32, 0u16, 255u16),
        };
        for sr in rmin..=rmax {
            for sc in cmin..=cmax {
                let saddr = Address::new(source_sheet, sc, sr);
                let Ok(cv) = src.get_cell(saddr) else {
                    continue;
                };
                if cv.value == Value::Empty && cv.formula.is_none() {
                    continue;
                }
                let taddr = Address::new(target_sheet, origin.col + sc, origin.row + sr);
                self.combine_apply(taddr, &cv, kind);
            }
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.mode = Mode::Ready;
    }

    /// Apply one source cell onto one target cell per `kind`.  Copy
    /// overwrites; Add/Subtract numerically combine, leaving
    /// non-numeric source or target cells untouched.
    fn combine_apply(&mut self, taddr: Address, src_cv: &CellView, kind: CombineKind) {
        match kind {
            CombineKind::Copy => {
                let input = match (&src_cv.formula, &src_cv.value) {
                    (Some(f), _) => format!("={f}"),
                    (None, Value::Number(n)) => l123_core::format_number_general(*n),
                    (None, Value::Text(s)) => format!("'{s}"),
                    (None, Value::Bool(b)) => {
                        if *b {
                            "TRUE".into()
                        } else {
                            "FALSE".into()
                        }
                    }
                    _ => return,
                };
                let _ = self.wb_mut().engine.set_user_input(taddr, &input);
                let contents = match &src_cv.value {
                    Value::Number(n) => CellContents::Constant(Value::Number(*n)),
                    Value::Text(s) => CellContents::Label {
                        prefix: LabelPrefix::Apostrophe,
                        text: s.clone(),
                    },
                    Value::Bool(b) => CellContents::Constant(Value::Bool(*b)),
                    _ => return,
                };
                self.wb_mut().cells.insert(taddr, contents);
            }
            CombineKind::Add | CombineKind::Subtract => {
                let Value::Number(src_n) = src_cv.value else {
                    return;
                };
                let target_now = self
                    .wb_mut()
                    .engine
                    .get_cell(taddr)
                    .ok()
                    .map(|cv| cv.value)
                    .unwrap_or(Value::Empty);
                let base = match target_now {
                    Value::Number(n) => n,
                    Value::Empty => 0.0,
                    _ => return,
                };
                let merged = match kind {
                    CombineKind::Add => base + src_n,
                    CombineKind::Subtract => base - src_n,
                    CombineKind::Copy => unreachable!(),
                };
                let input = l123_core::format_number_general(merged);
                let _ = self.wb_mut().engine.set_user_input(taddr, &input);
                self.wb_mut()
                    .cells
                    .insert(taddr, CellContents::Constant(Value::Number(merged)));
            }
        }
    }

    /// `/File Erase {Worksheet|Print|Graph|Other}` — prompt for the
    /// path to delete.  All four leaves share the same flow today; the
    /// kind would only change the directory listing filter, which we
    /// don't implement here.
    fn start_file_erase_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter file to erase:".into(),
            buffer: String::new(),
            next: PromptNext::FileEraseFilename,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn start_file_xtract_prompt(&mut self, kind: XtractKind) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter extract file name:".into(),
            buffer: String::new(),
            next: PromptNext::FileXtractFilename { kind },
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn start_file_retrieve_prompt(&mut self) {
        self.menu = None;
        let (buffer, fresh) = match &self.wb_mut().active_path {
            Some(p) => (p.to_string_lossy().into_owned(), true),
            None => (String::new(), false),
        };
        self.prompt = Some(PromptState {
            label: "Enter file to retrieve:".into(),
            buffer,
            next: PromptNext::FileRetrieveFilename,
            fresh,
        });
        self.mode = Mode::Menu;
    }

    /// Read `path` as CSV, wipe the in-memory workbook, and paint the
    /// parsed rows starting at A1 — the "retrieve" counterpart to
    /// `/File Import Numbers`. Fails closed on a read error so the
    /// current workbook survives a bad path.
    fn load_csv_workbook_from(&mut self, path: PathBuf) {
        let body = match std::fs::read_to_string(&path) {
            Ok(b) => b,
            Err(e) => {
                self.set_error(format!("Cannot read {}: {e}", path.display()));
                return;
            }
        };
        let rows = l123_io::csv::parse(&body);
        self.execute_file_new();
        let sheet = self.wb().pointer.sheet;
        for (dr, row) in rows.iter().enumerate() {
            for (dc, field) in row.iter().enumerate() {
                if field.is_empty() {
                    continue;
                }
                let addr = Address::new(sheet, dc as u16, dr as u32);
                let (contents, engine_input) = match field.parse::<f64>() {
                    Ok(n) => (
                        CellContents::Constant(Value::Number(n)),
                        l123_core::format_number_general(n),
                    ),
                    Err(_) => (
                        CellContents::Label {
                            prefix: LabelPrefix::Apostrophe,
                            text: field.clone(),
                        },
                        format!("'{field}"),
                    ),
                };
                let _ = self.wb_mut().engine.set_user_input(addr, &engine_input);
                self.wb_mut().cells.insert(addr, contents);
            }
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().active_path = Some(path);
        self.wb_mut().dirty = false;
        self.mode = Mode::Ready;
    }

    /// Load an xlsx from disk, wiping the current in-memory workbook
    /// and repopulating the UI cache from the loaded engine model.
    fn load_workbook_from(&mut self, path: PathBuf) {
        let is_wk3 = is_wk3_path(&path);
        let load_result = if is_wk3 {
            #[cfg(feature = "wk3")]
            {
                self.wb_mut().engine.load_wk3(&path)
            }
            #[cfg(not(feature = "wk3"))]
            unreachable!()
        } else {
            self.wb_mut().engine.load_xlsx(&path)
        };
        if let Err(e) = load_result {
            self.set_error(format!("Cannot open {}: {e}", path.display()));
            return;
        }
        self.repopulate_after_xlsx_load(path, is_wk3);
    }

    /// Post-engine-load workbook rebuild: wipe UI caches, pull cells
    /// and styles back out of the engine, apply the formula-source
    /// sidecar, and pin `active_path`. Shared by the sync CLI-startup
    /// path (`load_workbook_from`) and the §4.7 async `/File Retrieve`
    /// completion path (`apply_async_result::FileRetrieveXlsx`).
    fn repopulate_after_xlsx_load(&mut self, path: PathBuf, is_wk3: bool) {
        // Wipe UI state; the loaded engine is the new source of truth.
        self.wb_mut().cells.clear();
        self.wb_mut().clear_all_cell_formats();
        self.wb_mut().cell_text_styles.clear();
        self.wb_mut().cell_alignments.clear();
        self.wb_mut().cell_fills.clear();
        self.wb_mut().cell_font_styles.clear();
        self.wb_mut().cell_borders.clear();
        self.wb_mut().comments.clear();
        self.wb_mut().merges.clear();
        self.wb_mut().frozen.clear();
        self.wb_mut().sheet_states.clear();
        self.wb_mut().tables.clear();
        self.wb_mut().sheet_colors.clear();
        self.wb_mut().col_widths.clear();
        self.wb_mut().default_col_width = 9;
        self.wb_mut().hidden_cols.clear();
        self.wb_mut().named_ranges.clear();
        self.wb_mut().name_notes.clear();
        self.entry = None;
        self.wb_mut().pointer = Address::A1;
        self.wb_mut().viewport_col_offset = 0;
        self.wb_mut().viewport_row_offset = 0;
        self.recalc_pending = false;

        // Pull every non-empty cell into the UI cache.
        let sheet_names = self.wb().engine.all_sheet_names();
        let sheet_refs: Vec<&str> = sheet_names.iter().map(String::as_str).collect();
        for (addr, cv) in self.wb_mut().engine.used_cells() {
            if let Some(contents) = cell_view_to_contents(&cv, &sheet_refs) {
                self.wb_mut().cells.insert(addr, contents);
            }
        }
        // Apply the formula-source sidecar if present. The sidecar
        // is the source of truth for `expr` whenever it has an
        // entry — the cosmetic reverse translator above is the
        // fallback for cells without one (e.g. files originating
        // from Excel, or saved before this feature landed).
        if let Ok(sources) = l123_io::formula_sources::read_from_xlsx(&path) {
            for (addr, src) in sources {
                if let Some(CellContents::Formula { expr, .. }) = self.wb_mut().cells.get_mut(&addr)
                {
                    *expr = src;
                }
            }
        }
        for (addr, w) in self.wb_mut().engine.used_column_widths() {
            self.wb_mut().col_widths.insert((addr.sheet, addr.col), w);
        }
        for (addr, style) in self.wb_mut().engine.used_cell_text_styles() {
            self.wb_mut().cell_text_styles.insert(addr, style);
        }
        for (addr, fmt) in self.wb_mut().engine.used_cell_formats() {
            self.wb_mut().cell_formats.insert(addr, fmt);
        }
        for (addr, raw) in self.wb_mut().engine.used_cell_format_strings() {
            self.wb_mut().cell_format_overrides.insert(addr, raw);
        }
        // Layer the cell-format-extras sidecar on top of the engine's
        // num_fmt-based view (kind override for non-Excel kinds, parens
        // flag, negative-color). See `l123_io::cell_formats`.
        let format_extras = l123_io::cell_formats::read_from_xlsx(&path).unwrap_or_default();
        if let Some(g) = format_extras.global {
            self.wb_mut().global_format = g.apply_to(self.wb().global_format);
        }
        for (addr, fe) in &format_extras.cells {
            let base = self
                .wb()
                .cell_formats
                .get(addr)
                .copied()
                .unwrap_or(Format::GENERAL);
            self.wb_mut().cell_formats.insert(*addr, fe.apply_to(base));
        }
        for (addr, a) in self.wb_mut().engine.used_cell_alignments() {
            self.wb_mut().cell_alignments.insert(addr, a);
        }
        for (addr, f) in self.wb_mut().engine.used_cell_fills() {
            self.wb_mut().cell_fills.insert(addr, f);
        }
        for (addr, fs) in self.wb_mut().engine.used_cell_font_styles() {
            self.wb_mut().cell_font_styles.insert(addr, fs);
        }
        for (addr, b) in self.wb_mut().engine.used_cell_borders() {
            self.wb_mut().cell_borders.insert(addr, b);
        }
        for c in self.wb_mut().engine.used_comments() {
            self.wb_mut().comments.insert(c.addr, c);
        }
        for (sheet, m) in self.wb_mut().engine.used_merged_cells() {
            self.wb_mut().merges.entry(sheet).or_default().push(m);
        }
        let sheet_count = self.wb().engine.sheet_count();
        for sheet_idx in 0..sheet_count {
            let sid = SheetId(sheet_idx);
            let f = self.wb().engine.frozen_panes(sid);
            if f != (0, 0) {
                self.wb_mut().frozen.insert(sid, f);
            }
        }
        for sheet_idx in 0..sheet_count {
            let sid = SheetId(sheet_idx);
            let st = self.wb().engine.sheet_state(sid);
            if st != SheetState::Visible {
                self.wb_mut().sheet_states.insert(sid, st);
            }
        }
        for (sheet, t) in self.wb_mut().engine.used_tables() {
            self.wb_mut().tables.entry(sheet).or_default().push(t);
        }
        // Pull workbook-global defined names back into the UI map so
        // POINT typed-buffer name resolution and Alt-letter macro
        // dispatch (including \0 autoexec) survive a save → reload.
        // Keys are lowercased, matching the /RNC ingestion path.
        for (name, range) in self.wb_mut().engine.used_defined_names() {
            self.wb_mut()
                .named_ranges
                .insert(name.to_ascii_lowercase(), range);
        }
        redirect_pointer_off_hidden(self.wb_mut());
        for sheet_idx in 0..sheet_count {
            let sid = SheetId(sheet_idx);
            if let Some(c) = self.wb().engine.sheet_color(sid) {
                self.wb_mut().sheet_colors.insert(sid, c);
            }
        }

        // For a `.WK3` source, set the save target to "<orig>.WK3.xlsx"
        // so /File Save converts to xlsx without overwriting the legacy
        // file. The original WK3 stays untouched on disk; there is no
        // engine-side `save_wk3`.
        let active = if is_wk3 {
            let mut buf = path.into_os_string();
            buf.push(".xlsx");
            PathBuf::from(buf)
        } else {
            path
        };
        self.wb_mut().active_path = Some(active);
        self.wb_mut().dirty = false;
        self.mode = Mode::Ready;
    }

    /// Push every UI-side override into the engine so the saved
    /// xlsx carries column widths, text styles, formats, alignments,
    /// fills, font styles, borders, comments, merges, frozen panes,
    /// sheet states, tables, and tab colors. Mutates the engine in
    /// place; safe to call repeatedly. Factored out of the legacy
    /// sync `save_workbook_to` so the §4.7 async `/File Save` path
    /// can run the same prep before handing the engine to a worker.
    fn push_ui_overrides_into_engine(&mut self) {
        // Push UI-side column-width overrides into the engine so they
        // land in the xlsx. `col_widths` only contains non-default
        // entries; the engine default is preserved for every other
        // column.
        let widths: Vec<((SheetId, u16), u8)> =
            self.wb().col_widths.iter().map(|(k, v)| (*k, *v)).collect();
        for ((sheet, col), w) in widths {
            let _ = self.wb_mut().engine.set_column_width(sheet, col, w);
        }
        // Push per-cell WYSIWYG text styles (bold / italic / underline)
        // into the engine so they land in the xlsx font-run table.
        let styles: Vec<(Address, TextStyle)> = self
            .wb()
            .cell_text_styles
            .iter()
            .map(|(a, s)| (*a, *s))
            .collect();
        for (addr, style) in styles {
            let _ = self.wb_mut().engine.set_cell_text_style(addr, style);
        }
        // Push per-cell number formats so xlsx carries the num_fmt
        // that /File Retrieve reads back. Cells with an Excel-format
        // override (loaded verbatim from xlsx and untouched by the
        // user) bypass the canonical D-letter round-trip and write
        // their original num_fmt string back unchanged.
        let formats: Vec<(Address, Format)> = self
            .wb()
            .cell_formats
            .iter()
            .map(|(a, f)| (*a, *f))
            .collect();
        for (addr, fmt) in formats {
            let _ = self.wb_mut().engine.set_cell_format(addr, fmt);
        }
        let overrides: Vec<(Address, String)> = self
            .wb()
            .cell_format_overrides
            .iter()
            .map(|(a, s)| (*a, s.clone()))
            .collect();
        for (addr, raw) in overrides {
            let _ = self.wb_mut().engine.set_cell_format_string(addr, &raw);
        }
        // Push per-cell alignments so xlsx preserves the horizontal /
        // vertical / wrap settings imported (or assigned) on L123's side.
        let aligns: Vec<(Address, Alignment)> = self
            .wb()
            .cell_alignments
            .iter()
            .map(|(a, al)| (*a, *al))
            .collect();
        for (addr, align) in aligns {
            let _ = self.wb_mut().engine.set_cell_alignment(addr, align);
        }
        // Push per-cell fills so the background color round-trips.
        let fills: Vec<(Address, Fill)> =
            self.wb().cell_fills.iter().map(|(a, f)| (*a, *f)).collect();
        for (addr, fill) in fills {
            let _ = self.wb_mut().engine.set_cell_fill(addr, fill);
        }
        // Push per-cell font styles (color / size / strike).
        let font_styles: Vec<(Address, FontStyle)> = self
            .wb()
            .cell_font_styles
            .iter()
            .map(|(a, f)| (*a, *f))
            .collect();
        for (addr, fs) in font_styles {
            let _ = self.wb_mut().engine.set_cell_font_style(addr, fs);
        }
        // Push per-cell borders (all 4 sides preserve through xlsx).
        let borders: Vec<(Address, Border)> = self
            .wb()
            .cell_borders
            .iter()
            .map(|(a, b)| (*a, *b))
            .collect();
        for (addr, b) in borders {
            let _ = self.wb_mut().engine.set_cell_border(addr, b);
        }
        // Push per-cell comments.  IronCalc 0.7 doesn't actually
        // serialize these on xlsx save (upstream gap, pinned by
        // `comments_are_dropped_on_xlsx_save_upstream_gap` in the
        // engine adapter tests).  The setter is still the right UI
        // boundary — when upstream closes the gap no L123 work is
        // needed here.
        let comments: Vec<Comment> = self.wb().comments.values().cloned().collect();
        for c in comments {
            let _ = self.wb_mut().engine.set_comment(c);
        }
        // Push merged ranges.  IronCalc's xlsx exporter writes
        // <mergeCells> faithfully, so this round-trips end-to-end.
        let merges: Vec<Merge> = self
            .wb()
            .merges
            .values()
            .flat_map(|v| v.iter().copied())
            .collect();
        for m in merges {
            let _ = self.wb_mut().engine.set_merged_range(m);
        }
        // Push frozen-pane counts.  IronCalc round-trips these
        // natively via `<pane state="frozen" .../>` in sheet XML.
        let frozen: Vec<(SheetId, u32, u16)> = self
            .wb()
            .frozen
            .iter()
            .map(|(s, &(r, c))| (*s, r, c))
            .collect();
        for (sid, rows, cols) in frozen {
            let _ = self.wb_mut().engine.set_frozen_panes(sid, rows, cols);
        }
        // Push sheet visibility states.  Round-trips natively via the
        // workbook XML's `<sheet state="..."/>` attribute.
        let sheet_states: Vec<(SheetId, SheetState)> = self
            .wb()
            .sheet_states
            .iter()
            .map(|(s, &st)| (*s, st))
            .collect();
        for (sid, st) in sheet_states {
            let _ = self.wb_mut().engine.set_sheet_state(sid, st);
        }
        // Push tables.  IronCalc 0.7 doesn't actually serialize these
        // through xlsx export (upstream gap, pinned by
        // `tables_are_dropped_on_xlsx_save_upstream_gap`); the setter
        // is still the right UI boundary.
        let tables: Vec<(SheetId, Table)> = self
            .wb()
            .tables
            .iter()
            .flat_map(|(s, ts)| ts.iter().map(move |t| (*s, t.clone())))
            .collect();
        for (sid, t) in tables {
            let _ = self.wb_mut().engine.set_table(sid, t);
        }
        // Push sheet tab colors.  IronCalc 0.7's xlsx exporter drops
        // these (upstream gap), but the setter is still the right UI
        // boundary — when upstream closes the gap no code change is
        // needed here.
        let sheet_colors: Vec<(SheetId, RgbColor)> = self
            .wb()
            .sheet_colors
            .iter()
            .map(|(s, c)| (*s, *c))
            .collect();
        for (sid, color) in sheet_colors {
            let _ = self.wb_mut().engine.set_sheet_color(sid, Some(color));
        }
    }

    /// Synchronous `/File Save` — used by xlsx round-trip unit tests
    /// where stepping through tokio adds noise. Production /FS goes
    /// through `queue_file_save` so the UI doesn't freeze.
    #[cfg(test)]
    fn save_workbook_to(&mut self, path: PathBuf) {
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                let _ = std::fs::create_dir_all(parent);
            }
        }
        self.push_ui_overrides_into_engine();
        if self.wb_mut().engine.save_xlsx(&path).is_ok() {
            let sources = self.formula_sources_snapshot();
            let _ = l123_io::formula_sources::write_to_xlsx(&path, &sources);
            let extras = self.cell_format_extras_snapshot();
            let _ = l123_io::cell_formats::write_to_xlsx(&path, &extras);
            self.wb_mut().active_path = Some(path);
            self.wb_mut().dirty = false;
        }
    }

    /// Snapshot of every formula cell's user-typed Lotus source, used
    /// by the formula-source sidecar embedded in the xlsx zip so
    /// save → reload preserves shapes the cosmetic reverse
    /// translator can't recover (arg-fix wrappers, emulated
    /// functions like @CTERM, 3D-range expansions).
    fn formula_sources_snapshot(&self) -> HashMap<Address, String> {
        self.wb()
            .cells
            .iter()
            .filter_map(|(addr, c)| match c {
                CellContents::Formula { expr, .. } => Some((*addr, expr.clone())),
                _ => None,
            })
            .collect()
    }

    /// Snapshot of every cell-format-extra (kind override for non-Excel
    /// kinds, parens flag, negative-color override) plus the workbook
    /// global default's extras. Written into `l123/cell_formats.tsv`
    /// inside the xlsx zip so save → reload preserves what Excel's
    /// `num_fmt` system can't represent. See
    /// [`l123_io::cell_formats`] for the on-disk shape.
    fn cell_format_extras_snapshot(&self) -> l123_io::cell_formats::CellFormatExtras {
        use l123_io::cell_formats::{CellFormatExtras, FormatExtras};
        let global = FormatExtras::from_format(self.wb().global_format);
        let cells = self
            .wb()
            .cell_formats
            .iter()
            .filter_map(|(addr, fmt)| FormatExtras::from_format(*fmt).map(|fe| (*addr, fe)))
            .collect();
        CellFormatExtras { global, cells }
    }

    /// Queue an async `/File Save`. Pushes UI overrides into the
    /// engine on the main thread (fast O(N-overridden-cells)), then
    /// hands ownership of the engine to a worker that does the
    /// actual xlsx write. The worker also writes the formula-source
    /// sidecar on success. Engine is restored to the workbook when
    /// the worker returns (success, error, or cancel).
    fn queue_file_save(&mut self, path: PathBuf) {
        self.push_ui_overrides_into_engine();
        let formula_sources = self.formula_sources_snapshot();
        let cell_format_extras = self.cell_format_extras_snapshot();
        let placeholder = IronCalcEngine::new().expect("IronCalc placeholder engine init");
        let engine = std::mem::replace(&mut self.wb_mut().engine, placeholder);
        let name = display_basename(&path);
        self.queue_async_op(
            "Saving",
            name,
            QueuedOp::FileSave {
                engine,
                path,
                formula_sources,
                cell_format_extras,
            },
        );
    }

    /// Queue an async `/File Import {Numbers,Text}`. Engine is
    /// taken out of the workbook and travels with the op so the
    /// per-row `set_user_input` calls run off the UI thread.
    /// `numeric_split = true` corresponds to `/FIN` (CSV split with
    /// number coercion); `false` to `/FIT` (one label per line, no
    /// splitting).
    fn queue_file_import(&mut self, path: PathBuf, numeric_split: bool) {
        let origin = self.wb().pointer;
        let placeholder = IronCalcEngine::new().expect("IronCalc placeholder engine init");
        let engine = std::mem::replace(&mut self.wb_mut().engine, placeholder);
        let name = display_basename(&path);
        let queued = if numeric_split {
            QueuedOp::FileImportNumbers {
                engine,
                path,
                origin,
            }
        } else {
            QueuedOp::FileImportText {
                engine,
                path,
                origin,
            }
        };
        self.queue_async_op("Importing", name, queued);
    }

    fn commit_erase_confirm(&mut self, choice: usize) {
        let Some(ec) = self.erase_confirm.take() else {
            self.mode = Mode::Ready;
            return;
        };
        match choice {
            // No — leave the file alone.
            0 => self.mode = Mode::Ready,
            // Yes — delete it.  A failure (missing file, permission
            // denied) surfaces on line 3 via the standard error path
            // rather than panicking.
            1 => {
                if let Err(e) = std::fs::remove_file(&ec.path) {
                    self.set_error(format!("Cannot erase {}: {e}", ec.path.display()));
                } else {
                    self.mode = Mode::Ready;
                }
            }
            _ => self.mode = Mode::Ready,
        }
    }

    /// Execute the user's pick on the Cancel/Replace/Backup submenu.
    fn commit_save_confirm(&mut self, choice: usize) {
        let Some(sc) = self.save_confirm.take() else {
            self.mode = Mode::Ready;
            return;
        };
        match choice {
            0 => {
                // Cancel — no write.
                self.mode = Mode::Ready;
            }
            1 => {
                // Replace — overwrite. IronCalc's save_xlsx refuses
                // to clobber an existing file, so blow it away here
                // (the user explicitly chose Replace).
                let _ = std::fs::remove_file(&sc.path);
                self.queue_file_save(sc.path);
            }
            2 => {
                // Backup — rename existing to .BAK, then save.
                let backup = sc.path.with_extension("BAK");
                let _ = std::fs::rename(&sc.path, &backup);
                self.queue_file_save(sc.path);
            }
            _ => {
                self.mode = Mode::Ready;
            }
        }
    }

    /// Sheets targeted by a structural op at the current pointer. GROUP
    /// mode broadcasts to every sheet in the active file; otherwise
    /// only the pointer's sheet.
    fn target_sheets(&self) -> Vec<SheetId> {
        if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![self.wb().pointer.sheet]
        }
    }

    fn insert_row_at_pointer(&mut self, n: u32) {
        let at = self.wb().pointer.row;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            if self.wb_mut().engine.insert_rows(sheet, at, n).is_ok() {
                shift_cells_rows(&mut self.wb_mut().cells, sheet, at, n as i64);
                // One RowInsert entry per inserted row so undo can
                // replay them cleanly (delete_rows with n=1).
                for k in 0..n {
                    batch.push(JournalEntry::RowInsert { sheet, at: at + k });
                }
            }
        }
        if !batch.is_empty() {
            self.wb_mut().dirty = true;
        }
        self.push_journal_batch(batch);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.close_menu();
    }

    fn delete_row_at_pointer(&mut self, n: u32) {
        let at = self.wb_mut().pointer.row;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            // Capture the cells and formats about to be destroyed so
            // Alt-F4 can reinstate them. Only the first `n` rows on
            // this sheet are captured; deletion is always 1 for M5.
            let captured_cells: Vec<(Address, CellContents)> = self
                .wb_mut()
                .cells
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.row >= at && a.row < at + n)
                .map(|(a, c)| (*a, c.clone()))
                .collect();
            let captured_formats: Vec<(Address, Format)> = self
                .wb_mut()
                .cell_formats
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.row >= at && a.row < at + n)
                .map(|(a, f)| (*a, *f))
                .collect();
            let captured_text_styles: Vec<(Address, TextStyle)> = self
                .wb_mut()
                .cell_text_styles
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.row >= at && a.row < at + n)
                .map(|(a, s)| (*a, *s))
                .collect();
            if self.wb_mut().engine.delete_rows(sheet, at, n).is_ok() {
                self.wb_mut()
                    .cells
                    .retain(|a, _| !(a.sheet == sheet && a.row >= at && a.row < at + n));
                self.wb_mut()
                    .cell_formats
                    .retain(|a, _| !(a.sheet == sheet && a.row >= at && a.row < at + n));
                self.wb_mut()
                    .cell_text_styles
                    .retain(|a, _| !(a.sheet == sheet && a.row >= at && a.row < at + n));
                shift_cells_rows(&mut self.wb_mut().cells, sheet, at + n, -(n as i64));
                batch.push(JournalEntry::RowDelete {
                    sheet,
                    at,
                    cells: captured_cells,
                    formats: captured_formats,
                    text_styles: captured_text_styles,
                });
            }
        }
        if !batch.is_empty() {
            self.wb_mut().dirty = true;
        }
        self.push_journal_batch(batch);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.close_menu();
    }

    fn insert_col_at_pointer(&mut self, n: u16) {
        let at = self.wb().pointer.col;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            if self.wb_mut().engine.insert_cols(sheet, at, n).is_ok() {
                shift_cells_cols(&mut self.wb_mut().cells, sheet, at, n as i32);
                for k in 0..n {
                    batch.push(JournalEntry::ColInsert { sheet, at: at + k });
                }
            }
        }
        if !batch.is_empty() {
            self.wb_mut().dirty = true;
        }
        self.push_journal_batch(batch);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.close_menu();
    }

    /// /Worksheet Insert Sheet Before: a new empty sheet takes the
    /// current sheet's slot; the existing sheet shifts forward one
    /// position. The pointer follows the original data, so it ends up
    /// on the (shifted) original sheet rather than on the new blank.
    fn insert_sheet_before_current(&mut self) {
        let at = self.wb().pointer.sheet.0;
        let (col, row) = (self.wb().pointer.col, self.wb().pointer.row);
        let wb = self.wb_mut();
        if wb.engine.insert_sheet_at(at).is_ok() {
            shift_sheets_from(
                &mut wb.cells,
                &mut wb.cell_formats,
                &mut wb.cell_format_overrides,
                &mut wb.cell_text_styles,
                &mut wb.col_widths,
                at,
                1,
            );
            wb.pointer = Address::new(SheetId(at + 1), col, row);
            wb.engine.recalc();
            self.refresh_formula_caches();
        }
        self.close_menu();
    }

    /// /Worksheet Insert Sheet After: a new empty sheet is inserted at
    /// the position after the current one. The pointer stays on the
    /// current sheet; Ctrl-PgDn reveals the new blank.
    fn insert_sheet_after_current(&mut self) {
        let at = self.wb().pointer.sheet.0 + 1;
        let wb = self.wb_mut();
        if wb.engine.insert_sheet_at(at).is_ok() {
            shift_sheets_from(
                &mut wb.cells,
                &mut wb.cell_formats,
                &mut wb.cell_format_overrides,
                &mut wb.cell_text_styles,
                &mut wb.col_widths,
                at,
                1,
            );
            wb.engine.recalc();
            self.refresh_formula_caches();
        }
        self.close_menu();
    }

    /// /Worksheet Delete Sheet: drop the worksheet at the pointer.
    /// Sheets after it shift back one slot; the pointer stays at the
    /// same column/row on whatever sheet now occupies that slot
    /// (clamped to the last surviving sheet). The engine refuses to
    /// delete the only remaining sheet — that's silently a no-op
    /// here, leaving the workbook intact.
    fn delete_sheet_at_pointer(&mut self) {
        let at = self.wb().pointer.sheet.0;
        let (col, row) = (self.wb().pointer.col, self.wb().pointer.row);
        let wb = self.wb_mut();
        if wb.engine.delete_sheet_at(at).is_ok() {
            drop_sheet_from_caches(
                &mut wb.cells,
                &mut wb.cell_formats,
                &mut wb.cell_format_overrides,
                &mut wb.cell_text_styles,
                &mut wb.col_widths,
                at,
            );
            let new_count = wb.engine.sheet_count();
            let new_sheet = if new_count == 0 {
                0
            } else {
                at.min(new_count - 1)
            };
            wb.pointer = Address::new(SheetId(new_sheet), col, row);
            wb.engine.recalc();
            self.refresh_formula_caches();
        }
        self.close_menu();
    }

    /// /Worksheet Delete File: drop the foreground active file from
    /// memory. When more than one file is open, the previous file
    /// (or the first, if we were already on the first) takes focus.
    /// Deleting the only remaining active file resets the workspace
    /// to a single blank workbook — same end-state as
    /// `/Worksheet Erase Yes`.
    fn delete_current_file(&mut self) {
        if self.active_files.len() <= 1 {
            self.execute_worksheet_erase();
            return;
        }
        self.active_files.remove(self.current);
        if self.current >= self.active_files.len() {
            self.current = self.active_files.len() - 1;
        }
        self.close_menu();
    }

    /// Ctrl-PgDn / Ctrl-PgUp: jump to the next / previous sheet. Clamps
    /// at the bookends — no wrap.  Hidden / VeryHidden sheets are
    /// skipped: stepping with `delta=+1` past a hidden sheet lands on
    /// the next visible one, not on the hidden one itself.  When all
    /// sheets in the requested direction are hidden, the pointer stays
    /// put.
    fn move_sheet(&mut self, delta: i32) {
        let count = self.wb().engine.sheet_count();
        if count == 0 || delta == 0 {
            return;
        }
        let cur = self.wb().pointer.sheet.0 as i32;
        let max = count as i32 - 1;
        let step = delta.signum();
        let mut probe = cur + step;
        let mut landed: Option<u16> = None;
        while (0..=max).contains(&probe) {
            let sid = SheetId(probe as u16);
            if self
                .wb()
                .sheet_states
                .get(&sid)
                .copied()
                .unwrap_or(SheetState::Visible)
                .is_visible()
            {
                landed = Some(probe as u16);
                if probe - cur == delta {
                    break;
                }
                // Continue past this visible sheet only if we still
                // owe more steps in `delta`.  `delta = ±1` short-
                // circuits above; for ±N we keep stepping.
                if (probe - cur).signum() != step {
                    break;
                }
            }
            probe += step;
        }
        let Some(next) = landed else {
            return;
        };
        let wb = self.wb_mut();
        if next != wb.pointer.sheet.0 {
            wb.pointer = Address::new(SheetId(next), 0, 0);
            wb.viewport_col_offset = 0;
            wb.viewport_row_offset = 0;
        }
    }

    // ---------------- POINT mode ----------------

    fn begin_color(&mut self, target: ColorTarget, color: RgbColor) {
        self.begin_point(PendingCommand::RangeColor {
            target,
            color: Some(color),
        });
    }

    fn begin_point(&mut self, pending: PendingCommand) {
        self.menu = None;
        self.point = Some(PointState {
            anchor: Some(self.wb().pointer),
            pending,
            typed: String::new(),
        });
        self.mode = Mode::Point;
    }

    fn cancel_point(&mut self) {
        self.point = None;
        self.mode = Mode::Ready;
    }

    /// Range currently highlighted. If the user has unanchored (single Esc),
    /// this collapses to the pointer's cell.
    fn highlight_range(&self) -> Range {
        match self.point.as_ref().and_then(|p| p.anchor) {
            Some(anchor) => Range {
                start: anchor,
                end: self.wb().pointer,
            }
            .normalized(),
            None => Range::single(self.wb().pointer),
        }
    }

    /// Esc during POINT: with a non-empty typed range buffer, first
    /// clear the buffer (returning to highlight POINT). Otherwise the
    /// usual cascade — first press unanchors, second cancels back to
    /// READY.
    fn esc_in_point(&mut self) {
        let Some(ps) = self.point.as_mut() else {
            return;
        };
        if !ps.typed.is_empty() {
            ps.typed.clear();
            return;
        }
        if ps.anchor.is_some() {
            ps.anchor = None;
        } else {
            self.cancel_point();
        }
    }

    /// `.` during POINT: if unanchored, anchor at current pointer. If
    /// anchored, cycle the free corner clockwise.
    fn period_in_point(&mut self) {
        // Snapshot the pointer up-front so the inner match can touch
        // both the point state (via `self.point`) and the workbook
        // (via `self.wb_mut()`) without simultaneous borrows.
        let pointer = self.wb().pointer;
        let anchor = self.point.as_ref().and_then(|p| p.anchor);
        match anchor {
            None => {
                if let Some(ps) = self.point.as_mut() {
                    ps.anchor = Some(pointer);
                }
            }
            Some(anchor) => {
                let (min_c, max_c) = (pointer.col.min(anchor.col), pointer.col.max(anchor.col));
                let (min_r, max_r) = (pointer.row.min(anchor.row), pointer.row.max(anchor.row));
                let at_min_col = pointer.col == min_c;
                let at_min_row = pointer.row == min_r;
                let (new_col, new_row, new_anchor_col, new_anchor_row) =
                    match (at_min_col, at_min_row) {
                        // TL → TR
                        (true, true) => (max_c, min_r, min_c, max_r),
                        // TR → BR
                        (false, true) => (max_c, max_r, min_c, min_r),
                        // BR → BL
                        (false, false) => (min_c, max_r, max_c, min_r),
                        // BL → TL
                        (true, false) => (min_c, min_r, max_c, max_r),
                    };
                self.wb_mut().pointer = Address::new(pointer.sheet, new_col, new_row);
                if let Some(ps) = self.point.as_mut() {
                    ps.anchor = Some(Address::new(anchor.sheet, new_anchor_col, new_anchor_row));
                }
                self.scroll_into_view();
            }
        }
    }

    fn commit_point(&mut self) {
        // Typed range buffer takes precedence over the highlight
        // anchor+pointer pair. Resolution order: comma-separated range
        // list first (also handles plain `A1..D5`), then a defined
        // range name (case-insensitive, single name only). Both miss →
        // silent no-op clear-buffer-stay-in-POINT, same shape as F5
        // GOTO.
        let typed_input: Option<RangeInput> = match self.point.as_ref() {
            Some(ps) if !ps.typed.is_empty() => {
                let default_sheet = self.wb().pointer.sheet;
                let resolved = RangeInput::parse_with_default_sheet(&ps.typed, default_sheet)
                    .ok()
                    .or_else(|| {
                        // Single typed token that didn't parse as an
                        // address — try the named-ranges table.
                        if ps.typed.contains(',') {
                            None
                        } else {
                            self.wb()
                                .named_ranges
                                .get(&ps.typed.to_ascii_lowercase())
                                .copied()
                                .map(RangeInput::One)
                        }
                    });
                match resolved {
                    Some(ri) => Some(ri),
                    None => {
                        if let Some(ps) = self.point.as_mut() {
                            ps.typed.clear();
                        }
                        return;
                    }
                }
            }
            _ => None,
        };
        let Some(ps) = self.point.take() else {
            self.mode = Mode::Ready;
            return;
        };
        let ranges: Vec<Range> = match typed_input {
            Some(ri) => ri.into_vec(),
            None => match ps.anchor {
                Some(a) => vec![Range {
                    start: a,
                    end: self.wb_mut().pointer,
                }
                .normalized()],
                None => vec![Range::single(self.wb_mut().pointer)],
            },
        };
        self.apply_pending_with_ranges(ps.pending, &ranges);
    }

    /// Dispatch a [`PendingCommand`] with a fully-resolved list of
    /// ranges. Reused by [`Self::commit_point`] (highlight/typed-buffer
    /// path) and by F3 NAMES selection in POINT (named-range path) so
    /// per-command effects don't drift out of sync. Multi-range commands
    /// iterate; single-range commands take the first range only.
    fn apply_pending_with_ranges(&mut self, pending: PendingCommand, ranges: &[Range]) {
        if ranges.is_empty() {
            self.mode = Mode::Ready;
            return;
        }
        let first = ranges[0];
        match pending {
            PendingCommand::RangeErase => {
                for r in ranges {
                    self.execute_range_erase(*r);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::CopyFrom => {
                self.transition_point(PendingCommand::CopyTo { source: first })
            }
            PendingCommand::MoveFrom => {
                self.transition_point(PendingCommand::MoveTo { source: first })
            }
            PendingCommand::CopyTo { source } => {
                if self.execute_copy(source, first) {
                    self.wb_mut().dirty = true;
                    self.mode = Mode::Ready;
                }
                // On dim-mismatch error, set_error already put the app
                // in Mode::Error — leave it.
            }
            PendingCommand::MoveTo { source } => {
                if self.execute_move(source, first) {
                    self.wb_mut().dirty = true;
                    self.mode = Mode::Ready;
                }
            }
            PendingCommand::SpecialCopyFrom => {
                self.transition_point(PendingCommand::SpecialCopyTo { source: first })
            }
            PendingCommand::SpecialMoveFrom => {
                self.transition_point(PendingCommand::SpecialMoveTo { source: first })
            }
            PendingCommand::SpecialCopyTo { source } => {
                if self.execute_special_copy(source, first) {
                    self.wb_mut().dirty = true;
                    self.mode = Mode::Ready;
                }
            }
            PendingCommand::SpecialMoveTo { source } => {
                if self.execute_special_move(source, first) {
                    self.wb_mut().dirty = true;
                    self.mode = Mode::Ready;
                }
            }
            PendingCommand::RangeLabel { new_prefix } => {
                for r in ranges {
                    self.execute_range_label(*r, new_prefix);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeFormat { format } => {
                for r in ranges {
                    self.execute_range_format(*r, format);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeParens { value } => {
                for r in ranges {
                    self.execute_range_parens(*r, value);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNegColor { color } => {
                for r in ranges {
                    self.execute_range_neg_color(*r, color);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeTextStyle { bits, set } => {
                for r in ranges {
                    self.execute_range_text_style(*r, bits, set);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeAlignment { halign } => {
                for r in ranges {
                    self.execute_range_alignment(*r, halign);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeColor { target, color } => {
                for r in ranges {
                    self.execute_range_color(*r, target, color);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNameCreate => {
                if let Some(name) = self.pending_name.take() {
                    let _ = self.wb_mut().engine.define_name(&name, first);
                    self.wb_mut()
                        .named_ranges
                        .insert(name.to_ascii_lowercase(), first);
                    self.wb_mut().engine.recalc();
                    self.refresh_formula_caches();
                    self.wb_mut().dirty = true;
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNameLabels { direction } => {
                for r in ranges {
                    self.execute_range_name_labels(*r, direction);
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNameTable => {
                self.execute_range_name_table(first.start);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNameNoteTable => {
                self.execute_range_name_note_table(first.start);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeProtect { unprotected } => {
                for r in ranges {
                    self.execute_range_protection(*r, unprotected);
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeInput => {
                self.enter_input_mode(first);
            }
            PendingCommand::RangeValueFrom => {
                self.transition_point(PendingCommand::RangeValueTo { src: first });
            }
            PendingCommand::RangeValueTo { src } => {
                self.execute_range_value(src, first.start);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeTransFrom => {
                self.transition_point(PendingCommand::RangeTransTo { src: first });
            }
            PendingCommand::RangeTransTo { src } => {
                self.execute_range_trans(src, first.start);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeJustify => {
                for r in ranges {
                    self.execute_range_justify(*r);
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::FileXtractRange { kind } => {
                if let Some(path) = self.pending_xtract_path.take() {
                    self.execute_file_xtract(first, kind, path);
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::PrintFileRange => {
                if let Some(session) = self.print.as_mut() {
                    session.ranges = ranges.to_vec();
                }
                // Back to the /PF submenu for Options/Go/…
                self.enter_print_file_menu();
            }
            PendingCommand::RangeSearchRange { scope } => {
                self.start_range_search_string_prompt(scope, first);
            }
            PendingCommand::GraphSeries { series } => {
                self.wb_mut().current_graph.set(series, first);
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnRangeSetWidth { width } => {
                for r in ranges {
                    self.execute_col_range_width(*r, Some(width));
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnRangeResetWidth => {
                for r in ranges {
                    self.execute_col_range_width(*r, None);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnHide => {
                for r in ranges {
                    self.execute_col_hide_display(*r, true);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnDisplay => {
                for r in ranges {
                    self.execute_col_hide_display(*r, false);
                }
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            // Mouse-drag selection has no command to execute. Enter
            // just clears the POINT state and lands back in READY; any
            // command the user invokes next can read the highlight if
            // they entered POINT first via /…  or click an icon while
            // POINT is still active.
            PendingCommand::MouseSelect => {
                self.mode = Mode::Ready;
            }
            PendingCommand::WorksheetLearnRange => {
                self.learn_range = Some(first.normalized());
                self.mode = Mode::Ready;
            }
            PendingCommand::DataFillRange => {
                self.start_data_fill_start_prompt(first.normalized());
            }
            PendingCommand::DataSortDataRange => {
                self.data_sort.data_range = Some(first.normalized());
                self.enter_data_sort_menu();
            }
            PendingCommand::DataSortKey => {
                self.enter_data_sort_dir_menu(first.start.col);
            }
            PendingCommand::DataDistributionValues => {
                let values = first.normalized();
                self.transition_point(PendingCommand::DataDistributionBins { values });
            }
            PendingCommand::DataDistributionBins { values } => {
                self.execute_data_distribution(values, first.normalized());
            }
            PendingCommand::DataRegressionXRange => {
                self.data_regression.x_range = Some(first.normalized());
                self.enter_data_regression_menu();
            }
            PendingCommand::DataRegressionYRange => {
                self.data_regression.y_range = Some(first.normalized());
                self.enter_data_regression_menu();
            }
            PendingCommand::DataRegressionOutputRange => {
                // Only the cursor position matters for the output
                // anchor; ignore any extent the highlight picked up
                // from the prior commit's anchor.
                self.data_regression.output_anchor = Some(self.wb().pointer);
                self.enter_data_regression_menu();
            }
            PendingCommand::DataMatrixInvertInput => {
                let source = first.normalized();
                self.begin_point(PendingCommand::DataMatrixInvertOutput { source });
            }
            PendingCommand::DataMatrixInvertOutput { source } => {
                let anchor = self.wb().pointer;
                self.execute_matrix_invert(source, anchor);
            }
            PendingCommand::DataMatrixMultiplyA => {
                let a = first.normalized();
                self.begin_point(PendingCommand::DataMatrixMultiplyB { a });
            }
            PendingCommand::DataMatrixMultiplyB { a } => {
                let b = first.normalized();
                self.begin_point(PendingCommand::DataMatrixMultiplyOutput { a, b });
            }
            PendingCommand::DataMatrixMultiplyOutput { a, b } => {
                let anchor = self.wb().pointer;
                self.execute_matrix_multiply(a, b, anchor);
            }
            PendingCommand::DataParseInputColumn => {
                self.data_parse.input_range = Some(first.normalized());
                self.enter_data_parse_menu();
            }
            PendingCommand::DataParseOutputRange => {
                self.data_parse.output_anchor = Some(self.wb().pointer);
                self.enter_data_parse_menu();
            }
            PendingCommand::DataTable1Range => {
                let range = first.normalized();
                self.begin_point(PendingCommand::DataTable1Input1 { range });
            }
            PendingCommand::DataTable1Input1 { range } => {
                let input1 = self.wb().pointer;
                self.execute_data_table_1(range, input1);
            }
            PendingCommand::DataTable2Range => {
                let range = first.normalized();
                self.begin_point(PendingCommand::DataTable2Input1 { range });
            }
            PendingCommand::DataTable2Input1 { range } => {
                let input1 = self.wb().pointer;
                self.begin_point(PendingCommand::DataTable2Input2 { range, input1 });
            }
            PendingCommand::DataTable2Input2 { range, input1 } => {
                let input2 = self.wb().pointer;
                self.execute_data_table_2(range, input1, input2);
            }
            PendingCommand::DataQueryInput => {
                self.data_query.input = Some(first.normalized());
                self.enter_data_query_menu();
            }
            PendingCommand::DataQueryCriteria => {
                self.data_query.criteria = Some(first.normalized());
                self.enter_data_query_menu();
            }
            PendingCommand::DataQueryOutput => {
                self.data_query.output = Some(first.normalized());
                self.enter_data_query_menu();
            }
        }
    }

    fn start_range_search_string_prompt(&mut self, scope: SearchScope, range: Range) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter search string:".into(),
            buffer: String::new(),
            next: PromptNext::RangeSearchString { scope, range },
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn enter_range_search_find_replace_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::RANGE_SEARCH_FIND_REPLACE_MENU));
        self.mode = Mode::Menu;
    }

    fn start_range_search_replace_prompt(&mut self) {
        if self.search.is_none() {
            self.close_menu();
            return;
        }
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter replacement string:".into(),
            buffer: String::new(),
            next: PromptNext::RangeSearchReplacement,
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    /// Pick the first match, land in FIND mode. While in FIND,
    /// Enter advances (wrapping at the end), Esc exits. No matches
    /// → session is discarded and we return to READY.
    fn execute_range_search_find(&mut self) {
        let Some(mut session) = self.search.take() else {
            self.close_menu();
            return;
        };
        session.matches = self.find_matches(&session);
        if session.matches.is_empty() {
            self.close_menu();
            return;
        }
        session.cursor = 0;
        self.wb_mut().pointer = session.matches[0];
        self.scroll_into_view();
        self.menu = None;
        self.search = Some(session);
        self.mode = Mode::Find;
    }

    /// Collect the addresses within `session.range` whose content (per
    /// scope) contains `session.search` as a substring.
    fn find_matches(&self, session: &SearchSession) -> Vec<Address> {
        let r = session.range.normalized();
        let needle = &session.search;
        if needle.is_empty() {
            return Vec::new();
        }
        let mut out = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                let Some(contents) = self.wb().cells.get(&addr) else {
                    continue;
                };
                let matched = match (session.scope, contents) {
                    (SearchScope::Formulas, CellContents::Formula { expr, .. }) => {
                        expr.contains(needle)
                    }
                    (SearchScope::Labels, CellContents::Label { text, .. }) => {
                        text.contains(needle)
                    }
                    (SearchScope::Both, CellContents::Formula { expr, .. }) => {
                        expr.contains(needle)
                    }
                    (SearchScope::Both, CellContents::Label { text, .. }) => text.contains(needle),
                    _ => false,
                };
                if matched {
                    out.push(addr);
                }
            }
        }
        out
    }

    /// Build a fresh IronCalc workbook containing just `range`, then
    /// write it to `path`. Formulas variant preserves formulas;
    /// Values variant writes cached numeric/text values instead.
    fn execute_file_xtract(&mut self, range: Range, kind: XtractKind, path: PathBuf) {
        let r = range.normalized();
        let Ok(mut out) = IronCalcEngine::new() else {
            return;
        };
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                let Ok(cv) = self.wb_mut().engine.get_cell(addr) else {
                    continue;
                };
                if cv.value == Value::Empty && cv.formula.is_none() {
                    continue;
                }
                let input = xtract_cell_input(&cv, kind);
                let _ = out.set_user_input(addr, &input);
            }
        }
        out.recalc();
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                let _ = std::fs::create_dir_all(parent);
            }
        }
        let _ = out.save_xlsx(&path);
    }

    // ---------------- command-argument prompt ----------------

    fn enter_data_sort_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::DATA_SORT_MENU));
        self.mode = Mode::Menu;
    }

    fn enter_data_sort_dir_menu(&mut self, key_col: u16) {
        self.pending_sort_key_col = Some(key_col);
        self.menu = Some(MenuState::rooted_at(menu::DATA_SORT_DIR_MENU));
        self.mode = Mode::Menu;
    }

    fn bind_data_sort_dir(&mut self, dir: SortDir) {
        let slot = self.pending_sort_key_slot.take();
        let col = self.pending_sort_key_col.take();
        if let (Some(slot), Some(col)) = (slot, col) {
            match slot {
                SortKeySlot::Primary => self.data_sort.primary = Some((col, dir)),
                SortKeySlot::Secondary => self.data_sort.secondary = Some((col, dir)),
                SortKeySlot::Extra => self.data_sort.extra = Some((col, dir)),
            }
        }
        self.enter_data_sort_menu();
    }

    /// Sort the configured data range in place by primary (and
    /// optional secondary) key column. Empty/missing data range or
    /// missing primary key is a silent no-op back to READY — matches
    /// 1-2-3's behavior of refusing rather than erroring.
    fn execute_data_sort(&mut self) {
        let Some(range) = self.data_sort.data_range else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some((primary_col, primary_dir)) = self.data_sort.primary else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let secondary = self.data_sort.secondary;
        let extra = self.data_sort.extra;
        let r = range.normalized();
        let sheet = r.start.sheet;
        let row_lo = r.start.row;
        let row_hi = r.end.row;
        let col_lo = r.start.col;
        let col_hi = r.end.col;

        // Capture each row as a Vec of (col-offset, contents/format/style).
        type RowSnapshot = Vec<(u16, Option<CellContents>, Option<Format>, Option<TextStyle>)>;
        let mut rows: Vec<RowSnapshot> = Vec::with_capacity((row_hi - row_lo + 1) as usize);
        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for row in row_lo..=row_hi {
            let mut snap: RowSnapshot = Vec::with_capacity((col_hi - col_lo + 1) as usize);
            for col in col_lo..=col_hi {
                let addr = Address::new(sheet, col, row);
                let c = self.wb().cells.get(&addr).cloned();
                let f = self.wb().cell_formats.get(&addr).copied();
                let s = self.wb().cell_text_styles.get(&addr).copied();
                if let Some(ref cc) = c {
                    prev_cells.push((addr, cc.clone()));
                }
                if let Some(ff) = f {
                    prev_formats.push((addr, ff));
                }
                if let Some(ss) = s {
                    prev_text_styles.push((addr, ss));
                }
                snap.push((col - col_lo, c, f, s));
            }
            rows.push(snap);
        }

        let key_for = |snap: &RowSnapshot, key_col: u16| -> Option<CellContents> {
            let off = key_col.saturating_sub(col_lo);
            snap.iter()
                .find(|(o, _, _, _)| *o == off)
                .and_then(|(_, c, _, _)| c.clone())
        };

        rows.sort_by(|a, b| {
            let pa = key_for(a, primary_col);
            let pb = key_for(b, primary_col);
            let mut ord = compare_cell_contents(pa.as_ref(), pb.as_ref());
            if primary_dir == SortDir::Descending {
                ord = ord.reverse();
            }
            if ord == std::cmp::Ordering::Equal {
                if let Some((sec_col, sec_dir)) = secondary {
                    let sa = key_for(a, sec_col);
                    let sb = key_for(b, sec_col);
                    let mut sord = compare_cell_contents(sa.as_ref(), sb.as_ref());
                    if sec_dir == SortDir::Descending {
                        sord = sord.reverse();
                    }
                    ord = sord;
                }
            }
            if ord == std::cmp::Ordering::Equal {
                if let Some((ex_col, ex_dir)) = extra {
                    let ea = key_for(a, ex_col);
                    let eb = key_for(b, ex_col);
                    let mut eord = compare_cell_contents(ea.as_ref(), eb.as_ref());
                    if ex_dir == SortDir::Descending {
                        eord = eord.reverse();
                    }
                    ord = eord;
                }
            }
            ord
        });

        // Write the sorted rows back into the same rectangle.
        for (i, snap) in rows.iter().enumerate() {
            let row = row_lo + i as u32;
            for col in col_lo..=col_hi {
                let addr = Address::new(sheet, col, row);
                let off = col - col_lo;
                let entry = snap.iter().find(|(o, _, _, _)| *o == off);
                self.wb_mut().cells.remove(&addr);
                self.wb_mut().cell_formats.remove(&addr);
                self.wb_mut().cell_text_styles.remove(&addr);
                let _ = self.wb_mut().engine.clear_cell(addr);
                if let Some((_, contents, format, style)) = entry {
                    if let Some(c) = contents {
                        self.wb_mut().cells.insert(addr, c.clone());
                        self.push_to_engine_at(addr, c);
                    }
                    if let Some(f) = format {
                        self.wb_mut().cell_formats.insert(addr, *f);
                    }
                    if let Some(s) = style {
                        self.wb_mut().cell_text_styles.insert(addr, *s);
                    }
                }
            }
        }

        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    /// `/Data Matrix Invert` — Gauss-Jordan inverse of a square
    /// matrix read from `source`, written column-major into the
    /// rectangle anchored at `anchor`. Surfaces a status-line
    /// error on non-square or singular input.
    #[allow(clippy::needless_range_loop)]
    fn execute_matrix_invert(&mut self, source: Range, anchor: Address) {
        let s = source.normalized();
        let n_rows = (s.end.row - s.start.row + 1) as usize;
        let n_cols = (s.end.col - s.start.col + 1) as usize;
        if n_rows != n_cols {
            self.set_error("Matrix Invert: source range is not square");
            return;
        }
        let n = n_rows;
        let mut mat: Vec<Vec<f64>> = vec![vec![0.0; n]; n];
        for r in 0..n {
            for c in 0..n {
                let addr = Address::new(
                    s.start.sheet,
                    s.start.col + c as u16,
                    s.start.row + r as u32,
                );
                mat[r][c] = self.numeric_cell_value(addr).unwrap_or(0.0);
            }
        }
        let inv = match gauss_jordan_invert(mat) {
            Some(m) => m,
            None => {
                self.set_error("Matrix Invert: matrix is singular");
                return;
            }
        };
        self.write_matrix_at(anchor, &inv);
        self.mode = Mode::Ready;
    }

    /// `/Data Matrix Multiply` — write A*B into the rectangle
    /// anchored at `anchor`. Refuses with a status-line error when
    /// `cols(A) != rows(B)`.
    #[allow(clippy::needless_range_loop)]
    fn execute_matrix_multiply(&mut self, a_range: Range, b_range: Range, anchor: Address) {
        let ar = a_range.normalized();
        let br = b_range.normalized();
        let a_rows = (ar.end.row - ar.start.row + 1) as usize;
        let a_cols = (ar.end.col - ar.start.col + 1) as usize;
        let b_rows = (br.end.row - br.start.row + 1) as usize;
        let b_cols = (br.end.col - br.start.col + 1) as usize;
        if a_cols != b_rows {
            self.set_error("Matrix Multiply: cols(A) must equal rows(B)");
            return;
        }
        let mut a: Vec<Vec<f64>> = vec![vec![0.0; a_cols]; a_rows];
        for r in 0..a_rows {
            for c in 0..a_cols {
                let addr = Address::new(
                    ar.start.sheet,
                    ar.start.col + c as u16,
                    ar.start.row + r as u32,
                );
                a[r][c] = self.numeric_cell_value(addr).unwrap_or(0.0);
            }
        }
        let mut b: Vec<Vec<f64>> = vec![vec![0.0; b_cols]; b_rows];
        for r in 0..b_rows {
            for c in 0..b_cols {
                let addr = Address::new(
                    br.start.sheet,
                    br.start.col + c as u16,
                    br.start.row + r as u32,
                );
                b[r][c] = self.numeric_cell_value(addr).unwrap_or(0.0);
            }
        }
        let mut prod: Vec<Vec<f64>> = vec![vec![0.0; b_cols]; a_rows];
        for i in 0..a_rows {
            for j in 0..b_cols {
                let mut acc = 0.0_f64;
                for k in 0..a_cols {
                    acc += a[i][k] * b[k][j];
                }
                prod[i][j] = acc;
            }
        }
        self.write_matrix_at(anchor, &prod);
        self.mode = Mode::Ready;
    }

    /// Write a row-major matrix into the grid anchored at `anchor`.
    /// Captures previous cell contents so Alt-F4 reverts. Recalcs +
    /// marks the workbook dirty. Values are rounded to 12 significant
    /// decimal digits before storage to suppress floating-point noise
    /// from the linear-algebra kernels — `0.6000000000000001` becomes
    /// the exact `0.6` users expect to see.
    fn write_matrix_at(&mut self, anchor: Address, mat: &[Vec<f64>]) {
        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for (i, row) in mat.iter().enumerate() {
            for (j, value) in row.iter().enumerate() {
                let value = round_to_significant(*value, 12);
                let addr = Address::new(anchor.sheet, anchor.col + j as u16, anchor.row + i as u32);
                if let Some(c) = self.wb().cells.get(&addr) {
                    prev_cells.push((addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&addr) {
                    prev_formats.push((addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&addr) {
                    prev_text_styles.push((addr, *s));
                }
                let s = l123_core::format_number_general(value);
                let _ = self.wb_mut().engine.set_user_input(addr, &s);
                self.wb_mut()
                    .cells
                    .insert(addr, CellContents::Constant(Value::Number(value)));
            }
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
    }

    /// `/Data Table 1` — for each variable value in the left column
    /// of `range` (rows below the corner), substitute it into
    /// `input1`, recalc, and copy the resulting top-row formula
    /// values into the body cells. The original contents of `input1`
    /// are restored when the loop completes. Refuses degenerate
    /// (single-row or single-column) table ranges with no body
    /// cells.
    /// `/Data Table 2` — for each (var-1 in left column, var-2 in
    /// top row), substitute into Input cells 1 and 2, recalc, and
    /// write the value of the corner-cell formula into the body.
    /// Both Input cells are restored when the loop completes.
    fn execute_data_table_2(&mut self, range: Range, input1: Address, input2: Address) {
        let r = range.normalized();
        if r.start.row == r.end.row || r.start.col == r.end.col {
            self.set_error("Data Table 2: range must include at least one body cell");
            return;
        }
        let sheet = r.start.sheet;
        let formula_addr = Address::new(sheet, r.start.col, r.start.row);
        let original_input1 = self.wb().cells.get(&input1).cloned();
        let original_input2 = self.wb().cells.get(&input2).cloned();

        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for body_row in (r.start.row + 1)..=r.end.row {
            let var1_addr = Address::new(sheet, r.start.col, body_row);
            let Some(var1) = self.numeric_cell_value(var1_addr) else {
                continue;
            };
            let s1 = l123_core::format_number_general(var1);
            let _ = self.wb_mut().engine.set_user_input(input1, &s1);
            self.wb_mut()
                .cells
                .insert(input1, CellContents::Constant(Value::Number(var1)));

            for body_col in (r.start.col + 1)..=r.end.col {
                let var2_addr = Address::new(sheet, body_col, r.start.row);
                let Some(var2) = self.numeric_cell_value(var2_addr) else {
                    continue;
                };
                let s2 = l123_core::format_number_general(var2);
                let _ = self.wb_mut().engine.set_user_input(input2, &s2);
                self.wb_mut()
                    .cells
                    .insert(input2, CellContents::Constant(Value::Number(var2)));
                self.wb_mut().engine.recalc();
                self.refresh_formula_caches();

                let value = self
                    .wb()
                    .cells
                    .get(&formula_addr)
                    .map(|c| c.value())
                    .unwrap_or(Value::Empty);
                let body_addr = Address::new(sheet, body_col, body_row);
                if let Some(c) = self.wb().cells.get(&body_addr) {
                    prev_cells.push((body_addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&body_addr) {
                    prev_formats.push((body_addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&body_addr) {
                    prev_text_styles.push((body_addr, *s));
                }
                if let Value::Number(n) = value {
                    let s = l123_core::format_number_general(n);
                    let _ = self.wb_mut().engine.set_user_input(body_addr, &s);
                    self.wb_mut()
                        .cells
                        .insert(body_addr, CellContents::Constant(Value::Number(n)));
                } else {
                    self.wb_mut().cells.remove(&body_addr);
                    let _ = self.wb_mut().engine.clear_cell(body_addr);
                }
            }
        }

        // Restore both Input cells to their pre-call contents.
        if let Some(orig) = &original_input1 {
            prev_cells.push((input1, orig.clone()));
            self.wb_mut().cells.insert(input1, orig.clone());
            self.push_to_engine_at(input1, orig);
        } else {
            self.wb_mut().cells.remove(&input1);
            let _ = self.wb_mut().engine.clear_cell(input1);
        }
        if let Some(orig) = &original_input2 {
            prev_cells.push((input2, orig.clone()));
            self.wb_mut().cells.insert(input2, orig.clone());
            self.push_to_engine_at(input2, orig);
        } else {
            self.wb_mut().cells.remove(&input2);
            let _ = self.wb_mut().engine.clear_cell(input2);
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    fn execute_data_table_1(&mut self, range: Range, input1: Address) {
        let r = range.normalized();
        if r.start.row == r.end.row || r.start.col == r.end.col {
            self.set_error("Data Table 1: range must include at least one body cell");
            return;
        }
        let sheet = r.start.sheet;
        let formula_row = r.start.row;
        let var_col = r.start.col;
        let original_input1 = self.wb().cells.get(&input1).cloned();

        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for body_row in (formula_row + 1)..=r.end.row {
            let var_addr = Address::new(sheet, var_col, body_row);
            let Some(var_value) = self.numeric_cell_value(var_addr) else {
                continue;
            };
            let s = l123_core::format_number_general(var_value);
            let _ = self.wb_mut().engine.set_user_input(input1, &s);
            self.wb_mut()
                .cells
                .insert(input1, CellContents::Constant(Value::Number(var_value)));
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();

            for body_col in (var_col + 1)..=r.end.col {
                let formula_addr = Address::new(sheet, body_col, formula_row);
                let value = match self.wb().cells.get(&formula_addr) {
                    Some(c) => c.value(),
                    None => Value::Empty,
                };
                let body_addr = Address::new(sheet, body_col, body_row);
                if let Some(c) = self.wb().cells.get(&body_addr) {
                    prev_cells.push((body_addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&body_addr) {
                    prev_formats.push((body_addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&body_addr) {
                    prev_text_styles.push((body_addr, *s));
                }
                if let Value::Number(n) = value {
                    let s = l123_core::format_number_general(n);
                    let _ = self.wb_mut().engine.set_user_input(body_addr, &s);
                    self.wb_mut()
                        .cells
                        .insert(body_addr, CellContents::Constant(Value::Number(n)));
                } else {
                    self.wb_mut().cells.remove(&body_addr);
                    let _ = self.wb_mut().engine.clear_cell(body_addr);
                }
            }
        }

        // Capture the input cell's prior state for the journal, then
        // restore it (so the workbook visually returns to its
        // pre-/DT 1 state apart from the new body cells).
        if let Some(orig) = &original_input1 {
            prev_cells.push((input1, orig.clone()));
            self.wb_mut().cells.insert(input1, orig.clone());
            self.push_to_engine_at(input1, orig);
        } else {
            self.wb_mut().cells.remove(&input1);
            let _ = self.wb_mut().engine.clear_cell(input1);
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    fn enter_data_parse_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::DATA_PARSE_MENU));
        self.mode = Mode::Menu;
    }

    fn enter_data_query_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::DATA_QUERY_MENU));
        self.mode = Mode::Menu;
    }

    /// `/Data Query Find` — jump the pointer to the first record
    /// in the input range that matches the criteria. Silent
    /// no-op when input or criteria is unset, or when no record
    /// matches.
    fn execute_data_query_find(&mut self) {
        let Some((input, criteria)) = self.data_query.input.zip(self.data_query.criteria) else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let input = input.normalized();
        let criteria = criteria.normalized();
        let n_records = input.end.row.saturating_sub(input.start.row);
        for r in 0..n_records {
            if self.query_record_matches(input, criteria, r) {
                let row = input.start.row + 1 + r;
                self.wb_mut().pointer = Address::new(input.start.sheet, input.start.col, row);
                self.scroll_into_view();
                break;
            }
        }
        self.menu = None;
        self.mode = Mode::Ready;
    }

    /// `/Data Query Extract` (`unique=false`) and `/Data Query
    /// Unique` (`unique=true`). Walks the input range, collects
    /// matching records, and writes them into the output range
    /// below its header. When the output's row 1 has labels, only
    /// fields whose names match are copied (in output-header
    /// order); otherwise every input field is copied.
    fn execute_data_query_extract(&mut self, unique: bool) {
        let Some(input) = self.data_query.input else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(criteria) = self.data_query.criteria else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(output) = self.data_query.output else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let input = input.normalized();
        let criteria = criteria.normalized();
        let output = output.normalized();
        let in_field_count = (input.end.col - input.start.col + 1) as usize;
        let in_field_names: Vec<String> = (0..in_field_count)
            .map(|i| {
                let addr = Address::new(
                    input.start.sheet,
                    input.start.col + i as u16,
                    input.start.row,
                );
                self.cell_label_lower(addr).unwrap_or_default()
            })
            .collect();
        let out_field_count = (output.end.col - output.start.col + 1) as usize;
        let out_field_names: Vec<String> = (0..out_field_count)
            .map(|i| {
                let addr = Address::new(
                    output.start.sheet,
                    output.start.col + i as u16,
                    output.start.row,
                );
                self.cell_label_lower(addr).unwrap_or_default()
            })
            .collect();
        // For each output column, find the matching input column
        // by header label; or use the same column index when the
        // output header is empty.
        let header_present = out_field_names.iter().any(|s| !s.is_empty());
        let column_map: Vec<Option<usize>> = if header_present {
            out_field_names
                .iter()
                .map(|name| {
                    if name.is_empty() {
                        None
                    } else {
                        in_field_names.iter().position(|n| n == name)
                    }
                })
                .collect()
        } else {
            (0..out_field_count.min(in_field_count))
                .map(Some)
                .chain(std::iter::repeat_n(
                    None,
                    out_field_count.saturating_sub(in_field_count),
                ))
                .collect()
        };

        let n_records = input.end.row.saturating_sub(input.start.row);
        let mut emitted: Vec<Vec<Option<CellContents>>> = Vec::new();
        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();

        let max_out_rows = if output.end.row > output.start.row {
            (output.end.row - output.start.row) as usize
        } else {
            0
        };

        let mut out_row_idx: usize = 0;
        for r in 0..n_records {
            if !self.query_record_matches(input, criteria, r) {
                continue;
            }
            let in_row = input.start.row + 1 + r;
            let row_values: Vec<Option<CellContents>> = column_map
                .iter()
                .map(|maybe_in_col| {
                    maybe_in_col.and_then(|in_col| {
                        let addr = Address::new(
                            input.start.sheet,
                            input.start.col + in_col as u16,
                            in_row,
                        );
                        self.wb().cells.get(&addr).cloned()
                    })
                })
                .collect();
            if unique && emitted.iter().any(|r| r == &row_values) {
                continue;
            }
            if max_out_rows > 0 && out_row_idx >= max_out_rows {
                break;
            }
            emitted.push(row_values.clone());
            for (j, contents) in row_values.iter().enumerate() {
                let out_addr = Address::new(
                    output.start.sheet,
                    output.start.col + j as u16,
                    output.start.row + 1 + out_row_idx as u32,
                );
                if let Some(c) = self.wb().cells.get(&out_addr) {
                    prev_cells.push((out_addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&out_addr) {
                    prev_formats.push((out_addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&out_addr) {
                    prev_text_styles.push((out_addr, *s));
                }
                if let Some(c) = contents {
                    self.wb_mut().cells.insert(out_addr, c.clone());
                    self.push_to_engine_at(out_addr, c);
                } else {
                    self.wb_mut().cells.remove(&out_addr);
                    let _ = self.wb_mut().engine.clear_cell(out_addr);
                }
            }
            out_row_idx += 1;
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    /// `/Data Query Del` — drop matching records and shift the
    /// surviving records up so the input range stays compact below
    /// its header. Trailing rows in the original input range are
    /// cleared.
    fn execute_data_query_del(&mut self) {
        let Some(input) = self.data_query.input else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(criteria) = self.data_query.criteria else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let input = input.normalized();
        let criteria = criteria.normalized();
        let n_records = input.end.row.saturating_sub(input.start.row);
        let n_cols = (input.end.col - input.start.col + 1) as usize;
        let mut survivors: Vec<Vec<Option<CellContents>>> = Vec::new();
        for r in 0..n_records {
            if self.query_record_matches(input, criteria, r) {
                continue;
            }
            let in_row = input.start.row + 1 + r;
            let row_values: Vec<Option<CellContents>> = (0..n_cols)
                .map(|c| {
                    let addr = Address::new(input.start.sheet, input.start.col + c as u16, in_row);
                    self.wb().cells.get(&addr).cloned()
                })
                .collect();
            survivors.push(row_values);
        }
        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for r in 0..n_records {
            for c in 0..n_cols {
                let addr = Address::new(
                    input.start.sheet,
                    input.start.col + c as u16,
                    input.start.row + 1 + r,
                );
                if let Some(cc) = self.wb().cells.get(&addr) {
                    prev_cells.push((addr, cc.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&addr) {
                    prev_formats.push((addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&addr) {
                    prev_text_styles.push((addr, *s));
                }
            }
        }
        for (i, row) in survivors.iter().enumerate() {
            for (c, contents) in row.iter().enumerate() {
                let addr = Address::new(
                    input.start.sheet,
                    input.start.col + c as u16,
                    input.start.row + 1 + i as u32,
                );
                if let Some(cc) = contents {
                    self.wb_mut().cells.insert(addr, cc.clone());
                    self.push_to_engine_at(addr, cc);
                } else {
                    self.wb_mut().cells.remove(&addr);
                    let _ = self.wb_mut().engine.clear_cell(addr);
                }
            }
        }
        for r in survivors.len()..(n_records as usize) {
            for c in 0..n_cols {
                let addr = Address::new(
                    input.start.sheet,
                    input.start.col + c as u16,
                    input.start.row + 1 + r as u32,
                );
                self.wb_mut().cells.remove(&addr);
                let _ = self.wb_mut().engine.clear_cell(addr);
            }
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    /// True if record `r` (zero-based, below the input header)
    /// satisfies any of the criteria rows. A criterion row matches
    /// when every non-empty criterion cell in that row matches the
    /// corresponding input field. Field-name matching is
    /// case-insensitive label equality. Numeric criteria match by
    /// equality. Empty criterion cells impose no constraint. Any
    /// other criterion type (formula, date, ...) is treated as
    /// "no match" for this MVP slice.
    fn query_record_matches(&self, input: Range, criteria: Range, record_idx: u32) -> bool {
        let in_row = input.start.row + 1 + record_idx;
        let n_input_cols = (input.end.col - input.start.col + 1) as usize;
        let input_field_names: Vec<String> = (0..n_input_cols)
            .map(|i| {
                let addr = Address::new(
                    input.start.sheet,
                    input.start.col + i as u16,
                    input.start.row,
                );
                self.cell_label_lower(addr).unwrap_or_default()
            })
            .collect();
        let n_crit_cols = (criteria.end.col - criteria.start.col + 1) as usize;
        let n_crit_rows = criteria.end.row.saturating_sub(criteria.start.row);
        if n_crit_rows == 0 {
            return false;
        }
        for cr in 0..n_crit_rows {
            let crit_row = criteria.start.row + 1 + cr;
            let mut all_match = true;
            let mut any_constraint = false;
            for cc in 0..n_crit_cols {
                let crit_addr = Address::new(
                    criteria.start.sheet,
                    criteria.start.col + cc as u16,
                    crit_row,
                );
                let Some(crit_contents) = self.wb().cells.get(&crit_addr) else {
                    continue;
                };
                if matches!(crit_contents, CellContents::Empty) {
                    continue;
                }
                any_constraint = true;
                let crit_field_name_addr = Address::new(
                    criteria.start.sheet,
                    criteria.start.col + cc as u16,
                    criteria.start.row,
                );
                let Some(field_name) = self.cell_label_lower(crit_field_name_addr) else {
                    all_match = false;
                    break;
                };
                let Some(in_col_offset) = input_field_names.iter().position(|n| n == &field_name)
                else {
                    all_match = false;
                    break;
                };
                let in_addr = Address::new(
                    input.start.sheet,
                    input.start.col + in_col_offset as u16,
                    in_row,
                );
                let in_contents = self.wb().cells.get(&in_addr);
                if !cell_values_equal_for_query(crit_contents, in_contents) {
                    all_match = false;
                    break;
                }
            }
            if all_match && any_constraint {
                return true;
            }
        }
        false
    }

    /// Read a cell's label text, lowercased, for case-insensitive
    /// header / field-name matching. Returns `None` for non-label
    /// cells.
    fn cell_label_lower(&self, addr: Address) -> Option<String> {
        match self.wb().cells.get(&addr)? {
            CellContents::Label { text, .. } => Some(text.to_ascii_lowercase()),
            CellContents::Constant(Value::Text(s)) => Some(s.to_ascii_lowercase()),
            _ => None,
        }
    }

    /// `/Data Parse Format-Line Create` — read the first non-empty
    /// label below the input column's top row and emit a format
    /// line classifying each char run (digits/sign/dot → `V`,
    /// whitespace gaps stay as spaces, anything else → `L`).
    /// Writes the result as an apostrophe-prefixed label into the
    /// top of the input column. Refuses with a status-line error
    /// when no input column is set or the data row isn't a label.
    fn execute_parse_format_line_create(&mut self) {
        let Some(input) = self.data_parse.input_range else {
            self.set_error("Parse Format-Line: set Input-Column first");
            return;
        };
        let r = input.normalized();
        let sheet = r.start.sheet;
        let col = r.start.col;
        let fl_addr = Address::new(sheet, col, r.start.row);
        let mut data_text: Option<String> = None;
        for row in (r.start.row + 1)..=r.end.row {
            let addr = Address::new(sheet, col, row);
            if let Some(CellContents::Label { text, .. }) = self.wb().cells.get(&addr) {
                if !text.is_empty() {
                    data_text = Some(text.clone());
                    break;
                }
            }
        }
        let Some(text) = data_text else {
            self.set_error("Parse Format-Line: no label data row to derive from");
            return;
        };
        let fl = build_format_line(&text);
        let prev = self.wb().cells.get(&fl_addr).cloned();
        let new_cell = label_cell(&fl);
        self.wb_mut().cells.insert(fl_addr, new_cell.clone());
        self.push_to_engine_at(fl_addr, &new_cell);
        if self.undo_enabled {
            let mut prev_cells = Vec::new();
            if let Some(c) = prev {
                prev_cells.push((fl_addr, c));
            }
            if !prev_cells.is_empty() {
                self.wb_mut().journal.push(JournalEntry::RangeRestore {
                    cells: prev_cells,
                    formats: Vec::new(),
                    text_styles: Vec::new(),
                });
            }
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.enter_data_parse_menu();
    }

    /// `/Data Parse Format-Line Edit` — move the pointer to the
    /// format-line cell (top of the input column) and open it in
    /// EDIT mode. Refuses with a status-line error when no input
    /// column is set.
    fn execute_parse_format_line_edit(&mut self) {
        let Some(input) = self.data_parse.input_range else {
            self.set_error("Parse Format-Line: set Input-Column first");
            return;
        };
        let r = input.normalized();
        let fl_addr = Address::new(r.start.sheet, r.start.col, r.start.row);
        self.wb_mut().pointer = fl_addr;
        self.scroll_into_view();
        self.menu = None;
        self.begin_edit();
    }

    /// `/Data Parse Go` — split each label in `input_range` (rows
    /// 2..N; row 1 holds the format-line label) according to the
    /// fields encoded in the format line, and write the parsed
    /// fields starting at `output_anchor`. Silent no-op when the
    /// input range or output anchor is unset, when the format-line
    /// row isn't a label starting with `|`, or when the format
    /// line declares zero non-skip fields.
    fn execute_data_parse(&mut self) {
        let Some(input) = self.data_parse.input_range else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(anchor) = self.data_parse.output_anchor else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let r = input.normalized();
        let sheet = r.start.sheet;
        let col = r.start.col;
        let fl_row = r.start.row;
        let fl_addr = Address::new(sheet, col, fl_row);
        let fl_text = match self.wb().cells.get(&fl_addr) {
            Some(CellContents::Label { text, .. }) if text.starts_with('|') => text.clone(),
            _ => {
                self.set_error("Parse: top of input column must be a `|`-prefixed format line");
                return;
            }
        };
        let fields = parse_format_line(&fl_text);
        if fields.is_empty() {
            self.set_error("Parse: format line declares no fields");
            return;
        }

        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        let mut data_row_idx: u32 = 0;
        for in_row in (fl_row + 1)..=r.end.row {
            let in_addr = Address::new(sheet, col, in_row);
            let label_text = match self.wb().cells.get(&in_addr) {
                Some(CellContents::Label { text, .. }) => text.clone(),
                _ => {
                    data_row_idx += 1;
                    continue;
                }
            };
            let chars: Vec<char> = label_text.chars().collect();
            let mut out_col_idx: u16 = 0;
            for &(start, end, kind) in &fields {
                if kind == FormatField::Skip {
                    out_col_idx += 1;
                    continue;
                }
                let slice: String = chars
                    .iter()
                    .skip(start)
                    .take(end.saturating_sub(start))
                    .collect();
                let trimmed = slice.trim();
                let out_addr =
                    Address::new(sheet, anchor.col + out_col_idx, anchor.row + data_row_idx);
                if !trimmed.is_empty() {
                    if let Some(c) = self.wb().cells.get(&out_addr) {
                        prev_cells.push((out_addr, c.clone()));
                    }
                    if let Some(f) = self.wb().cell_formats.get(&out_addr) {
                        prev_formats.push((out_addr, *f));
                    }
                    if let Some(s) = self.wb().cell_text_styles.get(&out_addr) {
                        prev_text_styles.push((out_addr, *s));
                    }
                    let contents = match kind {
                        FormatField::Value => match trimmed.parse::<f64>() {
                            Ok(n) => CellContents::Constant(Value::Number(n)),
                            Err(_) => label_cell(trimmed),
                        },
                        FormatField::Label | FormatField::Date | FormatField::Time => {
                            label_cell(trimmed)
                        }
                        FormatField::Skip => unreachable!(),
                    };
                    self.wb_mut().cells.insert(out_addr, contents.clone());
                    self.push_to_engine_at(out_addr, &contents);
                }
                out_col_idx += 1;
            }
            data_row_idx += 1;
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    fn enter_data_regression_menu(&mut self) {
        self.menu = Some(MenuState::rooted_at(menu::DATA_REGRESSION_MENU));
        self.mode = Mode::Menu;
    }

    /// `/Data Regression` — univariate ordinary least-squares
    /// linear regression. Reads numeric values from the configured
    /// X and Y ranges (must be the same length), computes
    /// `y = a + b*x`, and writes a labeled output table at
    /// `output_anchor`. Silent no-op when X, Y, or output anchor
    /// is unset, or when the ranges have fewer than 2 numeric
    /// points (degrees of freedom would be non-positive).
    fn execute_data_regression(&mut self) {
        let Some(x_range) = self.data_regression.x_range else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(y_range) = self.data_regression.y_range else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };
        let Some(anchor) = self.data_regression.output_anchor else {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        };

        let xs = self.collect_numeric_column(x_range);
        let ys = self.collect_numeric_column(y_range);
        let n = xs.len().min(ys.len());
        if n < 2 {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        }
        let nf = n as f64;
        let sum_x: f64 = xs.iter().take(n).sum();
        let sum_y: f64 = ys.iter().take(n).sum();
        let mean_x = sum_x / nf;
        let mean_y = sum_y / nf;
        let mut sxx = 0.0_f64;
        let mut syy = 0.0_f64;
        let mut sxy = 0.0_f64;
        for i in 0..n {
            let dx = xs[i] - mean_x;
            let dy = ys[i] - mean_y;
            sxx += dx * dx;
            syy += dy * dy;
            sxy += dx * dy;
        }
        let force_zero = self.data_regression.intercept_zero;
        let (b, a) = if force_zero {
            let sxx_raw: f64 = xs.iter().take(n).map(|x| x * x).sum();
            let sxy_raw: f64 = (0..n).map(|i| xs[i] * ys[i]).sum();
            (sxy_raw / sxx_raw, 0.0_f64)
        } else if sxx == 0.0 {
            (0.0_f64, mean_y)
        } else {
            let b = sxy / sxx;
            (b, mean_y - b * mean_x)
        };
        let r_squared = if syy == 0.0 || sxx == 0.0 {
            1.0
        } else {
            (sxy * sxy) / (sxx * syy)
        };
        let mut rss = 0.0_f64;
        for i in 0..n {
            let pred = a + b * xs[i];
            let r = ys[i] - pred;
            rss += r * r;
        }
        let df = if force_zero {
            n - 1
        } else {
            n.saturating_sub(2)
        };
        let dff = df.max(1) as f64;
        let s_y_est = (rss / dff).sqrt();
        let se_b = if sxx > 0.0 {
            (s_y_est * s_y_est / sxx).sqrt()
        } else {
            0.0
        };

        let sheet = anchor.sheet;
        let label_col = anchor.col;
        let value_col = anchor.col + 1;
        let r0 = anchor.row;
        let writes: Vec<(Address, CellContents)> = vec![
            (
                Address::new(sheet, label_col, r0),
                label_cell("Regression Output:"),
            ),
            (
                Address::new(sheet, label_col, r0 + 2),
                label_cell("Constant"),
            ),
            (
                Address::new(sheet, value_col, r0 + 2),
                CellContents::Constant(Value::Number(a)),
            ),
            (
                Address::new(sheet, label_col, r0 + 3),
                label_cell("Std Err of Y Est"),
            ),
            (
                Address::new(sheet, value_col, r0 + 3),
                CellContents::Constant(Value::Number(s_y_est)),
            ),
            (
                Address::new(sheet, label_col, r0 + 4),
                label_cell("R Squared"),
            ),
            (
                Address::new(sheet, value_col, r0 + 4),
                CellContents::Constant(Value::Number(r_squared)),
            ),
            (
                Address::new(sheet, label_col, r0 + 5),
                label_cell("No. of Observations"),
            ),
            (
                Address::new(sheet, value_col, r0 + 5),
                CellContents::Constant(Value::Number(n as f64)),
            ),
            (
                Address::new(sheet, label_col, r0 + 6),
                label_cell("Degrees of Freedom"),
            ),
            (
                Address::new(sheet, value_col, r0 + 6),
                CellContents::Constant(Value::Number(df as f64)),
            ),
            (
                Address::new(sheet, label_col, r0 + 8),
                label_cell("X Coefficient(s)"),
            ),
            (
                Address::new(sheet, value_col, r0 + 8),
                CellContents::Constant(Value::Number(b)),
            ),
            (
                Address::new(sheet, label_col, r0 + 9),
                label_cell("Std Err of Coef."),
            ),
            (
                Address::new(sheet, value_col, r0 + 9),
                CellContents::Constant(Value::Number(se_b)),
            ),
        ];

        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for (addr, contents) in &writes {
            if let Some(c) = self.wb().cells.get(addr) {
                prev_cells.push((*addr, c.clone()));
            }
            if let Some(f) = self.wb().cell_formats.get(addr) {
                prev_formats.push((*addr, *f));
            }
            if let Some(s) = self.wb().cell_text_styles.get(addr) {
                prev_text_styles.push((*addr, *s));
            }
            self.wb_mut().cells.insert(*addr, contents.clone());
            self.push_to_engine_at(*addr, contents);
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.menu = None;
        self.mode = Mode::Ready;
    }

    /// Walk a range column-major and collect every numeric cell value.
    /// Used by /Data Regression to build its X / Y vectors.
    fn collect_numeric_column(&self, range: Range) -> Vec<f64> {
        let r = range.normalized();
        let mut out = Vec::new();
        for col in r.start.col..=r.end.col {
            for row in r.start.row..=r.end.row {
                let addr = Address::new(r.start.sheet, col, row);
                if let Some(n) = self.numeric_cell_value(addr) {
                    out.push(n);
                }
            }
        }
        out
    }

    /// `/Data Distribution` — count how many cells in `values`
    /// fall into each bin defined by ascending thresholds in
    /// `bins` (must be a single column). Writes counts to the
    /// column immediately right of the bins, plus one extra row
    /// at the bottom for the over-the-largest-bin overflow count.
    /// Multi-column bins are silently treated as their first
    /// column to match 1-2-3's "use the leftmost cell" behavior.
    /// Journals overwritten cells for Alt-F4.
    fn execute_data_distribution(&mut self, values: Range, bins: Range) {
        let v = values.normalized();
        let b = bins.normalized();
        let bin_col = b.start.col;
        let out_col = bin_col + 1;
        let bin_sheet = b.start.sheet;
        let val_sheet = v.start.sheet;

        let mut bin_thresholds: Vec<(u32, f64)> = Vec::new();
        for row in b.start.row..=b.end.row {
            let addr = Address::new(bin_sheet, bin_col, row);
            if let Some(n) = self.numeric_cell_value(addr) {
                bin_thresholds.push((row, n));
            }
        }
        if bin_thresholds.is_empty() {
            self.menu = None;
            self.mode = Mode::Ready;
            return;
        }

        let mut samples: Vec<f64> = Vec::new();
        for row in v.start.row..=v.end.row {
            for col in v.start.col..=v.end.col {
                let addr = Address::new(val_sheet, col, row);
                if let Some(n) = self.numeric_cell_value(addr) {
                    samples.push(n);
                }
            }
        }

        let mut counts: Vec<u64> = vec![0; bin_thresholds.len() + 1];
        for s in &samples {
            let mut placed = false;
            for (i, (_, thr)) in bin_thresholds.iter().enumerate() {
                if *s <= *thr {
                    counts[i] += 1;
                    placed = true;
                    break;
                }
            }
            if !placed {
                let last = counts.len() - 1;
                counts[last] += 1;
            }
        }

        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        let mut write = |app: &mut Self, addr: Address, count: u64| {
            if let Some(c) = app.wb().cells.get(&addr) {
                prev_cells.push((addr, c.clone()));
            }
            if let Some(f) = app.wb().cell_formats.get(&addr) {
                prev_formats.push((addr, *f));
            }
            if let Some(s) = app.wb().cell_text_styles.get(&addr) {
                prev_text_styles.push((addr, *s));
            }
            let s = count.to_string();
            let _ = app.wb_mut().engine.set_user_input(addr, &s);
            app.wb_mut()
                .cells
                .insert(addr, CellContents::Constant(Value::Number(count as f64)));
        };
        for (i, (row, _)) in bin_thresholds.iter().enumerate() {
            let addr = Address::new(bin_sheet, out_col, *row);
            write(self, addr, counts[i]);
        }
        let overflow_row = bin_thresholds.last().unwrap().0 + 1;
        let overflow_addr = Address::new(bin_sheet, out_col, overflow_row);
        let overflow_count = *counts.last().unwrap();
        write(self, overflow_addr, overflow_count);

        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.mode = Mode::Ready;
    }

    /// Numeric value at `addr` for /Data Distribution / Sort —
    /// reads from `cells` and unwraps cached formula values too.
    /// Returns `None` for blanks, labels, errors, and unevaluated
    /// formulas.
    fn numeric_cell_value(&self, addr: Address) -> Option<f64> {
        match self.wb().cells.get(&addr)? {
            CellContents::Constant(Value::Number(n)) => Some(*n),
            CellContents::Formula {
                cached_value: Some(Value::Number(n)),
                ..
            } => Some(*n),
            _ => None,
        }
    }

    fn start_data_fill_start_prompt(&mut self, range: Range) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter Start value:".into(),
            buffer: "0".into(),
            next: PromptNext::DataFillStart { range },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    fn start_data_fill_step_prompt(&mut self, range: Range, start: f64) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter Step value:".into(),
            buffer: "1".into(),
            next: PromptNext::DataFillStep { range, start },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    fn start_data_fill_stop_prompt(&mut self, range: Range, start: f64, step: f64) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter Stop value:".into(),
            buffer: "2047".into(),
            next: PromptNext::DataFillStop { range, start, step },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// Write `start, start+step, start+2*step, ...` into `range`
    /// column-major (down each column, then to the next column),
    /// stopping at either the end of the range or when the next
    /// computed value would cross `stop`. Captures previous cell
    /// contents into a `RangeRestore` journal entry so Alt-F4 reverts.
    fn execute_data_fill(&mut self, range: Range, start: f64, step: f64, stop: f64) {
        let r = range.normalized();
        let sheet = r.start.sheet;
        let mut prev_cells: Vec<(Address, CellContents)> = Vec::new();
        let mut prev_formats: Vec<(Address, Format)> = Vec::new();
        let mut prev_text_styles: Vec<(Address, TextStyle)> = Vec::new();
        let mut value = start;
        'outer: for col in r.start.col..=r.end.col {
            for row in r.start.row..=r.end.row {
                if step >= 0.0 && value > stop {
                    break 'outer;
                }
                if step < 0.0 && value < stop {
                    break 'outer;
                }
                let addr = Address::new(sheet, col, row);
                if let Some(c) = self.wb().cells.get(&addr) {
                    prev_cells.push((addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&addr) {
                    prev_formats.push((addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&addr) {
                    prev_text_styles.push((addr, *s));
                }
                let s = l123_core::format_number_general(value);
                let _ = self.wb_mut().engine.set_user_input(addr, &s);
                self.wb_mut()
                    .cells
                    .insert(addr, CellContents::Constant(Value::Number(value)));
                value += step;
            }
        }
        if self.undo_enabled
            && (!prev_cells.is_empty() || !prev_formats.is_empty() || !prev_text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells: prev_cells,
                formats: prev_formats,
                text_styles: prev_text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
        self.mode = Mode::Ready;
    }

    fn start_decimals_prompt(&mut self, kind: FormatKind) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter number of decimal places (0..15):".into(),
            buffer: "2".into(),
            next: PromptNext::RangeFormat { kind },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Global Format <Fixed|Sci|Currency|Comma|Percent>` —
    /// prompt for decimal places, then set the workbook's global format
    /// (no POINT step).
    fn start_global_decimals_prompt(&mut self, kind: FormatKind) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter number of decimal places (0..15):".into(),
            buffer: "2".into(),
            next: PromptNext::WorksheetGlobalFormat { kind },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    fn start_col_width_prompt(&mut self) {
        self.menu = None;
        let p = self.wb().pointer;
        let current = self.col_width_of(p.sheet, p.col);
        self.prompt = Some(PromptState {
            label: "Enter column width (1..240):".into(),
            buffer: current.to_string(),
            next: PromptNext::WorksheetColumnSetWidth,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Global Recalc Iteration` — prompt for iteration
    /// count (1..=50). Seeded with the current value so Enter-only
    /// is a no-op.
    fn start_recalc_iteration_prompt(&mut self) {
        self.menu = None;
        let current = self.recalc_iterations;
        self.prompt = Some(PromptState {
            label: "Enter iteration count (1..50):".into(),
            buffer: current.to_string(),
            next: PromptNext::WorksheetGlobalRecalcIteration,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Global Col-Width` — prompt for the new workbook-wide
    /// default column width (1..240). The prompt is seeded with the
    /// current default so Enter-only is a no-op.
    fn start_global_col_width_prompt(&mut self) {
        self.menu = None;
        let current = self.wb().default_col_width;
        self.prompt = Some(PromptState {
            label: "Enter default column width (1..240):".into(),
            buffer: current.to_string(),
            next: PromptNext::WorksheetGlobalColWidth,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Global Label <Left|Right|Center>` — change the
    /// default label prefix used for unprefixed label entries. Journals
    /// the previous prefix so Alt-F4 reverts.
    fn set_default_label_prefix(&mut self, new_prefix: LabelPrefix) {
        let prev = self.default_label_prefix;
        self.default_label_prefix = new_prefix;
        self.push_journal_batch(vec![JournalEntry::DefaultLabelPrefix { prev }]);
        self.close_menu();
    }

    /// `/Worksheet Global Format <…>` — change the workbook-wide
    /// default cell format. Journals the previous format so Alt-F4
    /// reverts. Does not touch per-cell `cell_formats` overrides.
    fn set_global_format(&mut self, new_format: Format) {
        let prev = self.wb().global_format;
        self.wb_mut().global_format = new_format;
        self.push_journal_batch(vec![JournalEntry::GlobalFormat { prev }]);
        self.wb_mut().dirty = true;
        self.close_menu();
    }

    /// `/Worksheet Global Format Other Parentheses Yes|No` — toggle the
    /// parens flag on the workbook-wide default format. Cells inheriting
    /// the global pick the new flag automatically; per-cell overrides
    /// keep their own setting.
    fn set_global_parens(&mut self, value: bool) {
        let mut next = self.wb().global_format;
        next.parens = value;
        self.set_global_format(next);
    }

    /// Open POINT mode for `/Range Format Other Color Negative <color>`
    /// (or `Reset` when `color: None`).
    fn begin_neg_color(&mut self, color: Option<RgbColor>) {
        self.begin_point(PendingCommand::RangeNegColor { color });
    }

    /// `/Worksheet Global Format Other Color Negative <color>` (or
    /// Reset). Modifies the global default's `negative_color` field.
    fn set_global_neg_color(&mut self, color: Option<RgbColor>) {
        let mut next = self.wb().global_format;
        next.negative_color = color;
        self.set_global_format(next);
    }

    /// `/Worksheet Global Default Other International <field>` — apply
    /// `mutator` to the workbook's `International` and journal a
    /// snapshot of the previous state for one-step undo.
    fn set_international(&mut self, mutator: impl FnOnce(&mut International)) {
        let prev = self.wb().international.clone();
        mutator(&mut self.wb_mut().international);
        self.push_journal_batch(vec![JournalEntry::GlobalInternational { prev }]);
        self.wb_mut().dirty = true;
        self.close_menu();
    }

    fn set_punctuation(&mut self, p: Punctuation) {
        self.set_international(|i| i.punctuation = p);
    }

    fn set_date_intl(&mut self, d: DateIntl) {
        self.set_international(|i| i.date_intl = d);
    }

    fn set_time_intl(&mut self, t: TimeIntl) {
        self.set_international(|i| i.time_intl = t);
    }

    fn set_negative_style(&mut self, n: NegativeStyle) {
        self.set_international(|i| i.negative_style = n);
    }

    /// `/Worksheet Global Default Other International Currency
    /// Prefix|Suffix` — open a string prompt seeded with the current
    /// symbol; on commit, apply both the new symbol and the chosen
    /// position.
    fn start_currency_symbol_prompt(&mut self, position: CurrencyPosition) {
        self.menu = None;
        let buffer = self.wb().international.currency.symbol.clone();
        self.prompt = Some(PromptState {
            label: "Enter currency symbol:".into(),
            buffer,
            next: PromptNext::WorksheetGlobalDefaultOtherIntlCurrencySymbol { position },
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Column Column-Range Set-Width` — prompt for the new
    /// width first; on Enter, enter POINT to pick the column range.
    fn start_col_range_width_prompt(&mut self) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: "Enter column width (1..240):".into(),
            buffer: "9".into(),
            next: PromptNext::WorksheetColumnRangeSetWidth,
            fresh: true,
        });
        self.mode = Mode::Menu;
    }

    /// `/Worksheet Column Reset-Width` — clear any width override on
    /// the current column (every target sheet when GROUP is on).
    fn execute_col_reset_width(&mut self) {
        let col = self.wb().pointer.col;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            let key = (sheet, col);
            let prev = self.wb().col_widths.get(&key).copied();
            if prev.is_some() {
                self.wb_mut().col_widths.remove(&key);
            }
            batch.push(JournalEntry::ColWidth {
                sheet,
                col,
                prev_width: prev,
            });
        }
        self.push_journal_batch(batch);
        self.wb_mut().dirty = true;
        self.close_menu();
    }

    /// Toggle the hidden flag for every column in `range` across every
    /// target sheet. `hide == true` hides (`/Worksheet Column Hide`);
    /// `hide == false` unhides (`/Worksheet Column Display`). Journal
    /// entries capture the prior state per (sheet, col) so Alt-F4 can
    /// invert the whole batch in one step.
    fn execute_col_hide_display(&mut self, range: Range, hide: bool) {
        let r = range.normalized();
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            for col in r.start.col..=r.end.col {
                let key = (sheet, col);
                let prev_hidden = self.wb().hidden_cols.contains(&key);
                if hide {
                    self.wb_mut().hidden_cols.insert(key);
                } else {
                    self.wb_mut().hidden_cols.remove(&key);
                }
                batch.push(JournalEntry::ColHidden {
                    sheet,
                    col,
                    prev_hidden,
                });
            }
        }
        self.push_journal_batch(batch);
    }

    /// Apply a width change to every column in `range` across every
    /// target sheet. `new_width == None` resets to the default; `Some(w)`
    /// sets an explicit width. Batched as a single journal entry so
    /// Alt-F4 undoes the whole range in one step.
    fn execute_col_range_width(&mut self, range: Range, new_width: Option<u8>) {
        let r = range.normalized();
        let default = self.wb().default_col_width;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            for col in r.start.col..=r.end.col {
                let key = (sheet, col);
                let prev = self.wb().col_widths.get(&key).copied();
                match new_width {
                    Some(w) if w != default => {
                        self.wb_mut().col_widths.insert(key, w);
                    }
                    _ => {
                        self.wb_mut().col_widths.remove(&key);
                    }
                }
                batch.push(JournalEntry::ColWidth {
                    sheet,
                    col,
                    prev_width: prev,
                });
            }
        }
        self.push_journal_batch(batch);
    }

    fn start_name_prompt(&mut self, label: &str, next: PromptNext) {
        self.menu = None;
        self.prompt = Some(PromptState {
            label: label.into(),
            buffer: String::new(),
            next,
            // An empty buffer has nothing to "replace" on first keystroke;
            // fresh only matters for defaults.
            fresh: false,
        });
        self.mode = Mode::Menu;
    }

    fn range_name_reset(&mut self) {
        let prev: Vec<(String, Range)> = self
            .wb()
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), *v))
            .collect();
        if !prev.is_empty() {
            for (name, _) in &prev {
                let _ = self.wb_mut().engine.delete_name(name);
            }
            self.wb_mut().named_ranges.clear();
            self.wb_mut().name_notes.clear();
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();
            self.push_journal_batch(vec![JournalEntry::RangeNameReset { prev }]);
            self.wb_mut().dirty = true;
        }
        self.close_menu();
    }

    fn range_name_note_reset(&mut self) {
        let prev: Vec<(String, String)> = self
            .wb()
            .name_notes
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect();
        if !prev.is_empty() {
            self.wb_mut().name_notes.clear();
            self.push_journal_batch(vec![JournalEntry::RangeNameNoteReset { prev }]);
            self.wb_mut().dirty = true;
        }
        self.close_menu();
    }

    fn execute_range_protection(&mut self, range: Range, unprotected: bool) {
        let r = range.normalized();
        let mut entries: Vec<(Address, bool)> = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                let was = self.wb().cell_unprotected.contains(&addr);
                if was == unprotected {
                    continue;
                }
                if unprotected {
                    self.wb_mut().cell_unprotected.insert(addr);
                } else {
                    self.wb_mut().cell_unprotected.remove(&addr);
                }
                entries.push((addr, was));
            }
        }
        if !entries.is_empty() {
            self.push_journal_batch(vec![JournalEntry::RangeProtection { entries }]);
            self.wb_mut().dirty = true;
        }
    }

    /// True when an edit to `addr` is currently refused — i.e. global
    /// protection is on and the cell is not in the unprotected set.
    /// `/Range Input` lifts the gate inside its range so the user can
    /// fill the form even while protection is active.
    fn is_cell_protected(&self, addr: Address) -> bool {
        if !self.global_protection {
            return false;
        }
        self.input_range
            .as_ref()
            .is_none_or(|r| !r.normalized().contains(addr))
            && !self.wb().cell_unprotected.contains(&addr)
    }

    fn enter_input_mode(&mut self, range: Range) {
        let r = range.normalized();
        if let Some(addr) = self.first_unprotected_in(r) {
            self.wb_mut().pointer = addr;
        }
        self.input_range = Some(r);
        self.mode = Mode::Ready;
    }

    fn exit_input_mode(&mut self) {
        self.input_range = None;
    }

    fn first_unprotected_in(&self, r: Range) -> Option<Address> {
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                if self.wb().cell_unprotected.contains(&addr) {
                    return Some(addr);
                }
            }
        }
        None
    }

    /// Find the next unprotected cell within the active input range
    /// in `(d_col, d_row)` direction from `from`. Stops at the range
    /// edge — does not wrap. Returns `None` when no unprotected cell
    /// exists in that direction.
    fn next_unprotected(&self, from: Address, d_col: i32, d_row: i32) -> Option<Address> {
        let r = self.input_range?.normalized();
        let mut cur = from;
        loop {
            cur = cur.shifted(d_col, d_row)?;
            if !r.contains(cur) {
                return None;
            }
            if self.wb().cell_unprotected.contains(&cur) {
                return Some(cur);
            }
        }
    }

    fn execute_range_value(&mut self, src: Range, dst: Address) {
        let s = src.normalized();
        let mut writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for row in s.start.row..=s.end.row {
            for col in s.start.col..=s.end.col {
                let src_addr = Address::new(s.start.sheet, col, row);
                let target = Address::new(
                    dst.sheet,
                    dst.col + (col - s.start.col),
                    dst.row + (row - s.start.row),
                );
                let new_contents = freeze_to_value(self.wb().cells.get(&src_addr).cloned());
                self.write_cell_with_undo(target, new_contents, &mut writes);
            }
        }
        self.finish_range_write(writes);
    }

    fn execute_range_trans(&mut self, src: Range, dst: Address) {
        let s = src.normalized();
        let mut writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for row in s.start.row..=s.end.row {
            for col in s.start.col..=s.end.col {
                let src_addr = Address::new(s.start.sheet, col, row);
                let dr = (col - s.start.col) as u32;
                let dc = row - s.start.row;
                let target =
                    Address::new(dst.sheet, dst.col.saturating_add(dc as u16), dst.row + dr);
                let new_contents = freeze_to_value(self.wb().cells.get(&src_addr).cloned());
                self.write_cell_with_undo(target, new_contents, &mut writes);
            }
        }
        self.finish_range_write(writes);
    }

    fn execute_range_justify(&mut self, range: Range) {
        let r = range.normalized();
        let sheet = r.start.sheet;
        let col = r.start.col;
        // Concatenate all label cells in the leftmost column.
        let mut text = String::new();
        for row in r.start.row..=r.end.row {
            let addr = Address::new(sheet, col, row);
            match self.wb().cells.get(&addr) {
                Some(CellContents::Label { text: t, .. }) => {
                    if !text.is_empty() {
                        text.push(' ');
                    }
                    text.push_str(t);
                }
                Some(CellContents::Empty) | None => break,
                _ => break,
            }
        }
        if text.is_empty() {
            return;
        }
        // Width = column width of the leftmost column.
        let width = self.col_width_of(sheet, col) as usize;
        let lines = wrap_text_to_width(&text, width.max(1));
        let max_rows = (r.end.row - r.start.row + 1) as usize;
        let to_write = lines.into_iter().take(max_rows).collect::<Vec<_>>();
        let mut writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for (i, line) in to_write.iter().enumerate() {
            let addr = Address::new(sheet, col, r.start.row + i as u32);
            self.write_cell_with_undo(
                addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: line.clone(),
                },
                &mut writes,
            );
        }
        // Clear any leftover rows in the original block.
        for i in to_write.len()..=(r.end.row - r.start.row) as usize {
            let addr = Address::new(sheet, col, r.start.row + i as u32);
            self.write_cell_with_undo(addr, CellContents::Empty, &mut writes);
        }
        self.finish_range_write(writes);
    }

    fn finish_range_write(&mut self, writes: Vec<(Address, Option<CellContents>)>) {
        if writes.is_empty() {
            return;
        }
        self.push_journal_batch(vec![JournalEntry::RangeRestore {
            cells: writes
                .into_iter()
                .map(|(addr, prev)| (addr, prev.unwrap_or(CellContents::Empty)))
                .collect(),
            formats: Vec::new(),
            text_styles: Vec::new(),
        }]);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.wb_mut().dirty = true;
    }

    fn execute_range_name_labels(&mut self, range: Range, direction: LabelDirection) {
        let r = range.normalized();
        let (dc, dr): (i32, i32) = match direction {
            LabelDirection::Right => (1, 0),
            LabelDirection::Down => (0, 1),
            LabelDirection::Left => (-1, 0),
            LabelDirection::Up => (0, -1),
        };
        let mut created: Vec<String> = Vec::new();
        let mut overwritten: Vec<(String, Range)> = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                let Some(CellContents::Label { text, .. }) = self.wb().cells.get(&addr).cloned()
                else {
                    continue;
                };
                if !is_valid_range_name(&text) {
                    continue;
                }
                let Some(target) = addr.shifted(dc, dr) else {
                    continue;
                };
                let target_range = Range {
                    start: target,
                    end: target,
                };
                let key = text.to_ascii_lowercase();
                if let Some(prior) = self.wb().named_ranges.get(&key).copied() {
                    overwritten.push((key.clone(), prior));
                    let _ = self.wb_mut().engine.delete_name(&key);
                }
                if self
                    .wb_mut()
                    .engine
                    .define_name(&text, target_range)
                    .is_ok()
                {
                    self.wb_mut().named_ranges.insert(key.clone(), target_range);
                    if !created.contains(&key) {
                        created.push(key);
                    }
                }
            }
        }
        if !created.is_empty() || !overwritten.is_empty() {
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();
            self.push_journal_batch(vec![JournalEntry::RangeNameLabels {
                created,
                overwritten,
            }]);
            self.wb_mut().dirty = true;
        }
    }

    fn execute_range_name_table(&mut self, anchor: Address) {
        let mut entries: Vec<(String, Range)> = self
            .wb()
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), *v))
            .collect();
        entries.sort_by(|a, b| a.0.cmp(&b.0));
        let mut writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for (i, (name, range)) in entries.iter().enumerate() {
            let row = anchor.row.saturating_add(i as u32);
            let name_addr = Address::new(anchor.sheet, anchor.col, row);
            let range_addr = Address::new(anchor.sheet, anchor.col.saturating_add(1), row);
            self.write_cell_with_undo(
                name_addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: name.clone(),
                },
                &mut writes,
            );
            self.write_cell_with_undo(
                range_addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: range_to_lotus_form(*range),
                },
                &mut writes,
            );
        }
        if !writes.is_empty() {
            self.push_journal_batch(vec![JournalEntry::RangeRestore {
                cells: writes
                    .into_iter()
                    .map(|(addr, prev)| (addr, prev.unwrap_or(CellContents::Empty)))
                    .collect(),
                formats: Vec::new(),
                text_styles: Vec::new(),
            }]);
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();
            self.wb_mut().dirty = true;
        }
    }

    fn execute_range_name_note_table(&mut self, anchor: Address) {
        let mut entries: Vec<(String, Range, String)> = self
            .wb()
            .named_ranges
            .iter()
            .map(|(k, v)| {
                let note = self.wb().name_notes.get(k).cloned().unwrap_or_default();
                (k.clone(), *v, note)
            })
            .collect();
        entries.sort_by(|a, b| a.0.cmp(&b.0));
        let mut writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for (i, (name, range, note)) in entries.iter().enumerate() {
            let row = anchor.row.saturating_add(i as u32);
            let name_addr = Address::new(anchor.sheet, anchor.col, row);
            let range_addr = Address::new(anchor.sheet, anchor.col.saturating_add(1), row);
            let note_addr = Address::new(anchor.sheet, anchor.col.saturating_add(2), row);
            self.write_cell_with_undo(
                name_addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: name.clone(),
                },
                &mut writes,
            );
            self.write_cell_with_undo(
                range_addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: range_to_lotus_form(*range),
                },
                &mut writes,
            );
            self.write_cell_with_undo(
                note_addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: note.clone(),
                },
                &mut writes,
            );
        }
        if !writes.is_empty() {
            self.push_journal_batch(vec![JournalEntry::RangeRestore {
                cells: writes
                    .into_iter()
                    .map(|(addr, prev)| (addr, prev.unwrap_or(CellContents::Empty)))
                    .collect(),
                formats: Vec::new(),
                text_styles: Vec::new(),
            }]);
            self.wb_mut().engine.recalc();
            self.refresh_formula_caches();
            self.wb_mut().dirty = true;
        }
    }

    fn write_cell_with_undo(
        &mut self,
        addr: Address,
        contents: CellContents,
        writes: &mut Vec<(Address, Option<CellContents>)>,
    ) {
        let prev = self.wb().cells.get(&addr).cloned();
        writes.push((addr, prev));
        self.wb_mut().cells.insert(addr, contents.clone());
        self.push_to_engine_at(addr, &contents);
    }

    fn restore_cell_contents(&mut self, addr: Address, prev: Option<CellContents>) {
        match prev {
            Some(c) => {
                self.wb_mut().cells.insert(addr, c.clone());
                self.push_to_engine_at(addr, &c);
            }
            None => {
                self.wb_mut().cells.remove(&addr);
                let _ = self.wb_mut().engine.clear_cell(addr);
            }
        }
    }

    fn execute_range_name_undefine(&mut self, name: &str) {
        let key = name.to_ascii_lowercase();
        let Some(range) = self.wb().named_ranges.get(&key).copied() else {
            return;
        };
        let prior_note = self.wb().name_notes.get(&key).cloned();
        let literal = range_to_lotus_form(range);
        let cell_addrs: Vec<Address> = self
            .wb()
            .cells
            .iter()
            .filter_map(|(addr, c)| match c {
                CellContents::Formula { expr, .. } if formula_uses_name(expr, &key) => Some(*addr),
                _ => None,
            })
            .collect();
        let mut cell_writes: Vec<(Address, Option<CellContents>)> = Vec::new();
        for addr in cell_addrs {
            let Some(CellContents::Formula { expr, .. }) = self.wb().cells.get(&addr).cloned()
            else {
                continue;
            };
            let new_expr = replace_name_in_formula(&expr, &key, &literal);
            let new_contents = CellContents::Formula {
                expr: new_expr,
                cached_value: None,
            };
            let prev = self.wb().cells.get(&addr).cloned();
            cell_writes.push((addr, prev));
            self.wb_mut().cells.insert(addr, new_contents.clone());
            self.push_to_engine_at(addr, &new_contents);
        }
        let _ = self.wb_mut().engine.delete_name(&key);
        self.wb_mut().named_ranges.remove(&key);
        self.wb_mut().name_notes.remove(&key);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.push_journal_batch(vec![JournalEntry::RangeNameUndefine {
            name: key,
            range,
            note: prior_note,
            cell_writes,
        }]);
        self.wb_mut().dirty = true;
    }

    fn set_range_name_note(&mut self, name: &str, note: String) {
        let key = name.to_ascii_lowercase();
        if !self.wb().named_ranges.contains_key(&key) {
            return;
        }
        let prev = self.wb().name_notes.get(&key).cloned();
        if note.is_empty() {
            self.wb_mut().name_notes.remove(&key);
        } else {
            self.wb_mut().name_notes.insert(key.clone(), note);
        }
        self.push_journal_batch(vec![JournalEntry::RangeNameNote { name: key, prev }]);
        self.wb_mut().dirty = true;
    }

    fn delete_range_name_note(&mut self, name: &str) {
        let key = name.to_ascii_lowercase();
        let Some(prev) = self.wb_mut().name_notes.remove(&key) else {
            return;
        };
        self.push_journal_batch(vec![JournalEntry::RangeNameNote {
            name: key,
            prev: Some(prev),
        }]);
        self.wb_mut().dirty = true;
    }

    /// Snapshot the workbook's defined range names into the F3 NAMES
    /// overlay, keyed by ascii-lowercase ordering. Underlying state
    /// (POINT or prompt) is left untouched so dismissal returns to it.
    fn open_name_list(&mut self, origin: NameListOrigin) {
        let mut entries: Vec<(String, Range)> = self
            .wb()
            .named_ranges
            .iter()
            .map(|(k, v)| (k.clone(), *v))
            .collect();
        entries.sort_by(|a, b| a.0.cmp(&b.0));
        self.name_list = Some(NameListState {
            entries,
            highlight: 0,
            view_offset: 0,
            origin,
        });
        self.mode = Mode::Names;
    }

    /// F3 from a command-argument prompt. Only opens the picker for
    /// prompts where a range-name selection is meaningful (GOTO and
    /// `/Range Name Delete`); other prompts ignore F3.
    fn open_name_list_from_prompt(&mut self) {
        let origin = match self.prompt.as_ref().map(|p| p.next) {
            Some(PromptNext::Goto) => NameListOrigin::Goto,
            Some(
                PromptNext::RangeNameDelete
                | PromptNext::RangeNameUndefine
                | PromptNext::RangeNameNoteCreate
                | PromptNext::RangeNameNoteDelete,
            ) => NameListOrigin::PromptName,
            _ => return,
        };
        self.open_name_list(origin);
    }

    /// Width of `col` in `sheet`. Returns the per-column override if
    /// one is set, otherwise the workbook's global default width.
    fn col_width_of(&self, sheet: SheetId, col: u16) -> u8 {
        self.wb()
            .col_widths
            .get(&(sheet, col))
            .copied()
            .unwrap_or(self.wb().default_col_width)
    }

    fn cancel_prompt(&mut self) {
        // Esc on a macro-driven prompt cancels the whole macro: the
        // user has bailed out of the input the macro asked for, so
        // resuming would be wrong. Match Lotus's "Esc aborts macro"
        // convention.
        let was_macro = matches!(
            self.prompt.as_ref().map(|p| p.next),
            Some(PromptNext::MacroGetInput { .. })
        );
        self.prompt = None;
        self.mode = Mode::Ready;
        if was_macro {
            self.pending_macro_input_loc = None;
            self.macro_state = None;
        }
    }

    fn commit_prompt(&mut self) {
        let Some(p) = self.prompt.take() else {
            self.mode = Mode::Ready;
            return;
        };
        match p.next {
            PromptNext::RangeFormat { kind } => {
                let decimals: u8 = p.buffer.parse().unwrap_or(2);
                let decimals = decimals.min(15);
                let format = Format {
                    kind,
                    decimals,
                    parens: false,
                    negative_color: None,
                };
                self.begin_point(PendingCommand::RangeFormat { format });
            }
            PromptNext::WorksheetGlobalFormat { kind } => {
                let decimals: u8 = p.buffer.parse().unwrap_or(2);
                let decimals = decimals.min(15);
                self.set_global_format(Format {
                    kind,
                    decimals,
                    parens: false,
                    negative_color: None,
                });
            }
            PromptNext::WorksheetGlobalDefaultOtherIntlCurrencySymbol { position } => {
                let symbol = p.buffer;
                self.set_international(|i| {
                    i.currency.symbol = symbol;
                    i.currency.position = position;
                });
            }
            PromptNext::WorksheetColumnSetWidth => {
                let default = self.wb().default_col_width;
                let width: u8 = p.buffer.parse().unwrap_or(default).clamp(1, 240);
                let col = self.wb().pointer.col;
                let mut batch: Vec<JournalEntry> = Vec::new();
                for sheet in self.target_sheets() {
                    let key = (sheet, col);
                    let prev = self.wb().col_widths.get(&key).copied();
                    if width == default {
                        self.wb_mut().col_widths.remove(&key);
                    } else {
                        self.wb_mut().col_widths.insert(key, width);
                    }
                    batch.push(JournalEntry::ColWidth {
                        sheet,
                        col,
                        prev_width: prev,
                    });
                }
                self.push_journal_batch(batch);
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PromptNext::WorksheetColumnRangeSetWidth => {
                let default = self.wb().default_col_width;
                let width: u8 = p.buffer.parse().unwrap_or(default).clamp(1, 240);
                self.begin_point(PendingCommand::ColumnRangeSetWidth { width });
            }
            PromptNext::WorksheetGlobalColWidth => {
                let default = self.wb().default_col_width;
                let width: u8 = p.buffer.parse().unwrap_or(default).clamp(1, 240);
                let prev = default;
                self.wb_mut().default_col_width = width;
                self.push_journal_batch(vec![JournalEntry::GlobalColWidth { prev }]);
                self.wb_mut().dirty = true;
                self.mode = Mode::Ready;
            }
            PromptNext::WorksheetGlobalRecalcIteration => {
                let current = self.recalc_iterations;
                let n: u16 = p.buffer.parse().unwrap_or(current).clamp(1, 50);
                self.recalc_iterations = n;
                self.mode = Mode::Ready;
            }
            PromptNext::DataFillStart { range } => {
                let start: f64 = p.buffer.parse().unwrap_or(0.0);
                self.start_data_fill_step_prompt(range, start);
            }
            PromptNext::DataFillStep { range, start } => {
                let step: f64 = p.buffer.parse().unwrap_or(1.0);
                self.start_data_fill_stop_prompt(range, start, step);
            }
            PromptNext::DataFillStop { range, start, step } => {
                let stop: f64 = p.buffer.parse().unwrap_or(2047.0);
                self.execute_data_fill(range, start, step, stop);
            }
            PromptNext::RangeNameCreate => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                self.pending_name = Some(p.buffer);
                self.begin_point(PendingCommand::RangeNameCreate);
            }
            PromptNext::RangeNameDelete => {
                if !p.buffer.is_empty() {
                    let _ = self.wb_mut().engine.delete_name(&p.buffer);
                    self.wb_mut()
                        .named_ranges
                        .remove(&p.buffer.to_ascii_lowercase());
                    self.wb_mut()
                        .name_notes
                        .remove(&p.buffer.to_ascii_lowercase());
                    self.wb_mut().engine.recalc();
                    self.refresh_formula_caches();
                    self.wb_mut().dirty = true;
                }
                self.mode = Mode::Ready;
            }
            PromptNext::RangeNameUndefine => {
                if !p.buffer.is_empty() {
                    let name = p.buffer.clone();
                    self.execute_range_name_undefine(&name);
                }
                self.mode = Mode::Ready;
            }
            PromptNext::RangeNameNoteCreate => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let key = p.buffer.to_ascii_lowercase();
                if !self.wb().named_ranges.contains_key(&key) {
                    self.mode = Mode::Ready;
                    return;
                }
                self.pending_name = Some(p.buffer);
                self.prompt = Some(PromptState {
                    label: "Enter note text:".into(),
                    buffer: String::new(),
                    next: PromptNext::RangeNameNoteCreateBody,
                    fresh: false,
                });
                self.mode = Mode::Menu;
            }
            PromptNext::RangeNameNoteCreateBody => {
                if let Some(name) = self.pending_name.take() {
                    self.set_range_name_note(&name, p.buffer);
                }
                self.mode = Mode::Ready;
            }
            PromptNext::RangeNameNoteDelete => {
                if !p.buffer.is_empty() {
                    let name = p.buffer.clone();
                    self.delete_range_name_note(&name);
                }
                self.mode = Mode::Ready;
            }
            PromptNext::FileSaveFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = resolve_save_path(&p.buffer);
                if path.exists() {
                    // Default highlight = Cancel, matching 1-2-3's
                    // "safe if you Enter by accident" convention.
                    self.save_confirm = Some(SaveConfirmState { path, highlight: 0 });
                    self.mode = Mode::Menu;
                } else {
                    self.queue_file_save(path);
                }
            }
            PromptNext::GraphSaveFilename => {
                let buf = p.buffer.clone();
                self.commit_graph_save(&buf);
            }
            PromptNext::Goto => {
                if let Ok(addr) = Address::parse(&p.buffer) {
                    self.move_pointer_to(addr);
                }
                self.mode = Mode::Ready;
            }
            PromptNext::FileRetrieveFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                let name = display_basename(&path);
                self.queue_async_op("Loading", name, QueuedOp::FileRetrieve { path });
            }
            PromptNext::FileXtractFilename { kind } => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = resolve_save_path(&p.buffer);
                self.pending_xtract_path = Some(path);
                self.begin_point(PendingCommand::FileXtractRange { kind });
            }
            PromptNext::FileImportNumbersFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.queue_file_import(path, /* numeric_split = */ true);
            }
            PromptNext::FileImportTextFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.queue_file_import(path, /* numeric_split = */ false);
            }
            PromptNext::FileEraseFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.erase_confirm = Some(EraseConfirmState { path, highlight: 0 });
                self.mode = Mode::Menu;
            }
            PromptNext::FileCombineFilename { kind, entire } => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                if entire {
                    self.combine_from(path, kind, None);
                } else {
                    self.pending_combine_path = Some(path);
                    self.start_file_combine_range_prompt(kind);
                }
            }
            PromptNext::FileCombineRange { kind } => {
                let Some(path) = self.pending_combine_path.take() else {
                    self.mode = Mode::Ready;
                    return;
                };
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                match Range::parse(&p.buffer) {
                    Ok(range) => self.combine_from(path, kind, Some(range)),
                    Err(e) => self.set_error(format!("Bad range {:?}: {e}", p.buffer)),
                }
            }
            PromptNext::FileDirPath => {
                if !p.buffer.is_empty() {
                    let _ = std::env::set_current_dir(PathBuf::from(&p.buffer));
                }
                self.mode = Mode::Ready;
            }
            PromptNext::FileOpenFilename { before } => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.open_file_alongside(path, before);
            }
            PromptNext::PrintFileFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.print = Some(PrintSession::new_file(path));
                self.enter_print_file_menu();
            }
            PromptNext::PrintEncodedFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.print = Some(PrintSession::new_encoded(path));
                self.enter_print_file_menu();
            }
            PromptNext::PrintFileHeader => {
                if let Some(s) = self.print.as_mut() {
                    s.header = p.buffer;
                }
                self.enter_print_options_menu();
            }
            PromptNext::PrintFileFooter => {
                if let Some(s) = self.print.as_mut() {
                    s.footer = p.buffer;
                }
                self.enter_print_options_menu();
            }
            PromptNext::PrintFileSetup => {
                if let Some(s) = self.print.as_mut() {
                    s.setup_string = p.buffer;
                }
                self.enter_print_options_menu();
            }
            PromptNext::RangeSearchString { scope, range } => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                self.search = Some(SearchSession {
                    scope,
                    range,
                    search: p.buffer,
                    matches: Vec::new(),
                    cursor: 0,
                });
                self.enter_range_search_find_replace_menu();
            }
            PromptNext::RangeSearchReplacement => {
                let Some(session) = self.search.take() else {
                    self.mode = Mode::Ready;
                    return;
                };
                self.execute_range_search_replace(session, p.buffer);
                self.mode = Mode::Ready;
            }
            PromptNext::PrintFileMarginLeft
            | PromptNext::PrintFileMarginRight
            | PromptNext::PrintFileMarginTop
            | PromptNext::PrintFileMarginBottom => {
                let v: u16 = p.buffer.parse::<u16>().unwrap_or(0).min(1000);
                if let Some(s) = self.print.as_mut() {
                    match p.next {
                        PromptNext::PrintFileMarginLeft => s.margin_left = v,
                        PromptNext::PrintFileMarginRight => s.margin_right = v,
                        PromptNext::PrintFileMarginTop => s.margin_top = v,
                        PromptNext::PrintFileMarginBottom => s.margin_bottom = v,
                        _ => {}
                    }
                }
                self.enter_print_margins_menu();
            }
            PromptNext::PrintFilePgLength => {
                let v: u16 = p.buffer.parse::<u16>().unwrap_or(0).min(1000);
                if let Some(s) = self.print.as_mut() {
                    s.pg_length = v;
                }
                self.enter_print_options_menu();
            }
            PromptNext::PrintSessionOptionsAdvancedDevice => {
                if let Some(s) = self.print.as_mut() {
                    s.lp_destination = p.buffer;
                }
                self.enter_print_advanced_menu();
            }
            PromptNext::WgdDir => {
                self.defaults.default_dir = p.buffer.trim().to_string();
                self.mode = Mode::Ready;
            }
            PromptNext::WgdTemp => {
                self.defaults.temp_dir = p.buffer.trim().to_string();
                self.mode = Mode::Ready;
            }
            PromptNext::WgdExtSave => {
                self.defaults.ext_save = p.buffer.trim().trim_start_matches('.').to_string();
                self.mode = Mode::Ready;
            }
            PromptNext::WgdExtList => {
                self.defaults.ext_list = p.buffer.trim().trim_start_matches('.').to_string();
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterInterface => {
                let prev = self.defaults.printer_interface;
                let n: u8 = p.buffer.parse().unwrap_or(prev).clamp(1, 9);
                self.defaults.printer_interface = n;
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterMarginLeft => {
                self.defaults.printer_left = parse_margin(&p.buffer, self.defaults.printer_left);
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterMarginRight => {
                self.defaults.printer_right = parse_margin(&p.buffer, self.defaults.printer_right);
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterMarginTop => {
                self.defaults.printer_top = parse_margin(&p.buffer, self.defaults.printer_top);
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterMarginBottom => {
                self.defaults.printer_bottom =
                    parse_margin(&p.buffer, self.defaults.printer_bottom);
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterPgLength => {
                let prev = self.defaults.printer_pg_length;
                let n: u16 = p.buffer.parse::<u16>().unwrap_or(prev).clamp(1, 1000);
                self.defaults.printer_pg_length = n;
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterSetup => {
                self.defaults.printer_setup = p.buffer;
                self.mode = Mode::Ready;
            }
            PromptNext::WgdPrinterName => {
                self.defaults.printer_name = p.buffer;
                self.mode = Mode::Ready;
            }
            PromptNext::MacroGetInput { numeric } => {
                let buf = p.buffer;
                let loc = self.pending_macro_input_loc.take().unwrap_or_default();
                if !loc.is_empty() {
                    let expr = if numeric {
                        // Numeric: pass through as-is so the source
                        // parser tries to make a number out of it;
                        // a non-numeric reply will fall back to a
                        // label, matching how Lotus' lenient {GETNUMBER}
                        // handler stores garbage as text.
                        buf
                    } else {
                        // Force-as-label: prepend the apostrophe so
                        // the source parser recognizes a leading
                        // value-starter (`+`, digit, ...) as part
                        // of a label rather than a value.
                        format!("'{buf}")
                    };
                    self.execute_macro_let(&loc, &expr);
                }
                if let Some(s) = self.macro_state.as_mut() {
                    s.suspend = None;
                }
                self.mode = Mode::Ready;
                self.pump_macro();
            }
        }
    }

    /// Replace every occurrence of `session.search` within `session.range`
    /// with `replacement`. Formulas use expr-string substring; labels
    /// use text substring. Both updates journaled as CellEdits.
    fn execute_range_search_replace(&mut self, session: SearchSession, replacement: String) {
        let matches = self.find_matches(&session);
        if matches.is_empty() {
            return;
        }
        let needle = session.search;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for addr in matches {
            let Some(contents) = self.wb().cells.get(&addr).cloned() else {
                continue;
            };
            let (new_contents, prev_contents, prev_format) = match contents {
                CellContents::Formula { expr, cached_value } => {
                    let new_expr = expr.replace(&needle, &replacement);
                    let prev = CellContents::Formula {
                        expr: expr.clone(),
                        cached_value: cached_value.clone(),
                    };
                    (
                        CellContents::Formula {
                            expr: new_expr,
                            cached_value: None,
                        },
                        Some(prev),
                        self.wb().cell_formats.get(&addr).copied(),
                    )
                }
                CellContents::Label { prefix, text } => {
                    let new_text = text.replace(&needle, &replacement);
                    let prev = CellContents::Label {
                        prefix,
                        text: text.clone(),
                    };
                    (
                        CellContents::Label {
                            prefix,
                            text: new_text,
                        },
                        Some(prev),
                        self.wb().cell_formats.get(&addr).copied(),
                    )
                }
                _ => continue,
            };
            batch.push(JournalEntry::CellEdit {
                addr,
                prev_contents,
                prev_format,
            });
            self.push_to_engine_at(addr, &new_contents);
            self.wb_mut().cells.insert(addr, new_contents);
        }
        self.push_journal_batch(batch);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
    }

    // ---------------- range-format execution ----------------

    fn execute_range_format(&mut self, range: Range, format: Format) {
        let r = range.normalized();
        // GROUP mode: broadcast to every sheet in the active file.
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<Format>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    prior.push((addr, self.wb().cell_formats.get(&addr).copied()));
                    if matches!(format.kind, FormatKind::Reset) {
                        self.wb_mut().clear_cell_format(addr);
                    } else {
                        self.wb_mut().set_cell_format(addr, format);
                    }
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeFormat { entries: prior });
        }
        // No recalc needed — format is presentation only.
    }

    /// `/Range Format Other Parentheses Yes|No` — toggle the parens
    /// flag on each cell's effective format. Cells without a per-cell
    /// override start from the global default; the modified format is
    /// then stored as a per-cell override (we don't try to clear back
    /// to the global if the result happens to equal it — keeps the
    /// model simple and predictable).
    fn execute_range_parens(&mut self, range: Range, value: bool) {
        let r = range.normalized();
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<Format>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    let prev = self.wb().cell_formats.get(&addr).copied();
                    prior.push((addr, prev));
                    let mut next = prev.unwrap_or(self.wb().global_format);
                    next.parens = value;
                    self.wb_mut().set_cell_format(addr, next);
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeFormat { entries: prior });
        }
    }

    /// `/Range Format Other Color Negative <color>` (or Reset, with
    /// `color: None`) — set the per-cell format's negative-color
    /// override, materializing the global default first if the cell
    /// has no per-cell format yet.
    fn execute_range_neg_color(&mut self, range: Range, color: Option<RgbColor>) {
        let r = range.normalized();
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<Format>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    let prev = self.wb().cell_formats.get(&addr).copied();
                    prior.push((addr, prev));
                    let mut next = prev.unwrap_or(self.wb().global_format);
                    next.negative_color = color;
                    self.wb_mut().set_cell_format(addr, next);
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeFormat { entries: prior });
        }
    }

    fn execute_range_color(&mut self, range: Range, target: ColorTarget, color: Option<RgbColor>) {
        let r = range.normalized();
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<Fill>, Option<FontStyle>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    let prev_fill = self.wb().cell_fills.get(&addr).copied();
                    let prev_font = self.wb().cell_font_styles.get(&addr).copied();
                    prior.push((addr, prev_fill, prev_font));
                    if matches!(target, ColorTarget::Background | ColorTarget::Both) {
                        let new_fill = match color {
                            Some(rgb) => Fill::solid(rgb),
                            None => Fill::DEFAULT,
                        };
                        if new_fill.is_default() {
                            self.wb_mut().cell_fills.remove(&addr);
                        } else {
                            self.wb_mut().cell_fills.insert(addr, new_fill);
                        }
                    }
                    if matches!(target, ColorTarget::Text | ColorTarget::Both) {
                        let mut next = prev_font.unwrap_or_default();
                        next.color = color;
                        if next.is_default() {
                            self.wb_mut().cell_font_styles.remove(&addr);
                        } else {
                            self.wb_mut().cell_font_styles.insert(addr, next);
                        }
                    }
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeColor { entries: prior });
        }
    }

    fn execute_range_alignment(&mut self, range: Range, halign: HAlign) {
        let r = range.normalized();
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<Alignment>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    let prev = self.wb().cell_alignments.get(&addr).copied();
                    prior.push((addr, prev));
                    let mut next = prev.unwrap_or_default();
                    next.horizontal = halign;
                    if next.is_default() {
                        self.wb_mut().cell_alignments.remove(&addr);
                    } else {
                        self.wb_mut().cell_alignments.insert(addr, next);
                    }
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeAlignment { entries: prior });
        }
    }

    fn execute_range_text_style(&mut self, range: Range, bits: TextStyle, set: bool) {
        let r = range.normalized();
        // GROUP mode broadcasts the style change to every sheet, matching
        // the existing `/Range Format` behavior.
        let sheets: Vec<SheetId> = if self.group_mode {
            (0..self.wb().engine.sheet_count()).map(SheetId).collect()
        } else {
            vec![r.start.sheet]
        };
        let mut prior: Vec<(Address, Option<TextStyle>)> = Vec::new();
        for sheet in &sheets {
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    let addr = Address::new(*sheet, col, row);
                    let prev = self.wb().cell_text_styles.get(&addr).copied();
                    prior.push((addr, prev));
                    let current = prev.unwrap_or_default();
                    let next = if set {
                        current.merge(bits)
                    } else {
                        current.without(bits)
                    };
                    if next.is_empty() {
                        self.wb_mut().cell_text_styles.remove(&addr);
                    } else {
                        self.wb_mut().cell_text_styles.insert(addr, next);
                    }
                }
            }
        }
        if self.undo_enabled && !prior.is_empty() {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeTextStyle { entries: prior });
        }
    }

    fn execute_range_label(&mut self, range: Range, new_prefix: LabelPrefix) {
        let r = range.normalized();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                if let Some(CellContents::Label { prefix, .. }) = self.wb_mut().cells.get_mut(&addr)
                {
                    *prefix = new_prefix;
                }
            }
        }
        // No engine push: label prefix is a display-layer property the
        // engine doesn't track. No recalc needed.
    }

    /// Transition to the next POINT step of a two-step command. Pointer
    /// returns to the source's top-left so the user can anchor and navigate
    /// to the destination.
    fn transition_point(&mut self, next: PendingCommand) {
        let source_tl = match next {
            PendingCommand::CopyTo { source } | PendingCommand::MoveTo { source } => source.start,
            PendingCommand::SpecialCopyTo { source } | PendingCommand::SpecialMoveTo { source } => {
                source.start
            }
            // /Data Distribution: spring back to the values-range
            // top-left so the user navigates from a familiar landmark
            // to the bin column.
            PendingCommand::DataDistributionBins { values } => values.start,
            PendingCommand::RangeValueTo { src } | PendingCommand::RangeTransTo { src } => {
                src.start
            }
            _ => self.wb_mut().pointer,
        };
        self.wb_mut().pointer = source_tl;
        self.scroll_into_view();
        // Copy/Move TO and DataDistribution Bins start with no anchor
        // so the user can freely navigate to the destination/bin
        // location. Pressing `.` anchors for a multi-cell extent
        // (the standard 1-2-3 POINT muscle memory).
        let anchor = match next {
            PendingCommand::CopyTo { .. }
            | PendingCommand::MoveTo { .. }
            | PendingCommand::SpecialCopyTo { .. }
            | PendingCommand::SpecialMoveTo { .. }
            | PendingCommand::DataDistributionBins { .. }
            | PendingCommand::RangeValueTo { .. }
            | PendingCommand::RangeTransTo { .. } => None,
            _ => Some(self.wb().pointer),
        };
        self.point = Some(PointState {
            anchor,
            pending: next,
            typed: String::new(),
        });
        self.mode = Mode::Point;
    }

    /// Apply the Lotus-tutorial /Copy dimension matrix:
    /// - single source × any-size dest → replicate the cell into every
    ///   dest position (including across all dest sheets for 3D dest)
    /// - multi source × single-cell dest → paste source block at dest's
    ///   top-left
    /// - source and dest same dimensions → cell-for-cell paste at
    ///   dest.start
    /// - both multi-cell with different dimensions → predictable error,
    ///   no mutation
    ///
    /// Returns `false` on the dimension-mismatch error path so the
    /// caller can preserve the error mode instead of bouncing back to
    /// READY.
    fn execute_copy(&mut self, source: Range, dest_range: Range) -> bool {
        let src = source.normalized();
        let dest = dest_range.normalized();
        let anchors = match copy_paste_anchors(src, dest) {
            Ok(a) => a,
            Err(msg) => {
                self.set_error(msg);
                return false;
            }
        };
        let src_cells = self.collect_cells_in_range(src);
        for anchor in anchors {
            self.write_cells_at_offset(&src_cells, src.start, anchor);
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        true
    }

    /// Same dim-rule contract as [`execute_copy`] but never replicates
    /// a single source — /Move is always 1:1 in Lotus.
    fn execute_move(&mut self, source: Range, dest_range: Range) -> bool {
        let src = source.normalized();
        let dest = dest_range.normalized();
        let src_cols = u32::from(src.end.col - src.start.col + 1);
        let src_rows = src.end.row - src.start.row + 1;
        let dst_cols = u32::from(dest.end.col - dest.start.col + 1);
        let dst_rows = dest.end.row - dest.start.row + 1;
        let same_size = src_cols == dst_cols && src_rows == dst_rows;
        let single_dest = dst_cols == 1 && dst_rows == 1;
        if !same_size && !single_dest {
            self.set_error("Move: source and destination ranges have different sizes");
            return false;
        }
        let src_cells = self.collect_cells_in_range(src);
        let dest_anchor = Address::new(dest.start.sheet, dest.start.col, dest.start.row);
        self.write_cells_at_offset(&src_cells, src.start, dest_anchor);
        let dest_block = Range {
            start: dest_anchor,
            end: Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (src.end.col - src.start.col),
                dest_anchor.row + (src.end.row - src.start.row),
            ),
        };
        for (s, _) in &src_cells {
            if !dest_block.contains(*s) {
                self.wb_mut().cells.remove(s);
                let _ = self.wb_mut().engine.clear_cell(*s);
            }
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        true
    }

    /// `:Special Copy` — replicate every formatting attribute (number
    /// format, text style, alignment, fill, font color/size/strike,
    /// borders) from each source cell onto the corresponding
    /// destination cell. Cell contents are untouched. Reuses
    /// [`copy_paste_anchors`] so the dim-rule matrix matches `/Copy`.
    fn execute_special_copy(&mut self, source: Range, dest_range: Range) -> bool {
        let src = source.normalized();
        let dest = dest_range.normalized();
        let anchors = match copy_paste_anchors(src, dest) {
            Ok(a) => a,
            Err(msg) => {
                self.set_error(msg);
                return false;
            }
        };
        let snap = self.collect_format_snapshots(src);
        for anchor in anchors {
            self.write_format_snapshots_at_offset(&snap, src.start, anchor);
        }
        true
    }

    /// `:Special Move` — like `execute_special_copy`, but after writing,
    /// every source cell outside the destination block has its
    /// formatting cleared. Cell contents are untouched.
    fn execute_special_move(&mut self, source: Range, dest_range: Range) -> bool {
        let src = source.normalized();
        let dest = dest_range.normalized();
        let src_cols = u32::from(src.end.col - src.start.col + 1);
        let src_rows = src.end.row - src.start.row + 1;
        let dst_cols = u32::from(dest.end.col - dest.start.col + 1);
        let dst_rows = dest.end.row - dest.start.row + 1;
        let same_size = src_cols == dst_cols && src_rows == dst_rows;
        let single_dest = dst_cols == 1 && dst_rows == 1;
        if !same_size && !single_dest {
            self.set_error("Move: source and destination ranges have different sizes");
            return false;
        }
        let snap = self.collect_format_snapshots(src);
        let dest_anchor = Address::new(dest.start.sheet, dest.start.col, dest.start.row);
        self.write_format_snapshots_at_offset(&snap, src.start, dest_anchor);
        let dest_block = Range {
            start: dest_anchor,
            end: Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (src.end.col - src.start.col),
                dest_anchor.row + (src.end.row - src.start.row),
            ),
        };
        for fs in &snap {
            if !dest_block.contains(fs.addr) {
                self.clear_formatting_at(fs.addr);
            }
        }
        true
    }

    fn collect_format_snapshots(&self, range: Range) -> Vec<FormatSnapshot> {
        let r = range.normalized();
        let mut out = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                out.push(FormatSnapshot {
                    addr,
                    format: self.wb().cell_formats.get(&addr).copied(),
                    text_style: self.wb().cell_text_styles.get(&addr).copied(),
                    alignment: self.wb().cell_alignments.get(&addr).copied(),
                    fill: self.wb().cell_fills.get(&addr).copied(),
                    font_style: self.wb().cell_font_styles.get(&addr).copied(),
                    border: self.wb().cell_borders.get(&addr).copied(),
                });
            }
        }
        out
    }

    fn write_format_snapshots_at_offset(
        &mut self,
        snaps: &[FormatSnapshot],
        src_origin: Address,
        dest_anchor: Address,
    ) {
        for fs in snaps {
            let dst = Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (fs.addr.col - src_origin.col),
                dest_anchor.row + (fs.addr.row - src_origin.row),
            );
            match fs.format {
                Some(v) => {
                    self.wb_mut().cell_formats.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_formats.remove(&dst);
                }
            }
            match fs.text_style {
                Some(v) => {
                    self.wb_mut().cell_text_styles.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_text_styles.remove(&dst);
                }
            }
            match fs.alignment {
                Some(v) => {
                    self.wb_mut().cell_alignments.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_alignments.remove(&dst);
                }
            }
            match fs.fill {
                Some(v) => {
                    self.wb_mut().cell_fills.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_fills.remove(&dst);
                }
            }
            match fs.font_style {
                Some(v) => {
                    self.wb_mut().cell_font_styles.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_font_styles.remove(&dst);
                }
            }
            match fs.border {
                Some(v) => {
                    self.wb_mut().cell_borders.insert(dst, v);
                }
                None => {
                    self.wb_mut().cell_borders.remove(&dst);
                }
            }
        }
    }

    fn clear_formatting_at(&mut self, addr: Address) {
        self.wb_mut().cell_formats.remove(&addr);
        self.wb_mut().cell_text_styles.remove(&addr);
        self.wb_mut().cell_alignments.remove(&addr);
        self.wb_mut().cell_fills.remove(&addr);
        self.wb_mut().cell_font_styles.remove(&addr);
        self.wb_mut().cell_borders.remove(&addr);
    }

    fn collect_cells_in_range(&self, range: Range) -> Vec<(Address, CellContents)> {
        let r = range.normalized();
        let mut out = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                if let Some(c) = self.wb().cells.get(&addr) {
                    out.push((addr, c.clone()));
                }
            }
        }
        out
    }

    /// Write cells to their new positions, offset so `src_origin` maps
    /// to `dest_anchor`. Formula references in copied cells shift by
    /// the same `(dx, dy)` so the new formulas refer to cells in the
    /// same relative positions as the originals.
    fn write_cells_at_offset(
        &mut self,
        cells: &[(Address, CellContents)],
        src_origin: Address,
        dest_anchor: Address,
    ) {
        let dx = dest_anchor.col as i32 - src_origin.col as i32;
        let dy = dest_anchor.row as i32 - src_origin.row as i32;
        for (src, contents) in cells {
            let dst = Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (src.col - src_origin.col),
                dest_anchor.row + (src.row - src_origin.row),
            );
            let to_write = match contents {
                CellContents::Formula { expr, .. } => CellContents::Formula {
                    expr: l123_parse::shift_refs(expr, dx, dy),
                    cached_value: None,
                },
                other => other.clone(),
            };
            self.wb_mut().cells.insert(dst, to_write.clone());
            self.push_to_engine_at(dst, &to_write);
        }
    }

    /// Like `push_to_engine` but for an arbitrary address (not
    /// `self.wb_mut().pointer`). Used during Copy/Move.
    fn push_to_engine_at(&mut self, addr: Address, contents: &CellContents) {
        let result = match contents {
            CellContents::Empty => self.wb_mut().engine.clear_cell(addr),
            CellContents::Label { text, .. } => self
                .wb_mut()
                .engine
                .set_user_input(addr, &format!("'{text}")),
            CellContents::Constant(Value::Number(n)) => self
                .wb_mut()
                .engine
                .set_user_input(addr, &l123_core::format_number_general(*n)),
            CellContents::Constant(Value::Text(s)) => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{s}"))
            }
            CellContents::Constant(_) => Ok(()),
            CellContents::Formula { expr, .. } => {
                let names = self.wb_mut().engine.all_sheet_names();
                let names_ref: Vec<&str> = names.iter().map(String::as_str).collect();
                let cfg = parse_config_from(&self.wb().international);
                let expanded = l123_parse::expand_cellpointer(expr, addr);
                let excel = l123_parse::to_engine_source_with_config(&expanded, &names_ref, &cfg);
                self.wb_mut().engine.set_user_input(addr, &excel)
            }
        };
        let _ = result;
    }

    fn execute_range_erase(&mut self, range: Range) {
        let r = range.normalized();
        // Single-sheet only for now; 3D ranges arrive with M5.
        let sheet = r.start.sheet;
        // Capture prior cell contents and format overrides for undo.
        let mut cells: Vec<(Address, CellContents)> = Vec::new();
        let mut formats: Vec<(Address, Format)> = Vec::new();
        let mut text_styles: Vec<(Address, TextStyle)> = Vec::new();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(sheet, col, row);
                if let Some(c) = self.wb().cells.get(&addr) {
                    cells.push((addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&addr) {
                    formats.push((addr, *f));
                }
                if let Some(s) = self.wb().cell_text_styles.get(&addr) {
                    text_styles.push((addr, *s));
                }
                self.wb_mut().cells.remove(&addr);
                self.wb_mut().cell_formats.remove(&addr);
                self.wb_mut().cell_text_styles.remove(&addr);
                let _ = self.wb_mut().engine.clear_cell(addr);
            }
        }
        if self.undo_enabled
            && (!cells.is_empty() || !formats.is_empty() || !text_styles.is_empty())
        {
            self.wb_mut().journal.push(JournalEntry::RangeRestore {
                cells,
                formats,
                text_styles,
            });
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
    }

    fn delete_col_at_pointer(&mut self, n: u16) {
        let at = self.wb().pointer.col;
        let mut batch: Vec<JournalEntry> = Vec::new();
        for sheet in self.target_sheets() {
            let captured_cells: Vec<(Address, CellContents)> = self
                .wb()
                .cells
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.col >= at && a.col < at + n)
                .map(|(a, c)| (*a, c.clone()))
                .collect();
            let captured_formats: Vec<(Address, Format)> = self
                .wb()
                .cell_formats
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.col >= at && a.col < at + n)
                .map(|(a, f)| (*a, *f))
                .collect();
            let captured_text_styles: Vec<(Address, TextStyle)> = self
                .wb()
                .cell_text_styles
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.col >= at && a.col < at + n)
                .map(|(a, s)| (*a, *s))
                .collect();
            if self.wb_mut().engine.delete_cols(sheet, at, n).is_ok() {
                self.wb_mut()
                    .cells
                    .retain(|a, _| !(a.sheet == sheet && a.col >= at && a.col < at + n));
                self.wb_mut()
                    .cell_formats
                    .retain(|a, _| !(a.sheet == sheet && a.col >= at && a.col < at + n));
                self.wb_mut()
                    .cell_text_styles
                    .retain(|a, _| !(a.sheet == sheet && a.col >= at && a.col < at + n));
                shift_cells_cols(&mut self.wb_mut().cells, sheet, at + n, -(n as i32));
                // One ColDelete per deleted column so undo restores in
                // the correct order via apply_undo.
                for k in 0..n {
                    let col_k = at + k;
                    let cells_k: Vec<_> = captured_cells
                        .iter()
                        .filter(|(a, _)| a.col == col_k)
                        .cloned()
                        .collect();
                    let formats_k: Vec<_> = captured_formats
                        .iter()
                        .filter(|(a, _)| a.col == col_k)
                        .cloned()
                        .collect();
                    let text_styles_k: Vec<_> = captured_text_styles
                        .iter()
                        .filter(|(a, _)| a.col == col_k)
                        .cloned()
                        .collect();
                    batch.push(JournalEntry::ColDelete {
                        sheet,
                        at: col_k,
                        cells: cells_k,
                        formats: formats_k,
                        text_styles: text_styles_k,
                    });
                }
            }
        }
        if !batch.is_empty() {
            self.wb_mut().dirty = true;
        }
        self.push_journal_batch(batch);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.close_menu();
    }

    fn descend_into(&mut self, item: &MenuItem) {
        match item.body {
            MenuBody::Submenu(_) => {
                if let Some(state) = self.menu.as_mut() {
                    state.path.push(item.letter);
                    state.highlight = 0;
                    state.message = None;
                }
            }
            MenuBody::Action(action) => self.execute_action(action),
            MenuBody::NotImplemented(tag) => {
                if let Some(state) = self.menu.as_mut() {
                    state.message = Some(tag);
                }
            }
        }
    }

    /// Translate the cell's contents into the form IronCalc expects and
    /// push it. Labels are stored with a `'` prefix so the engine treats
    /// them as text. Formulas are translated to Excel syntax.
    fn push_to_engine(&mut self, contents: &CellContents) {
        let addr = self.wb_mut().pointer;
        let result = match contents {
            CellContents::Empty => self.wb_mut().engine.clear_cell(addr),
            CellContents::Label { text, .. } => self
                .wb_mut()
                .engine
                .set_user_input(addr, &format!("'{text}")),
            CellContents::Constant(Value::Number(n)) => self
                .wb_mut()
                .engine
                .set_user_input(addr, &l123_core::format_number_general(*n)),
            CellContents::Constant(Value::Text(s)) => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{s}"))
            }
            CellContents::Constant(_) => Ok(()),
            CellContents::Formula { expr, .. } => {
                let names = self.wb_mut().engine.all_sheet_names();
                let names_ref: Vec<&str> = names.iter().map(String::as_str).collect();
                let cfg = parse_config_from(&self.wb().international);
                let expanded = l123_parse::expand_cellpointer(expr, addr);
                let excel = l123_parse::to_engine_source_with_config(&expanded, &names_ref, &cfg);
                self.wb_mut().engine.set_user_input(addr, &excel)
            }
        };
        // Engine errors are non-fatal for the UI — an ERR value will
        // surface on the next cache refresh. Swallow for M2; surfacing
        // in an error panel is its own milestone.
        let _ = result;
    }

    /// Walk every `Formula` cell and re-read its computed value from the
    /// engine.  Called after every recalc.
    fn refresh_formula_caches(&mut self) {
        let formula_addrs: Vec<Address> = self
            .wb_mut()
            .cells
            .iter()
            .filter_map(|(addr, c)| matches!(c, CellContents::Formula { .. }).then_some(*addr))
            .collect();
        for addr in formula_addrs {
            let Ok(view) = self.wb_mut().engine.get_cell(addr) else {
                continue;
            };
            if let Some(CellContents::Formula { cached_value, .. }) =
                self.wb_mut().cells.get_mut(&addr)
            {
                *cached_value = Some(view.value);
            }
        }
    }

    fn move_pointer(&mut self, d_col: i32, d_row: i32) {
        if self.input_range.is_some() {
            let from = self.wb().pointer;
            if let Some(next) = self.next_unprotected(from, d_col, d_row) {
                self.wb_mut().pointer = next;
                self.scroll_into_view();
            } else {
                self.request_beep();
            }
            return;
        }
        if let Some(next) = self.wb_mut().pointer.shifted(d_col, d_row) {
            self.wb_mut().pointer = next;
            self.scroll_into_view();
        } else {
            self.request_beep();
        }
    }

    /// Record an error-beep request. No-op when beep is disabled, so
    /// `beep_count` / `beep_pending` remain dormant and downstream
    /// emission is skipped. Internal use only — UI code chooses *when*
    /// to beep; the config choice gates whether we actually do.
    fn request_beep(&mut self) {
        if !self.beep_enabled {
            return;
        }
        self.beep_count = self.beep_count.saturating_add(1);
        self.beep_pending = true;
    }

    /// Monotonic count of beeps observed since the app was created.
    /// Acceptance transcripts use this; production code never reads it.
    pub fn beep_count(&self) -> u64 {
        self.beep_count
    }

    /// Whether the error-beep is currently active. Mirrors the config
    /// at startup; `/Worksheet Global Default Other Beep Enable|Disable`
    /// flips it at runtime.
    pub fn beep_enabled(&self) -> bool {
        self.beep_enabled
    }

    /// Seed the beep setting — called once at startup from the binary
    /// after resolving `Config`. Safe to call later too (tests use it).
    pub fn set_beep_enabled(&mut self, enabled: bool) {
        self.beep_enabled = enabled;
    }

    /// Active chrome theme.
    pub fn theme(&self) -> crate::Theme {
        self.theme
    }

    /// Seed the chrome theme — called once at startup after resolving
    /// `Config` and the `--theme` CLI override. Safe to call later
    /// (tests use it).
    pub fn set_theme(&mut self, theme: crate::Theme) {
        self.theme = theme;
    }

    /// Returns true if a beep has been requested since the last call,
    /// and clears the pending flag. The event loop reads this once per
    /// iteration to emit a single BEL no matter how many requests piled
    /// up inside a single keystroke handler.
    pub fn take_pending_beep(&mut self) -> bool {
        std::mem::take(&mut self.beep_pending)
    }

    fn scroll_into_view(&mut self) {
        if self.wb_mut().pointer.col < self.wb_mut().viewport_col_offset {
            self.wb_mut().viewport_col_offset = self.wb_mut().pointer.col;
        }
        if self.wb_mut().pointer.row < self.wb_mut().viewport_row_offset {
            self.wb_mut().viewport_row_offset = self.wb_mut().pointer.row;
        }

        // Down/right scroll requires viewport dimensions, which only
        // the renderer knows. We use the previous frame's grid rect —
        // the user always sees a frame before pressing a key, so the
        // cached value is correct in steady state. If no grid has been
        // rendered yet, leave the offsets alone; A1 is in view by
        // construction.
        let Some(area) = self.last_grid_area.get() else {
            return;
        };
        if area.width <= ROW_GUTTER || area.height < 2 {
            return;
        }
        let content_width = area.width - ROW_GUTTER;
        let visible_rows = (area.height - 1) as u32;
        // Frozen rows occupy the top of the body region; only the
        // remaining rows can scroll, so the effective scrolling
        // capacity shrinks by the frozen-row count when computing
        // viewport_row_offset.  When the pointer sits inside the
        // frozen prefix it never needs scroll.
        let sheet = self.wb().pointer.sheet;
        let frozen_rows: u32 = self.wb().frozen.get(&sheet).map(|f| f.0).unwrap_or(0);
        let frozen_cols: u16 = self.wb().frozen.get(&sheet).map(|f| f.1).unwrap_or(0);

        let pointer_row = self.wb().pointer.row;
        if pointer_row >= frozen_rows {
            let scrolling_rows = visible_rows.saturating_sub(frozen_rows).max(1);
            let effective_offset = self.wb().viewport_row_offset.max(frozen_rows);
            if pointer_row >= effective_offset + scrolling_rows {
                self.wb_mut().viewport_row_offset = pointer_row - scrolling_rows + 1;
            }
        }

        let pointer_col = self.wb().pointer.col;
        if pointer_col >= frozen_cols && !self.wb().hidden_cols.contains(&(sheet, pointer_col)) {
            let actual_w = self.col_width_of(sheet, pointer_col) as u16;
            let layout = self.visible_column_layout(content_width);
            let fully_visible = layout
                .iter()
                .any(|&(c, _, drawn)| c == pointer_col && drawn == actual_w);
            if !fully_visible {
                let frozen_width: u16 = (0..frozen_cols)
                    .filter(|c| !self.wb().hidden_cols.contains(&(sheet, *c)))
                    .map(|c| self.col_width_of(sheet, c) as u16)
                    .sum();
                let scrolling_width = content_width.saturating_sub(frozen_width).max(1);
                let new_offset = self
                    .ideal_left_for_rightmost(pointer_col, scrolling_width)
                    .max(frozen_cols);
                self.wb_mut().viewport_col_offset = new_offset;
            }
        }
    }
}

impl Default for App {
    fn default() -> Self {
        Self::new()
    }
}

fn shift_cells_rows(
    cells: &mut HashMap<Address, CellContents>,
    sheet: SheetId,
    at: u32,
    delta: i64,
) {
    let affected: Vec<Address> = cells
        .keys()
        .filter(|a| a.sheet == sheet && a.row >= at)
        .copied()
        .collect();
    // Shift in an order that avoids collisions: highest first for +delta,
    // lowest first for -delta.
    let mut sorted = affected;
    if delta >= 0 {
        sorted.sort_by_key(|a| std::cmp::Reverse(a.row));
    } else {
        sorted.sort_by_key(|a| a.row);
    }
    for addr in sorted {
        let contents = cells.remove(&addr).expect("present");
        let new_row = (addr.row as i64 + delta).max(0) as u32;
        let new_addr = Address::new(addr.sheet, addr.col, new_row);
        cells.insert(new_addr, contents);
    }
}

/// After deleting the sheet at index `at`, drop every cache entry on
/// that sheet and shift entries on later sheets back by one slot. The
/// inverse of [`shift_sheets_from`]; covers the same caches.
fn drop_sheet_from_caches(
    cells: &mut HashMap<Address, CellContents>,
    cell_formats: &mut HashMap<Address, Format>,
    cell_format_overrides: &mut HashMap<Address, String>,
    cell_text_styles: &mut HashMap<Address, TextStyle>,
    col_widths: &mut HashMap<(SheetId, u16), u8>,
    at: u16,
) {
    cells.retain(|a, _| a.sheet.0 != at);
    cell_formats.retain(|a, _| a.sheet.0 != at);
    cell_format_overrides.retain(|a, _| a.sheet.0 != at);
    cell_text_styles.retain(|a, _| a.sheet.0 != at);
    col_widths.retain(|(s, _), _| s.0 != at);

    let shift_addr = |a: Address| -> Address {
        if a.sheet.0 > at {
            Address::new(SheetId(a.sheet.0 - 1), a.col, a.row)
        } else {
            a
        }
    };
    let mut affected: Vec<Address> = cells.keys().filter(|a| a.sheet.0 > at).copied().collect();
    affected.sort_by_key(|a| a.sheet.0);
    for addr in affected {
        let contents = cells.remove(&addr).expect("present");
        cells.insert(shift_addr(addr), contents);
    }
    let mut fmt_affected: Vec<Address> = cell_formats
        .keys()
        .filter(|a| a.sheet.0 > at)
        .copied()
        .collect();
    fmt_affected.sort_by_key(|a| a.sheet.0);
    for addr in fmt_affected {
        let f = cell_formats.remove(&addr).expect("present");
        cell_formats.insert(shift_addr(addr), f);
    }
    let mut ovr_affected: Vec<Address> = cell_format_overrides
        .keys()
        .filter(|a| a.sheet.0 > at)
        .copied()
        .collect();
    ovr_affected.sort_by_key(|a| a.sheet.0);
    for addr in ovr_affected {
        let s = cell_format_overrides.remove(&addr).expect("present");
        cell_format_overrides.insert(shift_addr(addr), s);
    }
    let mut style_affected: Vec<Address> = cell_text_styles
        .keys()
        .filter(|a| a.sheet.0 > at)
        .copied()
        .collect();
    style_affected.sort_by_key(|a| a.sheet.0);
    for addr in style_affected {
        let s = cell_text_styles.remove(&addr).expect("present");
        cell_text_styles.insert(shift_addr(addr), s);
    }
    let mut cw_affected: Vec<(SheetId, u16)> = col_widths
        .keys()
        .filter(|(s, _)| s.0 > at)
        .copied()
        .collect();
    cw_affected.sort_by_key(|(s, _)| s.0);
    for key in cw_affected {
        let w = col_widths.remove(&key).expect("present");
        col_widths.insert((SheetId(key.0 .0 - 1), key.1), w);
    }
}

/// After inserting `delta` sheets at position `at`, every cell whose
/// sheet index is >= `at` moves forward by `delta`. Applies to the
/// three per-sheet caches App keeps in sync with the engine.
fn shift_sheets_from(
    cells: &mut HashMap<Address, CellContents>,
    cell_formats: &mut HashMap<Address, Format>,
    cell_format_overrides: &mut HashMap<Address, String>,
    cell_text_styles: &mut HashMap<Address, TextStyle>,
    col_widths: &mut HashMap<(SheetId, u16), u8>,
    at: u16,
    delta: u16,
) {
    if delta == 0 {
        return;
    }
    let shift_addr = |a: Address| -> Address {
        if a.sheet.0 >= at {
            Address::new(SheetId(a.sheet.0 + delta), a.col, a.row)
        } else {
            a
        }
    };
    let mut affected: Vec<Address> = cells.keys().filter(|a| a.sheet.0 >= at).copied().collect();
    affected.sort_by_key(|a| std::cmp::Reverse(a.sheet.0));
    for addr in affected {
        let contents = cells.remove(&addr).expect("present");
        cells.insert(shift_addr(addr), contents);
    }
    let mut fmt_affected: Vec<Address> = cell_formats
        .keys()
        .filter(|a| a.sheet.0 >= at)
        .copied()
        .collect();
    fmt_affected.sort_by_key(|a| std::cmp::Reverse(a.sheet.0));
    for addr in fmt_affected {
        let f = cell_formats.remove(&addr).expect("present");
        cell_formats.insert(shift_addr(addr), f);
    }
    let mut ovr_affected: Vec<Address> = cell_format_overrides
        .keys()
        .filter(|a| a.sheet.0 >= at)
        .copied()
        .collect();
    ovr_affected.sort_by_key(|a| std::cmp::Reverse(a.sheet.0));
    for addr in ovr_affected {
        let s = cell_format_overrides.remove(&addr).expect("present");
        cell_format_overrides.insert(shift_addr(addr), s);
    }
    let mut style_affected: Vec<Address> = cell_text_styles
        .keys()
        .filter(|a| a.sheet.0 >= at)
        .copied()
        .collect();
    style_affected.sort_by_key(|a| std::cmp::Reverse(a.sheet.0));
    for addr in style_affected {
        let s = cell_text_styles.remove(&addr).expect("present");
        cell_text_styles.insert(shift_addr(addr), s);
    }
    let mut cw_affected: Vec<(SheetId, u16)> = col_widths
        .keys()
        .filter(|(s, _)| s.0 >= at)
        .copied()
        .collect();
    cw_affected.sort_by_key(|(s, _)| std::cmp::Reverse(s.0));
    for key in cw_affected {
        let w = col_widths.remove(&key).expect("present");
        col_widths.insert((SheetId(key.0 .0 + delta), key.1), w);
    }
}

fn shift_cells_cols(
    cells: &mut HashMap<Address, CellContents>,
    sheet: SheetId,
    at: u16,
    delta: i32,
) {
    let affected: Vec<Address> = cells
        .keys()
        .filter(|a| a.sheet == sheet && a.col >= at)
        .copied()
        .collect();
    let mut sorted = affected;
    if delta >= 0 {
        sorted.sort_by_key(|a| std::cmp::Reverse(a.col));
    } else {
        sorted.sort_by_key(|a| a.col);
    }
    for addr in sorted {
        let contents = cells.remove(&addr).expect("present");
        let new_col = (addr.col as i32 + delta).max(0) as u16;
        let new_addr = Address::new(addr.sheet, new_col, addr.row);
        cells.insert(new_addr, contents);
    }
}

/// Mirrors ratatui-image's own `iterm2_from_env` list of hosts that
/// speak the OSC 1337 inline-image protocol. We re-check the same
/// environment in [`App::probe_image_picker`] as a workaround for a
/// quirk in `Picker::from_query_stdio`: when the font-size probe
/// fails (common in iTerm2), the library drops back to a default
/// Halfblocks picker and discards its own iTerm2 env hint.
/// Filename component used for the WAIT-mode line-3 noun, falling
/// back to the empty string for paths without a final component.
fn display_basename(path: &Path) -> String {
    path.file_name()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_string()
}

/// True when `path` ends in `.wk3`/`.WK3`, regardless of whether
/// the `wk3` cargo feature is on (callers gate the actual load
/// behind their own `#[cfg]`). Factored out so the §4.7 async
/// retrieve worker and the sync CLI-startup path agree.
fn is_wk3_path(path: &Path) -> bool {
    #[cfg(feature = "wk3")]
    {
        path.extension()
            .and_then(|e| e.to_str())
            .map(|s| s.eq_ignore_ascii_case("wk3"))
            .unwrap_or(false)
    }
    #[cfg(not(feature = "wk3"))]
    {
        let _ = path;
        false
    }
}

fn is_iterm2_compatible_env(term_program: Option<&str>, lc_terminal: Option<&str>) -> bool {
    const HINTS: &[&str] = &[
        "iTerm",
        "WezTerm",
        "mintty",
        "vscode",
        "Tabby",
        "Hyper",
        "rio",
        "Bobcat",
        "WarpTerminal",
    ];
    if let Some(tp) = term_program {
        if HINTS.iter().any(|h| tp.contains(h)) {
            return true;
        }
    }
    if let Some(lc) = lc_terminal {
        if lc.contains("iTerm") {
            return true;
        }
    }
    false
}
