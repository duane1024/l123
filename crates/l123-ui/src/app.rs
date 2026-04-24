//! The app loop: ratatui + crossterm, control panel + grid + status line.
//!
//! Scope as of M1 cycle 2:
//! - READY / LABEL / VALUE modes with first-character dispatch (LABEL only
//!   implemented this cycle; VALUE lands in cycle 3).
//! - `'` auto-prefixed labels. Enter commits; Ctrl-C quits.
//! - Three-line control panel, mode indicator, cell readout.

use std::collections::HashMap;
use std::io;
use std::path::{Path, PathBuf};
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
use l123_core::{
    address::col_to_letters, label::is_value_starter, Address, CellContents, Format, FormatKind,
    LabelPrefix, Mode, Range, SheetId, Value,
};
use l123_engine::{CellView, Engine, IronCalcEngine, RecalcMode};
use l123_menu::{self as menu, Action, MenuBody, MenuItem};

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

/// All per-file state — one instance per active file. Session-level
/// fields (mode, menu, entry buffer, …) live on [`App`].
struct Workbook {
    engine: IronCalcEngine,
    cells: HashMap<Address, CellContents>,
    cell_formats: HashMap<Address, Format>,
    col_widths: HashMap<(SheetId, u16), u8>,
    /// Last-saved-to path. Prefilled into `/FS` prompts so re-save is
    /// a single Enter. `None` until the file has been saved at least
    /// once.
    active_path: Option<PathBuf>,
    pointer: Address,
    viewport_col_offset: u16,
    viewport_row_offset: u32,
    /// Command journal for Undo (Alt-F4). Each mutating command
    /// pushes an inverse entry. Pop-and-apply reverts.
    journal: Vec<JournalEntry>,
}

impl Workbook {
    fn new() -> Self {
        Self {
            engine: IronCalcEngine::new().expect("IronCalc engine init"),
            cells: HashMap::new(),
            cell_formats: HashMap::new(),
            col_widths: HashMap::new(),
            active_path: None,
            pointer: Address::A1,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            journal: Vec::new(),
        }
    }
}

pub struct App {
    mode: Mode,
    running: bool,
    entry: Option<Entry>,
    default_label_prefix: LabelPrefix,
    recalc_mode: RecalcMode,
    recalc_pending: bool,
    /// 1-2-3 GROUP mode: when true, format and row/col operations
    /// propagate across all sheets of the active file. Toggled by
    /// `/Worksheet Global Group Enable|Disable`. Lights the GROUP
    /// indicator on the status line.
    group_mode: bool,
    /// True when `/Worksheet Global Default Other Undo` is enabled.
    /// While true, mutating commands push reverse entries onto the
    /// journal; Alt-F4 pops and applies. L123 defaults this to ON.
    undo_enabled: bool,
    menu: Option<MenuState>,
    point: Option<PointState>,
    prompt: Option<PromptState>,
    /// Transient slot for the two-step /Range Name Create flow — the
    /// typed name is stashed here after the prompt step and consumed by
    /// commit_point.
    pending_name: Option<String>,
    /// After committing a filename that already exists on disk, this
    /// carries the chosen path through the Cancel/Replace/Backup
    /// submenu. Mode stays MENU while present.
    save_confirm: Option<SaveConfirmState>,
    /// Transient slot for the two-step /File Xtract flow — the typed
    /// filename is stashed here after the prompt step and consumed by
    /// commit_point.
    pending_xtract_path: Option<PathBuf>,
    /// Overlay state for /File List. When present, the mode is Files
    /// and the grid is obscured by a horizontal picker on lines 2/3.
    file_list: Option<FileListState>,
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
}

/// Which cell kinds `/Range Search` walks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SearchScope {
    Formulas,
    Labels,
    Both,
}

/// Live state of a `/Range Search` session between scope selection
/// and the final Find or Replace leaf.
#[derive(Debug, Clone)]
struct SearchSession {
    scope: SearchScope,
    range: Range,
    search: String,
    /// Cached matches, populated when the user picks Find. Replace
    /// recomputes its own match set just before applying.
    matches: Vec<Address>,
    /// Index of the current highlighted match within `matches`.
    cursor: usize,
}

/// Live state of a `/Print File` session between filename commit and
/// final Go. Holds the destination path, the chosen range, and the
/// current page-decoration settings.
#[derive(Debug, Clone)]
struct PrintSession {
    path: PathBuf,
    range: Option<Range>,
    /// Three-part header string (`L|C|R`). Empty means no header.
    header: String,
    /// Three-part footer string (`L|C|R`). Empty means no footer.
    footer: String,
    /// As-Displayed (default) or Cell-Formulas.
    content_mode: PrintContentMode,
    /// Formatted (default — emits header/footer) or Unformatted
    /// (range content only).
    format_mode: PrintFormatMode,
    /// Left margin: N spaces prepended to every output line.
    margin_left: u16,
    /// Right margin — accepted but not yet honored (no wrapping
    /// implemented). Storing it keeps the menu muscle memory intact.
    margin_right: u16,
    /// Top margin: N blank lines above the first output line.
    margin_top: u16,
    /// Bottom margin — accepted but not yet honored (no pagination
    /// yet).
    margin_bottom: u16,
    /// Lines per page. 0 = no pagination.
    pg_length: u16,
    /// Next page number to print at the start of Go. Persists across
    /// successive Gos in the same session so headers using `#` count
    /// up; `/PF Align` resets it to 1.
    next_page: u32,
}

impl PrintSession {
    fn new(path: PathBuf) -> Self {
        Self {
            path,
            range: None,
            header: String::new(),
            footer: String::new(),
            content_mode: PrintContentMode::AsDisplayed,
            format_mode: PrintFormatMode::Formatted,
            margin_left: 0,
            margin_right: 0,
            margin_top: 0,
            margin_bottom: 0,
            pg_length: 0,
            next_page: 1,
        }
    }

    /// /PF Clear All: reset every per-session knob but keep the
    /// destination path, the chosen range, and the page counter —
    /// those are session identity, not settings.
    fn clear_all(&mut self) {
        self.header.clear();
        self.footer.clear();
        self.content_mode = PrintContentMode::AsDisplayed;
        self.format_mode = PrintFormatMode::Formatted;
        self.margin_left = 0;
        self.margin_right = 0;
        self.margin_top = 0;
        self.margin_bottom = 0;
        self.pg_length = 0;
    }
}

/// How cell contents are rendered into the print file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PrintContentMode {
    /// Default — render each cell as it appears on the screen.
    AsDisplayed,
    /// Render formula cells as their source expression; labels and
    /// constants unchanged.
    CellFormulas,
}

/// Whether page decorations (headers, footers, …) wrap the output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PrintFormatMode {
    Formatted,
    Unformatted,
}

/// Pre-resolved snapshot of the [`PrintSession`] fields the renderer
/// needs. Built at Go-time so the render helper doesn't re-borrow
/// `self.print`.
struct PrintSettings {
    header: String,
    footer: String,
    content_mode: PrintContentMode,
    format_mode: PrintFormatMode,
    margin_left: u16,
    margin_right: u16,
    margin_top: u16,
    margin_bottom: u16,
    pg_length: u16,
    /// Starting page number for the first page of this Go.
    start_page: u32,
}

impl Default for PrintSettings {
    fn default() -> Self {
        Self {
            header: String::new(),
            footer: String::new(),
            content_mode: PrintContentMode::AsDisplayed,
            format_mode: PrintFormatMode::Formatted,
            margin_left: 0,
            margin_right: 0,
            margin_top: 0,
            margin_bottom: 0,
            pg_length: 0,
            start_page: 1,
        }
    }
}

/// Inverse commands recorded before each mutating operation. See SPEC
/// §17 / PLAN §4.3.
#[derive(Debug, Clone)]
enum JournalEntry {
    /// Restore a single cell to its prior contents / format. A `None`
    /// field means the cell was unset before the recorded edit.
    CellEdit {
        addr: Address,
        prev_contents: Option<CellContents>,
        prev_format: Option<Format>,
    },
    /// Reinstate a deleted row on one sheet: insert a fresh row at
    /// `at`, then rewrite the captured cells.
    RowDelete {
        sheet: SheetId,
        at: u32,
        cells: Vec<(Address, CellContents)>,
        formats: Vec<(Address, Format)>,
    },
    /// Undo of a row insert: delete the row that was inserted.
    RowInsert { sheet: SheetId, at: u32 },
    /// Reinstate a deleted column on one sheet.
    ColDelete {
        sheet: SheetId,
        at: u16,
        cells: Vec<(Address, CellContents)>,
        formats: Vec<(Address, Format)>,
    },
    /// Undo of a column insert: delete the column that was inserted.
    ColInsert { sheet: SheetId, at: u16 },
    /// Restore a range's prior per-cell contents + formats. Captures
    /// the state that `/Range Erase` cleared.
    RangeRestore {
        cells: Vec<(Address, CellContents)>,
        formats: Vec<(Address, Format)>,
    },
    /// Restore per-cell format overrides after `/Range Format`. Each
    /// entry's `Option<Format>` is the pre-command format (None ==
    /// no override).
    RangeFormat {
        entries: Vec<(Address, Option<Format>)>,
    },
    /// Restore one column's width. `prev_width = None` means the
    /// column had no override (default width).
    ColWidth {
        sheet: SheetId,
        col: u16,
        prev_width: Option<u8>,
    },
    /// Group of entries popped and applied together — used when
    /// GROUP propagated a single command to multiple sheets.
    Batch(Vec<JournalEntry>),
}

#[derive(Debug, Clone)]
struct SaveConfirmState {
    path: PathBuf,
    /// 0=Cancel, 1=Replace, 2=Backup — matches `SAVE_CONFIRM_ITEMS` below.
    highlight: usize,
}

/// Items shown on line 2 of the Cancel/Replace/Backup submenu. The
/// first letter of each is the accelerator.
const SAVE_CONFIRM_ITEMS: &[(&str, &str)] = &[
    ("Cancel", "Abort the save"),
    ("Replace", "Overwrite the existing file"),
    ("Backup", "Rename existing to .BAK then save"),
];

/// Numeric (or short-text) prompt state for commands that need an argument
/// before descending into POINT. E.g. /RFC → "Enter number of decimal
/// places (0..15): 2" → then POINT for the range.
#[derive(Debug, Clone)]
struct PromptState {
    label: String,
    buffer: String,
    /// What the command wants to do once the prompt commits.
    next: PromptNext,
    /// True while the buffer still holds the auto-filled default. The
    /// first printable keystroke clears it (1-2-3 "typed input replaces
    /// the default" convention).
    fresh: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum PromptNext {
    /// Then go to POINT and apply `Format { kind, decimals: <buffer> }`.
    RangeFormat { kind: FormatKind },
    /// Set the current column's width to the buffered number.
    WorksheetColumnSetWidth,
    /// After the user types a name, stash it and go to POINT for the range.
    RangeNameCreate,
    /// After the user types a name, delete it from the engine.
    RangeNameDelete,
    /// After the user types a filename, save the workbook to that path
    /// as xlsx.
    FileSaveFilename,
    /// After the user types a filename, load that xlsx file, replacing
    /// all in-memory workbook state.
    FileRetrieveFilename,
    /// After the user types a filename, enter POINT to pick the range
    /// to extract with the given kind (Formulas or Values).
    FileXtractFilename { kind: XtractKind },
    /// After the user types a filename, parse it as CSV and paint the
    /// values into cells starting at the pointer.
    FileImportNumbersFilename,
    /// After the user types a directory path, make it the session's
    /// working directory.
    FileDirPath,
    /// After the user types a filename, load that xlsx file as a
    /// second active file. `before` controls whether the new file
    /// takes the current slot (and the old one is stashed ahead) or
    /// is appended after the current one.
    FileOpenFilename { before: bool },
    /// After the user types a print destination path, start a
    /// [`PrintSession`] and descend into the `/PF` submenu.
    PrintFileFilename,
    /// After the user types a header or footer string, store it on
    /// the active [`PrintSession`] and re-enter the Options submenu.
    PrintFileHeader,
    PrintFileFooter,
    /// Numeric margin prompts (0..=1000). Each stores onto the
    /// active [`PrintSession`] and re-enters the Margins submenu.
    PrintFileMarginLeft,
    PrintFileMarginRight,
    PrintFileMarginTop,
    PrintFileMarginBottom,
    /// Numeric page-length prompt (0..=1000). 0 means no pagination.
    PrintFilePgLength,
    /// After the user types the search string, open the Find|Replace
    /// submenu. `scope` and `range` were captured earlier.
    RangeSearchString { scope: SearchScope, range: Range },
    /// After the user types the replacement string, apply it to all
    /// matches in the active [`SearchSession`].
    RangeSearchReplacement,
}

/// /File Xtract sub-command: does the extracted file keep formulas,
/// or is each cell written as its current cached value?
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum XtractKind {
    Formulas,
    Values,
}

/// /File List sub-command: which set of files is in the overlay?
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FileListKind {
    /// xlsx files in the current session directory.
    Worksheet,
    /// Currently-loaded active files (single-file workbook today).
    Active,
}

#[derive(Debug, Clone)]
pub(crate) struct FileListState {
    kind: FileListKind,
    entries: Vec<PathBuf>,
    highlight: usize,
    /// Index of the first entry rendered on the overlay. Kept in sync
    /// with `highlight` so the selected row is always visible.
    view_offset: usize,
}

/// Visible window size (rows) for the /File List overlay. Also the
/// distance moved by PgUp / PgDn. Fixed rather than dynamic because the
/// key handler runs before render knows the real terminal height;
/// render clamps to the actual area anyway.
const FILE_LIST_PAGE_SIZE: usize = 10;

impl PromptNext {
    fn accepts_char(self, c: char) -> bool {
        match self {
            PromptNext::RangeFormat { .. } | PromptNext::WorksheetColumnSetWidth => {
                c.is_ascii_digit()
            }
            PromptNext::RangeNameCreate | PromptNext::RangeNameDelete => {
                c.is_ascii_alphanumeric() || c == '_'
            }
            PromptNext::FileSaveFilename
            | PromptNext::FileRetrieveFilename
            | PromptNext::FileXtractFilename { .. }
            | PromptNext::FileImportNumbersFilename
            | PromptNext::FileDirPath
            | PromptNext::FileOpenFilename { .. }
            | PromptNext::PrintFileFilename => is_path_char(c),
            // Header and footer are free-form text with the `|`
            // separator carving them into L|C|R.
            PromptNext::PrintFileHeader | PromptNext::PrintFileFooter => {
                c != '\n' && c != '\t'
            }
            // Search / replacement strings are free text.
            PromptNext::RangeSearchString { .. } | PromptNext::RangeSearchReplacement => {
                c != '\n' && c != '\t'
            }
            PromptNext::PrintFileMarginLeft
            | PromptNext::PrintFileMarginRight
            | PromptNext::PrintFileMarginTop
            | PromptNext::PrintFileMarginBottom
            | PromptNext::PrintFilePgLength => c.is_ascii_digit(),
        }
    }
}

/// Characters accepted inside a filename/path prompt. Deliberately
/// narrower than 1-2-3's "anything goes" — we exclude keys with menu
/// semantics (`/`, period-free submenus). `/` is fine; `.` is fine.
fn is_path_char(c: char) -> bool {
    c.is_ascii_alphanumeric()
        || matches!(c, '.' | '-' | '_' | '/' | '\\' | ' ' | '~')
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

/// Keep the /File List view window centered around the highlighted
/// row: `view_offset <= highlight < view_offset + FILE_LIST_PAGE_SIZE`.
fn adjust_file_list_view(fl: &mut FileListState) {
    if fl.highlight < fl.view_offset {
        fl.view_offset = fl.highlight;
    } else if fl.highlight >= fl.view_offset + FILE_LIST_PAGE_SIZE {
        fl.view_offset = fl.highlight + 1 - FILE_LIST_PAGE_SIZE;
    }
}

/// List every `.xlsx` file in `dir`, sorted by filename (case-
/// insensitive ordering isn't worth the extra code today). Hidden
/// files and non-file entries are skipped.
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
                .map(|x| x.eq_ignore_ascii_case("xlsx"))
                .unwrap_or(false)
        })
        .collect();
    entries.sort();
    entries
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
fn cell_view_to_contents(cv: &CellView) -> Option<CellContents> {
    if let Some(f) = &cv.formula {
        let expr = f.strip_prefix('=').unwrap_or(f).to_string();
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

/// Transient state while the user is selecting a cell/range in POINT mode.
#[derive(Debug, Clone)]
struct PointState {
    /// Anchor corner. `None` after a single Esc — in that state the
    /// pointer moves freely without growing a range; a second Esc cancels.
    anchor: Option<Address>,
    /// Which command initiated POINT, so that `Enter` routes the selected
    /// range back to the right handler.
    pending: PendingCommand,
}

/// Commands in progress that are waiting on one more POINT selection.
#[derive(Debug, Clone, Copy)]
enum PendingCommand {
    RangeErase,
    CopyFrom,
    CopyTo { source: Range },
    MoveFrom,
    MoveTo { source: Range },
    RangeLabel { new_prefix: LabelPrefix },
    RangeFormat { format: Format },
    /// `pending_name` on App carries the name; on commit, define it over
    /// the selected range.
    RangeNameCreate,
    /// `pending_xtract_path` on App carries the destination path; on
    /// commit, extract the selected range into a new workbook file.
    FileXtractRange { kind: XtractKind },
    /// The user is choosing the print range for the active
    /// [`PrintSession`]. On commit the range is stashed and the
    /// `/PF` submenu reopens.
    PrintFileRange,
    /// POINT step of `/Range Search`: on commit, prompt for the
    /// search string.
    RangeSearchRange { scope: SearchScope },
}

impl PendingCommand {
    fn prompt(self) -> &'static str {
        match self {
            PendingCommand::RangeErase => "Enter range to erase:",
            PendingCommand::CopyFrom => "Enter range to copy FROM:",
            PendingCommand::CopyTo { .. } => "Enter range to copy TO:",
            PendingCommand::MoveFrom => "Enter range to move FROM:",
            PendingCommand::MoveTo { .. } => "Enter range to move TO:",
            PendingCommand::RangeLabel { .. } => "Enter range for label-prefix change:",
            PendingCommand::RangeFormat { .. } => "Enter range to format:",
            PendingCommand::RangeNameCreate => "Enter range for the named range:",
            PendingCommand::FileXtractRange { .. } => "Enter range to extract:",
            PendingCommand::PrintFileRange => "Enter range to print:",
            PendingCommand::RangeSearchRange { .. } => "Enter search range:",
        }
    }
}

/// Transient state used while the user is navigating the slash menu.
#[derive(Debug, Clone)]
struct MenuState {
    /// Letters descended into, longest-ago first.
    path: Vec<char>,
    /// Index into the currently-visible level.
    highlight: usize,
    /// Message to display on line 3 — typically the last-selected leaf's
    /// identifier when it was `NotImplemented`.
    message: Option<&'static str>,
    /// Optional alternate root, used for nested menus (e.g. the
    /// `/Print File` submenu). When Some, `level()` resolves `path`
    /// against this slice instead of `l123_menu::ROOT`.
    override_root: Option<&'static [MenuItem]>,
}

impl MenuState {
    fn fresh() -> Self {
        Self { path: Vec::new(), highlight: 0, message: None, override_root: None }
    }

    /// New menu rooted at a specific submenu rather than the global
    /// root. Used when a command (e.g. `/PF` after filename) hands
    /// the user to a sub-tree.
    fn rooted_at(root: &'static [MenuItem]) -> Self {
        Self {
            path: Vec::new(),
            highlight: 0,
            message: None,
            override_root: Some(root),
        }
    }

    fn level(&self) -> &'static [MenuItem] {
        match self.override_root {
            Some(root) => menu::current_level_within(root, &self.path),
            None => menu::current_level(&self.path),
        }
    }

    fn highlighted(&self) -> Option<&'static MenuItem> {
        self.level().get(self.highlight)
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
            recalc_pending: false,
            group_mode: false,
            undo_enabled: true,
            menu: None,
            point: None,
            prompt: None,
            pending_name: None,
            save_confirm: None,
            pending_xtract_path: None,
            file_list: None,
            active_files: vec![Workbook::new()],
            current: 0,
            file_nav_pending: false,
            print: None,
            search: None,
        }
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
        self.wb().pointer
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
        if a.col < self.wb().viewport_col_offset || a.row < self.wb().viewport_row_offset {
            return None;
        }
        let dc = a.col - self.wb().viewport_col_offset;
        let dr = (a.row - self.wb().viewport_row_offset) as u16;
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
            _ => {}
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
            KeyCode::F(2) => self.begin_edit(),
            KeyCode::F(4) if k.modifiers.contains(KeyModifiers::ALT) => self.undo(),
            KeyCode::F(9) => self.do_recalc(),
            KeyCode::Char('/') => self.open_menu(),
            KeyCode::Char(c) => self.begin_entry(c),
            _ => {}
        }
    }

    fn begin_edit(&mut self) {
        let pointer = self.wb().pointer;
        let source = self
            .wb()
            .cells
            .get(&pointer)
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
        } else if matches!(c, '\'' | '"' | '^' | '\\' | '|') {
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
        let mut contents = match entry.kind {
            EntryKind::Label(prefix) => CellContents::Label { prefix, text: entry.buffer },
            EntryKind::Value => match entry.buffer.parse::<f64>() {
                Ok(n) => CellContents::Constant(Value::Number(n)),
                Err(_) => CellContents::Formula { expr: entry.buffer, cached_value: None },
            },
            // EDIT commits re-parse the full source buffer so the user can
            // change prefix or type (label ↔ value) via the first-char rule.
            EntryKind::Edit => {
                CellContents::from_source(&entry.buffer, self.default_label_prefix)
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
        let Some(entry) = self.wb_mut().journal.pop() else { return };
        self.apply_undo(entry);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
    }

    fn apply_undo(&mut self, entry: JournalEntry) {
        match entry {
            JournalEntry::CellEdit { addr, prev_contents, prev_format } => {
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
            JournalEntry::RowDelete { sheet, at, cells, formats } => {
                if self.wb_mut().engine.insert_rows(sheet, at, 1).is_ok() {
                    shift_cells_rows(&mut self.wb_mut().cells, sheet, at, 1);
                    for (addr, contents) in cells {
                        self.push_to_engine_at(addr, &contents);
                        self.wb_mut().cells.insert(addr, contents);
                    }
                    for (addr, fmt) in formats {
                        self.wb_mut().cell_formats.insert(addr, fmt);
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
                    shift_cells_rows(&mut self.wb_mut().cells, sheet, at + 1, -1);
                }
            }
            JournalEntry::ColDelete { sheet, at, cells, formats } => {
                if self.wb_mut().engine.insert_cols(sheet, at, 1).is_ok() {
                    shift_cells_cols(&mut self.wb_mut().cells, sheet, at, 1);
                    for (addr, contents) in cells {
                        self.push_to_engine_at(addr, &contents);
                        self.wb_mut().cells.insert(addr, contents);
                    }
                    for (addr, fmt) in formats {
                        self.wb_mut().cell_formats.insert(addr, fmt);
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
                    shift_cells_cols(&mut self.wb_mut().cells, sheet, at + 1, -1);
                }
            }
            JournalEntry::RangeRestore { cells, formats } => {
                for (addr, contents) in cells {
                    self.push_to_engine_at(addr, &contents);
                    self.wb_mut().cells.insert(addr, contents);
                }
                for (addr, fmt) in formats {
                    self.wb_mut().cell_formats.insert(addr, fmt);
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
            JournalEntry::ColWidth { sheet, col, prev_width } => {
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
    /// flag.
    fn do_recalc(&mut self) {
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.recalc_pending = false;
    }

    // ---------------- MENU mode ----------------

    fn open_menu(&mut self) {
        self.menu = Some(MenuState::fresh());
        self.mode = Mode::Menu;
    }

    fn close_menu(&mut self) {
        self.menu = None;
        self.mode = Mode::Ready;
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

    fn descend_highlighted(&mut self) {
        let Some(state) = self.menu.as_ref() else { return };
        let Some(item) = state.highlighted() else { return };
        self.descend_into(item);
    }

    fn descend_by_letter(&mut self, c: char) {
        let Some(state) = self.menu.as_ref() else { return };
        let level = state.level();
        let item = level.iter().find(|m| m.letter.eq_ignore_ascii_case(&c)).copied();
        if let Some(item) = item {
            self.descend_into(&item);
        }
    }

    fn execute_action(&mut self, action: Action) {
        match action {
            Action::Cancel => self.close_menu(),
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
            Action::WorksheetColumnSetWidth => self.start_col_width_prompt(),
            Action::RangeNameCreate => self.start_name_prompt(
                "Enter name:",
                PromptNext::RangeNameCreate,
            ),
            Action::RangeNameDelete => self.start_name_prompt(
                "Enter name to delete:",
                PromptNext::RangeNameDelete,
            ),
            Action::RangeErase => self.begin_point(PendingCommand::RangeErase),
            Action::Copy => self.begin_point(PendingCommand::CopyFrom),
            Action::Move => self.begin_point(PendingCommand::MoveFrom),
            Action::RangeLabelLeft => self
                .begin_point(PendingCommand::RangeLabel { new_prefix: LabelPrefix::Apostrophe }),
            Action::RangeLabelRight => self
                .begin_point(PendingCommand::RangeLabel { new_prefix: LabelPrefix::Quote }),
            Action::RangeLabelCenter => self
                .begin_point(PendingCommand::RangeLabel { new_prefix: LabelPrefix::Caret }),
            Action::RangeFormatFixed => self.start_decimals_prompt(FormatKind::Fixed),
            Action::RangeFormatScientific => self.start_decimals_prompt(FormatKind::Scientific),
            Action::RangeFormatCurrency => self.start_decimals_prompt(FormatKind::Currency),
            Action::RangeFormatComma => self.start_decimals_prompt(FormatKind::Comma),
            Action::RangeFormatPercent => self.start_decimals_prompt(FormatKind::Percent),
            Action::RangeFormatGeneral => self.begin_point(PendingCommand::RangeFormat {
                format: Format::GENERAL,
            }),
            Action::RangeFormatReset => self.begin_point(PendingCommand::RangeFormat {
                format: Format::RESET,
            }),
            Action::RangeFormatText => self.begin_point(PendingCommand::RangeFormat {
                format: Format { kind: FormatKind::Text, decimals: 0 },
            }),
            Action::FileSave => self.start_file_save_prompt(),
            Action::FileRetrieve => self.start_file_retrieve_prompt(),
            Action::FileXtractFormulas => self.start_file_xtract_prompt(XtractKind::Formulas),
            Action::FileXtractValues => self.start_file_xtract_prompt(XtractKind::Values),
            Action::FileImportNumbers => self.start_file_import_numbers_prompt(),
            Action::FileNew => self.execute_file_new(),
            Action::FileOpenBefore => self.start_file_open_prompt(true),
            Action::FileOpenAfter => self.start_file_open_prompt(false),
            Action::PrintFile => self.start_print_file_prompt(),
            Action::PrintFileRange => self.begin_point(PendingCommand::PrintFileRange),
            Action::PrintFileGo => self.execute_print_file_go(),
            Action::PrintFileQuit => self.finish_print_session(),
            Action::PrintFileAlign => {
                if let Some(s) = self.print.as_mut() {
                    s.next_page = 1;
                }
                self.enter_print_file_menu();
            }
            Action::PrintFileClear => {
                if let Some(s) = self.print.as_mut() {
                    s.clear_all();
                }
                self.enter_print_file_menu();
            }
            Action::PrintFileOptionsHeader => self.start_print_header_prompt(),
            Action::PrintFileOptionsFooter => self.start_print_footer_prompt(),
            Action::PrintFileOptionsQuit => self.enter_print_file_menu(),
            Action::PrintFileOptionsOtherAsDisplayed => {
                self.set_print_content_mode(PrintContentMode::AsDisplayed)
            }
            Action::PrintFileOptionsOtherCellFormulas => {
                self.set_print_content_mode(PrintContentMode::CellFormulas)
            }
            Action::PrintFileOptionsOtherFormatted => {
                self.set_print_format_mode(PrintFormatMode::Formatted)
            }
            Action::PrintFileOptionsOtherUnformatted => {
                self.set_print_format_mode(PrintFormatMode::Unformatted)
            }
            Action::PrintFileOptionsMarginLeft => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginLeft, "left")
            }
            Action::PrintFileOptionsMarginRight => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginRight, "right")
            }
            Action::PrintFileOptionsMarginTop => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginTop, "top")
            }
            Action::PrintFileOptionsMarginBottom => {
                self.start_print_margin_prompt(PromptNext::PrintFileMarginBottom, "bottom")
            }
            Action::PrintFileOptionsMarginsQuit => self.enter_print_options_menu(),
            Action::PrintFileOptionsPgLength => {
                self.start_print_pg_length_prompt();
            }
            Action::RangeSearchFormulas => {
                self.begin_point(PendingCommand::RangeSearchRange {
                    scope: SearchScope::Formulas,
                })
            }
            Action::RangeSearchLabels => {
                self.begin_point(PendingCommand::RangeSearchRange {
                    scope: SearchScope::Labels,
                })
            }
            Action::RangeSearchBoth => {
                self.begin_point(PendingCommand::RangeSearchRange {
                    scope: SearchScope::Both,
                })
            }
            Action::RangeSearchFind => self.execute_range_search_find(),
            Action::RangeSearchReplace => self.start_range_search_replace_prompt(),
            Action::FileDir => self.start_file_dir_prompt(),
            Action::FileListWorksheet => self.open_file_list(FileListKind::Worksheet),
            Action::FileListActive => self.open_file_list(FileListKind::Active),
            // Subsequent cycles wire these. For now, descent is a no-op.
            _ => self.close_menu(),
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
        };
        self.file_list = Some(FileListState {
            kind,
            entries,
            highlight: 0,
            view_offset: 0,
        });
        self.mode = Mode::Files;
    }

    fn handle_key_file_list(&mut self, k: KeyEvent) {
        let Some(fl) = self.file_list.as_mut() else { return };
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
                    fl.highlight = (fl.highlight + FILE_LIST_PAGE_SIZE)
                        .min(fl.entries.len() - 1);
                }
            }
            KeyCode::Home => fl.highlight = 0,
            KeyCode::End => {
                if !fl.entries.is_empty() {
                    fl.highlight = fl.entries.len() - 1;
                }
            }
            KeyCode::Enter => {
                let Some(fl) = self.file_list.take() else { return };
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
                }
            }
            _ => return,
        }
        if let Some(fl) = self.file_list.as_mut() {
            adjust_file_list_view(fl);
        }
    }

    /// /FN — wipe the current workbook back to a blank slate. Both the
    /// Before and After branches collapse to this same reset for now;
    /// true multi-file insertion is M5.
    fn execute_file_new(&mut self) {
        self.wb_mut().cells.clear();
        self.wb_mut().cell_formats.clear();
        self.wb_mut().col_widths.clear();
        self.entry = None;
        self.menu = None;
        self.prompt = None;
        self.point = None;
        self.save_confirm = None;
        self.pending_name = None;
        self.pending_xtract_path = None;
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

    fn finish_print_session(&mut self) {
        self.print = None;
        self.close_menu();
    }

    fn execute_print_file_go(&mut self) {
        let Some(session) = self.print.as_ref() else {
            self.close_menu();
            return;
        };
        let Some(range) = session.range else {
            // No range selected — bounce back to the menu without
            // writing anything. Matches 1-2-3's "Go with no range =
            // no-op".
            self.enter_print_file_menu();
            return;
        };
        let path = session.path.clone();
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                let _ = std::fs::create_dir_all(parent);
            }
        }
        let settings = PrintSettings {
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
        let (body, pages) = self.render_print_body(range, &settings);
        let _ = std::fs::write(&path, body);
        // Session stays alive so the user can issue further commands
        // (Options, another Go, Align, Clear, …). Quit is the way
        // out. Advance the page counter for the next Go.
        if let Some(s) = self.print.as_mut() {
            s.next_page = s.next_page.saturating_add(pages);
        }
        self.enter_print_file_menu();
    }

    /// Render the selected range as plain ASCII, one line per row.
    /// Each cell gets its configured column width; trailing spaces
    /// are trimmed per row. Optional `header` / `footer` surround
    /// the body as `|`-split L|C|R strings.
    /// Returns the rendered body and the number of pages written so
    /// the caller can advance its page counter.
    fn render_print_body(&self, range: Range, settings: &PrintSettings) -> (String, u32) {
        let r = range.normalized();
        // Page width = sum of column widths in the selected range.
        let page_width: usize = (r.start.col..=r.end.col)
            .map(|c| self.col_width_of(r.start.sheet, c) as usize)
            .sum();
        let page_width = page_width.max(1);
        // Right margin trims content lines (after the left pad) down
        // to `page_width - margin_right`. `0` = no limit.
        let effective_width = if settings.margin_right == 0 {
            page_width
        } else {
            page_width.saturating_sub(settings.margin_right as usize)
        };
        let (header, footer) = match settings.format_mode {
            PrintFormatMode::Formatted => (settings.header.as_str(), settings.footer.as_str()),
            PrintFormatMode::Unformatted => ("", ""),
        };
        let content_mode = settings.content_mode;
        let left_pad: String = " ".repeat(settings.margin_left as usize);

        // Collect content rows (after pipe-row suppression, content-
        // mode rendering, and right-margin truncation). Each entry
        // already has `left_pad` prepended and a trailing `\n`.
        let mut rows: Vec<String> = Vec::new();
        for row in r.start.row..=r.end.row {
            let first = Address::new(r.start.sheet, r.start.col, row);
            if let Some(CellContents::Label { prefix: LabelPrefix::Pipe, .. }) =
                self.wb().cells.get(&first)
            {
                continue;
            }
            let mut line = String::new();
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                let w = self.col_width_of(r.start.sheet, col) as usize;
                let piece = match self.wb().cells.get(&addr) {
                    Some(CellContents::Empty) | None => " ".repeat(w),
                    Some(CellContents::Label { prefix, text }) => {
                        render_label(*prefix, text, w)
                    }
                    Some(CellContents::Constant(v)) => {
                        let fmt = self.format_for_cell(addr);
                        render_value_in_cell(v, w, fmt).unwrap_or_else(|| " ".repeat(w))
                    }
                    Some(CellContents::Formula { expr, cached_value }) => {
                        match content_mode {
                            PrintContentMode::CellFormulas => {
                                let src = format!("@{expr}");
                                let pad = w.saturating_sub(src.chars().count());
                                let mut s = src;
                                s.extend(std::iter::repeat_n(' ', pad));
                                s
                            }
                            PrintContentMode::AsDisplayed => match cached_value {
                                Some(v) => {
                                    let fmt = self.format_for_cell(addr);
                                    render_value_in_cell(v, w, fmt)
                                        .unwrap_or_else(|| " ".repeat(w))
                                }
                                None => " ".repeat(w),
                            },
                        }
                    }
                };
                line.push_str(&piece);
            }
            let truncated: String = line.chars().take(effective_width).collect();
            let trimmed: String = truncated.trim_end().to_string();
            let mut entry = String::with_capacity(left_pad.len() + trimmed.len() + 1);
            entry.push_str(&left_pad);
            entry.push_str(&trimmed);
            entry.push('\n');
            rows.push(entry);
        }

        // Chunk into pages. pg_length == 0 means no pagination — one
        // page with every row.
        let per_page = if settings.pg_length == 0 {
            rows.len().max(1)
        } else {
            settings.pg_length as usize
        };
        let pages: Vec<&[String]> = if rows.is_empty() {
            vec![&[]]
        } else {
            rows.chunks(per_page).collect()
        };

        let today = today_ddmmmyy();
        let mut out = String::new();
        for (i, page) in pages.iter().enumerate() {
            let page_no = settings.start_page as usize + i;
            // Top margin.
            for _ in 0..settings.margin_top {
                out.push('\n');
            }
            if !header.is_empty() {
                let substituted = substitute_tokens(header, page_no, &today);
                out.push_str(&left_pad);
                out.push_str(&format_three_part(&substituted, page_width));
                out.push('\n');
                out.push('\n');
            }
            for row in page.iter() {
                out.push_str(row);
            }
            if !footer.is_empty() {
                let substituted = substitute_tokens(footer, page_no, &today);
                out.push('\n');
                out.push_str(&left_pad);
                out.push_str(&format_three_part(&substituted, page_width));
                out.push('\n');
            }
            // Bottom margin.
            for _ in 0..settings.margin_bottom {
                out.push('\n');
            }
            // Form feed between pages (not after the last).
            if i + 1 < pages.len() {
                out.push('\x0c');
            }
        }
        (out, pages.len() as u32)
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
        for (addr, cv) in engine.used_cells() {
            if let Some(contents) = cell_view_to_contents(&cv) {
                cells.insert(addr, contents);
            }
        }
        let new_file = Workbook {
            engine,
            cells,
            cell_formats: HashMap::new(),
            col_widths: HashMap::new(),
            active_path: Some(path),
            pointer: Address::A1,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            journal: Vec::new(),
        };
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

    /// Read `path` as CSV, paint values into cells starting at the
    /// pointer. Numeric tokens become `Constant(Number)`; everything
    /// else becomes `Label { Apostrophe, text }`. Empty fields are
    /// skipped (no overwrite).
    fn import_numbers_from(&mut self, path: PathBuf) {
        let Ok(body) = std::fs::read_to_string(&path) else {
            self.mode = Mode::Ready;
            return;
        };
        let rows = l123_io::csv::parse(&body);
        let origin = self.wb_mut().pointer;
        for (dr, row) in rows.iter().enumerate() {
            for (dc, field) in row.iter().enumerate() {
                if field.is_empty() {
                    continue;
                }
                let addr = Address::new(
                    origin.sheet,
                    origin.col + dc as u16,
                    origin.row + dr as u32,
                );
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
        self.mode = Mode::Ready;
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

    /// Load an xlsx from disk, wiping the current in-memory workbook
    /// and repopulating the UI cache from the loaded engine model.
    fn load_workbook_from(&mut self, path: PathBuf) {
        if self.wb_mut().engine.load_xlsx(&path).is_err() {
            self.mode = Mode::Ready;
            return;
        }
        // Wipe UI state; the loaded engine is the new source of truth.
        self.wb_mut().cells.clear();
        self.wb_mut().cell_formats.clear();
        self.wb_mut().col_widths.clear();
        self.entry = None;
        self.wb_mut().pointer = Address::A1;
        self.wb_mut().viewport_col_offset = 0;
        self.wb_mut().viewport_row_offset = 0;
        self.recalc_pending = false;

        // Pull every non-empty cell into the UI cache.
        for (addr, cv) in self.wb_mut().engine.used_cells() {
            if let Some(contents) = cell_view_to_contents(&cv) {
                self.wb_mut().cells.insert(addr, contents);
            }
        }

        self.wb_mut().active_path = Some(path);
        self.mode = Mode::Ready;
    }

    /// Create parent dirs and write the workbook as xlsx. On success,
    /// update `active_path` so the next /FS prefills this path.
    fn save_workbook_to(&mut self, path: PathBuf) {
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                let _ = std::fs::create_dir_all(parent);
            }
        }
        if self.wb_mut().engine.save_xlsx(&path).is_ok() {
            self.wb_mut().active_path = Some(path);
        }
    }

    /// Handle a keystroke while the Cancel/Replace/Backup confirm is up.
    fn handle_key_save_confirm(&mut self, k: KeyEvent) {
        let Some(sc) = self.save_confirm.as_mut() else { return };
        match k.code {
            KeyCode::Esc => {
                self.save_confirm = None;
                self.mode = Mode::Ready;
            }
            KeyCode::Left => {
                if sc.highlight > 0 {
                    sc.highlight -= 1;
                }
            }
            KeyCode::Right => {
                if sc.highlight + 1 < SAVE_CONFIRM_ITEMS.len() {
                    sc.highlight += 1;
                }
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
                // Replace — overwrite.
                self.save_workbook_to(sc.path);
                self.mode = Mode::Ready;
            }
            2 => {
                // Backup — rename existing to .BAK, then save.
                let backup = sc.path.with_extension("BAK");
                let _ = std::fs::rename(&sc.path, &backup);
                self.save_workbook_to(sc.path);
                self.mode = Mode::Ready;
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
            let captured_cells: Vec<(Address, CellContents)> = self.wb_mut().cells
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.row >= at && a.row < at + n)
                .map(|(a, c)| (*a, c.clone()))
                .collect();
            let captured_formats: Vec<(Address, Format)> = self.wb_mut().cell_formats
                .iter()
                .filter(|(a, _)| a.sheet == sheet && a.row >= at && a.row < at + n)
                .map(|(a, f)| (*a, *f))
                .collect();
            if self.wb_mut().engine.delete_rows(sheet, at, n).is_ok() {
                self.wb_mut().cells
                    .retain(|a, _| !(a.sheet == sheet && a.row >= at && a.row < at + n));
                self.wb_mut().cell_formats
                    .retain(|a, _| !(a.sheet == sheet && a.row >= at && a.row < at + n));
                shift_cells_rows(&mut self.wb_mut().cells, sheet, at + n, -(n as i64));
                batch.push(JournalEntry::RowDelete {
                    sheet,
                    at,
                    cells: captured_cells,
                    formats: captured_formats,
                });
            }
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
                &mut wb.col_widths,
                at,
                1,
            );
            wb.engine.recalc();
            self.refresh_formula_caches();
        }
        self.close_menu();
    }

    /// Ctrl-PgDn / Ctrl-PgUp: jump to the next / previous sheet. Clamps
    /// at the bookends — no wrap.
    fn move_sheet(&mut self, delta: i32) {
        let count = self.wb().engine.sheet_count();
        if count == 0 {
            return;
        }
        let cur = self.wb().pointer.sheet.0 as i32;
        let next = (cur + delta).clamp(0, count as i32 - 1) as u16;
        let wb = self.wb_mut();
        if next != wb.pointer.sheet.0 {
            wb.pointer = Address::new(SheetId(next), 0, 0);
            wb.viewport_col_offset = 0;
            wb.viewport_row_offset = 0;
        }
    }

    // ---------------- POINT mode ----------------

    fn begin_point(&mut self, pending: PendingCommand) {
        self.menu = None;
        self.point = Some(PointState { anchor: Some(self.wb().pointer), pending });
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
            Some(anchor) => Range { start: anchor, end: self.wb().pointer }.normalized(),
            None => Range::single(self.wb().pointer),
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
            KeyCode::Char('.') => self.period_in_point(),
            _ => {}
        }
    }

    /// Esc during POINT: first press unanchors (pointer moves free); second
    /// press cancels the pending command back to READY.
    fn esc_in_point(&mut self) {
        let Some(ps) = self.point.as_mut() else { return };
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
        let Some(ps) = self.point.take() else {
            self.mode = Mode::Ready;
            return;
        };
        let range = match ps.anchor {
            Some(a) => Range { start: a, end: self.wb_mut().pointer }.normalized(),
            None => Range::single(self.wb_mut().pointer),
        };
        match ps.pending {
            PendingCommand::RangeErase => {
                self.execute_range_erase(range);
                self.mode = Mode::Ready;
            }
            PendingCommand::CopyFrom => self.transition_point(PendingCommand::CopyTo { source: range }),
            PendingCommand::MoveFrom => self.transition_point(PendingCommand::MoveTo { source: range }),
            PendingCommand::CopyTo { source } => {
                // Destination anchor is the pointer's current position,
                // regardless of how the TO range was painted — for M3
                // this is the common case (single-cell destination anchor).
                let dest = self.wb_mut().pointer;
                self.execute_copy(source, dest);
                self.mode = Mode::Ready;
            }
            PendingCommand::MoveTo { source } => {
                let dest = self.wb_mut().pointer;
                self.execute_move(source, dest);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeLabel { new_prefix } => {
                self.execute_range_label(range, new_prefix);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeFormat { format } => {
                self.execute_range_format(range, format);
                self.mode = Mode::Ready;
            }
            PendingCommand::RangeNameCreate => {
                if let Some(name) = self.pending_name.take() {
                    let _ = self.wb_mut().engine.define_name(&name, range);
                    self.wb_mut().engine.recalc();
                    self.refresh_formula_caches();
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::FileXtractRange { kind } => {
                if let Some(path) = self.pending_xtract_path.take() {
                    self.execute_file_xtract(range, kind, path);
                }
                self.mode = Mode::Ready;
            }
            PendingCommand::PrintFileRange => {
                if let Some(session) = self.print.as_mut() {
                    session.range = Some(range);
                }
                // Back to the /PF submenu for Options/Go/…
                self.enter_print_file_menu();
            }
            PendingCommand::RangeSearchRange { scope } => {
                self.start_range_search_string_prompt(scope, range);
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
                let Some(contents) = self.wb().cells.get(&addr) else { continue };
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
                    (SearchScope::Both, CellContents::Label { text, .. }) => {
                        text.contains(needle)
                    }
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

    /// Width of `col` in `sheet`. Returns the stored width, or the default
    /// (9) if none is set.
    fn col_width_of(&self, sheet: SheetId, col: u16) -> u8 {
        self.wb().col_widths.get(&(sheet, col)).copied().unwrap_or(9)
    }

    fn handle_key_prompt(&mut self, k: KeyEvent) {
        match k.code {
            KeyCode::Enter => self.commit_prompt(),
            KeyCode::Esc => self.cancel_prompt(),
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

    fn cancel_prompt(&mut self) {
        self.prompt = None;
        self.mode = Mode::Ready;
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
                let format = Format { kind, decimals };
                self.begin_point(PendingCommand::RangeFormat { format });
            }
            PromptNext::WorksheetColumnSetWidth => {
                let width: u8 = p.buffer.parse().unwrap_or(9).clamp(1, 240);
                let col = self.wb().pointer.col;
                let mut batch: Vec<JournalEntry> = Vec::new();
                for sheet in self.target_sheets() {
                    let key = (sheet, col);
                    let prev = self.wb().col_widths.get(&key).copied();
                    if width == 9 {
                        self.wb_mut().col_widths.remove(&key);
                    } else {
                        self.wb_mut().col_widths.insert(key, width);
                    }
                    batch.push(JournalEntry::ColWidth { sheet, col, prev_width: prev });
                }
                self.push_journal_batch(batch);
                self.mode = Mode::Ready;
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
                    self.wb_mut().engine.recalc();
                    self.refresh_formula_caches();
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
                    self.save_workbook_to(path);
                    self.mode = Mode::Ready;
                }
            }
            PromptNext::FileRetrieveFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.load_workbook_from(path);
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
                self.import_numbers_from(path);
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
                self.print = Some(PrintSession::new(path));
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
            let Some(contents) = self.wb().cells.get(&addr).cloned() else { continue };
            let (new_contents, prev_contents, prev_format) = match contents {
                CellContents::Formula { expr, cached_value } => {
                    let new_expr = expr.replace(&needle, &replacement);
                    let prev = CellContents::Formula {
                        expr: expr.clone(),
                        cached_value: cached_value.clone(),
                    };
                    (
                        CellContents::Formula { expr: new_expr, cached_value: None },
                        Some(prev),
                        self.wb().cell_formats.get(&addr).copied(),
                    )
                }
                CellContents::Label { prefix, text } => {
                    let new_text = text.replace(&needle, &replacement);
                    let prev = CellContents::Label { prefix, text: text.clone() };
                    (
                        CellContents::Label { prefix, text: new_text },
                        Some(prev),
                        self.wb().cell_formats.get(&addr).copied(),
                    )
                }
                _ => continue,
            };
            batch.push(JournalEntry::CellEdit { addr, prev_contents, prev_format });
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
                        self.wb_mut().cell_formats.remove(&addr);
                    } else {
                        self.wb_mut().cell_formats.insert(addr, format);
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

    fn execute_range_label(&mut self, range: Range, new_prefix: LabelPrefix) {
        let r = range.normalized();
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(r.start.sheet, col, row);
                if let Some(CellContents::Label { prefix, .. }) = self.wb_mut().cells.get_mut(&addr) {
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
            _ => self.wb_mut().pointer,
        };
        self.wb_mut().pointer = source_tl;
        self.scroll_into_view();
        self.point = Some(PointState {
            anchor: Some(self.wb_mut().pointer),
            pending: next,
        });
        // mode stays POINT
    }

    fn execute_copy(&mut self, source: Range, dest_anchor: Address) {
        let src_cells = self.collect_cells_in_range(source);
        self.write_cells_at_offset(&src_cells, source.start, dest_anchor);
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
    }

    fn execute_move(&mut self, source: Range, dest_anchor: Address) {
        let src_cells = self.collect_cells_in_range(source);
        self.write_cells_at_offset(&src_cells, source.start, dest_anchor);
        // Clear the source cells (but only ones not overlapping the destination).
        let dest_range = Range {
            start: dest_anchor,
            end: Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (source.end.col - source.start.col),
                dest_anchor.row + (source.end.row - source.start.row),
            ),
        };
        for (src, _) in &src_cells {
            if !dest_range.contains(*src) {
                self.wb_mut().cells.remove(src);
                let _ = self.wb_mut().engine.clear_cell(*src);
            }
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
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

    /// Write cells to their new positions offset so that `src_origin` maps
    /// to `dest_anchor`. Formulas are pushed as-is (reference adjustment on
    /// copy is deferred to a later cycle).
    fn write_cells_at_offset(
        &mut self,
        cells: &[(Address, CellContents)],
        src_origin: Address,
        dest_anchor: Address,
    ) {
        for (src, contents) in cells {
            let dst = Address::new(
                dest_anchor.sheet,
                dest_anchor.col + (src.col - src_origin.col),
                dest_anchor.row + (src.row - src_origin.row),
            );
            self.wb_mut().cells.insert(dst, contents.clone());
            self.push_to_engine_at(dst, contents);
        }
    }

    /// Like `push_to_engine` but for an arbitrary address (not
    /// `self.wb_mut().pointer`). Used during Copy/Move.
    fn push_to_engine_at(&mut self, addr: Address, contents: &CellContents) {
        let result = match contents {
            CellContents::Empty => self.wb_mut().engine.clear_cell(addr),
            CellContents::Label { text, .. } => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{text}"))
            }
            CellContents::Constant(Value::Number(n)) => self.wb_mut().engine
                .set_user_input(addr, &l123_core::format_number_general(*n)),
            CellContents::Constant(Value::Text(s)) => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{s}"))
            }
            CellContents::Constant(_) => Ok(()),
            CellContents::Formula { expr, .. } => {
                let names = self.wb_mut().engine.all_sheet_names();
                let names_ref: Vec<&str> = names.iter().map(String::as_str).collect();
                let excel = l123_parse::to_engine_source(expr, &names_ref);
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
        for row in r.start.row..=r.end.row {
            for col in r.start.col..=r.end.col {
                let addr = Address::new(sheet, col, row);
                if let Some(c) = self.wb().cells.get(&addr) {
                    cells.push((addr, c.clone()));
                }
                if let Some(f) = self.wb().cell_formats.get(&addr) {
                    formats.push((addr, *f));
                }
                self.wb_mut().cells.remove(&addr);
                self.wb_mut().cell_formats.remove(&addr);
                let _ = self.wb_mut().engine.clear_cell(addr);
            }
        }
        if self.undo_enabled && (!cells.is_empty() || !formats.is_empty()) {
            self.wb_mut()
                .journal
                .push(JournalEntry::RangeRestore { cells, formats });
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
            if self.wb_mut().engine.delete_cols(sheet, at, n).is_ok() {
                self.wb_mut()
                    .cells
                    .retain(|a, _| !(a.sheet == sheet && a.col >= at && a.col < at + n));
                self.wb_mut()
                    .cell_formats
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
                    batch.push(JournalEntry::ColDelete {
                        sheet,
                        at: col_k,
                        cells: cells_k,
                        formats: formats_k,
                    });
                }
            }
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
            CellContents::Label { text, .. } => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{text}"))
            }
            CellContents::Constant(Value::Number(n)) => self.wb_mut().engine
                .set_user_input(addr, &l123_core::format_number_general(*n)),
            CellContents::Constant(Value::Text(s)) => {
                self.wb_mut().engine.set_user_input(addr, &format!("'{s}"))
            }
            CellContents::Constant(_) => Ok(()),
            CellContents::Formula { expr, .. } => {
                let names = self.wb_mut().engine.all_sheet_names();
                let names_ref: Vec<&str> = names.iter().map(String::as_str).collect();
                let excel = l123_parse::to_engine_source(expr, &names_ref);
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
        let formula_addrs: Vec<Address> = self.wb_mut().cells
            .iter()
            .filter_map(|(addr, c)| matches!(c, CellContents::Formula { .. }).then_some(*addr))
            .collect();
        for addr in formula_addrs {
            let Ok(view) = self.wb_mut().engine.get_cell(addr) else { continue };
            if let Some(CellContents::Formula { cached_value, .. }) = self.wb_mut().cells.get_mut(&addr) {
                *cached_value = Some(view.value);
            }
        }
    }

    fn move_pointer(&mut self, d_col: i32, d_row: i32) {
        if let Some(next) = self.wb_mut().pointer.shifted(d_col, d_row) {
            self.wb_mut().pointer = next;
            self.scroll_into_view();
        }
    }

    fn scroll_into_view(&mut self) {
        if self.wb_mut().pointer.col < self.wb_mut().viewport_col_offset {
            self.wb_mut().viewport_col_offset = self.wb_mut().pointer.col;
        }
        if self.wb_mut().pointer.row < self.wb_mut().viewport_row_offset {
            self.wb_mut().viewport_row_offset = self.wb_mut().pointer.row;
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
        if self.file_list.is_some() {
            self.render_file_list_overlay(chunks[1], buf);
        } else {
            self.render_grid(chunks[1], buf);
        }
        self.render_status(chunks[2], buf);
    }

    fn render_control_panel(&self, area: Rect, buf: &mut Buffer) {
        let block = Block::default().borders(Borders::BOTTOM);
        let inner = block.inner(area);
        block.render(area, buf);

        // Line 1: "<addr>: [(fmt) [Wn]] <readout>" left; mode indicator right.
        let readout = self.cell_readout_for_line1();
        let format_tag = self.format_tag_for_line1();
        let width_tag = self.width_tag_for_line1();
        let mut tags: Vec<&str> = Vec::new();
        if !format_tag.is_empty() {
            tags.push(&format_tag);
        }
        if !width_tag.is_empty() {
            tags.push(&width_tag);
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
        let pad = (area.width as usize)
            .saturating_sub(left.chars().count() + mode_str.len() + 1);
        let line1 = Line::from(vec![
            Span::raw(left),
            Span::raw(" ".repeat(pad)),
            Span::styled(mode_str, Style::default().fg(Color::Yellow).add_modifier(Modifier::BOLD)),
            Span::raw(" "),
        ]);

        // Line 2 & 3 depend on mode. File-list and save-confirm take
        // absolute precedence (they own the keyboard); then a
        // command-argument prompt; then mode-specific rendering.
        let (line2, line3) = if self.file_list.is_some() {
            self.render_file_list_lines()
        } else if self.save_confirm.is_some() {
            self.render_save_confirm_lines()
        } else if let Some(p) = self.prompt.as_ref() {
            (
                Line::from(format!(" {}", p.buffer)),
                Line::from(format!(" {}", p.label)),
            )
        } else {
            match self.mode {
                Mode::Menu => self.render_menu_lines(),
                Mode::Point => self.render_point_lines(),
                _ => {
                    let l2 = match self.entry.as_ref() {
                        Some(e) => Line::from(format!(" {}", e.buffer)),
                        None => Line::from(""),
                    };
                    (l2, Line::from(""))
                }
            }
        };

        Paragraph::new(vec![line1, line2, line3]).render(inner, buf);
    }

    fn render_file_list_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(fl) = self.file_list.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        // Panel lines are just the header + highlighted path / count.
        // The full picker lives in the overlay below.
        let header = match fl.kind {
            FileListKind::Worksheet => " File List — Worksheet",
            FileListKind::Active => " File List — Active",
        };
        let tail = if fl.entries.is_empty() {
            match fl.kind {
                FileListKind::Worksheet => " (no worksheet files in directory)".to_string(),
                FileListKind::Active => " (no active file)".to_string(),
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
    fn render_file_list_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(fl) = self.file_list.as_ref() else { return };
        let width = area.width as usize;
        let rows = area.height as usize;
        if rows == 0 || width == 0 {
            return;
        }

        let size_col_width: usize = 10;
        let name_col_width = width.saturating_sub(size_col_width + 3);

        // Header row.
        let header = format_file_list_row(
            "NAME",
            "SIZE",
            name_col_width,
            size_col_width,
            width,
        );
        set_line(buf, area.x, area.y, &header, area.width, Style::default());

        if fl.entries.is_empty() {
            let empty_msg = match fl.kind {
                FileListKind::Worksheet => "(no worksheet files in directory)",
                FileListKind::Active => "(no active file)",
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
            let row = format_file_list_row(
                &name,
                &size,
                name_col_width,
                size_col_width,
                width,
            );
            let style = if idx == fl.highlight {
                Style::default().add_modifier(Modifier::REVERSED)
            } else {
                Style::default()
            };
            set_line(buf, area.x, area.y + 1 + i as u16, &row, area.width, style);
        }
    }

    fn render_save_confirm_lines(&self) -> (Line<'_>, Line<'_>) {
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

    fn render_point_lines(&self) -> (Line<'_>, Line<'_>) {
        let Some(ps) = self.point.as_ref() else {
            return (Line::from(""), Line::from(""));
        };
        let range = self.highlight_range();
        let range_str = format!(
            "{}..{}",
            range.start.display_full(),
            range.end.display_full()
        );
        let line3_text = format!(" {} {}", ps.pending.prompt(), range_str);
        (Line::from(""), Line::from(line3_text))
    }

    fn render_menu_lines(&self) -> (Line<'_>, Line<'_>) {
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

    fn cell_readout_for_line1(&self) -> String {
        self.wb().cells
            .get(&self.wb().pointer)
            .map(|c| c.control_panel_readout())
            .unwrap_or_default()
    }

    /// Parenthesized format tag (e.g. `(G)`, `(C2)`, `(F3)`) for the current
    /// cell. Empty for labels, empty cells, or cells using `Reset` format.
    fn format_tag_for_line1(&self) -> String {
        match self.wb().cells.get(&self.wb().pointer) {
            Some(CellContents::Constant(Value::Number(_)))
            | Some(CellContents::Formula { .. }) => match self.format_for_cell(self.wb().pointer).tag() {
                Some(s) => format!("({s})"),
                None => String::new(),
            },
            _ => String::new(),
        }
    }

    /// Resolve the format for a given cell — the per-cell override if set,
    /// else General.
    fn format_for_cell(&self, addr: Address) -> Format {
        self.wb().cell_formats
            .get(&addr)
            .copied()
            .unwrap_or(Format::GENERAL)
    }

    /// `[Wn]` tag when the current column's width is non-default.
    fn width_tag_for_line1(&self) -> String {
        let w = self.col_width_of(self.wb().pointer.sheet, self.wb().pointer.col);
        if w == 9 {
            String::new()
        } else {
            format!("[W{w}]")
        }
    }

    fn render_grid(&self, area: Rect, buf: &mut Buffer) {
        if area.width <= ROW_GUTTER || area.height < 2 {
            return;
        }

        let visible_cols = ((area.width - ROW_GUTTER) / COL_WIDTH).max(1);
        let visible_rows = area.height - 1;

        // Column header row
        for i in 0..visible_cols {
            let c = self.wb().viewport_col_offset + i;
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
            let row_idx = self.wb().viewport_row_offset + r as u32;
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

            // During POINT, highlight the whole range; otherwise just the pointer.
            let highlight = if self.mode == Mode::Point {
                self.highlight_range()
            } else {
                Range::single(self.wb().pointer)
            };
            for c in 0..visible_cols {
                let col_idx = self.wb().viewport_col_offset + c;
                let x = area.x + ROW_GUTTER + c * COL_WIDTH;
                let addr = Address::new(self.wb().pointer.sheet, col_idx, row_idx);
                let highlighted = highlight.contains(addr);
                let cell_style = if highlighted {
                    Style::default().add_modifier(Modifier::REVERSED)
                } else {
                    Style::default()
                };
                // Blank background first
                for k in 0..COL_WIDTH {
                    buf[(x + k, y)].set_char(' ').set_style(cell_style);
                }
                // Content
                if let Some(contents) = self.wb().cells.get(&addr) {
                    let fmt = self.format_for_cell(addr);
                    draw_cell_contents(buf, x, y, COL_WIDTH, contents, cell_style, fmt);
                }
            }
        }
    }

    fn render_status(&self, area: Rect, buf: &mut Buffer) {
        let left = " *untitled*";
        let hint = "Ctrl-C to quit";
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
        let indicator_str = indicators.join(" ");
        let right_chunk = if indicator_str.is_empty() {
            hint.to_string()
        } else {
            format!("{indicator_str}  {hint}")
        };
        let pad = (area.width as usize)
            .saturating_sub(left.len() + right_chunk.len() + 1);
        let line = format!("{left}{}{right_chunk} ", " ".repeat(pad));
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

/// Shift cells within `sheet` whose row is `>= at` by `delta` rows.
/// Positive delta shifts down (for insert); negative shifts up (for delete,
/// after the deleted rows have already been removed).
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

/// After inserting `delta` sheets at position `at`, every cell whose
/// sheet index is >= `at` moves forward by `delta`. Applies to the
/// three per-sheet caches App keeps in sync with the engine.
fn shift_sheets_from(
    cells: &mut HashMap<Address, CellContents>,
    cell_formats: &mut HashMap<Address, Format>,
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
    let mut affected: Vec<Address> =
        cells.keys().filter(|a| a.sheet.0 >= at).copied().collect();
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
    let mut cw_affected: Vec<(SheetId, u16)> =
        col_widths.keys().filter(|(s, _)| s.0 >= at).copied().collect();
    cw_affected.sort_by_key(|(s, _)| std::cmp::Reverse(s.0));
    for key in cw_affected {
        let w = col_widths.remove(&key).expect("present");
        col_widths.insert((SheetId(key.0.0 + delta), key.1), w);
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

/// Substitute 1-2-3 header/footer tokens:
///   `#` → `page_no` (as a decimal number)
///   `@` → `today` (pre-formatted DD-MMM-YY date string)
/// Other characters pass through untouched. `\name` (named-range
/// substitution) is a later milestone.
fn substitute_tokens(s: &str, page_no: usize, today: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '#' => out.push_str(&page_no.to_string()),
            '@' => out.push_str(today),
            other => out.push(other),
        }
    }
    out
}

/// Today's date formatted as `DD-MMM-YY` (1-2-3's default D1 date
/// format). Falls back to an empty string if the system clock is
/// somehow earlier than the Unix epoch.
fn today_ddmmmyy() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = match SystemTime::now().duration_since(UNIX_EPOCH) {
        Ok(d) => d.as_secs() as i64,
        Err(_) => return String::new(),
    };
    // Convert Unix seconds to (year, month, day) via the civil
    // from-days algorithm from Howard Hinnant's date paper.
    let days = secs.div_euclid(86_400);
    let (y, m, d) = days_to_ymd(days);
    let month_name = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun",
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ][(m - 1) as usize];
    format!("{d:02}-{month_name}-{:02}", (y % 100 + 100) % 100)
}

/// Days since Unix epoch (1970-01-01) → (year, month [1..12],
/// day-of-month [1..31]). Hinnant's civil-from-days algorithm.
fn days_to_ymd(z: i64) -> (i32, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = (z - era * 146_097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = y + if m <= 2 { 1 } else { 0 };
    (y as i32, m, d)
}

/// Format a `|`-split three-part header/footer line to `width`
/// characters. Parts past the third are ignored; missing parts are
/// treated as empty. Left-aligned | centered | right-aligned,
/// truncated to width if over.
fn format_three_part(s: &str, width: usize) -> String {
    let mut parts = s.splitn(3, '|');
    let left = parts.next().unwrap_or("");
    let center = parts.next().unwrap_or("");
    let right = parts.next().unwrap_or("");
    let lcount = left.chars().count();
    let ccount = center.chars().count();
    let rcount = right.chars().count();
    // Lay out Left..padL..Center..padR..Right within `width`.
    if lcount + ccount + rcount >= width {
        // Truncate: concatenate without alignment, cut to width.
        let joined = format!("{left}{center}{right}");
        return joined.chars().take(width).collect();
    }
    let c_start = (width.saturating_sub(ccount)) / 2;
    let c_end = c_start + ccount;
    let r_start = width - rcount;
    let mut out = String::with_capacity(width);
    out.push_str(left);
    let pad1 = c_start.saturating_sub(lcount);
    out.extend(std::iter::repeat_n(' ', pad1));
    out.push_str(center);
    let pad2 = r_start.saturating_sub(c_end);
    out.extend(std::iter::repeat_n(' ', pad2));
    out.push_str(right);
    out
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

/// Render a cell's contents into a fixed-width slot. The cell's `format`
/// controls numeric display (decimals, currency, percent, etc.); labels
/// ignore the format and honor their prefix alignment only.
fn draw_cell_contents(
    buf: &mut Buffer,
    x: u16,
    y: u16,
    width: u16,
    contents: &CellContents,
    style: Style,
    format: Format,
) {
    let w = width as usize;
    let rendered = match contents {
        CellContents::Empty => return,
        CellContents::Label { prefix, text } => render_label(*prefix, text, w),
        CellContents::Constant(v) => match render_value_in_cell(v, w, format) {
            Some(s) => s,
            None => return,
        },
        CellContents::Formula { cached_value: Some(v), .. } => {
            match render_value_in_cell(v, w, format) {
                Some(s) => s,
                None => return,
            }
        }
        // Unevaluated formula: leave blank.
        CellContents::Formula { cached_value: None, .. } => return,
    };
    for (i, ch) in rendered.chars().enumerate().take(w) {
        buf[(x + i as u16, y)].set_char(ch).set_style(style);
    }
}

fn render_value_in_cell(v: &Value, width: usize, format: Format) -> Option<String> {
    match v {
        Value::Number(n) => {
            let s = l123_core::format_number(*n, format);
            // Overflow: if the number doesn't fit the column, fill with
            // asterisks (Authenticity Contract §20.9). General format
            // would normally switch to scientific first; for M3 we
            // short-circuit to asterisks as the visible failure.
            if s.chars().count() > width && !matches!(format.kind, FormatKind::General) {
                Some("*".repeat(width))
            } else {
                Some(right_pad(&s, width, true))
            }
        }
        Value::Text(s) => Some(right_pad(s, width, false)),
        Value::Bool(b) => Some(right_pad(if *b { "TRUE" } else { "FALSE" }, width, true)),
        Value::Error(e) => Some(right_pad(e.lotus_tag(), width, true)),
        Value::Empty => None,
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
    use std::path::Path;

    #[test]
    fn days_to_ymd_known_dates() {
        // Fixtures cover epoch, leap year, century non-leap, and
        // post-2000 leap dates to cross-check the civil algorithm.
        assert_eq!(days_to_ymd(0), (1970, 1, 1));
        assert_eq!(days_to_ymd(31), (1970, 2, 1));
        assert_eq!(days_to_ymd(365), (1971, 1, 1));
        // 2000-02-29 is 11016 days past 1970-01-01.
        assert_eq!(days_to_ymd(11016), (2000, 2, 29));
        assert_eq!(days_to_ymd(11017), (2000, 3, 1));
        // 2026-04-23
        assert_eq!(days_to_ymd(20566), (2026, 4, 23));
    }

    #[test]
    fn substitute_tokens_expands_hash_and_at() {
        assert_eq!(
            substitute_tokens("Page # of 5 (@)", 3, "23-Apr-26"),
            "Page 3 of 5 (23-Apr-26)"
        );
        assert_eq!(substitute_tokens("no tokens", 7, "X"), "no tokens");
    }

    #[test]
    fn starts_at_a1() {
        let app = App::new();
        assert_eq!(app.wb().pointer, Address::A1);
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.entry.is_none());
    }

    #[test]
    fn arrow_nav() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer.col, 1);
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer.row, 1);
        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer.col, 0);
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer.row, 0);
    }

    #[test]
    fn left_from_a1_stays_at_a1() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer, Address::A1);
    }

    #[test]
    fn home_resets_pointer() {
        let mut app = App::new();
        app.wb_mut().pointer = Address::new(SheetId::A, 10, 10);
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer, Address::A1);
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
        assert_eq!(app.wb().pointer.row, 20);
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

    fn make_label(text: &str) -> CellContents {
        CellContents::Label { prefix: LabelPrefix::Apostrophe, text: text.into() }
    }

    fn press(app: &mut App, code: KeyCode) {
        app.handle_key(KeyEvent::new(code, KeyModifiers::NONE));
    }

    fn press_ch(app: &mut App, c: char) {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }

    #[test]
    fn begin_point_auto_anchors_at_pointer() {
        let mut app = App::new();
        app.wb_mut().pointer = Address::new(SheetId::A, 1, 1); // B2
        app.begin_point(PendingCommand::RangeErase);
        assert_eq!(app.mode, Mode::Point);
        let anchor = app.point.as_ref().unwrap().anchor.unwrap();
        assert_eq!(anchor, Address::new(SheetId::A, 1, 1));
        assert_eq!(app.highlight_range(), Range::single(Address::new(SheetId::A, 1, 1)));
    }

    #[test]
    fn point_arrow_expands_range() {
        let mut app = App::new();
        app.begin_point(PendingCommand::RangeErase);
        press(&mut app, KeyCode::Right);
        press(&mut app, KeyCode::Down);
        let r = app.highlight_range();
        assert_eq!(r.start, Address::A1);
        assert_eq!(r.end, Address::new(SheetId::A, 1, 1));
    }

    #[test]
    fn point_esc_twice_cancels() {
        let mut app = App::new();
        app.begin_point(PendingCommand::RangeErase);
        press(&mut app, KeyCode::Down);
        press(&mut app, KeyCode::Esc);
        // Anchor cleared but still in POINT.
        assert_eq!(app.mode, Mode::Point);
        assert!(app.point.as_ref().unwrap().anchor.is_none());
        press(&mut app, KeyCode::Esc);
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.point.is_none());
    }

    #[test]
    fn point_period_anchors_when_unanchored() {
        let mut app = App::new();
        app.begin_point(PendingCommand::RangeErase);
        // Anchor manually cleared
        app.point.as_mut().unwrap().anchor = None;
        press(&mut app, KeyCode::Right);
        press_ch(&mut app, '.');
        let anchor = app.point.as_ref().unwrap().anchor.unwrap();
        assert_eq!(anchor, Address::new(SheetId::A, 1, 0));
    }

    #[test]
    fn point_period_cycles_corner() {
        let mut app = App::new();
        // Anchor at A1; extend to B2 → pointer at BR.
        app.begin_point(PendingCommand::RangeErase);
        press(&mut app, KeyCode::Right);
        press(&mut app, KeyCode::Down);
        let before = app.highlight_range();
        assert_eq!(before.start, Address::A1);
        assert_eq!(before.end, Address::new(SheetId::A, 1, 1));
        // Initially pointer is at BR corner of range. `.` rotates TL→TR→BR→BL.
        // Pointer was at BR(B2), so we're detecting at_min_col=false, at_min_row=false,
        // which the code treats as "BR → BL", moving pointer to BL (A2).
        press_ch(&mut app, '.');
        assert_eq!(app.wb().pointer, Address::new(SheetId::A, 0, 1)); // BL
        // Range unchanged.
        assert_eq!(app.highlight_range(), before);
        // Next `.` from BL → TL.
        press_ch(&mut app, '.');
        assert_eq!(app.wb().pointer, Address::A1); // TL
        assert_eq!(app.highlight_range(), before);
        // Next `.` from TL → TR.
        press_ch(&mut app, '.');
        assert_eq!(app.wb().pointer, Address::new(SheetId::A, 1, 0)); // TR
        assert_eq!(app.highlight_range(), before);
        // Next `.` from TR → BR. Full loop.
        press_ch(&mut app, '.');
        assert_eq!(app.wb().pointer, Address::new(SheetId::A, 1, 1)); // BR
        assert_eq!(app.highlight_range(), before);
    }

    #[test]
    fn range_erase_clears_all_cells_in_range() {
        let mut app = App::new();
        for (row, v) in [(0, "10"), (1, "20"), (2, "30")] {
            for c in v.chars() {
                press_ch(&mut app, c);
            }
            press(&mut app, KeyCode::Down);
            // After each commit, pointer moves; at end we're at row+1
            let _ = row;
        }
        // Go back to A1 and erase A1..A3
        app.wb_mut().pointer = Address::A1;
        app.begin_point(PendingCommand::RangeErase);
        press(&mut app, KeyCode::Down);
        press(&mut app, KeyCode::Down);
        press(&mut app, KeyCode::Enter);
        assert_eq!(app.mode, Mode::Ready);
        for row in 0..3 {
            assert!(
                !app.wb().cells.contains_key(&Address::new(SheetId::A, 0, row)),
                "A{} should be empty",
                row + 1
            );
        }
    }

    #[test]
    fn shift_cells_rows_insert_pushes_down() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 0), make_label("r0"));
        cells.insert(Address::new(SheetId::A, 0, 1), make_label("r1"));
        cells.insert(Address::new(SheetId::A, 0, 2), make_label("r2"));
        // Insert 1 row at row 1 — r1 and r2 move down by 1.
        shift_cells_rows(&mut cells, SheetId::A, 1, 1);
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 0)), Some(&make_label("r0")));
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 2)), Some(&make_label("r1")));
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 3)), Some(&make_label("r2")));
        assert!(!cells.contains_key(&Address::new(SheetId::A, 0, 1)));
    }

    #[test]
    fn shift_cells_rows_delete_pulls_up() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 0), make_label("r0"));
        cells.insert(Address::new(SheetId::A, 0, 2), make_label("r2"));
        // Simulate delete of row 1: remove it (already absent here) then pull.
        shift_cells_rows(&mut cells, SheetId::A, 2, -1);
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 0)), Some(&make_label("r0")));
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 1)), Some(&make_label("r2")));
    }

    #[test]
    fn shift_cells_cols_insert_pushes_right() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 0), make_label("A1"));
        cells.insert(Address::new(SheetId::A, 1, 0), make_label("B1"));
        shift_cells_cols(&mut cells, SheetId::A, 1, 1);
        assert_eq!(cells.get(&Address::new(SheetId::A, 0, 0)), Some(&make_label("A1")));
        assert_eq!(cells.get(&Address::new(SheetId::A, 2, 0)), Some(&make_label("B1")));
        assert!(!cells.contains_key(&Address::new(SheetId::A, 1, 0)));
    }

    #[test]
    fn shift_cells_rows_leaves_other_sheets_alone() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 1), make_label("a"));
        cells.insert(Address::new(SheetId(1), 0, 1), make_label("b"));
        shift_cells_rows(&mut cells, SheetId::A, 0, 1);
        assert!(cells.contains_key(&Address::new(SheetId::A, 0, 2)));
        // Sheet B unchanged.
        assert_eq!(
            cells.get(&Address::new(SheetId(1), 0, 1)),
            Some(&make_label("b"))
        );
    }

    #[test]
    fn manual_recalc_defers_computation_until_f9() {
        let mut app = App::new();
        app.set_recalc_mode(RecalcMode::Manual);

        // A1 = 10
        for c in "10".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        // B1 = +A1*2
        app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
        for c in "A1*2".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        // In Manual mode, the just-committed formula has no cached value yet.
        let formula = app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap();
        if let CellContents::Formula { cached_value, .. } = formula {
            assert!(cached_value.is_none(), "manual mode should not auto-eval");
        } else {
            panic!("expected Formula");
        }
        assert!(app.recalc_pending());

        // F9 computes and clears the pending flag.
        app.handle_key(KeyEvent::new(KeyCode::F(9), KeyModifiers::NONE));
        assert!(!app.recalc_pending());
        let formula = app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap();
        match formula {
            CellContents::Formula { cached_value, .. } => {
                assert_eq!(*cached_value, Some(Value::Number(20.0)));
            }
            other => panic!("expected Formula, got {other:?}"),
        }
    }

    #[test]
    fn calc_indicator_visible_only_when_pending() {
        let mut app = App::new();
        let buf = app.render_to_buffer(80, 25);
        let status_line = App::line_text(&buf, 24);
        assert!(!status_line.contains("CALC"), "should be absent: {status_line:?}");

        app.set_recalc_mode(RecalcMode::Manual);
        // Type a formula to get CALC to light up.
        app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('1'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        let buf = app.render_to_buffer(80, 25);
        let status_line = App::line_text(&buf, 24);
        assert!(status_line.contains("CALC"), "should contain CALC: {status_line:?}");

        app.handle_key(KeyEvent::new(KeyCode::F(9), KeyModifiers::NONE));
        let buf = app.render_to_buffer(80, 25);
        let status_line = App::line_text(&buf, 24);
        assert!(!status_line.contains("CALC"), "should be cleared: {status_line:?}");
    }

    #[test]
    fn formula_commit_populates_cached_value() {
        let mut app = App::new();
        // A1 = 10, A2 = 20
        for c in "10".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        for c in "20".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        // A3 = @SUM(A1..A2)
        for c in "@SUM(A1..A2)".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let a3 = app.wb().cells.get(&Address::new(SheetId::A, 0, 2)).unwrap();
        match a3 {
            CellContents::Formula { expr, cached_value } => {
                assert_eq!(expr, "@SUM(A1..A2)");
                assert_eq!(*cached_value, Some(Value::Number(30.0)));
            }
            other => panic!("expected Formula, got {other:?}"),
        }
    }

    #[test]
    fn upstream_edit_recomputes_dependent_formula() {
        let mut app = App::new();
        for c in "10".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        // B1 = +A1*3
        app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
        for c in "A1*3".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(
            app.wb().cells
                .get(&Address::new(SheetId::A, 0, 1))
                .and_then(|c| match c {
                    CellContents::Formula { cached_value, .. } => cached_value.clone(),
                    _ => None,
                }),
            Some(Value::Number(30.0))
        );
        // Change A1 to 5; B1 should recompute to 15.
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('5'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(
            app.wb().cells
                .get(&Address::new(SheetId::A, 0, 1))
                .and_then(|c| match c {
                    CellContents::Formula { cached_value, .. } => cached_value.clone(),
                    _ => None,
                }),
            Some(Value::Number(15.0))
        );
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
        app.wb_mut().cells.insert(
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
        app.wb_mut().cells.insert(
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
        match app.wb().cells.get(&Address::A1).unwrap() {
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
        assert!(!app.wb().cells.contains_key(&Address::A1));
    }

    #[test]
    fn arrow_commits_then_moves() {
        let mut app = App::new();
        for c in "hi".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(app.wb().pointer, Address::new(SheetId::A, 0, 1));
        assert!(matches!(
            app.wb().cells.get(&Address::A1),
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
        match app.wb().cells.get(&Address::A1).unwrap() {
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
        match app.wb().cells.get(&Address::A1).unwrap() {
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
        let stored = app.wb().cells.get(&Address::A1).unwrap();
        match stored {
            CellContents::Label { prefix, text } => {
                assert_eq!(*prefix, LabelPrefix::Apostrophe);
                assert_eq!(text, "hello");
            }
            other => panic!("expected Label, got {other:?}"),
        }
    }

    fn temp_test_dir(tag: &str) -> PathBuf {
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir = std::env::temp_dir().join(format!(
            "l123_test_{}_{}_{}",
            tag,
            process::id(),
            nanos,
        ));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    fn drive_save_keys(app: &mut App, seed_label: &str, path: &Path) {
        for c in seed_label.chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        for c in ['/', 'F', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in path.to_str().unwrap().chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    }

    /// End-to-end: /FS <path><Enter> writes an xlsx file at <path> that
    /// IronCalc can open. The temp dir is unique per test invocation so
    /// parallel runs don't collide.
    #[test]
    fn file_save_writes_xlsx_at_typed_path() {
        let dir = temp_test_dir("file_save");
        let target = dir.join("saved.xlsx");

        let mut app = App::new();
        drive_save_keys(&mut app, "42", &target);

        assert_eq!(app.mode, Mode::Ready);
        assert!(
            target.exists(),
            "expected xlsx at {target:?} — prompt did not write the file"
        );

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }

    /// Backup path: when the file already exists, picking Backup
    /// renames the existing file to `.BAK` and then writes a fresh one.
    #[test]
    fn file_save_backup_renames_existing_to_bak() {
        let dir = temp_test_dir("file_save_backup");
        let target = dir.join("sheet.xlsx");

        // First save populates the file.
        let mut app = App::new();
        drive_save_keys(&mut app, "42", &target);
        assert!(target.exists());

        // Modify, then /FS again — prefilled; Enter opens confirm.
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        for c in "99".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        for c in ['/', 'F', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert!(app.save_confirm.is_some(), "confirm submenu did not open");

        // Press B — Backup.
        app.handle_key(KeyEvent::new(KeyCode::Char('B'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        let bak = target.with_extension("BAK");
        assert!(bak.exists(), "expected {bak:?} after Backup");
        assert!(target.exists(), "fresh {target:?} after Backup");

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_file(&bak);
        let _ = std::fs::remove_dir(&dir);
    }

    /// Helper: drive /FX<kind> <path><Enter> HOME <Enter>. Assumes the
    /// pointer is at the bottom-right of the intended range before the
    /// call (POINT auto-anchors there; HOME slides the free corner to
    /// A1 → highlight covers A1..pointer).
    fn drive_xtract_keys(app: &mut App, kind: char, path: &Path) {
        for c in ['/', 'F', 'X', kind] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in path.to_str().unwrap().chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    }

    /// /FXF — Formulas extract keeps the formula string intact so the
    /// extracted file, when reloaded, still has a live formula in A3.
    #[test]
    fn file_xtract_formulas_preserves_formula() {
        let dir = temp_test_dir("xtract_f");
        let target = dir.join("x.xlsx");
        if target.exists() {
            std::fs::remove_file(&target).unwrap();
        }

        let mut app = App::new();
        // DOWN during entry commits-and-moves; ENTER commits-in-place.
        for line in ["10", "20"] {
            for c in line.chars() {
                app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
            }
            app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        }
        for c in "+A1+A2".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        drive_xtract_keys(&mut app, 'F', &target);
        assert_eq!(app.mode, Mode::Ready);
        assert!(target.exists(), "extract did not write the file");

        // Re-open and inspect.
        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&target).unwrap();
        let a3 = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
        assert_eq!(a3.value, Value::Number(30.0));
        assert!(
            a3.formula.is_some(),
            "Formulas variant should preserve the formula"
        );

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }

    /// /FXV — Values extract replaces the formula with its cached
    /// value; reloaded A3 has a number but no formula.
    #[test]
    fn file_xtract_values_strips_formula() {
        let dir = temp_test_dir("xtract_v");
        let target = dir.join("x.xlsx");
        if target.exists() {
            std::fs::remove_file(&target).unwrap();
        }

        let mut app = App::new();
        // DOWN during entry commits-and-moves; ENTER commits-in-place.
        for line in ["10", "20"] {
            for c in line.chars() {
                app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
            }
            app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        }
        for c in "+A1+A2".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        drive_xtract_keys(&mut app, 'V', &target);
        assert_eq!(app.mode, Mode::Ready);
        assert!(target.exists());

        let mut e = IronCalcEngine::new().unwrap();
        e.load_xlsx(&target).unwrap();
        let a3 = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
        assert_eq!(a3.value, Value::Number(30.0));
        assert!(
            a3.formula.is_none(),
            "Values variant should flatten formula to value"
        );

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }

    /// Worksheet listing is an alphabetical list of xlsx files in the
    /// given directory. Pure function of the directory's contents so
    /// we can test it without touching process CWD.
    #[test]
    fn list_worksheet_files_in_returns_xlsx_sorted() {
        let dir = temp_test_dir("list_ws");
        for name in ["zeta.xlsx", "alpha.xlsx", "other.txt", "mid.XLSX"] {
            std::fs::write(dir.join(name), b"placeholder").unwrap();
        }
        let got = list_worksheet_files_in(&dir);
        let names: Vec<String> = got
            .iter()
            .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
            .collect();
        assert_eq!(names, vec!["alpha.xlsx", "mid.XLSX", "zeta.xlsx"]);
        for name in ["zeta.xlsx", "alpha.xlsx", "other.txt", "mid.XLSX"] {
            let _ = std::fs::remove_file(dir.join(name));
        }
        let _ = std::fs::remove_dir(&dir);
    }

    /// Vertical navigation + scroll: with 20 files and a PAGE_SIZE of
    /// 10, pressing Down 15 times should advance the view so the
    /// highlight is still visible.
    #[test]
    fn file_list_vertical_nav_and_scroll() {
        let dir = temp_test_dir("list_scroll");
        for i in 0..20 {
            std::fs::write(dir.join(format!("f{i:02}.xlsx")), b"data").unwrap();
        }
        let entries = list_worksheet_files_in(&dir);
        assert_eq!(entries.len(), 20);

        let mut app = App::new();
        app.file_list = Some(FileListState {
            kind: FileListKind::Worksheet,
            entries,
            highlight: 0,
            view_offset: 0,
        });
        app.mode = Mode::Files;

        for _ in 0..15 {
            app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        }
        let fl = app.file_list.as_ref().unwrap();
        assert_eq!(fl.highlight, 15);
        assert!(
            fl.view_offset > 0,
            "view_offset should have advanced (got {})",
            fl.view_offset
        );
        assert!(
            fl.highlight >= fl.view_offset
                && fl.highlight < fl.view_offset + FILE_LIST_PAGE_SIZE,
            "highlight {} out of window starting at {}",
            fl.highlight,
            fl.view_offset
        );

        app.handle_key(KeyEvent::new(KeyCode::End, KeyModifiers::NONE));
        assert_eq!(app.file_list.as_ref().unwrap().highlight, 19);

        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        let fl = app.file_list.as_ref().unwrap();
        assert_eq!(fl.highlight, 0);
        assert_eq!(fl.view_offset, 0);

        for i in 0..20 {
            let _ = std::fs::remove_file(dir.join(format!("f{i:02}.xlsx")));
        }
        let _ = std::fs::remove_dir(&dir);
    }

    /// Worksheet branch: pressing Enter on the highlighted row loads
    /// that xlsx file into the workbook.
    #[test]
    fn file_list_worksheet_enter_retrieves_highlighted() {
        let dir = temp_test_dir("list_ws_enter");

        // Build two xlsx files with distinguishing contents. Driving
        // through the App would set CWD / active_path; use the engine
        // directly so the test stays hermetic.
        let a = dir.join("a.xlsx");
        let b = dir.join("b.xlsx");
        {
            let mut e = IronCalcEngine::new().unwrap();
            e.set_user_input(Address::A1, "111").unwrap();
            e.recalc();
            e.save_xlsx(&a).unwrap();
        }
        {
            let mut e = IronCalcEngine::new().unwrap();
            e.set_user_input(Address::A1, "222").unwrap();
            e.recalc();
            e.save_xlsx(&b).unwrap();
        }

        let mut app = App::new();
        app.file_list = Some(FileListState {
            kind: FileListKind::Worksheet,
            entries: vec![a.clone(), b.clone()],
            highlight: 1, // point at b.xlsx
            view_offset: 0,
        });
        app.mode = Mode::Files;

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.file_list.is_none());
        assert_eq!(app.wb().active_path.as_deref(), Some(b.as_path()));
        match app.wb().cells.get(&Address::A1).unwrap() {
            CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 222.0),
            other => panic!("A1 expected Number(222) from b.xlsx, got {other:?}"),
        }

        let _ = std::fs::remove_file(&a);
        let _ = std::fs::remove_file(&b);
        let _ = std::fs::remove_dir(&dir);
    }

    /// /FLA shows the single active_path on line 2. Enter / Esc both
    /// dismiss to READY without mutating the workbook.
    #[test]
    fn file_list_active_shows_active_path_and_esc_dismisses() {
        let mut app = App::new();
        app.wb_mut().active_path = Some(PathBuf::from("workbook.xlsx"));
        for c in ['/', 'F', 'L', 'A'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        assert_eq!(app.mode, Mode::Files);
        let fl = app.file_list.as_ref().expect("file_list populated");
        assert_eq!(fl.entries.len(), 1);
        assert_eq!(fl.entries[0], PathBuf::from("workbook.xlsx"));
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.file_list.is_none());
    }

    /// /FN wipes the entire in-memory workbook back to a blank sheet —
    /// cells, formats, active path, pointer, and the engine itself.
    #[test]
    fn file_new_wipes_in_memory_state() {
        let mut app = App::new();
        // Seed A1 and move pointer off origin so we can detect the reset.
        for c in "42".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        app.wb_mut().active_path = Some(PathBuf::from("/tmp/pretend.xlsx"));
        assert!(!app.wb().cells.is_empty());

        // /FNA (After).
        for c in ['/', 'F', 'N', 'A'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.wb().cells.is_empty(), "cells should be cleared");
        assert_eq!(app.wb().pointer, Address::A1);
        assert!(app.wb().active_path.is_none(), "active_path should be cleared");
    }

    /// /FIN drops the CSV rows into the grid starting at the pointer.
    /// Numbers become constants; strings become labels.
    #[test]
    fn file_import_numbers_populates_cells_from_csv() {
        let dir = temp_test_dir("import_n");
        let src = dir.join("in.csv");
        std::fs::write(&src, "10,20,30\n\"foo\",\"bar\",\"baz\"\n").unwrap();

        let mut app = App::new();
        // /FIN <path><Enter>
        for c in ['/', 'F', 'I', 'N'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in src.to_str().unwrap().chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        assert_eq!(app.mode, Mode::Ready);
        match app.wb().cells.get(&Address::new(SheetId::A, 0, 0)).unwrap() {
            CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 10.0),
            other => panic!("A1 expected Number(10), got {other:?}"),
        }
        match app.wb().cells.get(&Address::new(SheetId::A, 2, 0)).unwrap() {
            CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 30.0),
            other => panic!("C1 expected Number(30), got {other:?}"),
        }
        match app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap() {
            CellContents::Label { text, .. } => assert_eq!(text, "foo"),
            other => panic!("A2 expected Label(foo), got {other:?}"),
        }

        let _ = std::fs::remove_file(&src);
        let _ = std::fs::remove_dir(&dir);
    }

    /// /FR: save a workbook, dirty memory, retrieve — cells should
    /// show the saved values, not the dirty ones.
    #[test]
    fn file_retrieve_replaces_memory_with_saved_contents() {
        let dir = temp_test_dir("file_retrieve");
        let target = dir.join("sheet.xlsx");

        let mut app = App::new();
        drive_save_keys(&mut app, "42", &target);
        assert!(target.exists());

        // Dirty A1 in a fresh app, then /FR from disk.
        let mut app2 = App::new();
        for c in "99".chars() {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        for c in ['/', 'F', 'R'] {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in target.to_str().unwrap().chars() {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        assert_eq!(app2.mode, Mode::Ready);
        let stored = app2.wb().cells.get(&Address::A1).unwrap_or_else(|| {
            panic!("A1 not populated after /FR — have: {:?}", app2.wb().cells)
        });
        match stored {
            CellContents::Constant(Value::Number(n)) => assert_eq!(n, &42.0),
            other => panic!("expected Constant(42), got {other:?}"),
        }

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }

    /// Cancel path: pressing C on the confirm submenu leaves the
    /// existing file untouched.
    #[test]
    fn file_save_cancel_leaves_existing_untouched() {
        let dir = temp_test_dir("file_save_cancel");
        let target = dir.join("sheet.xlsx");
        let mut app = App::new();
        drive_save_keys(&mut app, "42", &target);
        let before_len = std::fs::metadata(&target).unwrap().len();

        // Second /FS opens confirm; Cancel (C) aborts.
        for c in ['/', 'F', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert!(app.save_confirm.is_some());
        app.handle_key(KeyEvent::new(KeyCode::Char('C'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.save_confirm.is_none());
        let after_len = std::fs::metadata(&target).unwrap().len();
        assert_eq!(before_len, after_len, "Cancel unexpectedly wrote the file");

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }
}
