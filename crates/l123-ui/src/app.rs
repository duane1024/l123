//! The app loop: ratatui + crossterm, control panel + grid + status line.
//!
//! Scope as of M1 cycle 2:
//! - READY / LABEL / VALUE modes with first-character dispatch (LABEL only
//!   implemented this cycle; VALUE lands in cycle 3).
//! - `'` auto-prefixed labels. Enter commits; `/QY` quits.
//! - Three-line control panel, mode indicator, cell readout.

use std::cell::Cell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io;
use std::path::{Path, PathBuf};
use std::time::Duration;

use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEvent, KeyEventKind,
        KeyModifiers, MouseButton, MouseEvent, MouseEventKind,
    },
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use l123_core::{
    address::col_to_letters, label::is_value_starter, plan_row_spill, render_value_in_cell,
    Address, CellContents, ErrKind, Format, FormatKind, LabelPrefix, Mode, Range, SheetId,
    SpillSlot, TextStyle, Value,
};
use l123_engine::{CellView, Engine, IronCalcEngine, RecalcMode};
use l123_graph::{GraphDef, GraphType, Series};
use l123_menu::{self as menu, Action, MenuBody, MenuItem};
use l123_print::{
    encode::lp::LpOptions, PrintContentMode, PrintFormatMode, PrintSettings, WorkbookView,
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
use ratatui_image::{picker::Picker, picker::ProtocolType, Image, Resize};

// Grid geometry — kept as consts so both render and cell-address-probe agree.
const ROW_GUTTER: u16 = 5;
const PANEL_HEIGHT: u16 = 4; // 3 content lines + 1 bottom border

/// Cap for the per-sheet name shown in the status line after the
/// filename. Longer xlsx tab names get truncated to keep the right-
/// side indicator zone (`FILE GROUP UNDO CALC CIRC MEM NUM …`) from
/// getting shoved off-screen on an 80-column terminal.
const STATUS_SHEET_NAME_MAX: usize = 20;

/// Emit a single BEL character to stdout and flush. The terminal's
/// own preferences decide whether that rings, flashes, or is ignored —
/// matching the soft, user-configurable behavior the user asked for.
fn emit_bell() {
    use std::io::Write;
    let mut out = io::stdout();
    let _ = out.write_all(b"\x07");
    let _ = out.flush();
}

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

/// `/Worksheet Global Recalc` direction. Natural is dependency-order
/// (IronCalc's default); Columnwise/Rowwise force the traversal shape.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum RecalcOrder {
    #[default]
    Natural,
    Columnwise,
    Rowwise,
}

impl RecalcOrder {
    pub fn label(self) -> &'static str {
        match self {
            RecalcOrder::Natural => "Natural",
            RecalcOrder::Columnwise => "Columnwise",
            RecalcOrder::Rowwise => "Rowwise",
        }
    }
}

/// `/Worksheet Global Zero` — whether numeric zero cells render blank.
/// R3.4a also has a `Label` mode where a custom string replaces the
/// zero; we keep the binary shape for now.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ZeroDisplay {
    #[default]
    No,
    Yes,
}

impl ZeroDisplay {
    pub fn label(self) -> &'static str {
        match self {
            ZeroDisplay::No => "No",
            ZeroDisplay::Yes => "Yes",
        }
    }
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
    /// Per-cell text-style overrides (bold / italic / underline) set
    /// by the WYSIWYG `:Format Bold|Italic|Underline Set|Clear`
    /// commands.  Empty style = no entry.
    cell_text_styles: HashMap<Address, TextStyle>,
    col_widths: HashMap<(SheetId, u16), u8>,
    /// Workbook-wide default column width (1..240). Applied to any
    /// column without a `col_widths` entry. Set by `/Worksheet Global
    /// Col-Width`. Initialized to 9 — the 1-2-3 R3 factory default.
    default_col_width: u8,
    /// Columns marked hidden by `/Worksheet Column Hide`. Skipped by
    /// the grid renderer; pointer and formulas can still address them.
    /// Not persisted through xlsx today — IronCalc 0.7 doesn't model
    /// a per-column hidden flag.
    hidden_cols: HashSet<(SheetId, u16)>,
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
    /// The unnamed working graph — target of every `/Graph` menu
    /// command until `/Graph Name Create` snapshots it by name.
    current_graph: GraphDef,
    /// Named graphs defined via `/Graph Name Create`. `Use` restores
    /// one into `current_graph`; `Delete` drops it; `Reset` wipes all.
    #[allow(dead_code)] // wired by `/Graph Name Create` in a later slice
    graphs: BTreeMap<String, GraphDef>,
}

impl Workbook {
    fn new() -> Self {
        Self {
            engine: IronCalcEngine::new().expect("IronCalc engine init"),
            cells: HashMap::new(),
            cell_formats: HashMap::new(),
            cell_text_styles: HashMap::new(),
            col_widths: HashMap::new(),
            default_col_width: 9,
            hidden_cols: HashSet::new(),
            active_path: None,
            pointer: Address::A1,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            journal: Vec::new(),
            current_graph: GraphDef::default(),
            graphs: BTreeMap::new(),
        }
    }
}

impl WorkbookView for Workbook {
    fn cell(&self, addr: Address) -> Option<&CellContents> {
        self.cells.get(&addr)
    }

    fn col_width(&self, sheet: SheetId, col: u16) -> u8 {
        self.col_widths.get(&(sheet, col)).copied().unwrap_or(9)
    }

    fn format_for_cell(&self, addr: Address) -> Format {
        self.cell_formats
            .get(&addr)
            .copied()
            .unwrap_or(Format::GENERAL)
    }
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
    /// `/Worksheet Global Protection` — disables edits to protected
    /// cells when On. Stored for the status panel; mutating commands
    /// don't consult it yet.
    global_protection: bool,
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
    /// Rect the icon panel last occupied on screen, stashed so mouse
    /// clicks can hit-test against it without recomputing the layout.
    /// Cleared at the top of each frame; re-set by `render_icon_panel`.
    icon_panel_area: Cell<Option<Rect>>,
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
    /// Startup welcome screen. `Some` while the splash is up; any
    /// keypress consumes the state and drops to READY without
    /// dispatching. Always `None` for `App::new()` so existing
    /// transcripts aren't blocked on a dismiss keystroke.
    splash: Option<SplashInfo>,
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
}

/// User-visible identity shown on the startup splash. The renderer
/// prints `user` after "User name:" and `organization` after
/// "Organization:", matching the 1-2-3 R3.4a licensing block.
#[derive(Debug, Clone)]
pub struct SplashInfo {
    pub user: String,
    pub organization: String,
}

/// State kept for the duration of a [`Mode::Graph`] overlay.
#[derive(Clone)]
struct GraphOverlay {
    /// Numeric values snapshotted off the engine at enter time.
    values: l123_graph::GraphValues,
    /// Pre-rendered PNG (via plotters) decoded into a DynamicImage.
    /// Populated only when the app has a graphical picker — feeds
    /// ratatui-image's `Image` widget at render time.
    img: Option<image::DynamicImage>,
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
/// Where the Go step sends its output.
#[derive(Debug, Clone)]
enum PrintDestination {
    /// `/Print File`: write the ASCII stream to this path.
    File(PathBuf),
    /// `/Print Printer`: pipe the ASCII stream (plus setup string) to
    /// CUPS `lp` with these options.
    Printer(LpOptions),
}

#[derive(Debug, Clone)]
struct PrintSession {
    destination: PrintDestination,
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
    fn new_file(path: PathBuf) -> Self {
        Self::with_destination(PrintDestination::File(path))
    }

    fn new_printer() -> Self {
        Self::with_destination(PrintDestination::Printer(LpOptions::default()))
    }

    fn with_destination(destination: PrintDestination) -> Self {
        Self {
            destination,
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
        text_styles: Vec<(Address, TextStyle)>,
    },
    /// Undo of a row insert: delete the row that was inserted.
    RowInsert { sheet: SheetId, at: u32 },
    /// Reinstate a deleted column on one sheet.
    ColDelete {
        sheet: SheetId,
        at: u16,
        cells: Vec<(Address, CellContents)>,
        formats: Vec<(Address, Format)>,
        text_styles: Vec<(Address, TextStyle)>,
    },
    /// Undo of a column insert: delete the column that was inserted.
    ColInsert { sheet: SheetId, at: u16 },
    /// Restore a range's prior per-cell contents + formats. Captures
    /// the state that `/Range Erase` cleared.
    RangeRestore {
        cells: Vec<(Address, CellContents)>,
        formats: Vec<(Address, Format)>,
        text_styles: Vec<(Address, TextStyle)>,
    },
    /// Restore per-cell format overrides after `/Range Format`. Each
    /// entry's `Option<Format>` is the pre-command format (None ==
    /// no override).
    RangeFormat {
        entries: Vec<(Address, Option<Format>)>,
    },
    /// Restore per-cell text-style overrides after `:Format
    /// Bold|Italic|Underline Set|Clear`.  `None` = no override before.
    RangeTextStyle {
        entries: Vec<(Address, Option<TextStyle>)>,
    },
    /// Restore one column's width. `prev_width = None` means the
    /// column had no override (default width).
    ColWidth {
        sheet: SheetId,
        col: u16,
        prev_width: Option<u8>,
    },
    /// Restore one column's hidden flag. Used by `/Worksheet Column
    /// Hide` and `/Worksheet Column Display`.
    ColHidden {
        sheet: SheetId,
        col: u16,
        prev_hidden: bool,
    },
    /// Restore the workbook-wide default column width.
    GlobalColWidth { prev: u8 },
    /// Restore the workbook-wide default label prefix.
    DefaultLabelPrefix { prev: LabelPrefix },
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
    RangeFormat {
        kind: FormatKind,
    },
    /// Set the current column's width to the buffered number.
    WorksheetColumnSetWidth,
    /// After the user types a width, enter POINT to pick the range of
    /// columns to apply it to.
    WorksheetColumnRangeSetWidth,
    /// Set the workbook-wide default column width.
    WorksheetGlobalColWidth,
    /// `/Worksheet Global Recalc Iteration` — numeric prompt clamped
    /// to 1..=50 iterations.
    WorksheetGlobalRecalcIteration,
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
    FileXtractFilename {
        kind: XtractKind,
    },
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
    FileOpenFilename {
        before: bool,
    },
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
    RangeSearchString {
        scope: SearchScope,
        range: Range,
    },
    /// After the user types the replacement string, apply it to all
    /// matches in the active [`SearchSession`].
    RangeSearchReplacement,
    /// After the user types a filename, save the current graph to
    /// that path as SVG.
    GraphSaveFilename,
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
            PromptNext::RangeFormat { .. }
            | PromptNext::WorksheetColumnSetWidth
            | PromptNext::WorksheetColumnRangeSetWidth
            | PromptNext::WorksheetGlobalColWidth
            | PromptNext::WorksheetGlobalRecalcIteration => c.is_ascii_digit(),
            PromptNext::RangeNameCreate | PromptNext::RangeNameDelete => {
                c.is_ascii_alphanumeric() || c == '_'
            }
            PromptNext::FileSaveFilename
            | PromptNext::FileRetrieveFilename
            | PromptNext::FileXtractFilename { .. }
            | PromptNext::FileImportNumbersFilename
            | PromptNext::FileDirPath
            | PromptNext::FileOpenFilename { .. }
            | PromptNext::PrintFileFilename
            | PromptNext::GraphSaveFilename => is_path_char(c),
            // Header and footer are free-form text with the `|`
            // separator carving them into L|C|R.
            PromptNext::PrintFileHeader | PromptNext::PrintFileFooter => c != '\n' && c != '\t',
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
    c.is_ascii_alphanumeric() || matches!(c, '.' | '-' | '_' | '/' | '\\' | ' ' | '~')
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
    CopyTo {
        source: Range,
    },
    MoveFrom,
    MoveTo {
        source: Range,
    },
    RangeLabel {
        new_prefix: LabelPrefix,
    },
    RangeFormat {
        format: Format,
    },
    /// `:Format Bold|Italic|Underline Set|Clear`: `bits` names which
    /// attributes the command touches; `set=true` ORs them in, `false`
    /// clears them.  `:Format Reset` sends `{bold,italic,underline}`
    /// with `set=false`.
    RangeTextStyle {
        bits: TextStyle,
        set: bool,
    },
    /// `pending_name` on App carries the name; on commit, define it over
    /// the selected range.
    RangeNameCreate,
    /// `pending_xtract_path` on App carries the destination path; on
    /// commit, extract the selected range into a new workbook file.
    FileXtractRange {
        kind: XtractKind,
    },
    /// The user is choosing the print range for the active
    /// [`PrintSession`]. On commit the range is stashed and the
    /// `/PF` submenu reopens.
    PrintFileRange,
    /// POINT step of `/Range Search`: on commit, prompt for the
    /// search string.
    RangeSearchRange {
        scope: SearchScope,
    },
    /// POINT step of `/Graph X` and `/Graph A`..`F`: on commit, the
    /// selected range is written into the named slot of the workbook's
    /// current graph and the menu returns to READY.
    GraphSeries {
        series: Series,
    },
    /// POINT step of `/Worksheet Column Column-Range Set-Width`. The
    /// width was captured from the prompt; on commit, apply it to every
    /// column in the selected range.
    ColumnRangeSetWidth {
        width: u8,
    },
    /// POINT step of `/Worksheet Column Column-Range Reset-Width`. On
    /// commit, clear width overrides for every column in the selected
    /// range.
    ColumnRangeResetWidth,
    /// POINT step of `/Worksheet Column Hide`. On commit, mark every
    /// column in the selected range as hidden.
    ColumnHide,
    /// POINT step of `/Worksheet Column Display`. On commit, unhide
    /// every column in the selected range.
    ColumnDisplay,
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
            PendingCommand::RangeTextStyle { .. } => "Enter range for style:",
            PendingCommand::RangeNameCreate => "Enter range for the named range:",
            PendingCommand::FileXtractRange { .. } => "Enter range to extract:",
            PendingCommand::PrintFileRange => "Enter range to print:",
            PendingCommand::RangeSearchRange { .. } => "Enter search range:",
            PendingCommand::GraphSeries { .. } => "Enter graph range:",
            PendingCommand::ColumnRangeSetWidth { .. } => "Enter range of columns to set:",
            PendingCommand::ColumnRangeResetWidth => "Enter range of columns to reset:",
            PendingCommand::ColumnHide => "Enter range of columns to hide:",
            PendingCommand::ColumnDisplay => "Enter range of columns to display:",
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
        Self {
            path: Vec::new(),
            highlight: 0,
            message: None,
            override_root: None,
        }
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
            recalc_order: RecalcOrder::Natural,
            recalc_iterations: 1,
            recalc_pending: false,
            zero_display: ZeroDisplay::No,
            global_protection: false,
            group_mode: false,
            undo_enabled: true,
            menu: None,
            point: None,
            prompt: None,
            error_message: None,
            pending_name: None,
            save_confirm: None,
            pending_xtract_path: None,
            file_list: None,
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
            splash: None,
            beep_enabled: true,
            beep_count: 0,
            beep_pending: false,
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

    pub fn run() -> anyhow::Result<()> {
        Self::run_with_file(None)
    }

    /// CLI entry point. When `path` is set the app opens that workbook
    /// and skips the splash; when `None` it greets the user with the
    /// licensing block until the first keypress.
    pub fn run_with_file(path: Option<PathBuf>) -> anyhow::Result<()> {
        let mut stdout = io::stdout();
        enable_raw_mode()?;
        execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        let cfg = crate::Config::resolve();
        let mut app = match path {
            Some(p) => App::new_with_file(p),
            None => App::new_with_splash(cfg.user.value.clone(), cfg.organization.value.clone()),
        };
        app.set_beep_enabled(cfg.error_beep_enabled());
        app.probe_image_picker();
        let result = app.event_loop(&mut terminal);

        disable_raw_mode()?;
        execute!(
            terminal.backend_mut(),
            LeaveAlternateScreen,
            DisableMouseCapture
        )?;
        terminal.show_cursor()?;

        result
    }

    fn event_loop<B: ratatui::backend::Backend>(
        &mut self,
        terminal: &mut Terminal<B>,
    ) -> anyhow::Result<()>
    where
        B::Error: Send + Sync + 'static,
    {
        while self.running {
            terminal.draw(|f| self.render(f.area(), f.buffer_mut()))?;
            if self.take_pending_beep() {
                emit_bell();
            }
            if event::poll(Duration::from_millis(100))? {
                match event::read()? {
                    Event::Key(k) if k.kind == KeyEventKind::Press => self.handle_key(k),
                    Event::Mouse(m) => self.handle_mouse(m),
                    _ => {}
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
        let dr = (a.row - self.wb().viewport_row_offset) as u16;
        let y = PANEL_HEIGHT + 1 + dr; // +1 skips column header row
        if y >= buf.area.height {
            return None;
        }
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

    // ---------------- key handling ----------------

    pub fn handle_key(&mut self, k: KeyEvent) {
        // Startup splash consumes the first keystroke and drops to
        // READY without dispatching — matches the 1-2-3 R3.4a behavior
        // where any key clears the welcome screen.
        if self.splash.is_some() {
            self.splash = None;
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
            Mode::Graph => self.handle_key_graph(k),
            Mode::Stat => self.handle_key_stat(k),
            Mode::Error => self.handle_key_error(k),
            _ => {}
        }
    }

    fn handle_key_error(&mut self, k: KeyEvent) {
        if matches!(k.code, KeyCode::Esc | KeyCode::Enter) {
            self.error_message = None;
            self.mode = Mode::Ready;
        }
    }

    fn set_error(&mut self, msg: impl Into<String>) {
        let msg = msg.into();
        tracing::error!(error = %msg, "user-visible error");
        self.error_message = Some(msg);
        self.mode = Mode::Error;
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
            KeyCode::F(2) => self.begin_edit(),
            KeyCode::F(4) if k.modifiers.contains(KeyModifiers::ALT) => self.undo(),
            KeyCode::F(9) => self.do_recalc(),
            KeyCode::F(10) => self.enter_graph_view(),
            KeyCode::Char('/') => self.open_menu(),
            KeyCode::Char(':') => self.open_wysiwyg_menu(),
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
        self.entry = Some(Entry {
            kind: EntryKind::Edit,
            buffer: source,
        });
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
            self.entry = Some(Entry {
                kind: EntryKind::Value,
                buffer: c.to_string(),
            });
            self.mode = Mode::Value;
        } else if matches!(c, '\'' | '"' | '^' | '\\' | '|') {
            // Explicit label prefix typed first: the char becomes the
            // LabelPrefix; the buffer starts empty.
            let prefix = LabelPrefix::from_char(c).expect("matched above");
            self.entry = Some(Entry {
                kind: EntryKind::Label(prefix),
                buffer: String::new(),
            });
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
            EntryKind::Label(prefix) => CellContents::Label {
                prefix,
                text: entry.buffer,
            },
            EntryKind::Value => match entry.buffer.parse::<f64>() {
                Ok(n) => CellContents::Constant(Value::Number(n)),
                Err(_) => CellContents::Formula {
                    expr: entry.buffer,
                    cached_value: None,
                },
            },
            // EDIT commits re-parse the full source buffer so the user can
            // change prefix or type (label ↔ value) via the first-char rule.
            EntryKind::Edit => CellContents::from_source(&entry.buffer, self.default_label_prefix),
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
            JournalEntry::DefaultLabelPrefix { prev } => {
                self.default_label_prefix = prev;
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
        self.mode = Mode::Stat;
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
            Action::RangeFormatText => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::Text,
                    decimals: 0,
                },
            }),
            Action::RangeFormatDateDmy => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::DateDmy,
                    decimals: 0,
                },
            }),
            Action::RangeFormatDateDm => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::DateDm,
                    decimals: 0,
                },
            }),
            Action::RangeFormatDateMy => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::DateMy,
                    decimals: 0,
                },
            }),
            Action::RangeFormatDateLongIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::DateLongIntl,
                    decimals: 0,
                },
            }),
            Action::RangeFormatDateShortIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::DateShortIntl,
                    decimals: 0,
                },
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
            Action::PrintPrinter => self.start_print_printer(),
            Action::PrintFileRange => self.begin_point(PendingCommand::PrintFileRange),
            Action::PrintFileGo => self.execute_print_go(),
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
                }
            }
            _ => return,
        }
        if let Some(fl) = self.file_list.as_mut() {
            adjust_file_list_view(fl);
        }
    }

    /// /FN — wipe the current workbook back to a blank slate. Both the
    /// `/Worksheet Erase Yes` — drop every active file and replace the
    /// workspace with a single blank workbook. Session-level prompts,
    /// menus, and modal overlays are also cleared so the user lands in
    /// a predictable READY state on A:A1.
    fn execute_worksheet_erase(&mut self) {
        self.entry = None;
        self.menu = None;
        self.prompt = None;
        self.point = None;
        self.save_confirm = None;
        self.pending_name = None;
        self.pending_xtract_path = None;
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
        self.wb_mut().cell_formats.clear();
        self.wb_mut().cell_text_styles.clear();
        self.wb_mut().col_widths.clear();
        self.wb_mut().default_col_width = 9;
        self.wb_mut().hidden_cols.clear();
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

    /// `/Print Printer`: open a printer session straight away — no path
    /// prompt. Shares the submenu with `/Print File`; the Go branch
    /// sends output through CUPS `lp` instead of writing a file.
    fn start_print_printer(&mut self) {
        self.print = Some(PrintSession::new_printer());
        self.enter_print_file_menu();
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

    fn execute_print_go(&mut self) {
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
        let grid = l123_print::render(self.wb(), range, &settings);
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
                    l123_print::to_ascii(&grid).into_bytes()
                };
                let _ = std::fs::write(path, bytes);
            }
            PrintDestination::Printer(lp_opts) => {
                #[cfg(unix)]
                {
                    let _ = l123_print::encode::lp::to_lp(&grid, lp_opts);
                }
                #[cfg(not(unix))]
                {
                    let _ = lp_opts; // hold field live on non-unix
                }
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
        for (addr, cv) in engine.used_cells() {
            if let Some(contents) = cell_view_to_contents(&cv) {
                cells.insert(addr, contents);
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
        let new_file = Workbook {
            engine,
            cells,
            cell_formats,
            cell_text_styles,
            col_widths,
            default_col_width: 9,
            hidden_cols: HashSet::new(),
            active_path: Some(path),
            pointer: Address::A1,
            viewport_col_offset: 0,
            viewport_row_offset: 0,
            journal: Vec::new(),
            current_graph: GraphDef::default(),
            graphs: BTreeMap::new(),
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
        let body = match std::fs::read_to_string(&path) {
            Ok(b) => b,
            Err(e) => {
                self.set_error(format!("Cannot read {}: {e}", path.display()));
                return;
            }
        };
        let rows = l123_io::csv::parse(&body);
        let origin = self.wb_mut().pointer;
        for (dr, row) in rows.iter().enumerate() {
            for (dc, field) in row.iter().enumerate() {
                if field.is_empty() {
                    continue;
                }
                let addr =
                    Address::new(origin.sheet, origin.col + dc as u16, origin.row + dr as u32);
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
        self.mode = Mode::Ready;
    }

    /// Load an xlsx from disk, wiping the current in-memory workbook
    /// and repopulating the UI cache from the loaded engine model.
    fn load_workbook_from(&mut self, path: PathBuf) {
        if let Err(e) = self.wb_mut().engine.load_xlsx(&path) {
            self.set_error(format!("Cannot open {}: {e}", path.display()));
            return;
        }
        // Wipe UI state; the loaded engine is the new source of truth.
        self.wb_mut().cells.clear();
        self.wb_mut().cell_formats.clear();
        self.wb_mut().cell_text_styles.clear();
        self.wb_mut().col_widths.clear();
        self.wb_mut().default_col_width = 9;
        self.wb_mut().hidden_cols.clear();
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
        for (addr, w) in self.wb_mut().engine.used_column_widths() {
            self.wb_mut().col_widths.insert((addr.sheet, addr.col), w);
        }
        for (addr, style) in self.wb_mut().engine.used_cell_text_styles() {
            self.wb_mut().cell_text_styles.insert(addr, style);
        }
        for (addr, fmt) in self.wb_mut().engine.used_cell_formats() {
            self.wb_mut().cell_formats.insert(addr, fmt);
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
        // that /File Retrieve reads back.
        let formats: Vec<(Address, Format)> = self
            .wb()
            .cell_formats
            .iter()
            .map(|(a, f)| (*a, *f))
            .collect();
        for (addr, fmt) in formats {
            let _ = self.wb_mut().engine.set_cell_format(addr, fmt);
        }
        if self.wb_mut().engine.save_xlsx(&path).is_ok() {
            self.wb_mut().active_path = Some(path);
        }
    }

    /// Handle a keystroke while the Cancel/Replace/Backup confirm is up.
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
        self.point = Some(PointState {
            anchor: Some(self.wb().pointer),
            pending,
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
        let Some(ps) = self.point.as_mut() else {
            return;
        };
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
            Some(a) => Range {
                start: a,
                end: self.wb_mut().pointer,
            }
            .normalized(),
            None => Range::single(self.wb_mut().pointer),
        };
        match ps.pending {
            PendingCommand::RangeErase => {
                self.execute_range_erase(range);
                self.mode = Mode::Ready;
            }
            PendingCommand::CopyFrom => {
                self.transition_point(PendingCommand::CopyTo { source: range })
            }
            PendingCommand::MoveFrom => {
                self.transition_point(PendingCommand::MoveTo { source: range })
            }
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
            PendingCommand::RangeTextStyle { bits, set } => {
                self.execute_range_text_style(range, bits, set);
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
            PendingCommand::GraphSeries { series } => {
                self.wb_mut().current_graph.set(series, range);
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnRangeSetWidth { width } => {
                self.execute_col_range_width(range, Some(width));
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnRangeResetWidth => {
                self.execute_col_range_width(range, None);
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnHide => {
                self.execute_col_hide_display(range, true);
                self.mode = Mode::Ready;
            }
            PendingCommand::ColumnDisplay => {
                self.execute_col_hide_display(range, false);
                self.mode = Mode::Ready;
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

    /// Width of `col` in `sheet`. Returns the per-column override if
    /// one is set, otherwise the workbook's global default width.
    fn col_width_of(&self, sheet: SheetId, col: u16) -> u8 {
        self.wb()
            .col_widths
            .get(&(sheet, col))
            .copied()
            .unwrap_or(self.wb().default_col_width)
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
                self.mode = Mode::Ready;
            }
            PromptNext::WorksheetGlobalRecalcIteration => {
                let current = self.recalc_iterations;
                let n: u16 = p.buffer.parse().unwrap_or(current).clamp(1, 50);
                self.recalc_iterations = n;
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
            PromptNext::GraphSaveFilename => {
                let buf = p.buffer.clone();
                self.commit_graph_save(&buf);
            }
            PromptNext::FileRetrieveFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                self.retrieve_by_extension(PathBuf::from(&p.buffer));
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
                self.print = Some(PrintSession::new_file(path));
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

        let pointer_row = self.wb().pointer.row;
        if pointer_row >= self.wb().viewport_row_offset + visible_rows {
            self.wb_mut().viewport_row_offset = pointer_row - visible_rows + 1;
        }

        let sheet = self.wb().pointer.sheet;
        let pointer_col = self.wb().pointer.col;
        if !self.wb().hidden_cols.contains(&(sheet, pointer_col)) {
            let actual_w = self.col_width_of(sheet, pointer_col) as u16;
            let layout = self.visible_column_layout(content_width);
            let fully_visible = layout
                .iter()
                .any(|&(c, _, drawn)| c == pointer_col && drawn == actual_w);
            if !fully_visible {
                self.wb_mut().viewport_col_offset =
                    self.ideal_left_for_rightmost(pointer_col, content_width);
            }
        }
    }

    /// Walk leftward from `target_col`, summing visible-column widths,
    /// and return the leftmost column that still leaves room for the
    /// target on the right edge. If `target_col`'s own width exceeds
    /// `content_width`, returns `target_col` so at least its left side
    /// is shown.
    fn ideal_left_for_rightmost(&self, target_col: u16, content_width: u16) -> u16 {
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

    fn render(&self, area: Rect, buf: &mut Buffer) {
        // Startup splash is full-screen — no control panel, no grid,
        // no status line. Draws until the first keystroke dismisses it.
        if let Some(info) = self.splash.as_ref() {
            self.render_splash(area, buf, info);
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
        if self.file_list.is_some() {
            self.render_file_list_overlay(chunks[1], buf);
        } else if self.mode == Mode::Graph {
            self.render_graph_overlay(chunks[1], buf);
        } else if self.mode == Mode::Stat {
            self.render_stat_overlay(chunks[1], buf);
        } else {
            self.render_grid(main_area, buf);
            if let Some(area) = icon_area {
                self.render_icon_panel(area, buf);
            }
        }
        self.render_status(chunks[2], buf);
    }

    fn render_splash(&self, area: Rect, buf: &mut Buffer, info: &SplashInfo) {
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

    fn split_for_icon_panel(&self, area: Rect) -> (Rect, Option<Rect>) {
        if self.icon_panel.is_none()
            || self.mode == Mode::Graph
            || self.file_list.is_some()
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

    fn render_icon_panel(&self, area: Rect, buf: &mut Buffer) {
        let (Some(picker), Some(img)) = (self.image_picker.as_ref(), self.icon_panel.as_ref())
        else {
            return;
        };
        if picker.protocol_type() == ProtocolType::Halfblocks {
            return;
        }
        if let Ok(protocol) = picker.new_protocol(img.clone(), area, Resize::Fit(None)) {
            // The PNG has a 1:17 aspect; `Resize::Fit` preserves it,
            // so the rendered icons usually occupy fewer rows than
            // the allocated strip. Stash the actual rendered rect so
            // hover/click hit-tests target where the icons really are.
            let rendered = protocol.area();
            Image::new(&protocol).render(area, buf);
            self.icon_panel_area.set(Some(Rect::new(
                area.x,
                area.y,
                rendered.width.min(area.width),
                rendered.height.min(area.height),
            )));
        }
    }

    /// Mouse handler. Icon-panel clicks dispatch the slot's action;
    /// grid clicks in READY move the pointer to the clicked cell.
    /// Move events update `hovered_icon` so control-panel line 3 can
    /// show the authentic R3.4a icon description.
    pub fn handle_mouse(&mut self, m: MouseEvent) {
        let icon_area = self.icon_panel_area.get();
        let inside_icon = icon_area.is_some_and(|a| {
            m.column >= a.x && m.column < a.x + a.width && m.row >= a.y && m.row < a.y + a.height
        });

        match m.kind {
            MouseEventKind::Moved => {
                // Slot 16 (pager) is excluded from hover tooltip —
                // its function is rendered on the slot itself.
                self.hovered_icon = match (inside_icon, icon_area) {
                    (true, Some(a)) => match Self::hit_test_slot(a, m.row) {
                        Some(slot) if slot < 16 => Some((self.current_panel, slot)),
                        _ => None,
                    },
                    _ => None,
                };
            }
            MouseEventKind::Down(MouseButton::Left) => {
                // Icon panel wins over the grid when both would hit —
                // it sits on top visually.
                if let (true, Some(a)) = (inside_icon, icon_area) {
                    self.dispatch_icon_click(a, m.column, m.row);
                } else if let Some(addr) = self.cell_at_screen(m.column, m.row) {
                    match self.mode {
                        // In POINT, the anchor is untouched by
                        // move_pointer_to, so an anchored range
                        // extends (or shrinks) to the clicked cell
                        // naturally; unanchored POINT just moves the
                        // pointer.
                        Mode::Ready | Mode::Point => self.move_pointer_to(addr),
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
            _ => {}
        }
    }

    /// Jump the cell pointer to `addr`, scrolling the viewport only
    /// if needed. Used by mouse click-to-move; keyboard `/RG` (F5)
    /// has its own prompt-driven path.
    fn move_pointer_to(&mut self, addr: Address) {
        self.wb_mut().pointer = addr;
        self.scroll_into_view();
    }

    /// Append `addr` (short form, e.g. `C3`) to the current entry
    /// buffer — but only when the buffer is in a position where a
    /// cell reference is grammatically valid (empty, or ending in an
    /// operator/paren/comma/dot/comparison/logical token). Called
    /// from mouse click-in-grid during VALUE/EDIT, so the user can
    /// build formulas by clicking instead of typing references.
    fn splice_cell_ref_into_entry(&mut self, addr: Address) {
        let Some(entry) = self.entry.as_mut() else {
            return;
        };
        if !matches!(entry.kind, EntryKind::Value | EntryKind::Edit) {
            return;
        }
        let accepts_ref = match entry.buffer.chars().last() {
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
        entry.buffer.push_str(&addr.display_short());
    }

    /// Map a screen coordinate to the cell address it sits on, or
    /// `None` for clicks on the column header, row-number gutter, or
    /// outside the grid. Requires that a grid has already rendered
    /// this session so `last_grid_area` is populated.
    fn cell_at_screen(&self, col: u16, row: u16) -> Option<Address> {
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
    /// Slot 16 is the panel navigator. Returns `None` when the area
    /// is zero-height or the row is above its origin.
    ///
    /// `Resize::Fit(None)` rarely produces a pixel-height that's an
    /// exact multiple of 17 rows, so computing `height / 17` first
    /// truncates and drifts the slot boundaries — worst near the
    /// bottom, where the drift rolls past slot 16. Multiply before
    /// dividing to keep the mapping proportional across the band.
    fn hit_test_slot(area: Rect, row: u16) -> Option<usize> {
        if area.height == 0 || row < area.y {
            return None;
        }
        let offset = (row - area.y) as usize;
        Some((offset * 17 / area.height as usize).min(16))
    }

    /// Map a mouse click on the icon panel to a slot and fire that
    /// slot's action. Slot 16 is the panel navigator; clicks on its
    /// left half go to the previous panel, right half to the next.
    fn dispatch_icon_click(&mut self, area: Rect, column: u16, row: u16) {
        let Some(slot) = Self::hit_test_slot(area, row) else {
            return;
        };

        if slot == 16 {
            // Pager: left half → previous panel, right half → next.
            let half = area.x + area.width / 2;
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
            l123_graph::IconAction::SysKey(act) => self.dispatch_sys_action(act),
            l123_graph::IconAction::PageNav => {} // reached only via slot 16
            l123_graph::IconAction::Noop => {}
        }
    }

    /// Open the slash menu and descend via the given accelerator
    /// letters — equivalent to the user typing "/" then each char.
    fn dispatch_menu_path(&mut self, path: &str) {
        self.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        for c in path.chars() {
            self.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    /// Open the WYSIWYG colon menu and descend via the given accelerator
    /// letters — equivalent to the user typing `:` then each char.
    fn dispatch_wysiwyg_menu_path(&mut self, path: &str) {
        self.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
        for c in path.chars() {
            self.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    fn dispatch_sys_action(&mut self, action: l123_graph::SysAction) {
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
        };
        self.handle_key(KeyEvent::new(code, mods));
    }

    fn render_graph_overlay(&self, area: Rect, buf: &mut Buffer) {
        let Some(overlay) = self.graph_view.as_ref() else {
            self.render_grid(area, buf);
            return;
        };
        // Graphical path: ratatui-image with a freshly-built Protocol
        // sized to this frame's content area. Protocol creation can
        // fail (encoding error, terminal query hiccup); on any failure
        // we fall through to the unicode path so the user still sees
        // something.
        if let (Some(picker), Some(img)) = (self.image_picker.as_ref(), overlay.img.as_ref()) {
            if picker.protocol_type() != ProtocolType::Halfblocks {
                if let Ok(protocol) = picker.new_protocol(img.clone(), area, Resize::Fit(None)) {
                    Image::new(&protocol).render(area, buf);
                    return;
                }
            }
        }
        l123_graph::render_unicode(&self.wb().current_graph, &overlay.values, area, buf);
    }

    fn render_stat_overlay(&self, area: Rect, buf: &mut Buffer) {
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
        // display). Lower band: the environment readout.
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
        Paragraph::new(vec![
            // Global default format isn't wired to `/Worksheet Global
            // Format` yet, so report the generic default — matches
            // what a cell with no explicit format gets.
            Line::from("Format:       (G)"),
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

    fn render_control_panel(&self, area: Rect, buf: &mut Buffer) {
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

        // Line 2 & 3 depend on mode. File-list and save-confirm take
        // absolute precedence (they own the keyboard); then a
        // command-argument prompt; then mode-specific rendering.
        let (line2, line3) = if self.file_list.is_some() {
            self.render_file_list_lines()
        } else if self.save_confirm.is_some() {
            self.render_save_confirm_lines()
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
                _ => {
                    let l2 = match self.entry.as_ref() {
                        Some(e) => Line::from(format!(" {}", e.buffer)),
                        None => Line::from(""),
                    };
                    // Icon-hover description — gated on entry being
                    // idle so an in-progress value/label/edit never
                    // has its canonical line-3 space (formula preview)
                    // clobbered by a stray mouse-over.
                    let l3 = match (self.entry.as_ref(), self.hovered_icon) {
                        (None, Some((panel, slot))) => {
                            Line::from(format!(" {}", l123_graph::slot_description(panel, slot)))
                        }
                        _ => Line::from(""),
                    };
                    (l2, l3)
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
        self.wb()
            .cells
            .get(&self.wb().pointer)
            .map(|c| c.control_panel_readout())
            .unwrap_or_default()
    }

    /// Parenthesized format tag (e.g. `(G)`, `(C2)`, `(F3)`) for the current
    /// cell. Empty for labels, empty cells, or cells using `Reset` format.
    fn format_tag_for_line1(&self) -> String {
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
    fn text_style_marker_for_line1(&self) -> String {
        match self.wb().cell_text_styles.get(&self.wb().pointer) {
            Some(style) => style.to_string(),
            None => String::new(),
        }
    }

    /// Resolve the format for a given cell — the per-cell override if set,
    /// else General.
    fn format_for_cell(&self, addr: Address) -> Format {
        self.wb()
            .cell_formats
            .get(&addr)
            .copied()
            .unwrap_or(Format::GENERAL)
    }

    /// `[Wn]` tag when the current column's width differs from the
    /// workbook's global default.
    fn width_tag_for_line1(&self) -> String {
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
    fn visible_column_layout(&self, content_width: u16) -> Vec<(u16, u16, u16)> {
        let mut out = Vec::new();
        if content_width == 0 {
            return out;
        }
        let sheet = self.wb().pointer.sheet;
        let mut x_off: u16 = 0;
        let mut col = self.wb().viewport_col_offset;
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

    fn render_grid(&self, area: Rect, buf: &mut Buffer) {
        self.last_grid_area.set(Some(area));
        if area.width <= ROW_GUTTER || area.height < 2 {
            return;
        }

        let content_width = area.width - ROW_GUTTER;
        let layout = self.visible_column_layout(content_width);
        let visible_rows = area.height - 1;

        // Column header row
        let header_style = Style::default().add_modifier(Modifier::REVERSED);
        for &(col_idx, x_off, w) in &layout {
            let letters = col_to_letters(col_idx);
            let x = area.x + ROW_GUTTER + x_off;
            write_centered(buf, x, area.y, w, &letters, header_style);
        }
        // Top-left gutter corner
        for k in 0..ROW_GUTTER {
            buf[(area.x + k, area.y)]
                .set_char(' ')
                .set_style(header_style);
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
                    match self.wb().cells.get(&addr) {
                        None | Some(CellContents::Empty) => RowInput::Empty,
                        Some(CellContents::Label { prefix, text }) => RowInput::Label {
                            prefix: *prefix,
                            text: text.clone(),
                        },
                        Some(other) => {
                            let fmt = self.format_for_cell(addr);
                            RowInput::Rendered(render_own_width(other, w as usize, fmt))
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
                    Style::default().add_modifier(Modifier::REVERSED)
                } else {
                    Style::default()
                };
                if let Some(style) = self.wb().cell_text_styles.get(&style_addr).copied() {
                    cell_style = cell_style.add_modifier(text_style_modifier(style));
                }
                let mut printed = 0u16;
                for ch in slot.text.chars().take(w as usize) {
                    buf[(x + printed, y)].set_char(ch).set_style(cell_style);
                    printed += 1;
                }
                // Pad any shortfall with blank cells so highlight still
                // fills the whole column.
                while printed < w {
                    buf[(x + printed, y)].set_char(' ').set_style(cell_style);
                    printed += 1;
                }
            }
        }
    }

    fn render_status(&self, area: Rect, buf: &mut Buffer) {
        // Left slot: the active workbook's base filename, or the
        // current local date/time in 1-2-3 `DD-Mon-YYYY HH:MM` style.
        let left_text = match self.wb().active_path.as_ref() {
            Some(p) => p
                .file_name()
                .and_then(|n| n.to_str())
                .map(|s| s.to_string())
                .unwrap_or_else(|| p.display().to_string()),
            None => crate::clock::format_ddmmmyyyy_hhmm(crate::clock::local_now()),
        };
        // Multi-sheet workbook → append "[<Letter>: <name>]" so users
        // can see which Excel tab their letter-addressed pointer is on.
        // Single-sheet workbooks skip this since there's nothing to
        // disambiguate and authentic 1-2-3 never showed a sheet name.
        let sheet_suffix = if self.wb().engine.sheet_count() > 1 {
            let sid = self.wb().pointer.sheet;
            let letter = sid.letter();
            let name = self
                .wb()
                .engine
                .sheet_name(sid)
                .unwrap_or_else(|| "?".to_string());
            let trimmed: String = name.chars().take(STATUS_SHEET_NAME_MAX).collect();
            format!("  [{letter}: {trimmed}]")
        } else {
            String::new()
        };
        let left = format!(" {left_text}{sheet_suffix}");
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
        let right_chunk = indicators.join(" ");
        let pad = (area.width as usize).saturating_sub(left.len() + right_chunk.len() + 1);
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
/// Map a per-cell [`TextStyle`] into the ratatui [`Modifier`] bits that
/// render it: bold → `BOLD`, italic → `ITALIC`, underline → `UNDERLINED`.
/// Empty style yields `Modifier::empty()` and adds no visible attributes.
fn text_style_modifier(style: TextStyle) -> Modifier {
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

/// Per-column classification fed to the spill planner. `Label` keeps
/// the owned text so the borrowed [`SpillSlot::Label`] that follows can
/// point into it for the planner's lifetime.
enum RowInput {
    Empty,
    Label { prefix: LabelPrefix, text: String },
    Rendered(String),
}

/// Render non-label cell contents into exactly `width` chars. An
/// `Empty` / `Value::Empty` / unevaluated formula produces blanks so
/// the result can slot directly into [`SpillSlot::Rendered`].
fn render_own_width(contents: &CellContents, width: usize, format: Format) -> String {
    match contents {
        CellContents::Empty => " ".repeat(width),
        CellContents::Label { .. } => {
            // Labels are routed through SpillSlot::Label; we should
            // never hit this branch from the planner caller.
            " ".repeat(width)
        }
        CellContents::Constant(v) => {
            render_value_in_cell(v, width, format).unwrap_or_else(|| " ".repeat(width))
        }
        CellContents::Formula {
            cached_value: Some(v),
            ..
        } => render_value_in_cell(v, width, format).unwrap_or_else(|| " ".repeat(width)),
        CellContents::Formula {
            cached_value: None, ..
        } => " ".repeat(width),
    }
}

/// Mirrors ratatui-image's own `iterm2_from_env` list of hosts that
/// speak the OSC 1337 inline-image protocol. We re-check the same
/// environment in [`App::probe_image_picker`] as a workaround for a
/// quirk in `Picker::from_query_stdio`: when the font-size probe
/// fails (common in iTerm2), the library drops back to a default
/// Halfblocks picker and discards its own iTerm2 env hint.
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    #[test]
    fn starts_at_a1() {
        let app = App::new();
        assert_eq!(app.wb().pointer, Address::A1);
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.entry.is_none());
    }

    #[test]
    fn new_with_file_routes_csv_by_extension() {
        use std::io::Write;
        let dir = std::env::temp_dir().join("l123_new_with_csv");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("foo.csv");
        std::fs::File::create(&path)
            .unwrap()
            .write_all(b"10,20\nfoo,bar\n")
            .unwrap();
        let app = App::new_with_file(path);
        assert_eq!(
            app.mode,
            Mode::Ready,
            "CLI-opening a .csv should land in READY, not ERROR (error={:?})",
            app.error_message,
        );
        match app.wb().cells.get(&Address::A1) {
            Some(CellContents::Constant(Value::Number(n))) => assert_eq!(*n, 10.0),
            other => panic!("A1 expected Number(10), got {other:?}"),
        }
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
    fn down_past_visible_area_advances_row_offset() {
        let mut app = App::new();
        // Render an 80x25 frame so scroll_into_view has a cached
        // grid rect to consult on the next move.
        let _ = app.render_to_buffer(80, 25);
        for _ in 0..25 {
            app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        }
        assert_eq!(app.wb().pointer.row, 25);
        assert!(
            app.wb().viewport_row_offset > 0,
            "viewport_row_offset should advance to keep pointer visible, got 0",
        );
        assert!(
            app.wb().pointer.row >= app.wb().viewport_row_offset,
            "pointer must be at or below the new top of viewport",
        );
    }

    #[test]
    fn right_past_visible_area_advances_col_offset() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        for _ in 0..12 {
            app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        }
        assert_eq!(app.wb().pointer.col, 12);
        assert!(
            app.wb().viewport_col_offset > 0,
            "viewport_col_offset should advance to keep pointer visible, got 0",
        );
    }

    #[test]
    fn up_after_scroll_pulls_viewport_back() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        for _ in 0..30 {
            app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        }
        assert!(app.wb().viewport_row_offset > 0);
        // Press UP enough to drop above the current viewport top.
        for _ in 0..30 {
            app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        }
        assert_eq!(app.wb().pointer.row, 0);
        assert_eq!(app.wb().viewport_row_offset, 0);
    }

    #[test]
    fn home_resets_pointer() {
        let mut app = App::new();
        app.wb_mut().pointer = Address::new(SheetId::A, 10, 10);
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer, Address::A1);
    }

    #[test]
    fn ctrl_c_does_not_quit() {
        // SPEC §7 Δ: 1-2-3 uses /QY to quit; Ctrl-C is unused.
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL));
        assert!(app.running);
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
        CellContents::Label {
            prefix: LabelPrefix::Apostrophe,
            text: text.into(),
        }
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
        assert_eq!(
            app.highlight_range(),
            Range::single(Address::new(SheetId::A, 1, 1))
        );
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
                !app.wb()
                    .cells
                    .contains_key(&Address::new(SheetId::A, 0, row)),
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
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 0)),
            Some(&make_label("r0"))
        );
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 2)),
            Some(&make_label("r1"))
        );
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 3)),
            Some(&make_label("r2"))
        );
        assert!(!cells.contains_key(&Address::new(SheetId::A, 0, 1)));
    }

    #[test]
    fn shift_cells_rows_delete_pulls_up() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 0), make_label("r0"));
        cells.insert(Address::new(SheetId::A, 0, 2), make_label("r2"));
        // Simulate delete of row 1: remove it (already absent here) then pull.
        shift_cells_rows(&mut cells, SheetId::A, 2, -1);
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 0)),
            Some(&make_label("r0"))
        );
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 1)),
            Some(&make_label("r2"))
        );
    }

    #[test]
    fn shift_cells_cols_insert_pushes_right() {
        let mut cells = HashMap::new();
        cells.insert(Address::new(SheetId::A, 0, 0), make_label("A1"));
        cells.insert(Address::new(SheetId::A, 1, 0), make_label("B1"));
        shift_cells_cols(&mut cells, SheetId::A, 1, 1);
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 0, 0)),
            Some(&make_label("A1"))
        );
        assert_eq!(
            cells.get(&Address::new(SheetId::A, 2, 0)),
            Some(&make_label("B1"))
        );
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
        assert!(
            !status_line.contains("CALC"),
            "should be absent: {status_line:?}"
        );

        app.set_recalc_mode(RecalcMode::Manual);
        // Type a formula to get CALC to light up.
        app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('1'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        let buf = app.render_to_buffer(80, 25);
        let status_line = App::line_text(&buf, 24);
        assert!(
            status_line.contains("CALC"),
            "should contain CALC: {status_line:?}"
        );

        app.handle_key(KeyEvent::new(KeyCode::F(9), KeyModifiers::NONE));
        let buf = app.render_to_buffer(80, 25);
        let status_line = App::line_text(&buf, 24);
        assert!(
            !status_line.contains("CALC"),
            "should be cleared: {status_line:?}"
        );
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
            app.wb()
                .cells
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
            app.wb()
                .cells
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
        let dir =
            std::env::temp_dir().join(format!("l123_test_{}_{}_{}", tag, process::id(), nanos,));
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
            fl.highlight >= fl.view_offset && fl.highlight < fl.view_offset + FILE_LIST_PAGE_SIZE,
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
        assert!(
            app.wb().active_path.is_none(),
            "active_path should be cleared"
        );
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
        let stored =
            app2.wb().cells.get(&Address::A1).unwrap_or_else(|| {
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

    /// Simulate a left-click at `(col, row)` by stashing a fake panel
    /// area and routing through [`App::handle_mouse`].
    fn click(app: &mut App, area: Rect, col: u16, row: u16) {
        app.icon_panel_area.set(Some(area));
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    /// Standard fixture panel: 17 icons × 3 rows each = 51 rows, so
    /// slot N starts at local row N * 3.
    const TEST_PANEL: Rect = Rect {
        x: 80,
        y: 4,
        width: 3,
        height: 51,
    };

    fn click_slot(app: &mut App, slot: u16) {
        // Click the middle row of the given slot.
        click(
            app,
            TEST_PANEL,
            TEST_PANEL.x + 1,
            TEST_PANEL.y + slot * 3 + 1,
        );
    }

    #[test]
    fn icon_click_save_opens_save_prompt() {
        let mut app = App::new();
        click_slot(&mut app, 0);
        // /FS enters Mode::Menu with an active prompt for the filename.
        assert_eq!(app.mode, Mode::Menu);
        assert!(app.prompt.is_some(), "save prompt should be active");
    }

    #[test]
    fn icon_click_retrieve_opens_retrieve_prompt() {
        let mut app = App::new();
        click_slot(&mut app, 1);
        assert_eq!(app.mode, Mode::Menu);
        assert!(app.prompt.is_some());
    }

    #[test]
    fn icon_click_graph_view_enters_graph_mode() {
        let mut app = App::new();
        click_slot(&mut app, 7);
        assert_eq!(app.mode, Mode::Graph);
    }

    #[test]
    fn icon_click_print_opens_print_prompt() {
        let mut app = App::new();
        click_slot(&mut app, 9);
        assert_eq!(app.mode, Mode::Menu);
        assert!(app.prompt.is_some());
    }

    #[test]
    fn icon_click_prev_sheet_is_noop_on_first_sheet() {
        let mut app = App::new();
        click_slot(&mut app, 5);
        // Only one sheet in a fresh workbook — prev-sheet clamps.
        assert_eq!(app.pointer().display_full(), "A:A1");
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn icon_click_help_is_safe_noop() {
        let mut app = App::new();
        click_slot(&mut app, 15);
        assert_eq!(app.pointer().display_full(), "A:A1");
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn icon_click_bold_enters_point_for_wysiwyg_bold_set() {
        // Panel 1 slot 11 = icon id 12 = :Format Bold Set.  The click
        // should drop the user into POINT mode auto-anchored at the
        // pointer; pressing Enter commits and applies bold.
        let mut app = App::new();
        click_slot(&mut app, 11);
        assert_eq!(app.mode, Mode::Point);
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
        );
    }

    #[test]
    fn icon_click_italic_applies_italic_on_commit() {
        // Panel 1 slot 12 = icon id 13 = :Format Italic Set.
        let mut app = App::new();
        click_slot(&mut app, 12);
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn icon_click_underline_applies_underline_on_commit() {
        // Panel 1 slot 13 = icon id 15 = :Format Underline Set.
        let mut app = App::new();
        click_slot(&mut app, 13);
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::UNDERLINE),
        );
    }

    #[test]
    fn mouse_click_outside_panel_is_ignored() {
        let mut app = App::new();
        click(&mut app, TEST_PANEL, 10, 10);
        assert_eq!(app.pointer().display_full(), "A:A1");
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn mouse_click_without_cached_panel_is_ignored() {
        let mut app = App::new();
        app.icon_panel_area.set(None);
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: 82,
            row: 5,
            modifiers: KeyModifiers::NONE,
        });
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_non_left_button_is_ignored() {
        let mut app = App::new();
        app.icon_panel_area.set(Some(TEST_PANEL));
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Right),
            column: 81,
            row: 5,
            modifiers: KeyModifiers::NONE,
        });
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.prompt.is_none());
    }

    #[test]
    fn pager_right_half_advances_panel() {
        let mut app = App::new();
        assert_eq!(app.current_panel, l123_graph::Panel::One);
        let mid_col = TEST_PANEL.x + TEST_PANEL.width / 2;
        click(&mut app, TEST_PANEL, mid_col + 1, TEST_PANEL.y + 49);
        assert_eq!(app.current_panel, l123_graph::Panel::Two);
    }

    #[test]
    fn pager_left_half_retreats_panel() {
        let mut app = App::new();
        let left_col = TEST_PANEL.x;
        click(&mut app, TEST_PANEL, left_col, TEST_PANEL.y + 49);
        // From panel 1, prev wraps to panel 7.
        assert_eq!(app.current_panel, l123_graph::Panel::Seven);
    }

    #[test]
    fn pager_full_cycle_returns_to_panel_one() {
        let mut app = App::new();
        let mid_col = TEST_PANEL.x + TEST_PANEL.width / 2;
        for _ in 0..7 {
            click(&mut app, TEST_PANEL, mid_col + 1, TEST_PANEL.y + 49);
        }
        assert_eq!(app.current_panel, l123_graph::Panel::One);
    }

    #[test]
    fn hit_test_slot_is_proportional_for_non_aligned_heights() {
        // Real rendering uses `Resize::Fit(None)`, which scales the
        // icon PNG (1:17 aspect) to preserve shape — so the actual
        // pixel height rarely lands on a multiple of 17 rows. A 26-row
        // panel is a realistic example: slots occupy fractional bands.
        //
        // Dividing `height / 17` first truncates to 1 row per slot and
        // then caps at slot 16, so every row past 16 collapses onto
        // the pager — this is the Bold→Help mis-hit users see.
        // The mapping must stay proportional over the full range.
        let area = Rect::new(80, 4, 3, 26);
        // Slot 11 (Bold) spans rows ~16.8–18.4 in this panel.
        assert_eq!(App::hit_test_slot(area, 4 + 17), Some(11));
        assert_eq!(App::hit_test_slot(area, 4 + 18), Some(11));
        // Slot 15 (Help) spans rows ~22.9–24.5.
        assert_eq!(App::hit_test_slot(area, 4 + 23), Some(15));
        // Slot 16 (Pager) spans rows ~24.5–26.
        assert_eq!(App::hit_test_slot(area, 4 + 25), Some(16));
    }

    /// Simulate a mouse-move at `(col, row)` by stashing a fake panel
    /// area (same fixture as click tests) and routing through
    /// [`App::handle_mouse`].
    fn hover(app: &mut App, area: Rect, col: u16, row: u16) {
        app.icon_panel_area.set(Some(area));
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Moved,
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    fn hover_slot(app: &mut App, slot: u16) {
        hover(
            app,
            TEST_PANEL,
            TEST_PANEL.x + 1,
            TEST_PANEL.y + slot * 3 + 1,
        );
    }

    #[test]
    fn mouse_move_over_slot_sets_hovered_icon() {
        let mut app = App::new();
        hover_slot(&mut app, 0);
        assert_eq!(app.hovered_icon, Some((l123_graph::Panel::One, 0)));
        hover_slot(&mut app, 7);
        assert_eq!(app.hovered_icon, Some((l123_graph::Panel::One, 7)));
    }

    #[test]
    fn mouse_move_outside_panel_clears_hovered_icon() {
        let mut app = App::new();
        hover_slot(&mut app, 0);
        assert!(app.hovered_icon.is_some());
        hover(&mut app, TEST_PANEL, 10, 10);
        assert_eq!(app.hovered_icon, None);
    }

    #[test]
    fn mouse_move_over_pager_slot_does_not_set_hovered_icon() {
        // Slot 16 is the panel navigator; we deliberately exclude it
        // from hover-tooltip since its function is already rendered on
        // the slot itself ("Panel N of 7").
        let mut app = App::new();
        hover_slot(&mut app, 16);
        assert_eq!(app.hovered_icon, None);
    }

    #[test]
    fn mouse_move_without_cached_panel_clears_hovered_icon() {
        let mut app = App::new();
        hover_slot(&mut app, 3);
        assert!(app.hovered_icon.is_some());
        app.icon_panel_area.set(None);
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Moved,
            column: 82,
            row: 5,
            modifiers: KeyModifiers::NONE,
        });
        assert_eq!(app.hovered_icon, None);
    }

    #[test]
    fn mouse_move_on_different_panel_tracks_current_panel() {
        let mut app = App::new();
        app.current_panel = l123_graph::Panel::Three;
        hover_slot(&mut app, 2);
        assert_eq!(app.hovered_icon, Some((l123_graph::Panel::Three, 2)));
    }

    // ---- Grid click-to-move (Phase 1: READY only) ----

    /// Synthesize a left-click at the given screen position. Unlike
    /// [`click`], this doesn't stash a fake icon-panel rect — the
    /// headless render path already populates `last_grid_area` when
    /// `render_to_buffer` is called, which is what grid-click uses.
    fn click_at(app: &mut App, col: u16, row: u16) {
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    // Default layout (80×25): control panel = 4 lines (y 0..4), grid
    // = 20 lines (y 4..24), status = 1 line (y 24). Column-header row
    // sits at area.y = 4; body rows start at y = 5. ROW_GUTTER = 5
    // columns; default col width = 9, so col A spans x [5, 14), col B
    // spans x [14, 23).
    const HEADER_Y: u16 = 4;
    const BODY_TOP_Y: u16 = 5;
    const COL_A_X: u16 = 7; // anywhere in [5, 14)
    const COL_B_X: u16 = 16; // anywhere in [14, 23)

    #[test]
    fn mouse_click_on_cell_moves_pointer() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 2);
        assert_eq!(app.pointer().display_full(), "A:B3");
    }

    #[test]
    fn mouse_click_on_column_header_does_not_move_pointer() {
        // Column-letter row sits at the top of the grid area; clicking
        // it is reserved for Phase 2 (full-column select).
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_B_X, HEADER_Y);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_click_on_row_gutter_does_not_move_pointer() {
        // x in [0, ROW_GUTTER) is the row-number gutter; reserved for
        // Phase 2 (full-row select).
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, 2, BODY_TOP_Y + 2);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_click_below_grid_does_not_move_pointer() {
        // Status line row; clicks there are not bound to anything.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, 24);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_click_without_rendered_grid_is_ignored() {
        // No render happened yet, so last_grid_area is None — the
        // hit-test must bail cleanly rather than guessing geometry.
        let mut app = App::new();
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_click_in_menu_mode_does_not_move_pointer() {
        // Phase 1 restricts click-to-move to READY. In MENU (and
        // entry modes) we leave the pointer alone; Phase 2 will
        // add context-appropriate behavior.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Menu);
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 2);
        assert_eq!(app.pointer().display_full(), "A:A1");
        assert_eq!(app.mode, Mode::Menu);
    }

    #[test]
    fn mouse_click_respects_scroll_offset() {
        // With the viewport scrolled down, clicking the top body row
        // lands on the first visible row, not row 1.
        let mut app = App::new();
        app.wb_mut().viewport_row_offset = 10;
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        assert_eq!(app.pointer().display_full(), "A:A11");
    }

    #[test]
    fn mouse_click_respects_custom_column_width() {
        // After widening column A to 15, col B's body shifts right.
        // Clicking the new B region must still resolve to B, not A.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        drive_set_col_width(&mut app, 15);
        let _ = app.render_to_buffer(80, 25);
        // ROW_GUTTER(5) + 15 = 20 → col B starts at x=20.
        click_at(&mut app, 22, BODY_TOP_Y);
        assert_eq!(app.pointer().display_full(), "A:B1");
    }

    // ---- Phase 2: POINT-mode click-to-extend ----

    // Body cell fixtures for POINT-mode tests. Default col width 9,
    // ROW_GUTTER 5 → col C at x [23,32), col D at x [32,41).
    const COL_C_X: u16 = 25; // anywhere in [23, 32)
    const COL_D_X: u16 = 35; // anywhere in [32, 41)

    /// Enter POINT via `/RE` (Range Erase) — the shortest path into
    /// auto-anchored POINT. The prompt expects a range; auto-anchor
    /// fires at the current pointer.
    fn enter_point_via_range_erase(app: &mut App) {
        for c in ['/', 'R', 'E'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    #[test]
    fn mouse_click_in_anchored_point_extends_range() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        assert_eq!(app.mode, Mode::Point);
        // Anchor is set at A1 (the pointer when POINT entered).
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));

        // Click at C3 — range should become A1..C3. Anchor stays at A1.
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:C3");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_click_in_anchored_point_can_shrink_range() {
        // After extending to C3, clicking back at B2 must shrink the
        // range (anchor still A1, pointer now B2).
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
        assert_eq!(app.pointer().display_full(), "A:B2");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_click_in_unanchored_point_moves_pointer_without_reanchoring() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        // First Esc unanchors.
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), None);

        click_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:D4");
        // Still unanchored after click.
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), None);
    }

    #[test]
    fn mouse_click_on_gutter_in_point_is_noop() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        click_at(&mut app, 2, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:A1");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_click_in_point_does_not_exit_mode() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Point, "click must not drop out of POINT");
    }

    // ---- Phase 3: mid-entry cell-reference splicing on click ----

    /// Type a sequence into a fresh app and return it in VALUE mode
    /// with `buffer` typed after the value-starter. The leading `+`
    /// forces VALUE (SPEC §20 #6 — `=` is not a value starter in L123).
    fn enter_value_with(app: &mut App, buffer: &str) {
        for c in buffer.chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    fn entry_buffer(app: &App) -> String {
        app.entry
            .as_ref()
            .map(|e| e.buffer.clone())
            .unwrap_or_default()
    }

    #[test]
    fn mouse_click_in_value_after_plus_splices_short_address() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+");
        assert_eq!(app.mode, Mode::Value);
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Value, "splice must keep VALUE mode");
        assert_eq!(entry_buffer(&app), "+C3");
    }

    #[test]
    fn mouse_click_in_value_after_open_paren_splices() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "@SUM(");
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        assert_eq!(entry_buffer(&app), "@SUM(A1");
    }

    #[test]
    fn mouse_click_in_value_after_comma_splices() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "@MIN(A1,");
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
        assert_eq!(entry_buffer(&app), "@MIN(A1,B2");
    }

    #[test]
    fn mouse_click_in_value_after_range_dots_splices() {
        // `..` is 1-2-3's range separator; a trailing `.` must count
        // as a cell-ref-accepting context so you can click the end
        // corner of a range.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+A1..");
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 4);
        assert_eq!(entry_buffer(&app), "+A1..B5");
    }

    #[test]
    fn mouse_click_in_value_after_digit_is_noop() {
        // Splicing after `5` would produce `5C3` — nonsense. The
        // click must not corrupt the buffer.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "5");
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Value);
        assert_eq!(entry_buffer(&app), "5");
    }

    #[test]
    fn mouse_click_in_value_after_close_paren_is_noop() {
        // Close paren closes a sub-expression; the parser wants an
        // operator next, not another cell ref.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "@SUM(A1..A3)");
        let before = entry_buffer(&app);
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(entry_buffer(&app), before);
    }

    #[test]
    fn mouse_click_in_label_is_noop() {
        // Labels hold literal text; a mid-label cell ref is almost
        // certainly not what the user meant. Leave the buffer alone.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "hello ");
        assert_eq!(app.mode, Mode::Label);
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Label);
        assert_eq!(entry_buffer(&app), "hello ");
    }

    #[test]
    fn mouse_click_in_edit_after_operator_splices() {
        // Put a formula referencing A1 into B1, then F2 to EDIT the
        // source. Appending `+` and clicking C3 must splice.
        let mut app = App::new();
        // Move pointer to B1 (right once from A1).
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        enter_value_with(&mut app, "+A1");
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let _ = app.render_to_buffer(80, 25);
        app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Edit);
        app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(entry_buffer(&app), "+A1+C3");
    }

    #[test]
    fn mouse_click_on_gutter_in_value_is_noop() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+");
        click_at(&mut app, 2, BODY_TOP_Y + 2);
        assert_eq!(entry_buffer(&app), "+");
    }

    #[test]
    fn mouse_click_in_value_does_not_move_pointer() {
        // Splicing is a buffer operation; the *cell pointer* must
        // stay put so the entry still belongs to the originally-
        // selected cell.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+");
        click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    /// Drive the `/Worksheet Column Set-Width` menu path, typing `width`
    /// at the prompt and pressing Enter.
    fn drive_set_col_width(app: &mut App, width: u8) {
        for c in ['/', 'W', 'C', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        // The prompt seeds the current width and is `fresh` — any digit
        // replaces the seed; subsequent digits append.
        for c in width.to_string().chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    }

    /// After /WCS 15 on column A, the grid must actually draw column A
    /// at 15 characters wide — pushing the B-column header and B1's
    /// content to x = ROW_GUTTER + 15.
    #[test]
    fn set_col_width_widens_column_on_screen() {
        let mut app = App::new();
        // A1 = "alpha", B1 = "beta".
        for c in "alpha".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        for c in "beta".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        // Widen column A back at the A1 pointer.
        app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
        drive_set_col_width(&mut app, 15);
        assert_eq!(app.col_width_of(SheetId::A, 0), 15);

        let buf = app.render_to_buffer(80, 25);
        // The body row for row 1 sits one line below the column-header
        // row (PANEL_HEIGHT + 1).
        let body_y = PANEL_HEIGHT + 1;
        // Column A's contents must start at x = ROW_GUTTER and run 15
        // chars; "alpha" is left-aligned (apostrophe prefix) and padded
        // with spaces to the full column width.
        let a_slot: String = (0..15)
            .map(|i| buf[(ROW_GUTTER + i, body_y)].symbol().to_string())
            .collect();
        assert_eq!(a_slot, "alpha          ");
        // Column B must start 15 characters later — not 9.
        let b_slot: String = (0..9)
            .map(|i| buf[(ROW_GUTTER + 15 + i, body_y)].symbol().to_string())
            .collect();
        assert_eq!(b_slot, "beta     ");

        // And the header row reflects the same geometry.
        let header_y = PANEL_HEIGHT;
        let b_header = buf[(ROW_GUTTER + 15 + 4, header_y)].symbol(); // center of 9-wide slot
        assert_eq!(b_header, "B");
    }

    #[test]
    fn iterm2_env_hint_matches_term_program_variants() {
        // Apple's iTerm2 sets TERM_PROGRAM=iTerm.app.
        assert!(is_iterm2_compatible_env(Some("iTerm.app"), None));
        // SSH into a shell from iTerm2: TERM_PROGRAM is whatever the
        // remote shell wants (often tmux / Apple_Terminal), but
        // LC_TERMINAL is forwarded.
        assert!(is_iterm2_compatible_env(Some("tmux"), Some("iTerm2")));
        // Other hosts that speak the OSC 1337 image protocol.
        assert!(is_iterm2_compatible_env(Some("WezTerm"), None));
        assert!(is_iterm2_compatible_env(Some("mintty"), None));
        assert!(is_iterm2_compatible_env(Some("WarpTerminal"), None));
    }

    #[test]
    fn iterm2_env_hint_rejects_other_terminals() {
        assert!(!is_iterm2_compatible_env(Some("ghostty"), None));
        assert!(!is_iterm2_compatible_env(Some("Apple_Terminal"), None));
        assert!(!is_iterm2_compatible_env(Some("xterm-kitty"), None));
        assert!(!is_iterm2_compatible_env(None, None));
        assert!(!is_iterm2_compatible_env(None, Some("")));
    }

    // --- :Format Bold/Italic/Underline Set|Clear (step 2 of the
    //     text-style slice: core storage, command execution, journaled
    //     undo). Menu wiring and rendering land in later steps.

    fn one_cell_range(addr: Address) -> Range {
        Range {
            start: addr,
            end: addr,
        }
    }

    #[test]
    fn range_text_style_set_bold_records_override() {
        let mut app = App::new();
        app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
        );
    }

    #[test]
    fn range_text_style_set_then_undo_restores_plain() {
        let mut app = App::new();
        app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
        app.undo();
        assert!(
            !app.wb().cell_text_styles.contains_key(&Address::A1),
            "bold should be gone after Alt-F4"
        );
    }

    #[test]
    fn range_text_style_bold_then_italic_composes_bits() {
        let mut app = App::new();
        let r = one_cell_range(Address::A1);
        app.execute_range_text_style(r, TextStyle::BOLD, true);
        app.execute_range_text_style(r, TextStyle::ITALIC, true);
        let s = app.wb().cell_text_styles.get(&Address::A1).copied();
        assert_eq!(
            s,
            Some(TextStyle {
                bold: true,
                italic: true,
                underline: false
            }),
        );
    }

    #[test]
    fn range_text_style_clear_drops_only_named_bits() {
        let mut app = App::new();
        let r = one_cell_range(Address::A1);
        app.execute_range_text_style(r, TextStyle::BOLD.merge(TextStyle::ITALIC), true);
        app.execute_range_text_style(r, TextStyle::BOLD, false);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn range_text_style_clearing_last_bit_removes_entry() {
        let mut app = App::new();
        let r = one_cell_range(Address::A1);
        app.execute_range_text_style(r, TextStyle::BOLD, true);
        app.execute_range_text_style(r, TextStyle::BOLD, false);
        assert!(
            !app.wb().cell_text_styles.contains_key(&Address::A1),
            "plain style should not leave an empty entry in the map"
        );
    }

    #[test]
    fn range_text_style_reset_clears_all_attributes() {
        let mut app = App::new();
        let r = one_cell_range(Address::A1);
        let all = TextStyle {
            bold: true,
            italic: true,
            underline: true,
        };
        app.execute_range_text_style(r, all, true);
        app.execute_range_text_style(r, all, false);
        assert!(!app.wb().cell_text_styles.contains_key(&Address::A1));
    }

    #[test]
    fn range_text_style_undo_restores_partial_prior_state() {
        let mut app = App::new();
        let r = one_cell_range(Address::A1);
        // Start: bold on A1 (the prior state we should restore to).
        app.execute_range_text_style(r, TextStyle::BOLD, true);
        // Clear the journal so we're only testing the next undo.
        app.wb_mut().journal.clear();
        // Now apply italic.
        app.execute_range_text_style(r, TextStyle::ITALIC, true);
        app.undo();
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
            "undoing the italic-set should leave the prior bold-only state"
        );
    }

    #[test]
    fn text_style_modifier_maps_bits_to_ratatui_modifier() {
        assert_eq!(text_style_modifier(TextStyle::PLAIN), Modifier::empty());
        assert_eq!(text_style_modifier(TextStyle::BOLD), Modifier::BOLD);
        assert_eq!(text_style_modifier(TextStyle::ITALIC), Modifier::ITALIC);
        assert_eq!(
            text_style_modifier(TextStyle::UNDERLINE),
            Modifier::UNDERLINED,
        );
        let all = TextStyle {
            bold: true,
            italic: true,
            underline: true,
        };
        assert_eq!(
            text_style_modifier(all),
            Modifier::BOLD | Modifier::ITALIC | Modifier::UNDERLINED,
        );
    }

    /// Read back the ratatui `Modifier` bits on the first cell of an
    /// address in the rendered buffer. Uses the same visible-column
    /// layout math as [`App::cell_rendered_text`].
    fn cell_modifier_at(app: &App, buf: &Buffer, addr: Address) -> Modifier {
        let dr = (addr.row - app.wb().viewport_row_offset) as u16;
        let y = PANEL_HEIGHT + 1 + dr;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = app.visible_column_layout(content_width);
        let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == addr.col).unwrap();
        let x = ROW_GUTTER + x_off;
        buf[(x, y)].style().add_modifier
    }

    #[test]
    fn bold_style_renders_with_ratatui_bold_modifier() {
        let mut app = App::new();
        // Put a label at B5 so there's a character to style.
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        for c in "hello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let target = Address::new(SheetId::A, 1, 4); // B5
        app.execute_range_text_style(one_cell_range(target), TextStyle::BOLD, true);

        let buf = app.render_to_buffer(80, 25);
        let modifier = cell_modifier_at(&app, &buf, target);
        assert!(
            modifier.contains(Modifier::BOLD),
            "expected BOLD in buffer cell's modifier, got {modifier:?}"
        );
    }

    #[test]
    fn compound_style_renders_with_all_three_modifiers() {
        let mut app = App::new();
        let target = Address::A1;
        let all = TextStyle {
            bold: true,
            italic: true,
            underline: true,
        };
        app.execute_range_text_style(one_cell_range(target), all, true);
        let buf = app.render_to_buffer(80, 25);
        let modifier = cell_modifier_at(&app, &buf, target);
        assert!(modifier.contains(Modifier::BOLD));
        assert!(modifier.contains(Modifier::ITALIC));
        assert!(modifier.contains(Modifier::UNDERLINED));
    }

    #[test]
    fn plain_cells_have_no_text_style_modifier() {
        let app = App::new();
        let buf = app.render_to_buffer(80, 25);
        let modifier = cell_modifier_at(&app, &buf, Address::new(SheetId::A, 2, 2));
        assert!(!modifier.contains(Modifier::BOLD));
        assert!(!modifier.contains(Modifier::ITALIC));
        assert!(!modifier.contains(Modifier::UNDERLINED));
    }

    #[test]
    fn text_style_follows_label_spill_into_adjacent_columns() {
        // Long italic label at A1 spills across B/C/D/E.  The
        // overflow characters should render italic too, not plain.
        let mut app = App::new();
        for c in "INCOME SUMMARY 1991: Sloane Camera and Video".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        keys(&mut app, ":FIS~");

        let buf = app.render_to_buffer(100, 25);

        // A1 itself is italic (its home column).
        assert!(
            cell_modifier_at(&app, &buf, Address::A1).contains(Modifier::ITALIC),
            "A1 (owner) should be italic"
        );
        // B1, C1, D1 are empty neighbors the label overflowed into —
        // they must pick up the owner's italic attribute.
        for col in 1..=3u16 {
            let addr = Address::new(SheetId::A, col, 0);
            let m = cell_modifier_at(&app, &buf, addr);
            assert!(
                m.contains(Modifier::ITALIC),
                "spill into column {col} should be italic, got {m:?}"
            );
        }
    }

    #[test]
    fn text_style_does_not_leak_past_spill_extent() {
        // Italic label at A1 long enough to fill B1 exactly; C1 must
        // remain plain (the label doesn't reach it).
        let mut app = App::new();
        // A1 + B1 default widths = 9 + 9 = 18.  "123456789abcdefgh" is
        // 17 chars — fits inside A1..B1 with no leftover for C1.
        for c in "123456789abcdefgh".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        keys(&mut app, ":FIS~");

        let buf = app.render_to_buffer(100, 25);
        // C1 is never touched by the spill, so it should stay plain.
        let c1 = Address::new(SheetId::A, 2, 0);
        let m = cell_modifier_at(&app, &buf, c1);
        assert!(
            !m.contains(Modifier::ITALIC),
            "C1 beyond spill extent should be plain, got {m:?}"
        );
    }

    #[test]
    fn bold_on_highlighted_cell_keeps_both_reversed_and_bold() {
        let mut app = App::new();
        // Pointer starts at A1, which is the highlighted cell in READY.
        app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
        let buf = app.render_to_buffer(80, 25);
        let modifier = cell_modifier_at(&app, &buf, Address::A1);
        assert!(
            modifier.contains(Modifier::REVERSED),
            "pointer reverse-video"
        );
        assert!(modifier.contains(Modifier::BOLD), "bold overlay");
    }

    /// Replay a string of characters as individual READY-mode key
    /// presses, just like a transcript.  `~` stands for Enter, matching
    /// the acceptance-transcript convention.
    fn keys(app: &mut App, s: &str) {
        for c in s.chars() {
            match c {
                '~' => app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE)),
                ch => app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::NONE)),
            }
        }
    }

    #[test]
    fn colon_enters_wysiwyg_menu_mode() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Menu);
        let state = app.menu.as_ref().expect("menu state");
        // First item under WYSIWYG_ROOT is "Worksheet".
        let first = state.level().first().expect("items visible");
        assert_eq!(first.name, "Worksheet");
    }

    #[test]
    fn colon_f_b_s_bolds_the_selected_range() {
        let mut app = App::new();
        keys(&mut app, ":FBS~");
        // After ~ commits POINT with default single-cell range at A1.
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
        );
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn colon_f_i_s_italicizes_range() {
        let mut app = App::new();
        keys(&mut app, ":FIS~");
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn colon_f_u_s_underlines_range() {
        let mut app = App::new();
        keys(&mut app, ":FUS~");
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::UNDERLINE),
        );
    }

    #[test]
    fn colon_f_b_c_clears_bold_left_other_bits_alone() {
        let mut app = App::new();
        keys(&mut app, ":FBS~");
        keys(&mut app, ":FIS~");
        keys(&mut app, ":FBC~");
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn colon_f_r_resets_all_attributes() {
        let mut app = App::new();
        keys(&mut app, ":FBS~");
        keys(&mut app, ":FIS~");
        keys(&mut app, ":FUS~");
        keys(&mut app, ":FR~");
        assert!(!app.wb().cell_text_styles.contains_key(&Address::A1));
    }

    #[test]
    fn line1_shows_bold_marker_on_labeled_cell() {
        let mut app = App::new();
        for c in "hello".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        keys(&mut app, ":FBS~");
        let buf = app.render_to_buffer(80, 25);
        let line1 = App::line_text(&buf, 0);
        assert!(
            line1.contains("{Bold}"),
            "expected {{Bold}} in line 1, got {line1:?}"
        );
    }

    #[test]
    fn line1_shows_compound_marker_with_space_separator() {
        let mut app = App::new();
        keys(&mut app, ":FBS~");
        keys(&mut app, ":FIS~");
        keys(&mut app, ":FUS~");
        let buf = app.render_to_buffer(80, 25);
        let line1 = App::line_text(&buf, 0);
        assert!(
            line1.contains("{Bold Italic Underline}"),
            "expected {{Bold Italic Underline}} in line 1, got {line1:?}"
        );
    }

    #[test]
    fn line1_has_no_style_marker_on_plain_cell() {
        let app = App::new();
        let buf = app.render_to_buffer(80, 25);
        let line1 = App::line_text(&buf, 0);
        assert!(
            !line1.contains('{'),
            "plain cell should not show a style marker, got {line1:?}"
        );
    }

    #[test]
    fn line1_marker_follows_format_tag_on_numeric_cell() {
        let mut app = App::new();
        // Type a number so the cell gets a (G) format tag.
        for c in "42".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        // Go back up to the just-entered cell before applying style.
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
        keys(&mut app, ":FBS~");
        let buf = app.render_to_buffer(80, 25);
        let line1 = App::line_text(&buf, 0);
        let tag_pos = line1.find("(G)").expect("format tag present");
        let marker_pos = line1.find("{Bold}").expect("style marker present");
        assert!(
            tag_pos < marker_pos,
            "format tag should precede style marker: {line1:?}"
        );
    }

    #[test]
    fn line1_marker_disappears_after_clearing_last_style_bit() {
        let mut app = App::new();
        keys(&mut app, ":FBS~");
        keys(&mut app, ":FBC~");
        let buf = app.render_to_buffer(80, 25);
        let line1 = App::line_text(&buf, 0);
        assert!(
            !line1.contains('{'),
            "marker should be gone after clear, got {line1:?}"
        );
    }

    #[test]
    fn text_style_survives_xlsx_save_and_retrieve() {
        use std::process;
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos())
            .unwrap_or(0);
        let dir =
            std::env::temp_dir().join(format!("l123_ui_style_rt_{}_{}", process::id(), nanos));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("style.xlsx");

        let mut app = App::new();
        // A1: label "hi" with bold + italic.
        for c in "hi".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        // Enter commits in place (SPEC §8) — pointer still at A1.
        keys(&mut app, ":FBS~");
        keys(&mut app, ":FIS~");
        // A2: label "bye" with underline.
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        for c in "bye".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        keys(&mut app, ":FUS~");

        // Probe the in-memory map right before save.
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD.merge(TextStyle::ITALIC)),
        );
        assert_eq!(
            app.wb()
                .cell_text_styles
                .get(&Address::new(SheetId::A, 0, 1))
                .copied(),
            Some(TextStyle::UNDERLINE),
        );

        app.save_workbook_to(path.clone());

        let mut reopened = App::new();
        reopened.load_workbook_from(path.clone());

        assert_eq!(
            reopened.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle {
                bold: true,
                italic: true,
                underline: false
            }),
            "A1 should round-trip bold + italic"
        );
        assert_eq!(
            reopened
                .wb()
                .cell_text_styles
                .get(&Address::new(SheetId::A, 0, 1))
                .copied(),
            Some(TextStyle::UNDERLINE),
            "A2 should round-trip underline"
        );

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    #[test]
    fn colon_q_closes_wysiwyg_menu() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('Q'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.menu.is_none());
    }

    #[test]
    fn range_text_style_applies_across_multi_cell_range() {
        let mut app = App::new();
        let r = Range {
            start: Address::new(SheetId::A, 0, 0),
            end: Address::new(SheetId::A, 1, 1),
        };
        app.execute_range_text_style(r, TextStyle::BOLD, true);
        for col in 0..=1 {
            for row in 0..=1 {
                let a = Address::new(SheetId::A, col, row);
                assert_eq!(
                    app.wb().cell_text_styles.get(&a).copied(),
                    Some(TextStyle::BOLD),
                    "cell ({col},{row}) should be bold"
                );
            }
        }
    }
}
