//! The app loop: ratatui + crossterm, control panel + grid + status line.
//!
//! Scope as of M1 cycle 2:
//! - READY / LABEL / VALUE modes with first-character dispatch (LABEL only
//!   implemented this cycle; VALUE lands in cycle 3).
//! - `'` auto-prefixed labels. Enter commits; `/QY` quits.
//! - Three-line control panel, mode indicator, cell readout.

use std::cell::Cell;
use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
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
use l123_core::cell_render::{apply_halign_to_rendered, halign_to_label_prefix, label_text_bounds};
use l123_core::{
    address::col_to_letters, label::is_value_starter, plan_row_spill, render_label,
    render_value_in_cell, Address, Alignment, Border, CellContents, Comment, CurrencyPosition,
    DateIntl, ErrKind, Fill, FontStyle, Format, FormatKind, HAlign, International, LabelPrefix,
    Merge, Mode, NegativeStyle, Punctuation, Range, RangeInput, RgbColor, SheetId, SheetState,
    SpillSlot, Table, TextStyle, TimeIntl, Value,
};
use l123_engine::{CellView, Engine, IronCalcEngine, RecalcMode};
use l123_graph::{GraphDef, GraphType, Series};
use l123_macro::{lex as lex_macro, lex_actions as lex_macro_actions, MacroAction, MacroKey};
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

use crate::help::HelpState;

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

/// `/Worksheet Global Default Graph Save` — default file format
/// /Graph Save writes when the user-typed name has no extension.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum GraphSaveFormat {
    #[default]
    Cgm,
    Pic,
}

impl GraphSaveFormat {
    pub fn label(self) -> &'static str {
        match self {
            GraphSaveFormat::Cgm => "Cgm",
            GraphSaveFormat::Pic => "Pic",
        }
    }
}

/// `/Worksheet Global Default Graph Group` — auto-graph orientation
/// applied by `/Graph Group`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum GraphGroupOrientation {
    #[default]
    Columnwise,
    Rowwise,
}

impl GraphGroupOrientation {
    pub fn label(self) -> &'static str {
        match self {
            GraphGroupOrientation::Columnwise => "Columnwise",
            GraphGroupOrientation::Rowwise => "Rowwise",
        }
    }
}

/// Workbook-wide defaults persisted to `L123.CNF` by `/Worksheet Global
/// Default Update`. Mirrors the 1-2-3 R3.4a `123R31.CNF` knobs that
/// every new session inherits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlobalDefaults {
    pub printer_interface: u8,
    pub printer_autolf: bool,
    pub printer_left: u16,
    pub printer_right: u16,
    pub printer_top: u16,
    pub printer_bottom: u16,
    pub printer_pg_length: u16,
    pub printer_wait: bool,
    pub printer_setup: String,
    pub printer_name: String,
    pub default_dir: String,
    pub temp_dir: String,
    pub ext_save: String,
    pub ext_list: String,
    pub autoexec: bool,
    pub graph_group: GraphGroupOrientation,
    pub graph_save: GraphSaveFormat,
}

impl Default for GlobalDefaults {
    fn default() -> Self {
        Self {
            printer_interface: 1,
            printer_autolf: false,
            printer_left: 4,
            printer_right: 76,
            printer_top: 2,
            printer_bottom: 2,
            printer_pg_length: 66,
            printer_wait: false,
            printer_setup: String::new(),
            printer_name: String::new(),
            default_dir: String::new(),
            temp_dir: String::new(),
            ext_save: "xlsx".into(),
            ext_list: String::new(),
            autoexec: true,
            graph_group: GraphGroupOrientation::Columnwise,
            graph_save: GraphSaveFormat::Cgm,
        }
    }
}

impl GlobalDefaults {
    /// Render this struct as the additive `# WGD defaults` block
    /// appended below an existing L123.CNF body. The block uses the
    /// same `key = value` syntax the CNF reader already accepts.
    pub fn render_cnf_block(&self) -> String {
        let mut out = String::new();
        out.push_str("# Persisted by /Worksheet Global Default Update\n");
        out.push_str(&format!(
            "wgd_printer_interface = {}\n",
            self.printer_interface
        ));
        out.push_str(&format!("wgd_printer_autolf = {}\n", self.printer_autolf));
        out.push_str(&format!("wgd_printer_left = {}\n", self.printer_left));
        out.push_str(&format!("wgd_printer_right = {}\n", self.printer_right));
        out.push_str(&format!("wgd_printer_top = {}\n", self.printer_top));
        out.push_str(&format!("wgd_printer_bottom = {}\n", self.printer_bottom));
        out.push_str(&format!(
            "wgd_printer_pg_length = {}\n",
            self.printer_pg_length
        ));
        out.push_str(&format!("wgd_printer_wait = {}\n", self.printer_wait));
        out.push_str(&format!(
            "wgd_printer_setup = \"{}\"\n",
            escape_cnf(&self.printer_setup)
        ));
        out.push_str(&format!(
            "wgd_printer_name = \"{}\"\n",
            escape_cnf(&self.printer_name)
        ));
        out.push_str(&format!(
            "wgd_dir = \"{}\"\n",
            escape_cnf(&self.default_dir)
        ));
        out.push_str(&format!("wgd_temp = \"{}\"\n", escape_cnf(&self.temp_dir)));
        out.push_str(&format!(
            "wgd_ext_save = \"{}\"\n",
            escape_cnf(&self.ext_save)
        ));
        out.push_str(&format!(
            "wgd_ext_list = \"{}\"\n",
            escape_cnf(&self.ext_list)
        ));
        out.push_str(&format!("wgd_autoexec = {}\n", self.autoexec));
        out.push_str(&format!(
            "wgd_graph_group = {}\n",
            match self.graph_group {
                GraphGroupOrientation::Columnwise => "columnwise",
                GraphGroupOrientation::Rowwise => "rowwise",
            }
        ));
        out.push_str(&format!(
            "wgd_graph_save = {}\n",
            match self.graph_save {
                GraphSaveFormat::Cgm => "cgm",
                GraphSaveFormat::Pic => "pic",
            }
        ));
        out
    }

    /// Write defaults to `path`, replacing any prior `# Persisted by
    /// /Worksheet Global Default Update` block while preserving every
    /// other line in the file (so user-managed `user`, `log_file`,
    /// etc. survive an Update).
    pub fn write_to_path(&self, path: &std::path::Path) -> std::io::Result<()> {
        if let Some(parent) = path.parent() {
            if !parent.as_os_str().is_empty() {
                std::fs::create_dir_all(parent)?;
            }
        }
        let prior = std::fs::read_to_string(path).unwrap_or_default();
        let preserved = strip_wgd_block(&prior);
        let mut body = preserved;
        if !body.is_empty() && !body.ends_with('\n') {
            body.push('\n');
        }
        body.push_str(&self.render_cnf_block());
        std::fs::write(path, body)
    }
}

fn escape_cnf(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"")
}

/// Drop every line of the persisted-WGD block from `body`, leaving
/// every other line intact. The block runs from the marker comment to
/// the next blank line or EOF.
fn strip_wgd_block(body: &str) -> String {
    let mut out = String::with_capacity(body.len());
    let mut skipping = false;
    for line in body.lines() {
        if line
            .trim_start()
            .starts_with("# Persisted by /Worksheet Global Default Update")
        {
            skipping = true;
            continue;
        }
        if skipping {
            if line.trim_start().starts_with("wgd_") {
                continue;
            }
            skipping = false;
        }
        out.push_str(line);
        out.push('\n');
    }
    out
}

/// Selects which screen `Mode::Stat` renders. `Worksheet` is the
/// `/Worksheet Status` panel, `Defaults` is `/Worksheet Global Default
/// Status`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
enum StatView {
    #[default]
    Worksheet,
    Defaults,
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

/// `/Worksheet Global Default Other Clock` — what occupies the
/// status-line clock slot.
///
/// Default is [`ClockDisplay::Filename`]: the active workbook's
/// filename takes the slot when one exists, falling back to the
/// 24-hour clock so an unsaved session still shows the date. This
/// keeps prior status-line behavior intact.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ClockDisplay {
    Standard,
    International,
    None,
    #[default]
    Filename,
}

impl ClockDisplay {
    pub fn label(self) -> &'static str {
        match self {
            ClockDisplay::Standard => "Standard",
            ClockDisplay::International => "International",
            ClockDisplay::None => "None",
            ClockDisplay::Filename => "Filename",
        }
    }
}

/// `:Display Mode` — picks the default style for cells with no
/// xlsx-imported fill or font color. Cells that *do* carry a fill or
/// font color always paint that color regardless of mode.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DisplayMode {
    /// Today's behavior — no default fg/bg, terminal defaults show
    /// through. Closest analog to R3.4a B&W mode in a TUI.
    #[default]
    BW,
    /// Paper look — white background, black text on otherwise-unstyled
    /// cells. Matches R3.4a's default WYSIWYG appearance.
    Color,
    /// Inverse paper — black background, white text on otherwise-
    /// unstyled cells.
    Reverse,
}

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

#[derive(Debug)]
struct Entry {
    kind: EntryKind,
    buffer: String,
    /// Byte index into `buffer`, 0..=buffer.len(). Always lands on a
    /// char boundary. Initialized to `buffer.len()` (cursor at end —
    /// matches typing-into-an-empty-buffer behavior).
    cursor: usize,
}

/// All per-file state — one instance per active file. Session-level
/// fields (mode, menu, entry buffer, …) live on [`App`].
struct Workbook {
    engine: IronCalcEngine,
    cells: HashMap<Address, CellContents>,
    cell_formats: HashMap<Address, Format>,
    /// Workbook-wide default cell format set by `/Worksheet Global
    /// Format`. Cells without a `cell_formats` entry inherit this.
    /// Initialized to General.
    global_format: Format,
    /// `/Worksheet Global Default Other International` — punctuation,
    /// date/time intl style, negative style, and currency symbol/
    /// position. Threaded into `format_number` and `parse_typed_value`
    /// so cell display and number entry honor the configured locale.
    /// Persistence to L123.CNF via `/WGDU` is out of scope; session-
    /// only for now.
    international: International,
    /// Per-cell text-style overrides (bold / italic / underline) set
    /// by the WYSIWYG `:Format Bold|Italic|Underline Set|Clear`
    /// commands.  Empty style = no entry.
    cell_text_styles: HashMap<Address, TextStyle>,
    /// Per-cell explicit alignment from an xlsx import (or a future
    /// /Range Alignment command).  Default alignment = no entry, so
    /// the label-prefix / number-right-align contract still governs
    /// uncharted cells.
    cell_alignments: HashMap<Address, Alignment>,
    /// Per-cell background-fill color from an xlsx import.  Default
    /// (no fill) = no entry; the terminal default shows through.
    /// Rendered via `Style::bg(Color::Rgb(...))` at grid-paint time.
    cell_fills: HashMap<Address, Fill>,
    /// Per-cell xlsx-derived font attributes (foreground color, size,
    /// strikethrough).  Sits alongside `cell_text_styles` (the 1-2-3
    /// WYSIWYG bold/italic/underline triple); the two maps can both
    /// apply to the same cell.  Size is preserve-only.
    cell_font_styles: HashMap<Address, FontStyle>,
    /// Per-cell border edges from xlsx imports.  All four sides + color
    /// round-trip; v1 renders only **right-edge** borders (overlaying
    /// a box-drawing glyph on the rightmost column of the cell's slot).
    /// Top, bottom, and left borders are preserved on save but not yet
    /// rendered — adding row-direction borders requires a grid layout
    /// change (seam rows) that's out of scope here.
    cell_borders: HashMap<Address, Border>,
    /// Per-cell comments from xlsx imports.  Renders as a small
    /// corner marker (`'`) on the cell's right-edge column; when the
    /// pointer lands on the cell, the author + text appears on
    /// control-panel line 3.  Note (IronCalc 0.7): the xlsx exporter
    /// drops comments — we preserve them in-memory and render them,
    /// but `/FS` will lose them until the upstream gap closes.
    comments: HashMap<Address, Comment>,
    /// Merged ranges by sheet.  The grid renderer paints the
    /// anchor's content across the merge's column span; non-anchor
    /// cells in the same row render blank (and block label-spill
    /// from neighbors).  Multi-row merges: the anchor's content
    /// shows on the anchor's row only; subsequent rows of the merge
    /// area render blank — top-aligned, like Excel.  Cursor
    /// navigation does NOT yet snap to the anchor (deferred).
    merges: HashMap<SheetId, Vec<Merge>>,
    /// Per-sheet frozen-pane counts: `(rows, cols)` indicate how many
    /// rows from the top and columns from the left stay pinned in the
    /// viewport while the rest of the grid scrolls.  `(0, 0)` (or no
    /// entry) means no freeze.  Round-trips natively through xlsx.
    frozen: HashMap<SheetId, (u32, u16)>,
    /// Per-sheet visibility from xlsx imports.  Sheets not in the
    /// map default to `Visible`.  Hidden / VeryHidden sheets are
    /// skipped by `Ctrl-PgUp/PgDn` navigation; loaded files with a
    /// hidden first sheet get the pointer redirected to the first
    /// `Visible` sheet on import so the user lands somewhere they
    /// can interact with.
    sheet_states: HashMap<SheetId, SheetState>,
    /// Excel tables (named ranges with header / autofilter / totals
    /// metadata) by sheet.  v1: round-trip-only — preserved through
    /// the engine on save, but no UI surface yet (no filter widgets,
    /// no `/Data Query Define` integration).  IronCalc 0.7's xlsx
    /// exporter doesn't write tables, so `/FS` drops them today
    /// (pinned by `tables_are_dropped_on_xlsx_save_upstream_gap`).
    tables: HashMap<SheetId, Vec<Table>>,
    /// Per-sheet tab color from an xlsx import.  When set, the sheet's
    /// letter in the status-line indicator renders with this fg color.
    /// Note (IronCalc 0.7): the xlsx *export* path drops tab colors —
    /// we preserve them in the model and render them, but `/FS` will
    /// not carry them back to disk until the upstream gap closes.
    sheet_colors: HashMap<SheetId, RgbColor>,
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
    /// True when the workbook has unsaved changes. Drives the `/QY`
    /// warn-on-quit second confirm. Flipped on by mutating commits;
    /// cleared on a successful `/FS`.
    dirty: bool,
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
    /// Range names defined via `/Range Name Create`. Keyed by the
    /// lowercased name (Lotus 1-2-3 names are case-insensitive). The
    /// engine also stores names for formula resolution; this UI-side
    /// mirror is what POINT typed-buffer name resolution reads, so we
    /// don't need to round-trip through the engine to look up a range.
    named_ranges: HashMap<String, Range>,
    /// Optional notes attached to named ranges by `/Range Name Note
    /// Create`. Keyed identically to `named_ranges` (lowercased name).
    name_notes: HashMap<String, String>,
    /// Cells that `/Range Unprot` has marked as writable. Cells not
    /// in this set are "protected" by default. Has no effect unless
    /// `App::global_protection` is on.
    cell_unprotected: HashSet<Address>,
}

impl Workbook {
    /// Find the merge containing `addr`, if any.  O(N) over the
    /// sheet's merge list — fine for the small N (~dozens) typical
    /// of real workbooks; revisit if a fixture pushes it past 1000.
    fn merge_at(&self, addr: Address) -> Option<Merge> {
        self.merges
            .get(&addr.sheet)?
            .iter()
            .find(|m| m.contains(addr))
            .copied()
    }

    fn new() -> Self {
        Self {
            engine: IronCalcEngine::new().expect("IronCalc engine init"),
            cells: HashMap::new(),
            cell_formats: HashMap::new(),
            global_format: Format::GENERAL,
            international: International::default(),
            cell_text_styles: HashMap::new(),
            cell_alignments: HashMap::new(),
            cell_fills: HashMap::new(),
            cell_font_styles: HashMap::new(),
            cell_borders: HashMap::new(),
            comments: HashMap::new(),
            merges: HashMap::new(),
            frozen: HashMap::new(),
            sheet_states: HashMap::new(),
            tables: HashMap::new(),
            sheet_colors: HashMap::new(),
            col_widths: HashMap::new(),
            default_col_width: 9,
            hidden_cols: HashSet::new(),
            active_path: None,
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
            .unwrap_or(self.global_format)
    }

    fn international(&self) -> &International {
        &self.international
    }
}

/// Geometry the icon panel last occupied: cell rect plus the actual
/// rendered image pixel height and the terminal cell pixel height. The
/// PNG's 1:17 aspect rarely lands on integer-cell boundaries, so each
/// icon spans a fractional cell — hit-testing must work in pixels and
/// then convert back to a slot index.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct IconPanelGeom {
    rect: Rect,
    rendered_px_h: u32,
    font_px_h: u16,
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

/// Adjacent-cell direction for `/Range Name Labels`. Each label in
/// the picked range gets a name pointing one cell in this direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LabelDirection {
    Right,
    Down,
    Left,
    Up,
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
    /// `/Print Encoded`: write `setup_string` followed by the ASCII
    /// page bytes to this path. Raw printer-ready output — no PDF
    /// branching, no `lp` invocation.
    Encoded(PathBuf),
}

#[derive(Debug, Clone)]
struct PrintSession {
    destination: PrintDestination,
    /// One or more print ranges. Empty until `/PF Range` runs. Multiple
    /// ranges (typed `A1..B2,C3..D4` in POINT) are emitted in order;
    /// each range is a separate "page" of the output (Lotus separates
    /// ranges with a form-feed when printed; here we emit a blank line
    /// between them in the unformatted/file path).
    ranges: Vec<Range>,
    /// Three-part header string (`L|C|R`). Empty means no header.
    header: String,
    /// Three-part footer string (`L|C|R`). Empty means no footer.
    footer: String,
    /// Printer init/escape string. Prepended verbatim to the output —
    /// ahead of the ASCII page bytes for a `.prn` file write, or piped
    /// to CUPS `lp` ahead of the ASCII stream for `/Print Printer`.
    /// Empty = no setup. PDF output ignores it (escape codes are
    /// meaningless inside a PDF body).
    setup_string: String,
    /// CUPS queue name passed as `lp -d <name>` for `/Print Printer`.
    /// Empty = use the system default printer. Stored regardless of
    /// destination kind; only consumed when destination is `Printer`.
    lp_destination: String,
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

    fn new_encoded(path: PathBuf) -> Self {
        Self::with_destination(PrintDestination::Encoded(path))
    }

    fn with_destination(destination: PrintDestination) -> Self {
        Self {
            destination,
            ranges: Vec::new(),
            header: String::new(),
            footer: String::new(),
            setup_string: String::new(),
            lp_destination: String::new(),
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
        self.setup_string.clear();
        self.lp_destination.clear();
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
    /// Restore per-cell alignment overrides after `:Format Alignment
    /// Left|Right|Center|General`.  `None` = no override before.
    RangeAlignment {
        entries: Vec<(Address, Option<Alignment>)>,
    },
    /// Restore per-cell fill and font-color overrides after `:Format
    /// Color Background|Text <color>` or `:Format Color Reset`. Each
    /// entry carries the prior fill *and* font style, since Reset
    /// touches both channels.
    RangeColor {
        entries: Vec<(Address, Option<Fill>, Option<FontStyle>)>,
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
    /// Restore the workbook-wide default cell format set by `/Worksheet
    /// Global Format`.
    GlobalFormat { prev: Format },
    /// Restore the workbook-wide international settings (punctuation,
    /// dates, times, negative style, currency) as a single snapshot.
    /// One entry per `/WGDOI ...` mutation.
    GlobalInternational { prev: International },
    /// Restore the workbook-wide default label prefix.
    DefaultLabelPrefix { prev: LabelPrefix },
    /// Restore one sheet's frozen-pane setting after `/Worksheet
    /// Titles`. `None` = no freeze before the command ran.
    Frozen {
        sheet: SheetId,
        prev: Option<(u32, u16)>,
    },
    /// Restore one sheet's visibility after `/Worksheet Hide`.
    SheetVisibility { sheet: SheetId, prev: SheetState },
    /// Restore the workbook's named-range map after `/Range Name
    /// Reset`. The captured pairs are re-defined wholesale on undo
    /// (engine + UI mirror), preserving names that pre-existed before
    /// the wipe.
    RangeNameReset { prev: Vec<(String, Range)> },
    /// Undo of `/Range Name Labels`: drop the names that were
    /// successfully created by the command (any pre-existing names
    /// that were overwritten are captured in `overwritten` and
    /// restored).
    RangeNameLabels {
        created: Vec<String>,
        overwritten: Vec<(String, Range)>,
    },
    /// Restore the workbook's named-range map and notes after
    /// `/Range Name Undefine`. The single dropped name + range is
    /// re-defined; the `cell_writes` block carries the cells that the
    /// formula-rewrite touched, so they can be restored to their
    /// pre-rewrite source.
    RangeNameUndefine {
        name: String,
        range: Range,
        note: Option<String>,
        cell_writes: Vec<(Address, Option<CellContents>)>,
    },
    /// Restore a single named-range note after Create or Delete.
    /// `prev = None` means the name had no note before the command.
    RangeNameNote { name: String, prev: Option<String> },
    /// Restore every named-range note after `/Range Name Note Reset`.
    RangeNameNoteReset { prev: Vec<(String, String)> },
    /// Restore the per-cell `cell_unprotected` set after `/Range Prot`
    /// or `/Range Unprot`. Each `(addr, was_unprotected)` pair records
    /// whether the cell was in the unprotected set before the command.
    RangeProtection { entries: Vec<(Address, bool)> },
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

#[derive(Debug, Clone)]
struct EraseConfirmState {
    path: PathBuf,
    /// 0=No, 1=Yes — matches `FILE_ERASE_CONFIRM_ITEMS` below.
    highlight: usize,
}

/// Items shown on line 2 of the No/Yes confirm submenu invoked by
/// `/File Erase` after the user types a path. First letter is the
/// accelerator.
const FILE_ERASE_CONFIRM_ITEMS: &[(&str, &str)] = &[
    ("No", "Do not erase the file"),
    ("Yes", "Permanently delete the file from disk"),
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
    /// Set the workbook-wide default format to `Format { kind, decimals:
    /// <buffer> }`. Unlike `RangeFormat`, no POINT step follows — the
    /// global is a single-target setting.
    WorksheetGlobalFormat {
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
    /// After the user types a name, drop it from the engine and rewrite
    /// every formula referencing it to use the literal Excel-form range
    /// (preserving the cells' values).
    RangeNameUndefine,
    /// After the user types a name, ask for a single-line note to attach.
    RangeNameNoteCreate,
    /// After the user types a name, ask for the note text body.
    RangeNameNoteCreateBody,
    /// After the user types a name, drop just that name's note (leaves
    /// the name itself untouched).
    RangeNameNoteDelete,
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
    /// After the user types a filename, read the file as plain text and
    /// paint each line as a label down a single column starting at the
    /// pointer (no CSV semantics — the whole line, including embedded
    /// commas, becomes one apostrophe-prefixed label).
    FileImportTextFilename,
    /// After the user types a filename for `/File Erase`, open the
    /// No/Yes confirm submenu.  The Worksheet/Print/Graph/Other leaves
    /// all share this prompt — the kind only differs in the unimplemented
    /// directory filter, not in the deletion semantics.
    FileEraseFilename,
    /// First step of `/File Combine` — the user types a source filename.
    /// `entire` distinguishes the Entire-File branch (commit immediately
    /// applies the merge) from Named-Or-Specified-Range (commit stashes
    /// the path and opens the second range-string prompt).
    FileCombineFilename {
        kind: CombineKind,
        entire: bool,
    },
    /// Second step of `/File Combine … Named/Specified-Range`. The
    /// filename is already stashed in `pending_combine_path`; this
    /// prompt collects the source range string (`A1..C5`).
    FileCombineRange {
        kind: CombineKind,
    },
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
    /// After the user types an encoded-output destination path, start
    /// a [`PrintSession`] (Encoded variant) and descend into the
    /// shared `/PF` submenu.
    PrintEncodedFilename,
    /// After the user types a header or footer string, store it on
    /// the active [`PrintSession`] and re-enter the Options submenu.
    PrintFileHeader,
    PrintFileFooter,
    /// After the user types a setup/escape string, store it on the
    /// active [`PrintSession`] and re-enter the Options submenu.
    PrintFileSetup,
    /// Numeric margin prompts (0..=1000). Each stores onto the
    /// active [`PrintSession`] and re-enters the Margins submenu.
    PrintFileMarginLeft,
    PrintFileMarginRight,
    PrintFileMarginTop,
    PrintFileMarginBottom,
    /// Numeric page-length prompt (0..=1000). 0 means no pagination.
    PrintFilePgLength,
    /// After the user types a CUPS queue name, store it on the active
    /// [`PrintSession`] and re-enter the Advanced submenu.
    PrintSessionOptionsAdvancedDevice,
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
    /// F5 GOTO: after the user types a cell address, move the pointer
    /// there. Silent no-op on parse failure (matches 1-2-3's "Esc back
    /// to READY" feel for an unrecognized address).
    Goto,
    /// `/Worksheet Global Default Other International Currency
    /// Prefix|Suffix` — after the user types the symbol string, store
    /// it on `International.currency` along with the chosen position.
    WorksheetGlobalDefaultOtherIntlCurrencySymbol {
        position: CurrencyPosition,
    },
    WgdDir,
    WgdTemp,
    WgdExtSave,
    WgdExtList,
    WgdPrinterInterface,
    WgdPrinterMarginLeft,
    WgdPrinterMarginRight,
    WgdPrinterMarginTop,
    WgdPrinterMarginBottom,
    WgdPrinterPgLength,
    WgdPrinterSetup,
    WgdPrinterName,
    /// Active macro is in `{GETLABEL}` / `{GETNUMBER}`. The dest
    /// cell lives on `App::pending_macro_input_loc` (PromptNext is
    /// `Copy` so it can't carry a `String`).
    MacroGetInput {
        numeric: bool,
    },
}

/// `/Worksheet Titles` axis selector.  Both freezes the rows above
/// and the columns left of the cell pointer; Horizontal freezes only
/// the rows above; Vertical freezes only the columns left.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum TitlesKind {
    Both,
    Horizontal,
    Vertical,
}

/// /File Xtract sub-command: does the extracted file keep formulas,
/// or is each cell written as its current cached value?
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum XtractKind {
    Formulas,
    Values,
}

/// `/File Combine` operation: how each source cell merges into the
/// matching target.  Copy overwrites; Add adds the source numerically;
/// Subtract subtracts.  Add/Subtract skip non-numeric source or target
/// cells (1-2-3 R3 semantics).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CombineKind {
    Copy,
    Add,
    Subtract,
}

/// /File List sub-command: which set of files is in the overlay?
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FileListKind {
    /// xlsx files in the current session directory.
    Worksheet,
    /// Currently-loaded active files (single-file workbook today).
    Active,
    /// Every regular file in the session directory, regardless of
    /// extension. Enter on a spreadsheet extension (xlsx, csv, and
    /// wk3 with `--features wk3`) retrieves the file; on anything
    /// else it just dismisses the overlay.
    Other,
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

/// Where F3 was pressed — determines what Enter on the name picker does.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NameListOrigin {
    /// F3 in POINT: Enter commits the picked range to the pending command.
    Point,
    /// F3 in the F5 GOTO prompt: Enter moves the pointer to the range's
    /// start corner and exits to READY.
    Goto,
    /// F3 in a name-typing prompt (e.g. `/Range Name Delete`): Enter
    /// fills the prompt buffer with the chosen name and returns to the
    /// underlying prompt.
    PromptName,
    /// Alt-F3 RUN from READY: Enter executes the macro stored at
    /// the picked range's start cell.
    RunMacro,
}

#[derive(Debug, Clone)]
pub(crate) struct NameListState {
    /// (name, range) sorted ascending by lowercased name.
    entries: Vec<(String, Range)>,
    highlight: usize,
    view_offset: usize,
    origin: NameListOrigin,
}

const NAME_LIST_PAGE_SIZE: usize = 10;

impl PromptNext {
    fn accepts_char(self, c: char) -> bool {
        match self {
            PromptNext::RangeFormat { .. }
            | PromptNext::WorksheetGlobalFormat { .. }
            | PromptNext::WorksheetColumnSetWidth
            | PromptNext::WorksheetColumnRangeSetWidth
            | PromptNext::WorksheetGlobalColWidth
            | PromptNext::WorksheetGlobalRecalcIteration => c.is_ascii_digit(),
            // 1-2-3 names accept letters, digits, `_`, `.`, and the
            // backslash that prefixes macro autonames (`\A`..`\Z`,
            // `\0`). 15-char max is enforced at commit time.
            PromptNext::RangeNameCreate
            | PromptNext::RangeNameDelete
            | PromptNext::RangeNameUndefine
            | PromptNext::RangeNameNoteCreate
            | PromptNext::RangeNameNoteDelete => {
                c.is_ascii_alphanumeric() || c == '_' || c == '\\' || c == '.'
            }
            // Note body is free text; allow anything printable.
            PromptNext::RangeNameNoteCreateBody => !c.is_control(),
            // GOTO accepts cell-address chars: letters (col), digits
            // (row), and `:` for the optional sheet prefix (`A:B5`).
            PromptNext::Goto => c.is_ascii_alphanumeric() || c == ':',
            PromptNext::FileSaveFilename
            | PromptNext::FileRetrieveFilename
            | PromptNext::FileXtractFilename { .. }
            | PromptNext::FileImportNumbersFilename
            | PromptNext::FileImportTextFilename
            | PromptNext::FileEraseFilename
            | PromptNext::FileCombineFilename { .. }
            | PromptNext::FileDirPath
            | PromptNext::FileOpenFilename { .. }
            | PromptNext::PrintFileFilename
            | PromptNext::PrintEncodedFilename
            | PromptNext::GraphSaveFilename => is_path_char(c),
            PromptNext::FileCombineRange { .. } => {
                c.is_ascii_alphanumeric() || c == ':' || c == '.' || c == '$'
            }
            // Header and footer are free-form text with the `|`
            // separator carving them into L|C|R.
            PromptNext::PrintFileHeader
            | PromptNext::PrintFileFooter
            | PromptNext::PrintFileSetup => c != '\n' && c != '\t',
            // CUPS queue names are conventionally alphanumeric with
            // `_`/`-`; reject whitespace so a stray space doesn't end
            // up as part of the `lp -d` argument.
            PromptNext::PrintSessionOptionsAdvancedDevice => {
                c.is_ascii_alphanumeric() || c == '_' || c == '-'
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
            // Macro input is free-form — labels accept anything;
            // numbers accept what `parse_typed_value` would (digits,
            // dot, comma, sign, etc.). For simplicity we accept all
            // non-control chars and let the commit handler reject a
            // bad number value.
            PromptNext::MacroGetInput { .. } => c != '\n' && c != '\t',
            // Currency symbols are short printable strings — accept
            // any printable graphic plus space. No tab/newline.
            PromptNext::WorksheetGlobalDefaultOtherIntlCurrencySymbol { .. } => {
                c != '\n' && c != '\t'
            }
            PromptNext::WgdDir | PromptNext::WgdTemp => is_path_char(c),
            PromptNext::WgdExtSave | PromptNext::WgdExtList => {
                c.is_ascii_alphanumeric() || c == '.'
            }
            PromptNext::WgdPrinterInterface
            | PromptNext::WgdPrinterMarginLeft
            | PromptNext::WgdPrinterMarginRight
            | PromptNext::WgdPrinterMarginTop
            | PromptNext::WgdPrinterMarginBottom
            | PromptNext::WgdPrinterPgLength => c.is_ascii_digit(),
            PromptNext::WgdPrinterSetup => c != '\n' && c != '\t',
            PromptNext::WgdPrinterName => c.is_ascii_alphanumeric() || c == '_' || c == '-',
        }
    }
}

fn parse_margin(buffer: &str, prev: u16) -> u16 {
    buffer.parse::<u16>().unwrap_or(prev).min(1000)
}

/// Characters accepted inside a filename/path prompt. Deliberately
/// narrower than 1-2-3's "anything goes" — we exclude keys with menu
/// semantics (`/`, period-free submenus). `/` is fine; `.` is fine.
fn is_path_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || matches!(c, '.' | '-' | '_' | '/' | '\\' | ' ' | '~')
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

/// `/System` — leave the alt screen + raw mode, run an interactive
/// shell, and on its exit restore the TUI. Mirrors the original 1-2-3
/// R3.4a behavior of suspending to a DOS shell ("Type EXIT to return
/// to 1-2-3"). Errors from the spawn are printed to the underlying
/// terminal and otherwise swallowed; we always try to restore the TUI.
fn suspend_to_shell<B: ratatui::backend::Backend>(terminal: &mut Terminal<B>) -> anyhow::Result<()>
where
    B::Error: Send + Sync + 'static,
{
    let mut stdout = io::stdout();
    disable_raw_mode()?;
    execute!(stdout, LeaveAlternateScreen, DisableMouseCapture)?;
    terminal.show_cursor()?;

    println!();
    println!("(Type 'exit' to return to 1-2-3.)");

    #[cfg(windows)]
    let shell = std::env::var("COMSPEC").unwrap_or_else(|_| "cmd.exe".to_string());
    #[cfg(not(windows))]
    let shell = std::env::var("SHELL").unwrap_or_else(|_| "/bin/sh".to_string());

    if let Err(e) = std::process::Command::new(&shell).status() {
        eprintln!("l123: /System: failed to launch {shell}: {e}");
    }

    enable_raw_mode()?;
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    terminal.clear()?;
    Ok(())
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

/// Transient state while the user is selecting a cell/range in POINT mode.
#[derive(Debug, Clone)]
struct PointState {
    /// Anchor corner. `None` after a single Esc — in that state the
    /// pointer moves freely without growing a range; a second Esc cancels.
    anchor: Option<Address>,
    /// Which command initiated POINT, so that `Enter` routes the selected
    /// range back to the right handler.
    pending: PendingCommand,
    /// Lotus-style typed range buffer (e.g. `c8..d12`). Empty in the
    /// usual highlight-by-arrows flow. When non-empty, line 3 shows the
    /// buffer in place of the auto-derived highlight, and `Enter` parses
    /// it (via [`Range::parse_with_default_sheet`]) to override the
    /// committed range.
    typed: String,
}

/// Which channel `:Format Color` is touching: cell background fill,
/// font foreground color, or both (Reset).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ColorTarget {
    Background,
    Text,
    Both,
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
    /// `:Format Alignment Left|Right|Center|General`. `General` is
    /// represented as `HAlign::General` and clears any per-cell
    /// override; other variants overwrite the horizontal alignment
    /// while preserving vertical and wrap.
    RangeAlignment {
        halign: HAlign,
    },
    /// `:Format Color Background|Text <color>` and `:Format Color
    /// Reset`. `target` selects which channel(s) to touch; `color`
    /// is `None` for a Reset-style clear.
    RangeColor {
        target: ColorTarget,
        color: Option<RgbColor>,
    },
    /// `pending_name` on App carries the name; on commit, define it over
    /// the selected range.
    RangeNameCreate,
    /// POINT step of `/Range Name Labels <Direction>`. On commit, walk
    /// every label cell in the selected range and define a 1-cell
    /// range name (the label's text) pointing at the adjacent cell in
    /// `direction`.
    RangeNameLabels {
        direction: LabelDirection,
    },
    /// POINT step of `/Range Name Table`. On commit, dump the active
    /// file's named-range table into a 2-column block anchored at the
    /// selected cell.
    RangeNameTable,
    /// POINT step of `/Range Name Note Table`. On commit, dump the
    /// names-with-notes table into a 3-column block.
    RangeNameNoteTable,
    /// POINT step of `/Range Prot` (`unprotected = false`) or
    /// `/Range Unprot` (`unprotected = true`). On commit, the
    /// per-cell `cell_unprotected` flag is updated for every cell
    /// in the selected range.
    RangeProtect {
        unprotected: bool,
    },
    /// POINT step of `/Range Input`. On commit, enter Input mode
    /// constrained to unprotected cells in the selected range.
    RangeInput,
    /// POINT step of `/Range Value`: pick the source range to copy
    /// (formulas → values).
    RangeValueFrom,
    /// POINT step of `/Range Value`: pick the destination anchor.
    RangeValueTo {
        src: Range,
    },
    /// POINT step of `/Range Trans`: pick the source range.
    RangeTransFrom,
    /// POINT step of `/Range Trans`: pick the destination anchor.
    RangeTransTo {
        src: Range,
    },
    /// POINT step of `/Range Justify`: pick the column block to
    /// reflow. Width is derived from the first cell's column width;
    /// height grows downward as needed (within the block).
    RangeJustify,
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
    /// Free-form selection started by a mouse drag from READY. There
    /// is no command to commit; Enter just returns to READY, leaving
    /// the highlight available for follow-up actions (e.g. SmartIcons
    /// Bold) before they happen.
    MouseSelect,
    /// POINT step of `/Worksheet Learn Range`. On commit, store the
    /// selected range as the destination for Alt-F5 LEARN
    /// recordings.
    WorksheetLearnRange,
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
            PendingCommand::RangeAlignment { .. } => "Enter range for alignment:",
            PendingCommand::RangeColor { .. } => "Enter range for color:",
            PendingCommand::RangeNameCreate => "Enter range for the named range:",
            PendingCommand::RangeNameLabels { .. } => "Enter range of labels:",
            PendingCommand::RangeNameTable => "Enter cell to write table to:",
            PendingCommand::RangeNameNoteTable => "Enter cell to write notes table to:",
            PendingCommand::RangeProtect { unprotected: true } => "Enter range to UNPROTECT:",
            PendingCommand::RangeProtect { unprotected: false } => "Enter range to RE-PROTECT:",
            PendingCommand::RangeInput => "Enter input range:",
            PendingCommand::RangeValueFrom => "Enter range to copy AS VALUES FROM:",
            PendingCommand::RangeValueTo { .. } => "Enter range to copy TO:",
            PendingCommand::RangeTransFrom => "Enter range to TRANSPOSE FROM:",
            PendingCommand::RangeTransTo { .. } => "Enter range to TRANSPOSE TO:",
            PendingCommand::RangeJustify => "Enter range to justify:",
            PendingCommand::FileXtractRange { .. } => "Enter range to extract:",
            PendingCommand::PrintFileRange => "Enter range to print:",
            PendingCommand::RangeSearchRange { .. } => "Enter search range:",
            PendingCommand::GraphSeries { .. } => "Enter graph range:",
            PendingCommand::ColumnRangeSetWidth { .. } => "Enter range of columns to set:",
            PendingCommand::ColumnRangeResetWidth => "Enter range of columns to reset:",
            PendingCommand::ColumnHide => "Enter range of columns to hide:",
            PendingCommand::ColumnDisplay => "Enter range of columns to display:",
            // Free mouse-drag selection has no prompt — line 3 keeps
            // showing the live range, but no command label is shown.
            PendingCommand::MouseSelect => "",
            PendingCommand::WorksheetLearnRange => "Enter range to record into:",
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
            input_range: None,
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
            if self.pending_system_suspend {
                self.pending_system_suspend = false;
                suspend_to_shell(terminal)?;
                continue;
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

    /// Find the buffer y coordinate for a given grid row, honoring
    /// frozen rows + the current row scroll.  Returns `None` when the
    /// row is outside the visible body region.
    fn cell_y_in_buffer(&self, buf: &Buffer, row: u32) -> Option<u16> {
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

    // ---------------- key handling ----------------

    pub fn handle_key(&mut self, k: KeyEvent) {
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
    fn run_named_macro(&mut self, name: &str) -> bool {
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
    fn read_macro_line(&self, addr: Address) -> Option<String> {
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
    fn resolve_macro_loc(&self, loc: &str) -> Option<Address> {
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
    fn eval_macro_condition(&mut self, expr: &str) -> bool {
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
    fn run_macro_at(&mut self, start: Address) {
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
    fn pump_macro(&mut self) {
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
    fn step_macro(&mut self) -> bool {
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
    fn open_custom_menu(&mut self, loc: &str, is_call: bool) {
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
    fn finish_custom_menu(&mut self, picked: Option<usize>) {
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
    fn cancel_learn(&mut self) {
        self.learn_range = None;
        self.learn_recording = false;
        self.learn_buffer.clear();
        self.close_menu();
    }

    /// `/Worksheet Learn Erase` — blank every cell in the learn
    /// range without dropping the range definition.
    fn erase_learn_range(&mut self) {
        if let Some(r) = self.learn_range {
            self.execute_range_erase(r);
            self.wb_mut().dirty = true;
        }
        self.close_menu();
    }

    /// Alt-F5 LEARN toggle. Off→On arms recording; On→Off flushes
    /// the buffered macro source to cells of the learn range.
    fn toggle_learn_recording(&mut self) {
        if self.learn_range.is_none() {
            // No range set — Lotus would beep with "no learn range
            // defined". Silent no-op for now.
            return;
        }
        if self.learn_recording {
            self.learn_recording = false;
            self.flush_learn_buffer();
        } else {
            self.learn_buffer.clear();
            self.learn_recording = true;
        }
    }

    /// Write `learn_buffer` into the cells of `learn_range`,
    /// splitting at 240 chars (the 1-2-3 label-cell capacity) so
    /// the recording wraps to subsequent rows.
    fn flush_learn_buffer(&mut self) {
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
    fn record_keystroke(&mut self, k: &KeyEvent) {
        if !self.learn_recording {
            return;
        }
        if let Some(token) = key_event_to_macro_source(k) {
            self.learn_buffer.push_str(&token);
        }
    }

    /// Custom-menu key handling: single-letter accelerators (item's
    /// first char, case-insensitive), Enter on highlight, Esc to
    /// cancel, arrows to move the highlight.
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

    /// Open a prompt in service of `{GETLABEL}` / `{GETNUMBER}` and
    /// park the macro until the user commits or cancels.
    fn start_macro_get_input(&mut self, prompt_text: String, loc: String, numeric: bool) {
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
    fn execute_macro_let(&mut self, loc: &str, expr: &str) {
        let Some(addr) = self.resolve_macro_loc(loc) else {
            self.set_error(format!("macro: {{LET}} bad loc `{loc}`"));
            return;
        };
        let intl = self.wb().international.clone();
        let (contents, format) =
            CellContents::from_source_with_format(expr, self.default_label_prefix, &intl);
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
    fn execute_macro_blank(&mut self, range_arg: &str) {
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
    fn parse_macro_range(&self, arg: &str) -> Option<Range> {
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
            self.wb_mut().cell_formats.insert(p, fmt);
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
            Action::WorksheetPage => self.insert_page_break_at_pointer(),
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
            Action::RangeFormatHidden => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::Hidden,
                    decimals: 0,
                },
            }),
            Action::RangeFormatTimeHmsAmPm => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::TimeHmsAmPm,
                    decimals: 0,
                },
            }),
            Action::RangeFormatTimeHmAmPm => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::TimeHmAmPm,
                    decimals: 0,
                },
            }),
            Action::RangeFormatTimeLongIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::TimeLongIntl,
                    decimals: 0,
                },
            }),
            Action::RangeFormatTimeShortIntl => self.begin_point(PendingCommand::RangeFormat {
                format: Format {
                    kind: FormatKind::TimeShortIntl,
                    decimals: 0,
                },
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
            Action::WorksheetGlobalFormatReset => self.set_global_format(Format::GENERAL),
            Action::WorksheetGlobalFormatText => self.set_global_format(Format {
                kind: FormatKind::Text,
                decimals: 0,
            }),
            Action::WorksheetGlobalFormatDateDmy => self.set_global_format(Format {
                kind: FormatKind::DateDmy,
                decimals: 0,
            }),
            Action::WorksheetGlobalFormatDateDm => self.set_global_format(Format {
                kind: FormatKind::DateDm,
                decimals: 0,
            }),
            Action::WorksheetGlobalFormatDateMy => self.set_global_format(Format {
                kind: FormatKind::DateMy,
                decimals: 0,
            }),
            Action::WorksheetGlobalFormatDateLongIntl => self.set_global_format(Format {
                kind: FormatKind::DateLongIntl,
                decimals: 0,
            }),
            Action::WorksheetGlobalFormatDateShortIntl => self.set_global_format(Format {
                kind: FormatKind::DateShortIntl,
                decimals: 0,
            }),
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

    fn insert_page_break_at_pointer(&mut self) {
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
        self.wb_mut().cell_formats.clear();
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
            global_format: Format::GENERAL,
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

    /// Read `path` as plain text; each line becomes an apostrophe-prefixed
    /// label down a single column starting at the pointer.  Counterpart to
    /// [`Self::import_numbers_from`]: no field splitting, no number coercion,
    /// embedded commas stay in the line. Empty lines are skipped (no
    /// overwrite of the existing target cell).
    fn import_text_from(&mut self, path: PathBuf) {
        let body = match std::fs::read_to_string(&path) {
            Ok(b) => b,
            Err(e) => {
                self.set_error(format!("Cannot read {}: {e}", path.display()));
                return;
            }
        };
        let origin = self.wb_mut().pointer;
        for (dr, line) in body.lines().enumerate() {
            if line.is_empty() {
                continue;
            }
            let addr = Address::new(origin.sheet, origin.col, origin.row + dr as u32);
            let engine_input = format!("'{line}");
            let _ = self.wb_mut().engine.set_user_input(addr, &engine_input);
            self.wb_mut().cells.insert(
                addr,
                CellContents::Label {
                    prefix: LabelPrefix::Apostrophe,
                    text: line.to_string(),
                },
            );
        }
        self.wb_mut().engine.recalc();
        self.refresh_formula_caches();
        self.mode = Mode::Ready;
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
        self.wb_mut().dirty = false;
        self.mode = Mode::Ready;
    }

    /// Load an xlsx from disk, wiping the current in-memory workbook
    /// and repopulating the UI cache from the loaded engine model.
    fn load_workbook_from(&mut self, path: PathBuf) {
        #[cfg(feature = "wk3")]
        let is_wk3 = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|s| s.eq_ignore_ascii_case("wk3"))
            .unwrap_or(false);
        #[cfg(not(feature = "wk3"))]
        let is_wk3 = false;
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
        // Wipe UI state; the loaded engine is the new source of truth.
        self.wb_mut().cells.clear();
        self.wb_mut().cell_formats.clear();
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
        if self.wb_mut().engine.save_xlsx(&path).is_ok() {
            // Embed the user-typed Lotus source per formula cell as
            // a sidecar inside the xlsx zip so save → reload
            // preserves shapes the cosmetic reverse translator
            // can't recover (arg-fix wrappers, emulated functions
            // like @CTERM, 3D-range expansions). A failure here is
            // best-effort — the xlsx itself is already saved.
            let sources: HashMap<Address, String> = self
                .wb()
                .cells
                .iter()
                .filter_map(|(addr, c)| match c {
                    CellContents::Formula { expr, .. } => Some((*addr, expr.clone())),
                    _ => None,
                })
                .collect();
            let _ = l123_io::formula_sources::write_to_xlsx(&path, &sources);
            self.wb_mut().active_path = Some(path);
            self.wb_mut().dirty = false;
        }
    }

    /// Handle a keystroke while the Cancel/Replace/Backup confirm is up.
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
                let format = Format { kind, decimals };
                self.begin_point(PendingCommand::RangeFormat { format });
            }
            PromptNext::WorksheetGlobalFormat { kind } => {
                let decimals: u8 = p.buffer.parse().unwrap_or(2);
                let decimals = decimals.min(15);
                self.set_global_format(Format { kind, decimals });
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
                    self.save_workbook_to(path);
                    self.mode = Mode::Ready;
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
            PromptNext::FileImportTextFilename => {
                if p.buffer.is_empty() {
                    self.mode = Mode::Ready;
                    return;
                }
                let path = PathBuf::from(&p.buffer);
                self.import_text_from(path);
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
            PendingCommand::RangeValueTo { src } | PendingCommand::RangeTransTo { src } => {
                src.start
            }
            _ => self.wb_mut().pointer,
        };
        self.wb_mut().pointer = source_tl;
        self.scroll_into_view();
        // Copy/Move TO start with no anchor so the user's pointer
        // movement defaults to a single-cell destination (matches the
        // M3 single-anchor flow). Pressing `.` anchors a multi-cell TO,
        // which is what the Lotus tutorial "single source → fill multi
        // destination" replicate flow needs.
        let anchor = match next {
            PendingCommand::CopyTo { .. }
            | PendingCommand::MoveTo { .. }
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
                let excel = l123_parse::to_engine_source_with_config(expr, &names_ref, &cfg);
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
                let excel = l123_parse::to_engine_source_with_config(expr, &names_ref, &cfg);
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
        if self.help.is_some() {
            self.render_help_overlay(chunks[1], buf);
        } else if self.name_list.is_some() {
            self.render_name_list_overlay(chunks[1], buf);
        } else if self.file_list.is_some() {
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

    fn render_icon_panel(&self, area: Rect, buf: &mut Buffer) {
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
    fn fit_rendered_px_h(area: Rect, font: (u16, u16)) -> u32 {
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

    /// Mouse handler. Icon-panel clicks dispatch the slot's action;
    /// grid clicks in READY move the pointer to the clicked cell.
    /// Move events update `hovered_icon` so control-panel line 3 can
    /// show the authentic R3.4a icon description.
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
    fn move_pointer_to(&mut self, addr: Address) {
        self.wb_mut().pointer = addr;
        self.scroll_into_view();
    }

    /// Insert `addr` (short form, e.g. `C3`) at the entry buffer
    /// cursor — but only when the position before the cursor is one
    /// where a cell reference is grammatically valid (start of buffer,
    /// or after an operator/paren/comma/dot/comparison/logical token).
    /// Called from mouse click-in-grid during VALUE/EDIT, so the user
    /// can build formulas by clicking instead of typing references.
    fn splice_cell_ref_into_entry(&mut self, addr: Address) {
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
    fn hit_test_slot(geom: &IconPanelGeom, row: u16) -> Option<usize> {
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
    fn dispatch_icon_click(&mut self, geom: &IconPanelGeom, column: u16, row: u16) {
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

    /// SmartIcons toggle for Bold / Italic / Underline. Acts on the
    /// active POINT highlight if one is in progress, otherwise on the
    /// single cell at the pointer. If every cell in the target already
    /// carries `bits`, clear them; otherwise set on all. Always lands
    /// back in Ready.
    fn dispatch_icon_text_style(&mut self, bits: TextStyle) {
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
        // display). Mid band: International box. Lower band: the
        // environment readout. Heights chosen so the whole overlay
        // fits inside an 80x30 terminal (the standard transcript size).
        let band = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(6),
                Constraint::Length(1),
                Constraint::Length(7),
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

        let intl_box = Block::default()
            .borders(Borders::ALL)
            .border_style(text_style)
            .title("International")
            .style(text_style);
        let intl_inner = intl_box.inner(band[2]);
        intl_box.render(band[2], buf);
        let intl = &self.wb().international;
        let currency_display = match intl.currency.position {
            CurrencyPosition::Prefix => format!("{}n", intl.currency.symbol),
            CurrencyPosition::Suffix => format!("n{}", intl.currency.symbol),
        };
        Paragraph::new(vec![
            Line::from(format!(
                "Punctuation: {} (decimal {} arg {} thousands {})",
                intl.punctuation.label(),
                intl.punctuation.decimal_char(),
                intl.punctuation.argument_sep(),
                if intl.punctuation.thousands_sep() == ' ' {
                    "(space)".to_string()
                } else {
                    intl.punctuation.thousands_sep().to_string()
                },
            )),
            Line::from(format!(
                "Date:        {} ({} long, {} short)",
                intl.date_intl.label(),
                intl.date_intl.long_label(),
                intl.date_intl.short_label(),
            )),
            Line::from(format!(
                "Time:        {} ({} long, {} short)",
                intl.time_intl.label(),
                intl.time_intl.long_label(),
                intl.time_intl.short_label(),
            )),
            Line::from(format!("Negative:    {}", intl.negative_style.label())),
            Line::from(format!(
                "Currency:    {} ({})",
                currency_display,
                intl.currency.position.label()
            )),
        ])
        .style(text_style)
        .render(intl_inner, buf);

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
        .render(band[3], buf);
    }

    fn render_defaults_overlay(&self, area: Rect, buf: &mut Buffer) {
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

        // Line 2 & 3 depend on mode. F3 NAMES, file-list and
        // save-confirm take absolute precedence (they own the
        // keyboard); then a command-argument prompt; then
        // mode-specific rendering.
        let (line2, line3) = if self.name_list.is_some() {
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

    fn render_file_list_lines(&self) -> (Line<'_>, Line<'_>) {
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

    fn render_name_list_lines(&self) -> (Line<'_>, Line<'_>) {
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
    fn render_name_list_overlay(&self, area: Rect, buf: &mut Buffer) {
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

    /// Draw the F1 help overlay: header bar with the page title, body
    /// (with hyperlinks colorized — focused link reversed), footer with
    /// key hints. Body rows are clipped to the available height; the
    /// renderer auto-scrolls so the focused link stays visible.
    fn render_help_overlay(&self, area: Rect, buf: &mut Buffer) {
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

    fn render_erase_confirm_lines(&self) -> (Line<'_>, Line<'_>) {
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
    fn render_custom_menu_lines(&self) -> (Line<'_>, Line<'_>) {
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

    /// Resolve the format for a given cell — the per-cell override if
    /// set, else the workbook's global default ([`Workbook::global_format`]).
    fn format_for_cell(&self, addr: Address) -> Format {
        self.wb().format_for_cell(addr)
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
    fn visible_row_layout(&self, body_rows: u16) -> Vec<(u32, u16)> {
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

        // Body rows: frozen rows pinned at the top, then scrolling
        // rows starting at the viewport offset (clamped to skip past
        // the frozen prefix).
        let row_layout = self.visible_row_layout(visible_rows);
        for &(row_idx, r) in &row_layout {
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
                            let s =
                                render_own_width(other, w as usize, fmt, &self.wb().international);
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
                    Style::default().add_modifier(Modifier::REVERSED)
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
                    let resolved_fg = explicit_fg
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
                                let s = render_own_width(
                                    other,
                                    span_w as usize,
                                    fmt,
                                    &self.wb().international,
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
                    // always apply.  The pointer-on-anchor case keeps
                    // the wide span REVERSED across the whole merge.
                    let anchor_highlighted = highlight.contains(m.anchor);
                    let mut astyle = if anchor_highlighted {
                        Style::default().add_modifier(Modifier::REVERSED)
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

    fn render_status(&self, area: Rect, buf: &mut Buffer) {
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

impl Default for App {
    fn default() -> Self {
        Self::new()
    }
}

/// One frame of the macro call stack. The interpreter executes
/// actions out of `remaining` and refills it from the cell at `pc`
/// whenever the buffer empties — this lets a subroutine call resume
/// the caller mid-line on `{RETURN}`.
struct MacroFrame {
    pc: Address,
    remaining: VecDeque<MacroAction>,
}

impl MacroFrame {
    fn starting_at(addr: Address) -> Self {
        Self {
            pc: addr,
            remaining: VecDeque::new(),
        }
    }
}

/// Top-level macro execution state. Lives on [`App`] while a macro
/// is running. `suspend = Some(...)` parks the interpreter so the
/// user can interact with the workbook between actions; the
/// `handle_key` post-dispatch hook clears the parked reason and
/// resumes via `pump_macro`.
struct MacroState {
    frames: Vec<MacroFrame>,
    /// Total actions executed so far across all frames; safety guard
    /// against runaway loops (cap at `MAX_MACRO_STEPS`).
    steps: u32,
    /// Why the interpreter is parked, if it is. `None` when ready
    /// to advance.
    suspend: Option<MacroSuspend>,
    /// One-shot flag: when true, [`step_macro`] skips the STEP-mode
    /// pre-pause so a single action can fire. Set by the user
    /// pressing Space at a STEP pause, cleared once consumed.
    step_advance: bool,
}

/// Reasons a macro can pause mid-execution.
enum MacroSuspend {
    /// `{?}` — resume on the next Enter the user presses.
    WaitEnter,
    /// `{GETLABEL p, loc}` / `{GETNUMBER p, loc}` — a prompt is up.
    /// The dest cell and numeric flag live on `App` because
    /// `PromptNext` is `Copy` and can't carry an owned `String`.
    GetInput,
    /// `{MENUBRANCH loc}` / `{MENUCALL loc}` — a custom menu is up.
    /// On commit, BRANCH (or CALL) the macro's PC to the picked
    /// item's action cell.
    MenuPick,
    /// STEP mode is on: paused before the next action. Space
    /// advances one step; Esc aborts.
    StepPause,
}

/// State backing a `{MENUBRANCH}` / `{MENUCALL}` overlay.
struct CustomMenuState {
    /// Display name + description for each menu column (item).
    items: Vec<CustomMenuItem>,
    /// Address of the "row 2" (action) cell of the leftmost menu
    /// column. The action cell for item `i` is `(action_row.col +
    /// i, action_row.row)`.
    action_row: Address,
    /// True when this was opened via `{MENUCALL}` — the macro will
    /// CALL (push frame) instead of BRANCH (replace PC).
    is_call: bool,
    /// Currently highlighted item index.
    highlight: usize,
}

struct CustomMenuItem {
    name: String,
    description: String,
}

const MAX_MACRO_STEPS: u32 = 100_000;

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

/// After deleting the sheet at index `at`, drop every cache entry on
/// that sheet and shift entries on later sheets back by one slot. The
/// inverse of [`shift_sheets_from`]; covers the same caches.
fn drop_sheet_from_caches(
    cells: &mut HashMap<Address, CellContents>,
    cell_formats: &mut HashMap<Address, Format>,
    cell_text_styles: &mut HashMap<Address, TextStyle>,
    col_widths: &mut HashMap<(SheetId, u16), u8>,
    at: u16,
) {
    cells.retain(|a, _| a.sheet.0 != at);
    cell_formats.retain(|a, _| a.sheet.0 != at);
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
fn render_own_width(
    contents: &CellContents,
    width: usize,
    format: Format,
    intl: &International,
) -> String {
    match contents {
        CellContents::Empty => " ".repeat(width),
        CellContents::Label { .. } => {
            // Labels are routed through SpillSlot::Label; we should
            // never hit this branch from the planner caller.
            " ".repeat(width)
        }
        CellContents::Constant(v) => {
            render_value_in_cell(v, width, format, intl).unwrap_or_else(|| " ".repeat(width))
        }
        CellContents::Formula {
            cached_value: Some(v),
            ..
        } => render_value_in_cell(v, width, format, intl).unwrap_or_else(|| " ".repeat(width)),
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
    fn wgd_update_writes_cnf_block_and_preserves_other_lines() {
        let dir = std::env::temp_dir().join("l123_wgd_update_test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("L123.CNF");
        std::fs::write(&path, "user = \"Pre-existing\"\nlog_file = /tmp/x.log\n").unwrap();

        let d = GlobalDefaults {
            printer_interface: 5,
            printer_pg_length: 88,
            default_dir: "/tmp/sheets".into(),
            autoexec: false,
            graph_group: GraphGroupOrientation::Rowwise,
            graph_save: GraphSaveFormat::Pic,
            ..GlobalDefaults::default()
        };
        d.write_to_path(&path).unwrap();

        let body = std::fs::read_to_string(&path).unwrap();
        assert!(body.contains("user = \"Pre-existing\""), "{body}");
        assert!(body.contains("log_file = /tmp/x.log"), "{body}");
        assert!(body.contains("wgd_printer_interface = 5"), "{body}");
        assert!(body.contains("wgd_printer_pg_length = 88"), "{body}");
        assert!(body.contains("wgd_dir = \"/tmp/sheets\""), "{body}");
        assert!(body.contains("wgd_autoexec = false"), "{body}");
        assert!(body.contains("wgd_graph_group = rowwise"), "{body}");
        assert!(body.contains("wgd_graph_save = pic"), "{body}");

        // Re-running update should not duplicate the block.
        d.write_to_path(&path).unwrap();
        let body2 = std::fs::read_to_string(&path).unwrap();
        let count = body2.matches("wgd_printer_interface").count();
        assert_eq!(count, 1, "block duplicated on re-run:\n{body2}");
    }

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

    #[cfg(feature = "wk3")]
    #[test]
    fn new_with_file_routes_wk3_by_extension() {
        // Open a `.WK3` via the CLI entry point: it should land in
        // READY (not ERROR) with the workbook content visible, and
        // `active_path` swapped to "<original>.WK3.xlsx" so /File
        // Save defaults to writing xlsx alongside the legacy file.
        let dir = temp_test_dir("new_with_wk3");
        let src = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .parent()
            .and_then(|p| p.parent())
            .expect("workspace root above crates/l123-ui")
            .join("tests/acceptance/fixtures/wk3/FILE0001.WK3");
        if !src.exists() {
            eprintln!(
                "skipping new_with_file_routes_wk3_by_extension: {} missing",
                src.display()
            );
            let _ = std::fs::remove_dir(&dir);
            return;
        }
        let wk3_path = dir.join("legacy.WK3");
        std::fs::copy(&src, &wk3_path).unwrap();

        let app = App::new_with_file(wk3_path.clone());
        assert_eq!(
            app.mode,
            Mode::Ready,
            "CLI-opening a .WK3 should land in READY (error={:?})",
            app.error_message,
        );
        match app.wb().cells.get(&Address::A1) {
            Some(CellContents::Label { text, .. }) => assert_eq!(text, "Hello"),
            other => panic!("A1 expected Label(Hello), got {other:?}"),
        }
        let expected_save = wk3_path.with_file_name("legacy.WK3.xlsx");
        assert_eq!(
            app.wb().active_path.as_deref(),
            Some(expected_save.as_path()),
            "active_path should be original.WK3.xlsx so /File Save writes xlsx",
        );

        let _ = std::fs::remove_file(&wk3_path);
        let _ = std::fs::remove_dir(&dir);
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
    fn f5_goto_moves_pointer() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::F(5), KeyModifiers::NONE));
        for c in "C5".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer, Address::new(SheetId::A, 2, 4));
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn f5_goto_esc_leaves_pointer_unchanged() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::F(5), KeyModifiers::NONE));
        for c in "Z99".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.wb().pointer, Address::A1);
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn typed_point_range_commits() {
        let mut app = App::new();
        // /Range Erase enters POINT in one step.
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        for c in "B2..D4".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(app.wb().pointer, Address::A1);
    }

    #[test]
    fn typed_point_esc_clears_buffer_first() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        for c in "B2".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        assert!(app.point.as_ref().unwrap().typed.is_empty());
        // Two more Esc presses to fully cancel (un-anchor, then exit).
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
    }

    #[test]
    fn typed_point_bad_input_stays_in_point() {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        for c in "WAT".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        assert!(app.point.as_ref().unwrap().typed.is_empty());
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

    #[test]
    fn backslash_prefix_repeat_fills_when_xlsx_halign_is_left() {
        // Reproduces the post-import regression: an xlsx that came from
        // a Lotus-saved sheet stores `\-` as Label{Backslash, "-"} and
        // separately carries HAlign::Left (Excel's text default). The
        // grid renderer was letting Left override the stored Backslash
        // prefix, so the cell rendered as a single "-" left-padded
        // rather than as a span of dashes.
        let mut app = App::new();
        app.wb_mut().cells.insert(
            Address::A1,
            CellContents::Label {
                prefix: LabelPrefix::Backslash,
                text: "-".into(),
            },
        );
        app.wb_mut().cell_alignments.insert(
            Address::A1,
            Alignment {
                horizontal: HAlign::Left,
                ..Alignment::DEFAULT
            },
        );
        let buf = app.render_to_buffer(80, 25);
        let got = app
            .cell_rendered_text(&buf, "A:A1")
            .expect("A1 must be in viewport");
        assert_eq!(
            got, "---------",
            "Backslash prefix must override imported HAlign::Left, got {got:?}"
        );
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

    /// Worksheet listing is an alphabetical list of `.xlsx` files
    /// (plus `.WK3` when built with `--features wk3`) in the given
    /// directory. Pure function of the directory's contents so we can
    /// test it without touching process CWD.
    #[test]
    fn list_worksheet_files_in_returns_xlsx_sorted() {
        let dir = temp_test_dir("list_ws");
        let names_in = [
            "zeta.xlsx",
            "alpha.xlsx",
            "other.txt",
            "mid.XLSX",
            "legacy.WK3",
            "lower.wk3",
        ];
        for name in names_in {
            std::fs::write(dir.join(name), b"placeholder").unwrap();
        }
        let got = list_worksheet_files_in(&dir);
        let names: Vec<String> = got
            .iter()
            .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
            .collect();
        #[cfg(feature = "wk3")]
        let expected = vec![
            "alpha.xlsx",
            "legacy.WK3",
            "lower.wk3",
            "mid.XLSX",
            "zeta.xlsx",
        ];
        #[cfg(not(feature = "wk3"))]
        let expected = vec!["alpha.xlsx", "mid.XLSX", "zeta.xlsx"];
        assert_eq!(names, expected);
        for name in names_in {
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

    /// `list_all_files_in` includes every regular file (any extension)
    /// and skips dotfiles, sorted by filename.
    #[test]
    fn list_all_files_in_returns_every_file_no_dotfiles() {
        let dir = temp_test_dir("list_other");
        let names_in = [
            "zeta.xlsx",
            "notes.txt",
            "data.csv",
            "alpha.bin",
            ".hidden",
            "README",
        ];
        for name in names_in {
            std::fs::write(dir.join(name), b"x").unwrap();
        }
        let got = list_all_files_in(&dir);
        let names: Vec<String> = got
            .iter()
            .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
            .collect();
        assert_eq!(
            names,
            vec!["README", "alpha.bin", "data.csv", "notes.txt", "zeta.xlsx"]
        );
        for name in names_in {
            let _ = std::fs::remove_file(dir.join(name));
        }
        let _ = std::fs::remove_dir(&dir);
    }

    /// /File List Other → Enter on a `.csv` file routes through the CSV
    /// loader: the workbook ends up populated with that file's rows.
    #[test]
    fn file_list_other_enter_loads_csv() {
        let dir = temp_test_dir("list_other_csv");
        let path = dir.join("data.csv");
        std::fs::write(&path, b"7,8,9\n").unwrap();

        let mut app = App::new();
        app.file_list = Some(FileListState {
            kind: FileListKind::Other,
            entries: vec![path.clone()],
            highlight: 0,
            view_offset: 0,
        });
        app.mode = Mode::Files;

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.file_list.is_none());
        match app.wb().cells.get(&Address::A1).unwrap() {
            CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 7.0),
            other => panic!("A1 expected Number(7) from data.csv, got {other:?}"),
        }

        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_dir(&dir);
    }

    /// /File List Other → Enter on a non-spreadsheet file just dismisses
    /// the overlay, leaving the workbook untouched.
    #[test]
    fn file_list_other_enter_on_unsupported_extension_dismisses() {
        let dir = temp_test_dir("list_other_dismiss");
        let path = dir.join("readme.txt");
        std::fs::write(&path, b"hello").unwrap();

        let mut app = App::new();
        app.wb_mut().active_path = Some(PathBuf::from("orig.xlsx"));
        app.file_list = Some(FileListState {
            kind: FileListKind::Other,
            entries: vec![path.clone()],
            highlight: 0,
            view_offset: 0,
        });
        app.mode = Mode::Files;

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        assert!(app.file_list.is_none());
        // Workbook untouched: active_path still points at the pre-existing
        // file, no cells were planted.
        assert_eq!(
            app.wb().active_path.as_deref(),
            Some(PathBuf::from("orig.xlsx").as_path())
        );
        assert!(app.wb().cells.is_empty());

        let _ = std::fs::remove_file(&path);
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

    /// /FR clears the dirty bit: the in-memory workbook now matches
    /// what's on disk, so a follow-up /Q should not warn.
    #[test]
    fn file_retrieve_clears_dirty_bit() {
        let dir = temp_test_dir("file_retrieve_dirty");
        let target = dir.join("sheet.xlsx");

        let mut app = App::new();
        drive_save_keys(&mut app, "42", &target);
        assert!(target.exists());

        let mut app2 = App::new();
        for c in "hi".chars() {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert!(app2.is_dirty(), "label commit should mark dirty");

        for c in ['/', 'F', 'R'] {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in target.to_str().unwrap().chars() {
            app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

        assert!(!app2.is_dirty(), "successful /FR should clear dirty bit");

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

    fn drive_chord(app: &mut App, chord: &[char]) {
        for c in chord {
            app.handle_key(KeyEvent::new(KeyCode::Char(*c), KeyModifiers::NONE));
        }
    }

    fn drive_chars(app: &mut App, s: &str) {
        for c in s.chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
    }

    fn enter(app: &mut App) {
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    }

    /// Commit `1` at A1 and force the workbook back to a clean
    /// baseline. Used by dirty-bit tests for mutators that need an
    /// existing cell to operate on.
    fn seed_a1_clean(app: &mut App) {
        drive_chord(app, &['1']);
        enter(app);
        app.wb_mut().dirty = false;
    }

    #[test]
    fn range_erase_marks_dirty() {
        let mut app = App::new();
        seed_a1_clean(&mut app);
        // /RE auto-anchors POINT at the pointer; Enter erases the
        // single-cell range A1..A1.
        drive_chord(&mut app, &['/', 'R', 'E']);
        enter(&mut app);
        assert!(app.is_dirty(), "/RE should mark dirty");
    }

    #[test]
    fn copy_marks_dirty() {
        let mut app = App::new();
        seed_a1_clean(&mut app);
        // /C: FROM-Enter (A1..A1), RIGHT, TO-Enter (B1).
        drive_chord(&mut app, &['/', 'C']);
        enter(&mut app);
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        enter(&mut app);
        assert!(app.is_dirty(), "/C should mark dirty");
    }

    #[test]
    fn move_marks_dirty() {
        let mut app = App::new();
        seed_a1_clean(&mut app);
        drive_chord(&mut app, &['/', 'M']);
        enter(&mut app);
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        enter(&mut app);
        assert!(app.is_dirty(), "/M should mark dirty");
    }

    #[test]
    fn range_format_marks_dirty() {
        let mut app = App::new();
        seed_a1_clean(&mut app);
        // /RFG → General format → POINT → Enter applies to A1..A1.
        drive_chord(&mut app, &['/', 'R', 'F', 'G']);
        enter(&mut app);
        assert!(app.is_dirty(), "/RFG should mark dirty");
    }

    #[test]
    fn range_name_create_marks_dirty() {
        let mut app = App::new();
        // /RNC → name "sales" → Enter → POINT (A1..A1) → Enter.
        drive_chord(&mut app, &['/', 'R', 'N', 'C']);
        drive_chars(&mut app, "sales");
        enter(&mut app);
        enter(&mut app);
        assert!(app.is_dirty(), "/RNC should mark dirty");
    }

    #[test]
    fn range_name_delete_marks_dirty() {
        let mut app = App::new();
        // Create first.
        drive_chord(&mut app, &['/', 'R', 'N', 'C']);
        drive_chars(&mut app, "sales");
        enter(&mut app);
        enter(&mut app);
        app.wb_mut().dirty = false;
        // /RND → name "sales" → Enter.
        drive_chord(&mut app, &['/', 'R', 'N', 'D']);
        drive_chars(&mut app, "sales");
        enter(&mut app);
        assert!(app.is_dirty(), "/RND should mark dirty");
    }

    #[test]
    fn ws_column_set_width_marks_dirty() {
        let mut app = App::new();
        // /WCS prompts for width; type 15 + Enter.
        drive_chord(&mut app, &['/', 'W', 'C', 'S']);
        drive_chars(&mut app, "15");
        enter(&mut app);
        assert!(app.is_dirty(), "/WCS should mark dirty");
    }

    #[test]
    fn ws_column_reset_width_marks_dirty() {
        let mut app = App::new();
        drive_chord(&mut app, &['/', 'W', 'C', 'S']);
        drive_chars(&mut app, "15");
        enter(&mut app);
        app.wb_mut().dirty = false;
        // /WCR resets the current column to the default width.
        drive_chord(&mut app, &['/', 'W', 'C', 'R']);
        assert!(app.is_dirty(), "/WCR should mark dirty");
    }

    #[test]
    fn ws_column_range_set_width_marks_dirty() {
        let mut app = App::new();
        // /WCCS prompts for width; type 12 + Enter, then POINT, Enter applies.
        drive_chord(&mut app, &['/', 'W', 'C', 'C', 'S']);
        drive_chars(&mut app, "12");
        enter(&mut app);
        enter(&mut app);
        assert!(app.is_dirty(), "/WCCS should mark dirty");
    }

    #[test]
    fn ws_column_range_reset_width_marks_dirty() {
        let mut app = App::new();
        // First set a non-default width so the reset has something to do.
        drive_chord(&mut app, &['/', 'W', 'C', 'C', 'S']);
        drive_chars(&mut app, "12");
        enter(&mut app);
        enter(&mut app);
        app.wb_mut().dirty = false;
        // /WCCR is the column-range reset; POINT, Enter.
        drive_chord(&mut app, &['/', 'W', 'C', 'C', 'R']);
        enter(&mut app);
        assert!(app.is_dirty(), "/WCCR should mark dirty");
    }

    #[test]
    fn ws_column_hide_marks_dirty() {
        let mut app = App::new();
        drive_chord(&mut app, &['/', 'W', 'C', 'H']);
        enter(&mut app);
        assert!(app.is_dirty(), "/WCH should mark dirty");
    }

    #[test]
    fn wysiwyg_display_mode_color_paints_white_bg_on_empty_cell() {
        let mut app = App::new();
        drive_chord(&mut app, &[':', 'D', 'M', 'C']);
        assert_eq!(app.mode, Mode::Ready);
        let buf = app.render_to_buffer(80, 25);
        assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), Some((0xFF, 0xFF, 0xFF)));
        assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), Some((0x00, 0x00, 0x00)));
    }

    #[test]
    fn wysiwyg_display_mode_reverse_paints_black_bg() {
        let mut app = App::new();
        drive_chord(&mut app, &[':', 'D', 'M', 'R']);
        let buf = app.render_to_buffer(80, 25);
        assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), Some((0x00, 0x00, 0x00)));
        assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), Some((0xFF, 0xFF, 0xFF)));
    }

    #[test]
    fn wysiwyg_display_mode_bw_leaves_terminal_default() {
        let mut app = App::new();
        // Switch to Color, then back to B&W — B&W should clear the
        // RGB BG so the terminal default shows through (read-back
        // returns None for non-RGB cells).
        drive_chord(&mut app, &[':', 'D', 'M', 'C']);
        drive_chord(&mut app, &[':', 'D', 'M', 'B']);
        let buf = app.render_to_buffer(80, 25);
        assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), None);
        assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), None);
    }

    #[test]
    fn wysiwyg_display_grid_yes_paints_dotted_right_edges() {
        let mut app = App::new();
        // Default off — no gridline glyphs anywhere.
        let buf = app.render_to_buffer(80, 25);
        let body_row = App::line_text(&buf, PANEL_HEIGHT + 1);
        assert!(
            !body_row.contains('┊'),
            "default body row should not contain gridline dots: {body_row:?}"
        );

        // :DOGY turns gridlines on — the rightmost column of each
        // 9-char-wide empty cell becomes `┊`.
        drive_chord(&mut app, &[':', 'D', 'O', 'G', 'Y']);
        let buf = app.render_to_buffer(80, 25);
        // A body row past the pointer cell so no REVERSED highlight
        // suppresses the overlay.
        let body_row = App::line_text(&buf, PANEL_HEIGHT + 2);
        assert!(
            body_row.matches('┊').count() >= 4,
            "expected dotted gridline at each cell right edge: {body_row:?}"
        );

        // :DOGN turns them back off.
        drive_chord(&mut app, &[':', 'D', 'O', 'G', 'N']);
        let buf = app.render_to_buffer(80, 25);
        let body_row = App::line_text(&buf, PANEL_HEIGHT + 2);
        assert!(
            !body_row.contains('┊'),
            "grid=No should suppress gridline dots: {body_row:?}"
        );
    }

    #[test]
    fn wysiwyg_display_grid_skips_cells_with_full_width_content() {
        let mut app = App::new();
        // Type a 9-char value that fills the default column width
        // exactly. With gridlines on, the right-edge `┊` would
        // overwrite the last digit — verify we skip the overlay.
        drive_chord(&mut app, &['9', '8', '7', '6', '5', '4', '3', '2', '1']);
        enter(&mut app);
        drive_chord(&mut app, &[':', 'D', 'O', 'G', 'Y']);
        let buf = app.render_to_buffer(80, 25);
        // Move the pointer off A1 so the highlight doesn't suppress
        // anything for an unrelated reason; assert by reading the
        // actual cell content.
        let painted = (0..9)
            .map(|i| buf[(ROW_GUTTER + i, PANEL_HEIGHT + 1)].symbol().to_string())
            .collect::<String>();
        assert_eq!(painted, "987654321", "filled cell must not be clipped");
    }

    #[test]
    fn ws_column_display_marks_dirty() {
        let mut app = App::new();
        drive_chord(&mut app, &['/', 'W', 'C', 'H']);
        enter(&mut app);
        app.wb_mut().dirty = false;
        drive_chord(&mut app, &['/', 'W', 'C', 'D']);
        enter(&mut app);
        assert!(app.is_dirty(), "/WCD should mark dirty");
    }

    #[test]
    fn ws_titles_set_marks_dirty() {
        let mut app = App::new();
        // /WTH freezes the rows above the pointer.
        drive_chord(&mut app, &['/', 'W', 'T', 'H']);
        assert!(app.is_dirty(), "/WTH should mark dirty");
    }

    #[test]
    fn ws_titles_clear_marks_dirty() {
        let mut app = App::new();
        drive_chord(&mut app, &['/', 'W', 'T', 'H']);
        app.wb_mut().dirty = false;
        // /WTC on a sheet that has frozen titles should clear them.
        drive_chord(&mut app, &['/', 'W', 'T', 'C']);
        assert!(app.is_dirty(), "/WTC (with prior titles) should mark dirty");
    }

    #[test]
    fn ws_titles_clear_no_op_does_not_dirty() {
        // /WTC on a sheet without titles is a no-op — no journal, no
        // dirty flip. Matches the existing M5_ws_titles transcript's
        // "quiet no-op" expectation.
        let mut app = App::new();
        drive_chord(&mut app, &['/', 'W', 'T', 'C']);
        assert!(!app.is_dirty(), "/WTC on clean sheet should remain clean");
    }

    #[test]
    fn wg_global_format_marks_dirty() {
        let mut app = App::new();
        // /WGFG sets global format to General — routes through
        // set_global_format and journals a GlobalFormat entry.
        drive_chord(&mut app, &['/', 'W', 'G', 'F', 'G']);
        assert!(app.is_dirty(), "/WGFG should mark dirty");
    }

    #[test]
    fn wg_global_col_width_marks_dirty() {
        let mut app = App::new();
        // /WGCS prompts for the new default column width; type 12 + Enter.
        drive_chord(&mut app, &['/', 'W', 'G', 'C']);
        drive_chord(&mut app, &['S']);
        drive_chars(&mut app, "12");
        enter(&mut app);
        assert!(app.is_dirty(), "/WGCS should mark dirty");
    }

    #[test]
    fn wgd_intl_punctuation_marks_dirty() {
        let mut app = App::new();
        // /WGDOIPB switches Punctuation from the default (A) to B,
        // mutating Workbook.international.
        drive_chord(&mut app, &['/', 'W', 'G', 'D', 'O', 'I', 'P', 'B']);
        assert!(app.is_dirty(), "/WGDOIPB should mark dirty");
    }

    #[test]
    fn ws_erase_yields_clean_workbook() {
        // /WEY discards the workbook for a fresh blank one — there are
        // no changes left to save, so /Q should not warn afterwards.
        let mut app = App::new();
        drive_chars(&mut app, "hi");
        enter(&mut app);
        assert!(app.is_dirty());
        drive_chord(&mut app, &['/', 'W', 'E', 'Y']);
        assert!(!app.is_dirty(), "/WEY should leave the workbook clean");
    }

    #[test]
    fn wysiwyg_bold_marks_dirty() {
        let mut app = App::new();
        drive_chars(&mut app, "hi");
        enter(&mut app);
        app.wb_mut().dirty = false;
        // :FBS → POINT → Enter applies Bold to A1..A1.
        drive_chord(&mut app, &[':', 'F', 'B', 'S']);
        enter(&mut app);
        assert!(app.is_dirty(), ":FBS should mark dirty");
    }

    /// Structural mutators (insert/delete row & col) flip the dirty
    /// bit. Each runs against a fresh `App::new()` so the assertion
    /// isolates one mutator.
    #[test]
    fn structural_mutators_mark_dirty() {
        for chord in [
            &['/', 'W', 'I', 'R'][..],
            &['/', 'W', 'I', 'C'][..],
            &['/', 'W', 'D', 'R'][..],
            &['/', 'W', 'D', 'C'][..],
        ] {
            let mut app = App::new();
            assert!(!app.is_dirty(), "fresh App should be clean");
            drive_chord(&mut app, chord);
            let chord_str: String = chord.iter().collect();
            assert!(app.is_dirty(), "{chord_str} should mark dirty");
        }
    }

    /// Dirty-bit lifecycle: a fresh App is clean; committing a label
    /// flips it dirty; a successful `/FS` clears it. Drives the
    /// `/QY` warn-on-quit guard.
    #[test]
    fn dirty_bit_tracks_modifications_and_save() {
        let mut app = App::new();
        assert!(!app.is_dirty(), "fresh workbook should be clean");

        for c in "hi".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert!(app.is_dirty(), "label commit should mark dirty");

        let dir = temp_test_dir("dirty_bit");
        let target = dir.join("clean.xlsx");
        for c in ['/', 'F', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in target.to_str().unwrap().chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert!(!app.is_dirty(), "successful /FS should clear the dirty bit");

        let _ = std::fs::remove_file(&target);
        let _ = std::fs::remove_dir(&dir);
    }

    /// Simulate a left-click at `(col, row)` by stashing a fake panel
    /// geometry and routing through [`App::handle_mouse`].
    fn click(app: &mut App, area: Rect, col: u16, row: u16) {
        app.icon_panel_area.set(Some(test_geom(area)));
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    /// Standard fixture panel: 17 icons × 3 rows each = 51 rows, so
    /// slot N's middle row is `N * 3 + 1`.
    const TEST_PANEL: Rect = Rect {
        x: 80,
        y: 4,
        width: 3,
        height: 51,
    };

    /// Build a hit-test fixture where the rendered image is assumed
    /// to fill the cell rect exactly (no aspect-ratio slack). Using
    /// `font_px_h = 2` lets the cell-midpoint pixel formula resolve
    /// to integer slot boundaries on multiples of `2 * 17 = 34`.
    fn test_geom(rect: Rect) -> IconPanelGeom {
        IconPanelGeom {
            rect,
            rendered_px_h: rect.height as u32 * 2,
            font_px_h: 2,
        }
    }

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
    fn icon_click_bold_applies_to_pointer_cell_instantly() {
        // Panel 1 slot 11 = icon id 12 = SmartIcons Bold. The click
        // applies bold to the cursor cell immediately — no menu, no
        // POINT prompt, mode stays Ready.
        let mut app = App::new();
        click_slot(&mut app, 11);
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
        );
    }

    #[test]
    fn icon_click_italic_applies_to_pointer_cell_instantly() {
        let mut app = App::new();
        click_slot(&mut app, 12);
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn icon_click_underline_applies_to_pointer_cell_instantly() {
        let mut app = App::new();
        click_slot(&mut app, 13);
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::UNDERLINE),
        );
    }

    #[test]
    fn icon_click_bold_applies_to_active_point_highlight() {
        // While in POINT mode with an extended highlight (e.g. mid /Range
        // command), clicking Bold should apply to that highlight rather
        // than just A1, then return to Ready.
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        // Now in POINT for /Range Erase, anchored at A1. Extend to B2.
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Point);
        click_slot(&mut app, 11);
        assert_eq!(app.mode, Mode::Ready);
        for addr in [
            Address::A1,
            Address::new(SheetId(0), 1, 0),
            Address::new(SheetId(0), 0, 1),
            Address::new(SheetId(0), 1, 1),
        ] {
            assert_eq!(
                app.wb().cell_text_styles.get(&addr).copied(),
                Some(TextStyle::BOLD),
                "{addr:?} should be bold",
            );
        }
    }

    #[test]
    fn icon_click_bold_toggles_off_when_pointer_cell_already_bold() {
        // Second click on Bold over an already-bold cell removes bold.
        let mut app = App::new();
        click_slot(&mut app, 11);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::BOLD),
        );
        click_slot(&mut app, 11);
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(app.wb().cell_text_styles.get(&Address::A1).copied(), None);
    }

    #[test]
    fn icon_click_bold_toggle_only_clears_bold_keeping_other_styles() {
        // Toggling Bold off must leave Italic/Underline intact on the cell.
        let mut app = App::new();
        app.execute_range_text_style(
            Range::single(Address::A1),
            TextStyle::BOLD.merge(TextStyle::ITALIC),
            true,
        );
        click_slot(&mut app, 11);
        assert_eq!(
            app.wb().cell_text_styles.get(&Address::A1).copied(),
            Some(TextStyle::ITALIC),
        );
    }

    #[test]
    fn icon_click_bold_toggles_off_when_every_cell_in_highlight_is_bold() {
        // Pre-bold A1:B2, then enter POINT over A1:B2 and click Bold —
        // the whole range should clear.
        let mut app = App::new();
        let r = Range {
            start: Address::A1,
            end: Address::new(SheetId(0), 1, 1),
        };
        app.execute_range_text_style(r, TextStyle::BOLD, true);
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        click_slot(&mut app, 11);
        assert_eq!(app.mode, Mode::Ready);
        for addr in [
            Address::A1,
            Address::new(SheetId(0), 1, 0),
            Address::new(SheetId(0), 0, 1),
            Address::new(SheetId(0), 1, 1),
        ] {
            assert_eq!(
                app.wb().cell_text_styles.get(&addr).copied(),
                None,
                "{addr:?} should no longer be bold",
            );
        }
    }

    #[test]
    fn icon_click_bold_sets_bold_on_mixed_highlight() {
        // Mixed highlight (only A1 bold, others plain): one click bolds
        // every cell rather than clearing.
        let mut app = App::new();
        app.execute_range_text_style(Range::single(Address::A1), TextStyle::BOLD, true);
        app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        click_slot(&mut app, 11);
        for addr in [
            Address::A1,
            Address::new(SheetId(0), 1, 0),
            Address::new(SheetId(0), 0, 1),
            Address::new(SheetId(0), 1, 1),
        ] {
            assert_eq!(
                app.wb().cell_text_styles.get(&addr).copied(),
                Some(TextStyle::BOLD),
                "{addr:?} should be bold",
            );
        }
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
        app.icon_panel_area.set(Some(test_geom(TEST_PANEL)));
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
    fn hit_test_maps_cell_to_visually_dominant_icon() {
        // Realistic geometry from a 3-col panel at font (9, 18): the
        // PNG's 1:17 aspect renders width-constrained at 27×459 px,
        // and the ceiled cell rect is 3×26 — so the bottom 9 px of
        // row 25 is empty slack and each icon is visually 1.5 cells
        // tall. Sampling at cell midpoint pixels then dividing by the
        // true rendered pixel height pins each cell to the icon that
        // covers most of its pixels. Naive (cells-only) mapping drifts
        // by half an icon at the bottom — mis-reporting Help as Bold,
        // Save as spilling into Retrieve, etc.
        let geom = IconPanelGeom {
            rect: Rect::new(80, 4, 3, 26),
            rendered_px_h: 459,
            font_px_h: 18,
        };
        // Row 4 (cell 0) sits entirely inside Save's pixel band [0,27).
        assert_eq!(App::hit_test_slot(&geom, 4), Some(0));
        // Row 5 (cell 1, midpt = 27 px) — right on the Save/Retrieve
        // boundary. The formula treats it as Retrieve, so Save's hit-
        // test does not spill into Retrieve's visual region.
        assert_eq!(App::hit_test_slot(&geom, 5), Some(1));
        // Row 21 (cell 17, midpt = 315 px) is entirely inside Bold.
        assert_eq!(App::hit_test_slot(&geom, 21), Some(11));
        // Row 22 (cell 18, midpt = 333 px) is entirely Italic.
        assert_eq!(App::hit_test_slot(&geom, 22), Some(12));
        // Row 27 (cell 23, midpt = 423 px) is entirely Help.
        assert_eq!(App::hit_test_slot(&geom, 27), Some(15));
        // Row 28 (cell 24, midpt = 441 px) is entirely Pager.
        assert_eq!(App::hit_test_slot(&geom, 28), Some(16));
        // Row 29 (cell 25, midpt = 459 px) is in the empty slack at
        // the bottom; map to pager so the last row isn't a dead zone.
        assert_eq!(App::hit_test_slot(&geom, 29), Some(16));
    }

    /// Simulate a mouse-move at `(col, row)` by stashing a fake panel
    /// geometry (same fixture as click tests) and routing through
    /// [`App::handle_mouse`].
    fn hover(app: &mut App, area: Rect, col: u16, row: u16) {
        app.icon_panel_area.set(Some(test_geom(area)));
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

    // ---- Phase 4: drag-to-select ----

    /// Simulate a left-button drag at `(col, row)` — used after a
    /// preceding [`click_at`] to drive the drag-to-select path.
    fn drag_at(app: &mut App, col: u16, row: u16) {
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Drag(MouseButton::Left),
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    /// Simulate the left-button release ending a drag.
    fn release_at(app: &mut App, col: u16, row: u16) {
        app.handle_mouse(MouseEvent {
            kind: MouseEventKind::Up(MouseButton::Left),
            column: col,
            row,
            modifiers: KeyModifiers::NONE,
        });
    }

    #[test]
    fn mouse_drag_from_ready_enters_point_anchored_at_press_cell() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y); // Press at A1.
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2); // Drag to C3.
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:C3");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_drag_extends_then_shrinks_with_anchor_fixed() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        drag_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
        assert_eq!(app.pointer().display_full(), "A:D4");
        drag_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
        assert_eq!(app.pointer().display_full(), "A:B2");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_release_does_not_exit_point() {
        // Up just ends the drag; selection persists for follow-up
        // commands (Bold icon, /Range Format, …).
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        release_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:C3");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_drag_inside_existing_point_does_not_reanchor() {
        // Already in POINT via /RE — anchor at A1. Press on B2 (Phase 2
        // moves pointer; anchor stays), then drag to C3. The /RE anchor
        // must persist — we never overwrite an existing POINT anchor
        // with the mouse-press cell.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_point_via_range_erase(&mut app);
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
        click_at(&mut app, COL_B_X, BODY_TOP_Y + 1); // Press at B2.
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2); // Drag to C3.
        assert_eq!(app.mode, Mode::Point);
        assert_eq!(app.pointer().display_full(), "A:C3");
        assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    }

    #[test]
    fn mouse_drag_without_press_on_grid_is_ignored() {
        // A drag that wasn't preceded by a press inside the grid (e.g.
        // mouse-down landed on the column header, then drag onto the
        // body) must not promote into POINT.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_B_X, HEADER_Y); // Header click is ignored.
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Ready);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn mouse_drag_in_value_mode_is_ignored() {
        // Mid-formula entry: a stray drag must not corrupt the buffer
        // or change mode. Phase 3 splicing is press-driven, not drag.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+");
        assert_eq!(app.mode, Mode::Value);
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        assert_eq!(app.mode, Mode::Value);
        assert_eq!(entry_buffer(&app), "+");
    }

    #[test]
    fn mouse_drag_off_grid_does_not_move_pointer() {
        // While dragging, if the cursor leaves the grid (onto the
        // column header or row gutter), the pointer freezes — it does
        // not snap to the last in-grid cell or jump to the gutter.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        let pinned = app.pointer().display_full();
        drag_at(&mut app, 2, BODY_TOP_Y + 2); // Onto row gutter.
        assert_eq!(app.pointer().display_full(), pinned);
        assert_eq!(app.mode, Mode::Point);
    }

    #[test]
    fn mouse_release_clears_drag_state_so_next_drag_needs_a_press() {
        // After Up, a follow-up Drag without a fresh press is a no-op:
        // we shouldn't keep extending the previous selection just
        // because the OS sends spurious motion events.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        click_at(&mut app, COL_A_X, BODY_TOP_Y);
        drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        release_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
        // Now exit POINT — Esc, Esc.
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Ready);
        // Bare Drag (no press): must NOT promote to POINT.
        drag_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
        assert_eq!(app.mode, Mode::Ready);
    }

    // ---- Phase 5: scroll wheel ----

    fn wheel(app: &mut App, kind: MouseEventKind) {
        app.handle_mouse(MouseEvent {
            kind,
            column: 10,
            row: 10,
            modifiers: KeyModifiers::NONE,
        });
    }

    #[test]
    fn scroll_down_advances_viewport_row_offset_by_scroll_step() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        assert_eq!(app.wb().viewport_row_offset, 0);
        wheel(&mut app, MouseEventKind::ScrollDown);
        assert_eq!(app.wb().viewport_row_offset, MOUSE_SCROLL_STEP);
    }

    #[test]
    fn scroll_up_retreats_viewport_row_offset_saturating_at_zero() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        app.wb_mut().viewport_row_offset = 50;
        wheel(&mut app, MouseEventKind::ScrollUp);
        assert_eq!(app.wb().viewport_row_offset, 50 - MOUSE_SCROLL_STEP);
        // Many up-scrolls saturate at 0 (no underflow / panic).
        for _ in 0..100 {
            wheel(&mut app, MouseEventKind::ScrollUp);
        }
        assert_eq!(app.wb().viewport_row_offset, 0);
    }

    #[test]
    fn scroll_does_not_move_pointer() {
        // Modern spreadsheet convention: the wheel scrolls the
        // viewport only, leaving the cell pointer where it sits — even
        // if it ends up off-screen. The next keyboard arrow press will
        // pull it back into view via scroll_into_view.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        assert_eq!(app.pointer().display_full(), "A:A1");
        wheel(&mut app, MouseEventKind::ScrollDown);
        wheel(&mut app, MouseEventKind::ScrollDown);
        assert_eq!(app.pointer().display_full(), "A:A1");
    }

    #[test]
    fn scroll_does_not_change_mode() {
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        wheel(&mut app, MouseEventKind::ScrollDown);
        assert_eq!(app.mode, Mode::Ready);

        // In POINT, scrolling must keep POINT alive — selection
        // persists, viewport just moves.
        enter_point_via_range_erase(&mut app);
        assert_eq!(app.mode, Mode::Point);
        wheel(&mut app, MouseEventKind::ScrollDown);
        assert_eq!(app.mode, Mode::Point);
    }

    #[test]
    fn scroll_works_during_value_entry_without_corrupting_buffer() {
        // The wheel is a navigation gesture, not an entry gesture — it
        // must never touch the entry buffer or commit/cancel the
        // current entry.
        let mut app = App::new();
        let _ = app.render_to_buffer(80, 25);
        enter_value_with(&mut app, "+A1");
        assert_eq!(app.mode, Mode::Value);
        wheel(&mut app, MouseEventKind::ScrollDown);
        assert_eq!(app.mode, Mode::Value);
        assert_eq!(entry_buffer(&app), "+A1");
        assert!(app.wb().viewport_row_offset > 0);
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
        // UNDERLINED applies only over actual glyphs, so the cell
        // needs visible text — an empty cell carries no underline
        // even when one is set on its style.
        let mut app = App::new();
        for c in "x".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
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

    /// Read the buffer cell at `addr.col`'s base x-position plus
    /// `col_off` columns, on the row of `addr`.  Lets a test inspect
    /// the trailing padding columns inside a wider cell.
    fn buf_cell_at<'a>(
        app: &App,
        buf: &'a Buffer,
        addr: Address,
        col_off: u16,
    ) -> &'a ratatui::buffer::Cell {
        let dr = (addr.row - app.wb().viewport_row_offset) as u16;
        let y = PANEL_HEIGHT + 1 + dr;
        let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
        let layout = app.visible_column_layout(content_width);
        let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == addr.col).unwrap();
        let x = ROW_GUTTER + x_off + col_off;
        &buf[(x, y)]
    }

    #[test]
    fn underline_does_not_extend_into_trailing_padding() {
        // Short underlined label "hi" in a width-9 cell: 'h' and 'i'
        // carry UNDERLINED; the seven trailing padding spaces do not.
        let mut app = App::new();
        // Move pointer off A1 so the test cell isn't REVERSED.
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        for c in "hi".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let target = Address::new(SheetId::A, 0, 2); // A3
        app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

        let buf = app.render_to_buffer(80, 25);
        let h = buf_cell_at(&app, &buf, target, 0);
        assert_eq!(h.symbol(), "h");
        assert!(h.style().add_modifier.contains(Modifier::UNDERLINED));
        let i = buf_cell_at(&app, &buf, target, 1);
        assert_eq!(i.symbol(), "i");
        assert!(i.style().add_modifier.contains(Modifier::UNDERLINED));
        for off in 2..9u16 {
            let cell = buf_cell_at(&app, &buf, target, off);
            assert_eq!(cell.symbol(), " ", "padding at +{off} should be a space");
            assert!(
                !cell.style().add_modifier.contains(Modifier::UNDERLINED),
                "padding at +{off} should NOT carry UNDERLINED, got {:?}",
                cell.style().add_modifier,
            );
        }
    }

    #[test]
    fn underline_does_not_extend_past_spilled_label_text() {
        // Underlined label long enough to spill A→B→C; the tail of
        // the spill (padding past the last glyph) must not carry
        // UNDERLINED, even though the spill cells inherit the
        // owner's text style.
        let mut app = App::new();
        // Move pointer off A1.
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        // 11 chars at width 9 → spills one column into B3, leaving
        // 7 trailing pad columns inside B3.
        for c in "hello world".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let target = Address::new(SheetId::A, 0, 2); // A3
        app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

        let buf = app.render_to_buffer(80, 25);
        let neighbor = Address::new(SheetId::A, 1, 2); // B3
                                                       // First two B3 columns hold the spilled "ld" tail — underlined.
        let l = buf_cell_at(&app, &buf, neighbor, 0);
        assert_eq!(l.symbol(), "l");
        assert!(l.style().add_modifier.contains(Modifier::UNDERLINED));
        let d = buf_cell_at(&app, &buf, neighbor, 1);
        assert_eq!(d.symbol(), "d");
        assert!(d.style().add_modifier.contains(Modifier::UNDERLINED));
        // Remaining padding columns of B3 are space and unstyled.
        for off in 2..9u16 {
            let cell = buf_cell_at(&app, &buf, neighbor, off);
            assert_eq!(
                cell.symbol(),
                " ",
                "spill-tail padding at B3+{off} should be a space",
            );
            assert!(
                !cell.style().add_modifier.contains(Modifier::UNDERLINED),
                "spill-tail padding at B3+{off} should NOT carry UNDERLINED",
            );
        }
    }

    #[test]
    fn underline_continues_through_internal_space_at_cell_boundary() {
        // Long underlined label spills across A→B; the space between
        // "1991:" and "Sloane" lands at the A/B column boundary.
        // That space is internal to the original text, so the leading
        // column of B must still carry UNDERLINED — earlier per-slot
        // trim heuristics dropped it as if it were B's leading
        // padding.
        let mut app = App::new();
        // Set A's width to 20 so "INCOME SUMMARY 1991:" exactly fills
        // it and the spill boundary lands on the space character.
        // Pointer starts at A1 — /WCS 20 widens column A.
        for c in ['/', 'W', 'C', 'S'] {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        for c in "20".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        assert_eq!(app.col_width_of(SheetId::A, 0), 20);
        // Move pointer to A3 so it's not REVERSED-highlighting our
        // test cells.
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        for c in "INCOME SUMMARY 1991: Sloane Camera and Video".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let target = Address::new(SheetId::A, 0, 2); // A3
        app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

        let buf = app.render_to_buffer(120, 25);
        let neighbor = Address::new(SheetId::A, 1, 2); // B3
                                                       // First column of B3 is the boundary space — it's the
                                                       // internal " " of "...1991: Sloane..." and must stay
                                                       // underlined for the run to read continuously.
        let boundary = buf_cell_at(&app, &buf, neighbor, 0);
        assert_eq!(boundary.symbol(), " ");
        assert!(
            boundary.style().add_modifier.contains(Modifier::UNDERLINED),
            "internal-text space at A/B cell seam should keep UNDERLINED",
        );
        // Second column of B3 is 'S' — clearly part of text.
        let s_glyph = buf_cell_at(&app, &buf, neighbor, 1);
        assert_eq!(s_glyph.symbol(), "S");
        assert!(s_glyph.style().add_modifier.contains(Modifier::UNDERLINED));
    }

    #[test]
    fn underline_covers_internal_whitespace_between_glyphs() {
        // An internal space between two glyphs in a label is part of
        // the text run and should keep its underline so the line
        // reads as continuous under "a b".
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
        for c in "a b".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
        let target = Address::new(SheetId::A, 0, 2); // A3
        app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

        let buf = app.render_to_buffer(80, 25);
        // "a", " " (internal), "b" all carry UNDERLINED.
        let a = buf_cell_at(&app, &buf, target, 0);
        assert_eq!(a.symbol(), "a");
        assert!(a.style().add_modifier.contains(Modifier::UNDERLINED));
        let mid = buf_cell_at(&app, &buf, target, 1);
        assert_eq!(mid.symbol(), " ");
        assert!(
            mid.style().add_modifier.contains(Modifier::UNDERLINED),
            "internal space between 'a' and 'b' should remain underlined",
        );
        let b = buf_cell_at(&app, &buf, target, 2);
        assert_eq!(b.symbol(), "b");
        assert!(b.style().add_modifier.contains(Modifier::UNDERLINED));
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
