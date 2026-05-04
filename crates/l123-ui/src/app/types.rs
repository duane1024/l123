//! Leaf types used by the App state machine.
//!
//! Anything in this module is shared by `app::mod.rs` and (eventually)
//! by sibling submodules. Types keep their original visibility; struct
//! fields are exposed at `pub(super)` so the parent module can construct
//! and mutate them directly without going through accessors.

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::Arc;

use tokio::sync::oneshot;

use l123_core::{
    Address, Alignment, Border, CellContents, CurrencyPosition, Fill, FontStyle, Format,
    FormatKind, HAlign, International, LabelPrefix, Merge, Range, RgbColor, SheetId, SheetState,
    Table, TextStyle,
};
use l123_engine::IronCalcEngine;
use l123_graph::{GraphDef, Series};
use l123_macro::MacroAction;
use l123_menu::{self as menu, MenuItem};
use l123_print::{encode::lp::LpOptions, PrintContentMode, PrintFormatMode, WorkbookView};
use ratatui::layout::Rect;

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum EntryKind {
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
pub(super) enum StatView {
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

#[derive(Debug)]
pub(super) struct Entry {
    pub(super) kind: EntryKind,
    pub(super) buffer: String,
    /// Byte index into `buffer`, 0..=buffer.len(). Always lands on a
    /// char boundary. Initialized to `buffer.len()` (cursor at end —
    /// matches typing-into-an-empty-buffer behavior).
    pub(super) cursor: usize,
}

/// All per-file state — one instance per active file. Session-level
/// fields (mode, menu, entry buffer, …) live on [`super::App`].
pub(super) struct Workbook {
    pub(super) engine: IronCalcEngine,
    pub(super) cells: HashMap<Address, CellContents>,
    pub(super) cell_formats: HashMap<Address, Format>,
    /// Workbook-wide default cell format set by `/Worksheet Global
    /// Format`. Cells without a `cell_formats` entry inherit this.
    /// Initialized to General.
    pub(super) global_format: Format,
    /// `/Worksheet Global Default Other International` — punctuation,
    /// date/time intl style, negative style, and currency symbol/
    /// position. Threaded into `format_number` and `parse_typed_value`
    /// so cell display and number entry honor the configured locale.
    /// Persistence to L123.CNF via `/WGDU` is out of scope; session-
    /// only for now.
    pub(super) international: International,
    /// Per-cell text-style overrides (bold / italic / underline) set
    /// by the WYSIWYG `:Format Bold|Italic|Underline Set|Clear`
    /// commands.  Empty style = no entry.
    pub(super) cell_text_styles: HashMap<Address, TextStyle>,
    /// Per-cell explicit alignment from an xlsx import (or a future
    /// /Range Alignment command).  Default alignment = no entry, so
    /// the label-prefix / number-right-align contract still governs
    /// uncharted cells.
    pub(super) cell_alignments: HashMap<Address, Alignment>,
    /// Per-cell background-fill color from an xlsx import.  Default
    /// (no fill) = no entry; the terminal default shows through.
    /// Rendered via `Style::bg(Color::Rgb(...))` at grid-paint time.
    pub(super) cell_fills: HashMap<Address, Fill>,
    /// Per-cell xlsx-derived font attributes (foreground color, size,
    /// strikethrough).  Sits alongside `cell_text_styles` (the 1-2-3
    /// WYSIWYG bold/italic/underline triple); the two maps can both
    /// apply to the same cell.  Size is preserve-only.
    pub(super) cell_font_styles: HashMap<Address, FontStyle>,
    /// Per-cell border edges from xlsx imports.  All four sides + color
    /// round-trip; v1 renders only **right-edge** borders (overlaying
    /// a box-drawing glyph on the rightmost column of the cell's slot).
    /// Top, bottom, and left borders are preserved on save but not yet
    /// rendered — adding row-direction borders requires a grid layout
    /// change (seam rows) that's out of scope here.
    pub(super) cell_borders: HashMap<Address, Border>,
    /// Per-cell comments from xlsx imports.  Renders as a small
    /// corner marker (`'`) on the cell's right-edge column; when the
    /// pointer lands on the cell, the author + text appears on
    /// control-panel line 3.  Note (IronCalc 0.7): the xlsx exporter
    /// drops comments — we preserve them in-memory and render them,
    /// but `/FS` will lose them until the upstream gap closes.
    pub(super) comments: HashMap<Address, l123_core::Comment>,
    /// Merged ranges by sheet.  The grid renderer paints the
    /// anchor's content across the merge's column span; non-anchor
    /// cells in the same row render blank (and block label-spill
    /// from neighbors).  Multi-row merges: the anchor's content
    /// shows on the anchor's row only; subsequent rows of the merge
    /// area render blank — top-aligned, like Excel.  Cursor
    /// navigation does NOT yet snap to the anchor (deferred).
    pub(super) merges: HashMap<SheetId, Vec<Merge>>,
    /// Per-sheet frozen-pane counts: `(rows, cols)` indicate how many
    /// rows from the top and columns from the left stay pinned in the
    /// viewport while the rest of the grid scrolls.  `(0, 0)` (or no
    /// entry) means no freeze.  Round-trips natively through xlsx.
    pub(super) frozen: HashMap<SheetId, (u32, u16)>,
    /// Per-sheet visibility from xlsx imports.  Sheets not in the
    /// map default to `Visible`.  Hidden / VeryHidden sheets are
    /// skipped by `Ctrl-PgUp/PgDn` navigation; loaded files with a
    /// hidden first sheet get the pointer redirected to the first
    /// `Visible` sheet on import so the user lands somewhere they
    /// can interact with.
    pub(super) sheet_states: HashMap<SheetId, SheetState>,
    /// Excel tables (named ranges with header / autofilter / totals
    /// metadata) by sheet.  v1: round-trip-only — preserved through
    /// the engine on save, but no UI surface yet (no filter widgets,
    /// no `/Data Query Define` integration).  IronCalc 0.7's xlsx
    /// exporter doesn't write tables, so `/FS` drops them today
    /// (pinned by `tables_are_dropped_on_xlsx_save_upstream_gap`).
    pub(super) tables: HashMap<SheetId, Vec<Table>>,
    /// Per-sheet tab color from an xlsx import.  When set, the sheet's
    /// letter in the status-line indicator renders with this fg color.
    /// Note (IronCalc 0.7): the xlsx *export* path drops tab colors —
    /// we preserve them in the model and render them, but `/FS` will
    /// not carry them back to disk until the upstream gap closes.
    pub(super) sheet_colors: HashMap<SheetId, RgbColor>,
    pub(super) col_widths: HashMap<(SheetId, u16), u8>,
    /// Workbook-wide default column width (1..240). Applied to any
    /// column without a `col_widths` entry. Set by `/Worksheet Global
    /// Col-Width`. Initialized to 9 — the 1-2-3 R3 factory default.
    pub(super) default_col_width: u8,
    /// Columns marked hidden by `/Worksheet Column Hide`. Skipped by
    /// the grid renderer; pointer and formulas can still address them.
    /// Not persisted through xlsx today — IronCalc 0.7 doesn't model
    /// a per-column hidden flag.
    pub(super) hidden_cols: HashSet<(SheetId, u16)>,
    /// Last-saved-to path. Prefilled into `/FS` prompts so re-save is
    /// a single Enter. `None` until the file has been saved at least
    /// once.
    pub(super) active_path: Option<PathBuf>,
    /// True when the workbook has unsaved changes. Drives the `/QY`
    /// warn-on-quit second confirm. Flipped on by mutating commits;
    /// cleared on a successful `/FS`.
    pub(super) dirty: bool,
    pub(super) pointer: Address,
    pub(super) viewport_col_offset: u16,
    pub(super) viewport_row_offset: u32,
    /// Command journal for Undo (Alt-F4). Each mutating command
    /// pushes an inverse entry. Pop-and-apply reverts.
    pub(super) journal: Vec<JournalEntry>,
    /// The unnamed working graph — target of every `/Graph` menu
    /// command until `/Graph Name Create` snapshots it by name.
    pub(super) current_graph: GraphDef,
    /// Named graphs defined via `/Graph Name Create`. `Use` restores
    /// one into `current_graph`; `Delete` drops it; `Reset` wipes all.
    #[allow(dead_code)] // wired by `/Graph Name Create` in a later slice
    pub(super) graphs: BTreeMap<String, GraphDef>,
    /// Range names defined via `/Range Name Create`. Keyed by the
    /// lowercased name (Lotus 1-2-3 names are case-insensitive). The
    /// engine also stores names for formula resolution; this UI-side
    /// mirror is what POINT typed-buffer name resolution reads, so we
    /// don't need to round-trip through the engine to look up a range.
    pub(super) named_ranges: HashMap<String, Range>,
    /// Optional notes attached to named ranges by `/Range Name Note
    /// Create`. Keyed identically to `named_ranges` (lowercased name).
    pub(super) name_notes: HashMap<String, String>,
    /// Cells that `/Range Unprot` has marked as writable. Cells not
    /// in this set are "protected" by default. Has no effect unless
    /// `super::App::global_protection` is on.
    pub(super) cell_unprotected: HashSet<Address>,
}

impl Workbook {
    /// Find the merge containing `addr`, if any.  O(N) over the
    /// sheet's merge list — fine for the small N (~dozens) typical
    /// of real workbooks; revisit if a fixture pushes it past 1000.
    pub(super) fn merge_at(&self, addr: Address) -> Option<Merge> {
        self.merges
            .get(&addr.sheet)?
            .iter()
            .find(|m| m.contains(addr))
            .copied()
    }

    pub(super) fn new() -> Self {
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
pub(super) struct IconPanelGeom {
    pub(super) rect: Rect,
    pub(super) rendered_px_h: u32,
    pub(super) font_px_h: u16,
}

/// User-visible identity shown on the startup splash. The renderer
/// prints `user` after "User name:" and `organization` after
/// "Organization:", matching the 1-2-3 R3.4a licensing block.
#[derive(Debug, Clone)]
pub struct SplashInfo {
    pub user: String,
    pub organization: String,
}

/// State kept for the duration of a [`l123_core::Mode::Graph`] overlay.
#[derive(Clone)]
pub(super) struct GraphOverlay {
    /// Numeric values snapshotted off the engine at enter time.
    pub(super) values: l123_graph::GraphValues,
    /// Pre-rendered PNG (via plotters) decoded into a DynamicImage.
    /// Populated only when the app has a graphical picker — feeds
    /// ratatui-image's `Image` widget at render time.
    pub(super) img: Option<image::DynamicImage>,
}

/// `/Data Sort` direction: ascending = low→high, descending = high→low.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SortDir {
    Ascending,
    Descending,
}

/// Which key slot a Primary-Key / Secondary-Key / Extra-Key
/// Asc/Desc submenu is about to write into.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SortKeySlot {
    Primary,
    Secondary,
    Extra,
}

/// Persisted `/Data Sort` settings. Sticky across the Sort menu and
/// across separate `/DS` invocations — Reset is the only way to
/// clear it short of restarting the session.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct DataSortState {
    pub(super) data_range: Option<Range>,
    pub(super) primary: Option<(u16, SortDir)>,
    pub(super) secondary: Option<(u16, SortDir)>,
    pub(super) extra: Option<(u16, SortDir)>,
}

/// Persisted `/Data Regression` settings, sticky across the
/// Regression menu and across separate `/DR` invocations.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct DataRegressionState {
    pub(super) x_range: Option<Range>,
    pub(super) y_range: Option<Range>,
    pub(super) output_anchor: Option<Address>,
    pub(super) intercept_zero: bool,
}

/// Persisted `/Data Parse` settings — the input column (whose top
/// row is the format line) and the output anchor.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct DataParseState {
    pub(super) input_range: Option<Range>,
    pub(super) output_anchor: Option<Address>,
}

/// Persisted `/Data Query` settings — three rectangular ranges
/// that drive Find / Extract / Unique / Del.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct DataQueryState {
    pub(super) input: Option<Range>,
    pub(super) criteria: Option<Range>,
    pub(super) output: Option<Range>,
}

/// Which cell kinds `/Range Search` walks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum SearchScope {
    Formulas,
    Labels,
    Both,
}

/// Adjacent-cell direction for `/Range Name Labels`. Each label in
/// the picked range gets a name pointing one cell in this direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum LabelDirection {
    Right,
    Down,
    Left,
    Up,
}

/// Live state of a `/Range Search` session between scope selection
/// and the final Find or Replace leaf.
#[derive(Debug, Clone)]
pub(super) struct SearchSession {
    pub(super) scope: SearchScope,
    pub(super) range: Range,
    pub(super) search: String,
    /// Cached matches, populated when the user picks Find. Replace
    /// recomputes its own match set just before applying.
    pub(super) matches: Vec<Address>,
    /// Index of the current highlighted match within `matches`.
    pub(super) cursor: usize,
}

/// Live state of a `/Print File` session between filename commit and
/// final Go. Holds the destination path, the chosen range, and the
/// current page-decoration settings.
/// Where the Go step sends its output.
#[derive(Debug, Clone)]
pub(super) enum PrintDestination {
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
pub(super) struct PrintSession {
    pub(super) destination: PrintDestination,
    /// One or more print ranges. Empty until `/PF Range` runs. Multiple
    /// ranges (typed `A1..B2,C3..D4` in POINT) are emitted in order;
    /// each range is a separate "page" of the output (Lotus separates
    /// ranges with a form-feed when printed; here we emit a blank line
    /// between them in the unformatted/file path).
    pub(super) ranges: Vec<Range>,
    /// Three-part header string (`L|C|R`). Empty means no header.
    pub(super) header: String,
    /// Three-part footer string (`L|C|R`). Empty means no footer.
    pub(super) footer: String,
    /// Printer init/escape string. Prepended verbatim to the output —
    /// ahead of the ASCII page bytes for a `.prn` file write, or piped
    /// to CUPS `lp` ahead of the ASCII stream for `/Print Printer`.
    /// Empty = no setup. PDF output ignores it (escape codes are
    /// meaningless inside a PDF body).
    pub(super) setup_string: String,
    /// CUPS queue name passed as `lp -d <name>` for `/Print Printer`.
    /// Empty = use the system default printer. Stored regardless of
    /// destination kind; only consumed when destination is `Printer`.
    pub(super) lp_destination: String,
    /// As-Displayed (default) or Cell-Formulas.
    pub(super) content_mode: PrintContentMode,
    /// Formatted (default — emits header/footer) or Unformatted
    /// (range content only).
    pub(super) format_mode: PrintFormatMode,
    /// Left margin: N spaces prepended to every output line.
    pub(super) margin_left: u16,
    /// Right margin — accepted but not yet honored (no wrapping
    /// implemented). Storing it keeps the menu muscle memory intact.
    pub(super) margin_right: u16,
    /// Top margin: N blank lines above the first output line.
    pub(super) margin_top: u16,
    /// Bottom margin — accepted but not yet honored (no pagination
    /// yet).
    pub(super) margin_bottom: u16,
    /// Lines per page. 0 = no pagination.
    pub(super) pg_length: u16,
    /// Next page number to print at the start of Go. Persists across
    /// successive Gos in the same session so headers using `#` count
    /// up; `/PF Align` resets it to 1.
    pub(super) next_page: u32,
}

impl PrintSession {
    pub(super) fn new_file(path: PathBuf) -> Self {
        Self::with_destination(PrintDestination::File(path))
    }

    pub(super) fn new_printer() -> Self {
        Self::with_destination(PrintDestination::Printer(LpOptions::default()))
    }

    pub(super) fn new_encoded(path: PathBuf) -> Self {
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
    pub(super) fn clear_all(&mut self) {
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
pub(super) enum JournalEntry {
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
    /// Restore per-cell `Border` overrides after the Outline SmartIcon
    /// (icon 20) toggles a perimeter outline on a range. `None` =
    /// no override before.
    RangeBorder {
        entries: Vec<(Address, Option<Border>)>,
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
pub(super) struct SaveConfirmState {
    pub(super) path: PathBuf,
    /// 0=Cancel, 1=Replace, 2=Backup — matches `SAVE_CONFIRM_ITEMS` below.
    pub(super) highlight: usize,
}

/// Items shown on line 2 of the Cancel/Replace/Backup submenu. The
/// first letter of each is the accelerator.
pub(super) const SAVE_CONFIRM_ITEMS: &[(&str, &str)] = &[
    ("Cancel", "Abort the save"),
    ("Replace", "Overwrite the existing file"),
    ("Backup", "Rename existing to .BAK then save"),
];

#[derive(Debug, Clone)]
pub(super) struct EraseConfirmState {
    pub(super) path: PathBuf,
    /// 0=No, 1=Yes — matches `FILE_ERASE_CONFIRM_ITEMS` below.
    pub(super) highlight: usize,
}

/// Items shown on line 2 of the No/Yes confirm submenu invoked by
/// `/File Erase` after the user types a path. First letter is the
/// accelerator.
pub(super) const FILE_ERASE_CONFIRM_ITEMS: &[(&str, &str)] = &[
    ("No", "Do not erase the file"),
    ("Yes", "Permanently delete the file from disk"),
];

/// Numeric (or short-text) prompt state for commands that need an argument
/// before descending into POINT. E.g. /RFC → "Enter number of decimal
/// places (0..15): 2" → then POINT for the range.
#[derive(Debug, Clone)]
pub(super) struct PromptState {
    pub(super) label: String,
    pub(super) buffer: String,
    /// What the command wants to do once the prompt commits.
    pub(super) next: PromptNext,
    /// True while the buffer still holds the auto-filled default. The
    /// first printable keystroke clears it (1-2-3 "typed input replaces
    /// the default" convention).
    pub(super) fresh: bool,
}

/// §4.7 — a long-running op queued while WAIT mode is active. The
/// op begins life in `OpState::Queued` (not yet on the tokio
/// thread); the next `tick()` moves it to `OpState::Running` by
/// spawning a `spawn_blocking` task. Subsequent ticks poll the
/// oneshot for completion. Drop the whole struct to cancel — the
/// shared `progress.cancel` flag lets cooperative workers bail
/// early.
pub(super) struct PendingAsyncOp {
    /// Verb prefix on control-panel line 3 — e.g. "Loading",
    /// "Saving", "Importing", "Recalculating". Joined with
    /// `display_name` for the rendered "Loading foo.csv…" line.
    pub(super) verb: &'static str,
    /// Filename basename or other short noun — rendered after `verb`.
    /// Empty for ops with no associated file (e.g. recalc).
    pub(super) display_name: String,
    /// Shared progress / cancel state, written by the worker and
    /// read by the renderer.
    pub(super) progress: AsyncProgress,
    /// State of the op — Queued (not yet on tokio) or Running
    /// (worker spawned, waiting for result).
    pub(super) state: OpState,
}

impl PendingAsyncOp {
    /// Control-panel line 3 text. With a `total > 0` progress reading
    /// we render a fixed-width bar `[████░░] N%`; otherwise just the
    /// verb + name.
    pub(super) fn render_line3(&self) -> String {
        let head = if self.display_name.is_empty() {
            format!(" {}…", self.verb)
        } else {
            format!(" {} {}…", self.verb, self.display_name)
        };
        let total = self.progress.total.load(Ordering::Relaxed);
        if total == 0 {
            return head;
        }
        let done = self.progress.done.load(Ordering::Relaxed).min(total);
        let pct = ((done * 100) / total) as u16;
        let bar = render_progress_bar(done, total, PROGRESS_BAR_CELLS);
        format!("{head} [{bar}] {pct}%")
    }
}

/// Width (in cells) of the rendered `[████░░]` progress bar. Each
/// cell = `total / N` bytes; partial cells round down.
pub(super) const PROGRESS_BAR_CELLS: u32 = 20;

pub(super) fn render_progress_bar(done: u64, total: u64, cells: u32) -> String {
    if total == 0 || cells == 0 {
        return String::new();
    }
    let filled = ((done.saturating_mul(cells as u64)) / total).min(cells as u64) as u32;
    let mut s = String::with_capacity(cells as usize * 3);
    for _ in 0..filled {
        s.push('\u{2588}');
    }
    for _ in filled..cells {
        s.push('\u{2591}');
    }
    s
}

/// Shared state between the UI thread and a worker task — the worker
/// writes progress and reads `cancel`; the UI thread does the inverse.
#[derive(Clone, Default)]
pub(super) struct AsyncProgress {
    /// Bytes (or rows) processed so far.
    pub(super) done: Arc<AtomicU64>,
    /// Total bytes (or rows). Zero = indeterminate; render hides the
    /// bar and shows the verb-only line.
    pub(super) total: Arc<AtomicU64>,
    /// Set by Ctrl-Break. Workers in cooperative loops (CSV row
    /// scanner, import row scanner) check this between rows and
    /// return early; ops backed by a single opaque IronCalc call
    /// (xlsx load, xlsx save, recalc) only honor it before/after.
    pub(super) cancel: Arc<AtomicBool>,
}

/// State machine for `PendingAsyncOp`. The `Queued` payload is
/// boxed because some variants embed an `IronCalcEngine` (~1KB on
/// the stack) that would otherwise dominate the enum size.
pub(super) enum OpState {
    /// Worker hasn't been spawned yet. The next `tick()` moves it to
    /// `Running` (unless `block_next_async_op` is set, which lets
    /// transcripts observe pre-flight WAIT state).
    Queued(Box<QueuedOp>),
    /// Worker spawned; waiting for the oneshot result.
    Running(oneshot::Receiver<AsyncResult>),
}

/// What the worker should do — variants own the inputs they need
/// (path, taken-out engine, etc.) so the spawn-blocking closure has
/// nothing to borrow from `super::App`.
pub(super) enum QueuedOp {
    /// `/File Retrieve` — dispatched by extension inside the worker.
    FileRetrieve { path: PathBuf },
    /// `/File Save` — engine has been taken out of the workbook
    /// (placeholder swapped in) and travels with the op.
    FileSave {
        engine: IronCalcEngine,
        path: PathBuf,
        formula_sources: HashMap<Address, String>,
    },
    /// `/File Import Numbers` — workbook engine taken out; the
    /// worker fills it from the parsed CSV starting at `origin`.
    FileImportNumbers {
        engine: IronCalcEngine,
        path: PathBuf,
        origin: Address,
    },
    /// `/File Import Text` — same as Numbers but each line is a
    /// single label (no comma-splitting).
    FileImportText {
        engine: IronCalcEngine,
        path: PathBuf,
        origin: Address,
    },
    /// F9 recalc, gated on cell count > `super::RECALC_WAIT_CELL_THRESHOLD`.
    Recalc { engine: IronCalcEngine },
}

/// What the worker hands back over the oneshot. Each variant carries
/// the engine (or `None` if construction failed before we got one)
/// plus any cells-cache deltas the main thread should apply.
pub(super) enum AsyncResult {
    /// `/File Retrieve` of an xlsx — the worker built a fresh engine
    /// from the file. Cells/styles are pulled out of it on the main
    /// thread (re-using the existing post-load cache rebuild).
    FileRetrieveXlsx {
        engine: IronCalcEngine,
        path: PathBuf,
        is_wk3: bool,
    },
    /// `/File Retrieve` of a csv — the worker built a fresh engine
    /// pre-populated with the parsed cells; `cells` is the matching
    /// UI-side cache.
    FileRetrieveCsv {
        engine: IronCalcEngine,
        path: PathBuf,
        cells: Vec<(Address, CellContents)>,
    },
    /// `/File Save` — engine returned, with success or error string.
    FileSave {
        engine: IronCalcEngine,
        path: PathBuf,
        result: std::result::Result<(), String>,
    },
    /// Import op completed (text or numbers). `cells` is the run of
    /// UI-side entries to merge into `Workbook::cells`.
    FileImport {
        engine: IronCalcEngine,
        cells: Vec<(Address, CellContents)>,
    },
    /// F9 recalc done — engine carries the recomputed values.
    Recalc { engine: IronCalcEngine },
    /// Worker bailed because of a cancel flag or pre-spawn check.
    /// Engine is returned so the main thread can put it back.
    Cancelled { engine: Option<IronCalcEngine> },
    /// Worker hit an error (file open failed, etc.). Engine is
    /// returned (when one was taken out) so the workbook isn't
    /// stranded with a placeholder.
    Errored {
        engine: Option<IronCalcEngine>,
        message: String,
    },
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum PromptNext {
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
    /// cell lives on `super::App::pending_macro_input_loc` (PromptNext is
    /// `Copy` so it can't carry a `String`).
    MacroGetInput {
        numeric: bool,
    },
    /// `/Data Fill` — first prompt: starting value of the sequence.
    /// Default is `0`. After commit, descends to `DataFillStep`.
    DataFillStart {
        range: Range,
    },
    /// `/Data Fill` — second prompt: per-cell increment. Default is
    /// `1`. After commit, descends to `DataFillStop`.
    DataFillStep {
        range: Range,
        start: f64,
    },
    /// `/Data Fill` — third prompt: clamp value. Default is `2047`
    /// (R3.4a's documented default). After commit, the sequence is
    /// written into `range` column-major.
    DataFillStop {
        range: Range,
        start: f64,
        step: f64,
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
    pub(super) kind: FileListKind,
    pub(super) entries: Vec<PathBuf>,
    pub(super) highlight: usize,
    /// Index of the first entry rendered on the overlay. Kept in sync
    /// with `highlight` so the selected row is always visible.
    pub(super) view_offset: usize,
}

/// Visible window size (rows) for the /File List overlay. Also the
/// distance moved by PgUp / PgDn. Fixed rather than dynamic because the
/// key handler runs before render knows the real terminal height;
/// render clamps to the actual area anyway.
pub(super) const FILE_LIST_PAGE_SIZE: usize = 10;

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
    pub(super) entries: Vec<(String, Range)>,
    pub(super) highlight: usize,
    pub(super) view_offset: usize,
    pub(super) origin: NameListOrigin,
}

pub(super) const NAME_LIST_PAGE_SIZE: usize = 10;

impl PromptNext {
    pub(super) fn accepts_char(self, c: char) -> bool {
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
            // /Data Fill takes signed real numbers — digits, decimal
            // point, and a leading sign. We don't validate placement
            // here; the commit handler parses with f64::from_str.
            PromptNext::DataFillStart { .. }
            | PromptNext::DataFillStep { .. }
            | PromptNext::DataFillStop { .. } => c.is_ascii_digit() || matches!(c, '.' | '-' | '+'),
        }
    }
}

/// Characters accepted inside a filename/path prompt. Deliberately
/// narrower than 1-2-3's "anything goes" — we exclude keys with menu
/// semantics (`/`, period-free submenus). `/` is fine; `.` is fine.
pub(super) fn is_path_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || matches!(c, '.' | '-' | '_' | '/' | '\\' | ' ' | '~')
}

/// `/Data Parse` format-line field kind. Each new field marker
/// (L/V/D/T/S) opens a field that extends through subsequent `>`
/// continuation chars until the next marker (or end of line).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormatField {
    Label,
    Value,
    Date,
    Time,
    Skip,
}

/// Transient state while the user is selecting a cell/range in POINT mode.
#[derive(Debug, Clone)]
pub(super) struct PointState {
    /// Anchor corner. `None` after a single Esc — in that state the
    /// pointer moves freely without growing a range; a second Esc cancels.
    pub(super) anchor: Option<Address>,
    /// Which command initiated POINT, so that `Enter` routes the selected
    /// range back to the right handler.
    pub(super) pending: PendingCommand,
    /// Lotus-style typed range buffer (e.g. `c8..d12`). Empty in the
    /// usual highlight-by-arrows flow. When non-empty, line 3 shows the
    /// buffer in place of the auto-derived highlight, and `Enter` parses
    /// it (via [`Range::parse_with_default_sheet`]) to override the
    /// committed range.
    pub(super) typed: String,
}

/// Which channel `:Format Color` is touching: cell background fill,
/// font foreground color, or both (Reset).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ColorTarget {
    Background,
    Text,
    Both,
}

/// One source cell's frozen formatting state, used by `:Special
/// Copy` / `:Special Move`. `None` on a field means the source had no
/// override — when written, the destination's matching entry is
/// removed (so the dest ends up *equal* to the source, not merged).
#[derive(Debug, Clone, Copy)]
pub(super) struct FormatSnapshot {
    pub(super) addr: Address,
    pub(super) format: Option<Format>,
    pub(super) text_style: Option<TextStyle>,
    pub(super) alignment: Option<Alignment>,
    pub(super) fill: Option<Fill>,
    pub(super) font_style: Option<FontStyle>,
    pub(super) border: Option<Border>,
}

/// Commands in progress that are waiting on one more POINT selection.
#[derive(Debug, Clone, Copy)]
pub(super) enum PendingCommand {
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
    /// First POINT of `:Special Copy` — pick the source range whose
    /// formatting will be replicated.
    SpecialCopyFrom,
    /// Second POINT of `:Special Copy` — pick the destination. On
    /// commit, every formatting attribute on each source cell
    /// overwrites the corresponding destination cell's formatting.
    /// Cell contents are untouched.
    SpecialCopyTo {
        source: Range,
    },
    /// First POINT of `:Special Move` — pick the source range.
    SpecialMoveFrom,
    /// Second POINT of `:Special Move` — pick the destination. On
    /// commit, formatting is copied to the destination and then
    /// cleared from any source cells outside the destination block.
    SpecialMoveTo {
        source: Range,
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
    /// POINT step of `/Data Fill`. On commit, kick off the
    /// Start → Step → Stop prompt chain that culminates in the
    /// sequence write.
    DataFillRange,
    /// POINT step of `/Data Sort Data-Range`. On commit, store the
    /// selected range on `App.data_sort` and re-enter the Sort menu.
    DataSortDataRange,
    /// POINT step of `/Data Sort Primary-Key` / `Secondary-Key` —
    /// the selected cell's column becomes the key column. Which slot
    /// is being set lives on `App.pending_sort_key_slot`.
    DataSortKey,
    /// First POINT of `/Data Distribution` — the values range. On
    /// commit, descend into [`PendingCommand::DataDistributionBins`]
    /// to collect the bin column.
    DataDistributionValues,
    /// Second POINT of `/Data Distribution` — the bin range. Must
    /// be a single column; on commit, write the frequency counts to
    /// the column immediately right of the bins.
    DataDistributionBins {
        values: Range,
    },
    /// POINT step of `/Data Regression X-Range`.
    DataRegressionXRange,
    /// POINT step of `/Data Regression Y-Range`.
    DataRegressionYRange,
    /// POINT step of `/Data Regression Output-Range`. Only the
    /// top-left of the picked range is used as the output anchor.
    DataRegressionOutputRange,
    /// First POINT of `/Data Matrix Invert` — square matrix to
    /// invert. On commit, descend into
    /// [`PendingCommand::DataMatrixInvertOutput`].
    DataMatrixInvertInput,
    /// Output-anchor POINT of `/Data Matrix Invert` — only the
    /// pointer position is used.
    DataMatrixInvertOutput {
        source: Range,
    },
    /// First POINT of `/Data Matrix Multiply` — matrix A.
    DataMatrixMultiplyA,
    /// Second POINT of `/Data Matrix Multiply` — matrix B
    /// (`cols(A)` must equal `rows(B)`).
    DataMatrixMultiplyB {
        a: Range,
    },
    /// Output-anchor POINT of `/Data Matrix Multiply` — only the
    /// pointer position is used.
    DataMatrixMultiplyOutput {
        a: Range,
        b: Range,
    },
    /// POINT step of `/Data Parse Input-Column` — single column
    /// containing the format-line label at the top and data rows
    /// below.
    DataParseInputColumn,
    /// POINT step of `/Data Parse Output-Range` — only the pointer
    /// position is used as the output top-left.
    DataParseOutputRange,
    /// First POINT of `/Data Table 1` — the rectangular table range.
    /// Top row is the corner + formulas; left column is the corner +
    /// variable values; body is filled by the executor.
    DataTable1Range,
    /// Second POINT of `/Data Table 1` — Input cell 1. Variable
    /// values from the table's left column are substituted here
    /// before each formula re-evaluation.
    DataTable1Input1 {
        range: Range,
    },
    /// First POINT of `/Data Table 2` — table range. Corner cell
    /// holds the formula; left column = var-1, top row = var-2.
    DataTable2Range,
    /// Second POINT of `/Data Table 2` — Input cell 1 (left-column
    /// values substitute here).
    DataTable2Input1 {
        range: Range,
    },
    /// Third POINT of `/Data Table 2` — Input cell 2 (top-row
    /// values substitute here). On commit, run the executor.
    DataTable2Input2 {
        range: Range,
        input1: Address,
    },
    /// POINT step of `/Data Query Input`.
    DataQueryInput,
    /// POINT step of `/Data Query Criteria`.
    DataQueryCriteria,
    /// POINT step of `/Data Query Output`.
    DataQueryOutput,
}

impl PendingCommand {
    pub(super) fn prompt(self) -> &'static str {
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
            PendingCommand::SpecialCopyFrom => "Enter range to copy formatting FROM:",
            PendingCommand::SpecialCopyTo { .. } => "Enter range to copy formatting TO:",
            PendingCommand::SpecialMoveFrom => "Enter range to move formatting FROM:",
            PendingCommand::SpecialMoveTo { .. } => "Enter range to move formatting TO:",
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
            PendingCommand::DataFillRange => "Enter fill range:",
            PendingCommand::DataSortDataRange => "Enter data-range to sort:",
            PendingCommand::DataSortKey => "Enter cell in primary/secondary key column:",
            PendingCommand::DataDistributionValues => "Enter values range:",
            PendingCommand::DataDistributionBins { .. } => "Enter bin range (single column):",
            PendingCommand::DataRegressionXRange => "Enter X-range (independent variable):",
            PendingCommand::DataRegressionYRange => "Enter Y-range (dependent variable):",
            PendingCommand::DataRegressionOutputRange => "Enter output-range top-left:",
            PendingCommand::DataMatrixInvertInput => "Enter square matrix to invert:",
            PendingCommand::DataMatrixInvertOutput { .. } => "Enter output-range top-left:",
            PendingCommand::DataMatrixMultiplyA => "Enter first matrix:",
            PendingCommand::DataMatrixMultiplyB { .. } => "Enter second matrix:",
            PendingCommand::DataMatrixMultiplyOutput { .. } => "Enter output-range top-left:",
            PendingCommand::DataParseInputColumn => "Enter input column (with format-line row):",
            PendingCommand::DataParseOutputRange => "Enter output-range top-left:",
            PendingCommand::DataTable1Range => "Enter table range:",
            PendingCommand::DataTable1Input1 { .. } => "Enter Input cell 1:",
            PendingCommand::DataTable2Range => "Enter table range:",
            PendingCommand::DataTable2Input1 { .. } => "Enter Input cell 1:",
            PendingCommand::DataTable2Input2 { .. } => "Enter Input cell 2:",
            PendingCommand::DataQueryInput => "Enter input range (with field-name header row):",
            PendingCommand::DataQueryCriteria => "Enter criteria range:",
            PendingCommand::DataQueryOutput => "Enter output range:",
            // Free mouse-drag selection has no prompt — line 3 keeps
            // showing the live range, but no command label is shown.
            PendingCommand::MouseSelect => "",
            PendingCommand::WorksheetLearnRange => "Enter range to record into:",
        }
    }
}

/// Transient state used while the user is navigating the slash menu.
#[derive(Debug, Clone)]
pub(super) struct MenuState {
    /// Letters descended into, longest-ago first.
    pub(super) path: Vec<char>,
    /// Index into the currently-visible level.
    pub(super) highlight: usize,
    /// Message to display on line 3 — typically the last-selected leaf's
    /// identifier when it was `NotImplemented`.
    pub(super) message: Option<&'static str>,
    /// Optional alternate root, used for nested menus (e.g. the
    /// `/Print File` submenu). When Some, `level()` resolves `path`
    /// against this slice instead of `l123_menu::ROOT`.
    pub(super) override_root: Option<&'static [MenuItem]>,
}

impl MenuState {
    pub(super) fn fresh() -> Self {
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
    pub(super) fn rooted_at(root: &'static [MenuItem]) -> Self {
        Self {
            path: Vec::new(),
            highlight: 0,
            message: None,
            override_root: Some(root),
        }
    }

    pub(super) fn level(&self) -> &'static [MenuItem] {
        match self.override_root {
            Some(root) => menu::current_level_within(root, &self.path),
            None => menu::current_level(&self.path),
        }
    }

    pub(super) fn highlighted(&self) -> Option<&'static MenuItem> {
        self.level().get(self.highlight)
    }
}

/// One frame of the macro call stack. The interpreter executes
/// actions out of `remaining` and refills it from the cell at `pc`
/// whenever the buffer empties — this lets a subroutine call resume
/// the caller mid-line on `{RETURN}`.
pub(super) struct MacroFrame {
    pub(super) pc: Address,
    pub(super) remaining: VecDeque<MacroAction>,
}

impl MacroFrame {
    pub(super) fn starting_at(addr: Address) -> Self {
        Self {
            pc: addr,
            remaining: VecDeque::new(),
        }
    }
}

/// Top-level macro execution state. Lives on [`super::App`] while a macro
/// is running. `suspend = Some(...)` parks the interpreter so the
/// user can interact with the workbook between actions; the
/// `handle_key` post-dispatch hook clears the parked reason and
/// resumes via `pump_macro`.
pub(super) struct MacroState {
    pub(super) frames: Vec<MacroFrame>,
    /// Total actions executed so far across all frames; safety guard
    /// against runaway loops (cap at `MAX_MACRO_STEPS`).
    pub(super) steps: u32,
    /// Why the interpreter is parked, if it is. `None` when ready
    /// to advance.
    pub(super) suspend: Option<MacroSuspend>,
    /// One-shot flag: when true, [`super::App::step_macro`] skips the STEP-mode
    /// pre-pause so a single action can fire. Set by the user
    /// pressing Space at a STEP pause, cleared once consumed.
    pub(super) step_advance: bool,
}

/// Reasons a macro can pause mid-execution.
pub(super) enum MacroSuspend {
    /// `{?}` — resume on the next Enter the user presses.
    WaitEnter,
    /// `{GETLABEL p, loc}` / `{GETNUMBER p, loc}` — a prompt is up.
    /// The dest cell and numeric flag live on `super::App` because
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
pub(super) struct CustomMenuState {
    /// Display name + description for each menu column (item).
    pub(super) items: Vec<CustomMenuItem>,
    /// Address of the "row 2" (action) cell of the leftmost menu
    /// column. The action cell for item `i` is `(action_row.col +
    /// i, action_row.row)`.
    pub(super) action_row: Address,
    /// True when this was opened via `{MENUCALL}` — the macro will
    /// CALL (push frame) instead of BRANCH (replace PC).
    pub(super) is_call: bool,
    /// Currently highlighted item index.
    pub(super) highlight: usize,
}

pub(super) struct CustomMenuItem {
    pub(super) name: String,
    pub(super) description: String,
}

pub(super) const MAX_MACRO_STEPS: u32 = 100_000;

/// Per-column classification fed to the spill planner. `Label` keeps
/// the owned text so the borrowed [`l123_core::SpillSlot::Label`] that follows can
/// point into it for the planner's lifetime.
pub(super) enum RowInput {
    Empty,
    Label { prefix: LabelPrefix, text: String },
    Rendered(String),
}
