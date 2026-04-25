# XLSX Import Feature-Gap Plan

Companion to `docs/SPEC.md` and `docs/PLAN.md`. This document is the
execution plan for closing the gap between what IronCalc parses out of
`.xlsx` files and what L123 currently consumes. It is structured so a
fresh session can pick up any tier (or any single feature within a tier)
without reading prior chat history.

The gap analysis behind this plan was done against:
- L123 at `/Users/ddmoore/dev/l123/`
- IronCalc at `/Users/ddmoore/dev/IronCalc/`

When an IronCalc API path drifts (e.g. struct field renamed), re-verify
against the IronCalc source before coding. Treat the file/line citations
below as “last known good” pointers, not invariants.

---

## 1. Architecture pattern (read first)

Every XLSX feature in L123 follows the same five-layer wiring. When you
add a new feature, you walk these five layers in this order:

```
1.  l123-core            — define the data type (no external deps)
2.  l123-engine          — extend the Engine trait with get/set methods
3.  l123-engine adapter  — read from / write to IronCalc's model
4.  l123-ui (app.rs)     — mirror state into WorkbookState; drain on load,
                           push on save; render in the grid / control panel
5.  tests/acceptance     — keystroke transcript proves end-to-end behavior
```

Reference implementations to copy the shape from:

- **Per-cell scalar attribute** (single value per address):
  `TextStyle` — see `crates/l123-core/src/text_style.rs`,
  `IronCalcEngine::used_cell_text_styles` (`ironcalc_adapter.rs:340`),
  `Engine::set_cell_text_style` (`engine.rs:132`),
  app.rs `cell_text_styles` HashMap (`app.rs:132`).
- **Per-cell enum** (mapped to IronCalc's `num_fmt`): `Format` — see
  `crates/l123-core/src/format.rs`, `used_cell_formats`
  (`ironcalc_adapter.rs:380`), `set_cell_format` (`engine.rs:141`).
- **Per-column scalar**: column width — `used_column_widths`
  (`ironcalc_adapter.rs:315`), `set_column_width` (`engine.rs:119`),
  app.rs `col_widths` (`app.rs:133`).

The xlsx load fan-out lives at `app.rs:2715–2723` and `app.rs:2936–2942`.
Every new accessor we add on `IronCalcEngine` needs a matching drain
call there. The save path lives at `app.rs:2989` (and the Xtract path at
`app.rs:3551`); every UI-side state map needs a push back into the
engine before save.

### Layering rules to honor (CLAUDE.md §“Crate layering”)

- IronCalc types **never** appear in `l123-core` or above `l123-engine`.
- New domain types (alignment, border, fill, merge, etc.) are
  zero-dependency Rust enums/structs in `l123-core`.
- The adapter (`ironcalc_adapter.rs`) is the only place that imports
  `ironcalc_base::*`.
- Conversions between L123 types and IronCalc types live in private
  helpers next to the accessor methods, not in `l123-core`.

### Strict TDD (CLAUDE.md §“red / green / refactor”)

For each feature below:

1. Write the failing **engine integration test** under
   `#[cfg(test)] mod tests` in `ironcalc_adapter.rs` that loads a
   fixture XLSX and asserts the new accessor returns the expected data.
2. Write the failing **unit test** on the new core type (round-trip,
   merge, marker rendering, etc.).
3. Write the failing **acceptance transcript** under
   `tests/acceptance/xlsx_<feature>.tsv` driving load → assert → save →
   reload → assert.
4. Implement until green. No speculative helpers.
5. Refactor with green tests. Clippy clean
   (`cargo clippy --workspace --all-targets -- -D warnings`).

### Fixtures (convention to establish — not yet in the repo)

Today the only test fixture in the repo is
`tests/acceptance/fixtures/m4_import.csv`, referenced from
`tests/acceptance/M4_file_retrieve_csv.tsv:23` and
`tests/acceptance/M4_file_import_numbers.tsv:19`. Acceptance transcripts
read fixture paths relative to the workspace root (`tests/acceptance/
fixtures/...`).

This plan needs XLSX fixtures, none of which exist yet. Establish the
convention as a one-time setup step before starting §2.1:

1. **Where to put them.** New subdirectory `tests/acceptance/fixtures/
   xlsx/` (sibling of `m4_import.csv`). Each fixture file is named for
   the feature it exercises: `alignment.xlsx`, `borders.xlsx`,
   `merges.xlsx`, etc.
2. **How to build them.** Authoring `.xlsx` by hand in Excel and
   committing the binary makes the fixture opaque and hard to update.
   Recommend a reproducible build script — pick *one* of:
   - **(a) Python + openpyxl** at `tests/acceptance/fixtures/xlsx/
     build.py`. Pro: openpyxl is the standard, scripts are short. Con:
     adds a non-Rust dev-time dep that contributors must install.
   - **(b) Rust + IronCalc** at `tests/acceptance/fixtures/xlsx/build.
     rs` or as a small `--bin` in `l123-engine`. Pro: keeps the
     toolchain pure-Rust. Con: round-trips through the same library
     we're testing, so a fixture can't catch IronCalc bugs that
     symmetrically affect both read and write.
   - **(c) Borrow from upstream.** IronCalc's `base/tests/` directory
     has hand-curated XLSX files. Where one fits a feature, copy it
     and document the provenance in a `README.md` next to the
     fixtures.

   No decision is forced now — the first feature to land (§2.1
   alignment) picks the approach and §6 row 1 inherits the cost.
3. **Commit policy.** If you adopt option (a) or (b), the script is
   the source of truth; the `.xlsx` binaries are also committed (so
   tests don't require Python or a separate build step) but treated
   as generated. Add a CI check that re-running the script produces
   byte-identical fixtures, or accept drift and document.

When this plan refers below to e.g. `alignment.xlsx`, read it as
`tests/acceptance/fixtures/xlsx/alignment.xlsx` once the directory
exists.

---

## 2. Tier 1 — High-value, IronCalc supports, L123 drops

These are the features users will miss first when they open an Excel
workbook in L123. Order within the tier reflects implementation order:
each item builds on conventions established by the previous one.

### 2.1 Cell alignment (horizontal, vertical, wrap-text)

**IronCalc source**
- `base/src/types.rs` — `Alignment` struct (~line 549) with
  `horizontal`, `vertical`, `wrap_text`, `text_rotation`, `indent`
  fields.
- Imported via `base/src/import/styles.rs`; lives on `CellXfs[i].alignment`.
- Read with `model.workbook.styles.get_style(idx).alignment`.

**L123 work**

1. `crates/l123-core/src/alignment.rs` — new file:
   ```rust
   pub enum HAlign { General, Left, Center, Right, Fill, Justify, CenterAcross }
   pub enum VAlign { Top, Center, Bottom }
   pub struct Alignment {
       pub horizontal: HAlign,    // General == default
       pub vertical: VAlign,      // Bottom == default
       pub wrap_text: bool,
   }
   ```
   Derive `Default`, `Copy`, `Clone`, `Debug`, `PartialEq`, `Eq`. Map
   IronCalc's `general | left | center | right | fill | justify |
   centerContinuous` strings explicitly; unknown values fall back to
   `General` (no panics on weird inputs).
2. `l123-core/src/lib.rs` — re-export `Alignment`, `HAlign`, `VAlign`.
3. `l123-engine/src/ironcalc_adapter.rs`:
   - `pub fn used_cell_alignments(&self) -> Vec<(Address, Alignment)>`
     mirroring `used_cell_text_styles` shape — only emit non-default.
4. `l123-engine/src/engine.rs`:
   - `fn set_cell_alignment(&mut self, addr: Address, a: Alignment) -> Result<()>`
     default-stub that returns `Unsupported`; implement on
     `IronCalcEngine` by mutating the cell's `Alignment` on its style.
5. `l123-ui/src/app.rs`:
   - Add `cell_alignments: HashMap<Address, Alignment>` to
     `WorkbookState`.
   - Drain in xlsx-load fan-out (`app.rs:~2940`).
   - Push back to engine before save (`app.rs:~2989`).
   - Apply during cell render: respect horizontal alignment in the grid
     (overrides 1-2-3 default which is left for labels, right for
     numbers). Wrap-text deferred — see §2.1.1 below.
6. Tests (red-first):
   - `alignment.rs` unit tests: default, parse-from-str, round-trip.
   - `ironcalc_adapter.rs` integration: load fixture
     `tests/acceptance/fixtures/xlsx/alignment.xlsx` (cells with each
     h/v combo);
     assert vector contents.
   - `tests/acceptance/xlsx_alignment.tsv`: `/File Retrieve` the fixture;
     assert grid renders A1 right-aligned text, B1 center, etc.

**Edge cases & non-goals**
- *Wrap text* (§2.1.1): rendering wrapped text requires multi-line cell
  display in the grid. Park this as a follow-up; for now, parse and
  preserve on save but render single-line.
- *Text rotation* and *indent* are read-but-not-rendered for v1.
- *Center-across-selection* (`centerContinuous`) is parsed; rendering
  defers to a later pass since it requires cross-cell layout.

**Effort:** medium (~1 day for parse + read + render + transcripts).

---

### 2.2 Cell borders

**IronCalc source**
- `base/src/types.rs:657` — `Border` and `BorderItem` structs. Fields:
  `left`, `right`, `top`, `bottom`, `diagonal_up`, `diagonal_down`. Each
  `BorderItem` is `{ style: BorderStyle, color: Option<String> }`.
- `BorderStyle` enum: None, Thin, Medium, Thick, Double, Dashed, Dotted,
  DashDot, DashDotDot, Hair, MediumDashed, MediumDashDot,
  MediumDashDotDot, SlantDashDot.
- `base/src/import/styles.rs` populates `CellXfs[i].border_id`; resolve
  via the styles table.

**L123 work**

1. `l123-core/src/border.rs`:
   ```rust
   pub enum BorderStyle { None, Hairline, Thin, Medium, Thick, Double,
                          Dashed, Dotted }
   pub struct BorderEdge {
       pub style: BorderStyle,
       pub color: Option<RgbColor>,   // see §2.3 for RgbColor
   }
   pub struct Border {
       pub left: BorderEdge,
       pub right: BorderEdge,
       pub top: BorderEdge,
       pub bottom: BorderEdge,
   }
   ```
   Map IronCalc’s 14 border styles down to L123's 8 (collapse
   dash-variants into `Dashed`, dot-variants into `Dotted`). Round-trip
   loses fidelity by design; document it in a doc comment on the enum.
   Skip diagonals for v1 (note in struct).
2. `l123-engine` — `used_cell_borders()` accessor + `set_cell_border()`
   trait method.
3. `l123-ui` — `cell_borders: HashMap<Address, Border>` in
   `WorkbookState`; render via ratatui's `Block` borders or by drawing
   into the cell's own characters. Borders sit *between* cells; the
   grid renderer needs to consult both the cell on each side of each
   edge and pick the heavier of the two (Excel's tie-break rule:
   later-applied wins, but on import we only see the result).
4. Tests:
   - Unit: BorderStyle round-trip; “heavier wins” edge merge.
   - Integration: load `borders.xlsx` (a cell with all four sides at
     varying styles); assert structure.
   - Acceptance: `tests/acceptance/xlsx_borders.tsv` with a snapshot of
     the rendered grid showing box-drawing characters.

**Edge cases**
- *Diagonal borders* — out of scope for v1; preserve on round-trip via a
  raw passthrough string stored alongside the L123 `Border` (or accept
  data loss and document it).
- *Border-of-empty-cell* — Excel allows borders on cells with no value.
  Our state map must allow `Address` → `Border` without a corresponding
  value entry.
- *Rendering interaction with cell selection highlight* — the cursor
  highlight should still be visible; borders draw underneath the
  inverted selection.

**Effort:** large (~2–3 days). Rendering is the hard part, not the data
model. Recommend prototyping with a single-edge case first (just bottom
borders) before generalizing.

---

### 2.3 Cell fill / background color

**IronCalc source**
- `base/src/types.rs:460` — `Fill` struct with `pattern_type`,
  `fg_color`, `bg_color`. Pattern types: `none`, `solid`, `gray125`,
  ... (≈18 patterns). 95% of real spreadsheets use `solid`.
- Colors are Excel theme-or-RGB. IronCalc resolves theme→RGB on import
  and stores hex strings like `"FFCCEEFF"` (ARGB).

**L123 work**

1. `l123-core/src/color.rs`:
   ```rust
   pub struct RgbColor { pub r: u8, pub g: u8, pub b: u8 }
   impl RgbColor {
       pub fn from_hex(s: &str) -> Option<Self>;  // accepts RGB or ARGB
       pub fn to_hex(&self) -> String;
   }
   ```
   No alpha channel — terminal cells can't render transparency. Drop
   alpha during import.
2. `l123-core/src/fill.rs`:
   ```rust
   pub struct Fill {
       pub bg: Option<RgbColor>,   // None == no fill
       pub fg: Option<RgbColor>,   // pattern fg, used only when patterned
       pub pattern: FillPattern,   // Solid by default
   }
   pub enum FillPattern { Solid, None, /* others mapped to Solid for v1 */ }
   ```
3. `l123-engine` — `used_cell_fills()`, `set_cell_fill()`.
4. `l123-ui` — `cell_fills: HashMap<Address, Fill>`; render via
   `ratatui::Style::bg`. RGB → nearest 256-color, with truecolor
   passthrough when terminal advertises it (check `COLORTERM`).
5. Tests as per pattern.

**Color-quantization detail.** Many terminals lie about truecolor
support. Default to the conservative ANSI-256 quantization; offer a
`--truecolor` CLI flag (later moved to L123.CNF per `docs/CONFIG.md`)
to opt in. Quantization logic goes in `l123-ui` since it depends on
the rendering target — keep `RgbColor` pure in `l123-core`.

**Edge cases**
- *Theme colors that change with workbook theme*: IronCalc resolves
  these to concrete RGB at import. We don't preserve the theme link;
  document as round-trip lossy.
- *Patterned fills (gray125 etc.)*: Map to `Solid` with an RGB blend
  for v1. Document the loss.
- *Conditional formatting fills*: out of scope (IronCalc doesn't parse
  CF — see Tier 4).

**Effort:** medium (~1.5 days). The quantization function deserves its
own unit test fixture (a table of (RGB, expected ANSI-256) pairs).

---

### 2.4 Merged cells

**IronCalc source**
- `base/src/worksheets.rs:178` — `load_merge_cells()`. Stored as
  `merge_cells: Vec<String>` on `Worksheet`, e.g. `"K7:L10"`.

**L123 work**

1. `l123-core/src/merge.rs`:
   ```rust
   pub struct Merge {
       pub anchor: Address,    // top-left
       pub end: Address,       // bottom-right
   }
   impl Merge {
       pub fn contains(&self, addr: Address) -> bool;
       pub fn is_anchor(&self, addr: Address) -> bool;
   }
   ```
2. `l123-engine` —
   `pub fn used_merged_cells(&self) -> Vec<(SheetId, Merge)>`. Parse
   IronCalc’s `"K7:L10"` strings via `Range::parse` or a thin local
   helper.
3. Trait additions: `set_merged_range`, `unset_merged_range` (for the
   eventual `/Range Unmerge` command — out of scope here, but the
   setter must exist for save round-trip).
4. `l123-ui`:
   - `merges: HashMap<SheetId, Vec<Merge>>` in `WorkbookState`.
   - **Renderer change** (the hard part): when drawing the grid, if a
     cell falls inside a merge but isn't the anchor, skip drawing it.
     The anchor cell renders across the union of column widths and row
     heights of the merge. Cursor navigation must skip non-anchor cells
     of a merge (cursor lands on the anchor only).
   - Selection model: selecting a merged cell selects the whole region.
5. Tests:
   - Unit: `Merge::contains`, `Merge::is_anchor`.
   - Integration: load `merges.xlsx`; assert vector of merges.
   - Acceptance: `xlsx_merges.tsv` — load, navigate with arrow keys,
     verify cursor jumps over the merge body; verify rendered text
     spans columns.

**Edge cases**
- *Overlapping merges*: Excel forbids them. Treat as malformed; log a
  warning and drop later merges.
- *Merge containing formulas*: Excel only stores the value at the
  anchor; other cells are blank. Verify our load path matches.
- *Insert/delete row/col through a merge*: out of scope for the import
  feature; revisit when `/Range Unmerge` lands.

**Effort:** large (~2–3 days). The renderer and cursor changes are
where the bugs hide. Write the acceptance transcript first.

---

### 2.5 Frozen panes

**IronCalc source**
- `base/src/worksheets.rs:646–699` — pane element parsing. Stored as
  `frozen_rows: i32` and `frozen_columns: i32` on `Worksheet`.

**L123 work**

1. `l123-engine`:
   - `pub fn frozen_panes(&self, sheet: SheetId) -> (u32, u16)` returning
     `(rows, cols)`. Default `(0, 0)` when not frozen.
   - Trait setter: `fn set_frozen_panes(&mut self, sheet: SheetId,
     rows: u32, cols: u16) -> Result<()>`.
2. `l123-core`: no new types; reuse primitive `(u32, u16)` keyed by
   `SheetId` in UI state.
3. `l123-ui`:
   - `frozen: HashMap<SheetId, (u32, u16)>` in `WorkbookState`.
   - Grid renderer splits the viewport into up to four regions:
     top-left frozen corner, top frozen row band, left frozen column
     band, scrolling main region. Scrolling only translates the main
     region.
   - This duplicates the L123 `/Worksheet Titles Both|Horizontal|
     Vertical` behavior; if that's already implemented, plug into the
     same render path. Otherwise, this implementation *is* the
     `/WT` implementation.
4. Tests:
   - Integration: load fixture with `frozen_rows=2, frozen_columns=1`;
     assert accessor.
   - Acceptance: `xlsx_frozen.tsv` — load, scroll right with arrow
     keys, verify column A stays put; scroll down, verify rows 1–2 stay
     put.

**Edge cases**
- *Split panes* (without freeze): IronCalc parses `xSplit/ySplit` but
  may store them in a different field; verify before relying on
  freeze-only semantics. v1 ignores split-without-freeze.
- *Per-view freezes*: IronCalc may store views per-user; we use the
  first view only.

**Effort:** large (~3 days) if `/Worksheet Titles` isn't yet
implemented; medium (~1 day) if we can plug into existing infrastructure.
Check `crates/l123-ui/src/app.rs` for any `titles_*` state before
starting.

---

### 2.6 Sheet tab color

**IronCalc source**
- `base/src/worksheets.rs:197` — `load_sheet_color()`. Stored as
  `color: Option<String>` (RGB hex) on `Worksheet`.

**L123 work**

1. `l123-engine`:
   - `pub fn sheet_color(&self, sheet: SheetId) -> Option<RgbColor>`.
   - `fn set_sheet_color(&mut self, sheet: SheetId, color: Option<RgbColor>) -> Result<()>`.
2. `l123-ui`:
   - Drain into `sheet_colors: HashMap<SheetId, RgbColor>` on the
     `WorkbookState`.
   - In the sheet-tab indicator (status line right side), tint the
     sheet letter with the color. This requires the ANSI-256
     quantization from §2.3 — gate this feature on §2.3 landing first.

**Effort:** small (~half day after §2.3 lands).

---

### 2.7 Cell comments / notes

**IronCalc source**
- `base/src/worksheets.rs:218` — `load_comments()`. Stored as
  `comments: Vec<Comment>` per worksheet. `Comment` has `text`,
  `author_name`, `cell_ref` (string like `"B5"`), and `author_id`.

**L123 work**

1. `l123-core/src/comment.rs`:
   ```rust
   pub struct Comment {
       pub addr: Address,
       pub author: String,
       pub text: String,
   }
   ```
2. `l123-engine`:
   - `pub fn used_comments(&self, sheet: SheetId) -> Vec<Comment>`.
   - `fn set_comment(&mut self, comment: Comment) -> Result<()>`,
     `fn delete_comment(&mut self, addr: Address) -> Result<()>`.
3. `l123-ui`:
   - `comments: HashMap<Address, Comment>` in `WorkbookState`.
   - Cell-corner indicator: a small character (`*` or `‘`) in the
     top-right of cells with comments, similar to Excel's red triangle.
   - **Viewer command**: `/Range Note Show` (1-2-3 R3.4a calls these
     "cell notes" — see SPEC §15 §"@notes" if mentioned). For v1 a
     simpler hook: F1 on a cell with a comment opens a panel showing
     `<author>: <text>`.
4. Tests:
   - Integration: load `comments.xlsx`; assert vector.
   - Acceptance: `xlsx_comments.tsv` — load, navigate to commented
     cell, hit F1, verify panel content.

**Edge cases**
- *Threaded comments* (Excel's modern reply-chain comments) — IronCalc
  parses the legacy `comments1.xml` only. Threaded comments live in
  `threadedComments.xml` and are not parsed. Document as known
  limitation.
- *Comment formatting* (rich text inside the comment) — flatten to
  plain text on import. Round-trip lossy.

**Effort:** medium (~1.5 days). Most of the work is the F1 panel UX.

---

## 3. Tier 2 — Mainstream, IronCalc supports, L123 drops

Implement after Tier 1. Each is smaller in scope but touches existing
state more invasively.

### 3.1 Extended font properties (size, color, name, family, scheme)

**IronCalc source**
- `base/src/styles.rs:81` — `Font` struct with `sz`, `name`, `color`,
  `family`, `scheme`, `strike`, `b`, `i`, `u`.

**Strategy**

The current `TextStyle` in `l123-core/src/text_style.rs` is a fixed
3-bit struct. Don't try to grow it in place — that breaks the existing
`{Bold Italic}` marker contract. Instead:

1. Introduce a parallel `FontStyle` struct in
   `l123-core/src/font_style.rs`:
   ```rust
   pub struct FontStyle {
       pub size: Option<u8>,        // points; None == default (10)
       pub color: Option<RgbColor>, // None == default (white)
       pub family: Option<String>,  // round-trip only; not rendered
       pub strike: bool,
   }
   ```
   `TextStyle` keeps its current bold/italic/underline role for the
   `:Format` command. `FontStyle` is the broader xlsx-only state.
2. Adapter accessors `used_font_styles()` / `set_font_style()`.
3. `l123-ui`:
   - `font_styles: HashMap<Address, FontStyle>` in `WorkbookState`.
   - Render `color` via foreground color in the grid (subject to the
     same ANSI-256 quantization as §2.3).
   - `family` and `size` are read-and-preserved-only; terminal cells
     are uniform-size. Document this as "round-trip preserves font
     name and size; rendering uses the terminal's font".
   - `strike` adds a fourth bit to control-panel marker:
     `{Bold Italic Strike}`. Update `TextStyle` *and* `FontStyle` to
     coordinate, or fold strike into `TextStyle` (recommend the
     latter: it's a single bit and matches the existing pattern).

**Effort:** medium (~1.5 days). Worth it because color is a high-impact
visual cue.

---

### 3.2 Tables and autofilter metadata

**IronCalc source**
- `base/src/tables.rs` — `Table` struct: name, range, columns
  (each with name, optional totals function), header/totals row flags,
  filter metadata.

**Strategy**

L123's MENU.md lists `/Data Query` and `/Data External` (for tables) —
review whether those should be the surface for displayed tables. For
v1: parse and preserve on round-trip; do not render special UI yet.

1. `l123-core/src/table.rs` — mirror the IronCalc `Table` structure
   minus the IronCalc-specific bits.
2. Adapter: `used_tables(&self) -> Vec<(SheetId, Table)>`,
   `set_table(...)`.
3. `l123-ui`: `tables: HashMap<SheetId, Vec<Table>>` in
   `WorkbookState`. No render changes for v1; the data sits in state
   waiting for a future `/Data Query Define` integration. Add a debug
   command (`/Data Query Tables` or similar) that lists table names —
   this proves the round-trip works.

**Effort:** medium (~1 day for parse + preserve; weeks if we want
real filter UI — out of scope here).

---

### 3.3 Sheet visibility (hidden, very-hidden)

**IronCalc source**
- `base/src/types.rs:68` — `enum SheetState { Visible, Hidden, VeryHidden }`.

**Strategy**

1. `l123-core/src/sheet_state.rs`:
   ```rust
   pub enum SheetState { Visible, Hidden, VeryHidden }
   ```
2. Adapter: `sheet_state(sheet)` / `set_sheet_state(sheet, state)`.
3. `l123-ui`:
   - `sheet_states: HashMap<SheetId, SheetState>` in `WorkbookState`.
   - Sheet-tab indicator skips non-visible sheets in the bar.
   - `Ctrl-PgUp/PgDn` navigation skips them.
   - Add a `/Worksheet Hide`/`/Worksheet Unhide` command (already in
     MENU.md? — check). Without an unhide UX, hidden sheets are
     orphaned forever.
4. Tests:
   - Acceptance: load workbook with one hidden sheet, assert
     `Ctrl-PgDn` skips it; unhide, assert it appears.

**Effort:** small (~half day).

---

## 4. Tier 3 — Niche, IronCalc supports, L123 drops

These are lower-priority. Plan stubs only — implement on demand.

### 4.1 Row heights

**IronCalc source**: `base/src/types.rs:128` — `Row` struct with
`custom_height` and `height` (points).

**Strategy**: mirror `col_widths` exactly. Add `row_heights:
HashMap<(SheetId, u32), u8>` keyed by sheet+row. Convert points →
terminal lines (heuristic: 15pt ≈ 1 line). `/Worksheet Row Set-Height`
in MENU.md is the L123 surface. Renderer must honor variable row
heights — non-trivial when combined with merged cells (§2.4).

**Effort:** medium (~2 days, mostly renderer).

### 4.2 Dynamic arrays / spill ranges

**IronCalc source**: `base/src/worksheets.rs:332` — `CellArrayKind`,
`base/src/types.rs:240` — `ArrayFormula`, `SpillCell`.

**Strategy**: render is mostly automatic since IronCalc evaluates and
fills spill cells. The gap is *visual*: no indication of which cells
are spill members vs. user-entered. Add a per-cell "is spilled" bit;
render with a subtle dim style or italics.

**Effort:** small (~half day, after the engine exposes `is_spilled`
per cell).

### 4.3 Strikethrough

Fold into §3.1's `FontStyle::strike`. No separate work.

### 4.4 Cell protection / locked status

**Strategy**: parse `apply_protection` and `locked` flags on cell
styles; preserve on round-trip; do not enforce. Real sheet-protection
needs SPEC additions and a password-prompt UX — punt entirely.

**Effort:** small to preserve; large to enforce. Recommend
preserve-only for now.

### 4.5 Workbook metadata (creator, app, timestamps)

**Strategy**: add a `Metadata` struct in `l123-core`; expose via
adapter; surface in a `/Worksheet Status` info panel. Update
`modified` on save; preserve `created` from import.

**Effort:** small (~half day).

### 4.6 Workbook views

**Strategy**: persist `selected_sheet` and per-sheet selection on
save so reopening lands the user where they left off. Window
geometry is meaningless in a TUI — drop it.

**Effort:** small (~half day).

### 4.7 Gradient fills, advanced patterns

IronCalc itself flattens gradients to solid fallback. No L123 work
beyond §2.3 — we get whatever IronCalc gives us.

---

## 5. Tier 4 — IronCalc doesn't support (upstream gap)

These are not L123 work in the normal sense. To support them we'd need
to patch IronCalc itself.

| Feature | Why it matters | Path forward |
|---|---|---|
| Conditional formatting | Common in real spreadsheets; rules drive cell colors | Either (a) wait for upstream, (b) fork IronCalc and add a parser for `<conditionalFormatting>` blocks, or (c) parse it ourselves in `l123-io` outside of IronCalc and apply post-load |
| Data validation | Drop-down lists, numeric ranges | Same options as CF; option (c) is most realistic |
| Charts | Visual artifacts; SPEC §M7 envisions L123-side graphs anyway | Out of scope for *import*; users re-create via `/Graph` |
| Images / embedded objects | Rare in workflow spreadsheets | Out of scope indefinitely |
| Pivot tables | Common but complex | Way out of scope |

### Recommended approach for Tier 4

Don't take any of these on until Tier 1 and Tier 2 are done. When the
demand is real:

1. **Conditional formatting** is the highest-value Tier 4 item. Plan to
   add it via a parallel parser in `l123-io/src/xlsx_extras.rs` that
   opens the xlsx as a zip, reads `xl/worksheets/sheetN.xml`, and
   extracts `<conditionalFormatting>` rules. Apply them as a
   post-processing pass over the engine's cell view at render time.
   Estimated effort: **2–3 weeks**.
2. **Data validation** can use the same parallel-parser approach.
   Estimated: **1 week** for the dropdown-list case (most common).

---

## 6. Implementation order (recommended)

Tackle features in this order. Each row is one PR.

| # | Tier | Feature | Why this order | Approx. effort |
|---|---|---|---|---|
| 1 | 1 | Cell alignment (§2.1) | Smallest Tier 1; establishes pattern | 1 day |
| 2 | 1 | Sheet tab color (§2.6) | Tiny, but needs §2.3 RgbColor — defer until #4 | (defer) |
| 3 | 1 | Cell fill color (§2.3) | Unlocks tab color and font color | 1.5 days |
| 4 | 1 | Sheet tab color (§2.6) | Now that RgbColor exists | 0.5 day |
| 5 | 2 | Extended font (color, size) (§3.1) | Same RgbColor infrastructure | 1.5 days |
| 6 | 1 | Cell borders (§2.2) | Hardest renderer change; do after color is settled | 2–3 days |
| 7 | 1 | Comments (§2.7) | Independent of styling work | 1.5 days |
| 8 | 1 | Merged cells (§2.4) | Touches grid renderer + cursor model | 2–3 days |
| 9 | 1 | Frozen panes (§2.5) | Touches grid scrolling | 2–3 days |
| 10 | 2 | Sheet visibility (§3.3) | Quick win | 0.5 day |
| 11 | 2 | Tables/autofilter (§3.2) | Round-trip preserve only | 1 day |
| 12 | 3 | Row heights (§4.1) | After merge + freeze land | 2 days |
| 13 | 3 | Spill rendering (§4.2) | Cosmetic | 0.5 day |
| 14 | 3 | Metadata + views (§4.5, §4.6) | Quick wins | 1 day |

**Total Tier 1+2:** ≈ 17 person-days.
**With Tier 3:** ≈ 21 person-days.

---

## 7. Cross-cutting concerns

### 7.1 The grid renderer is the bottleneck

Most Tier 1 features (borders, fills, merges, freezes, alignment) all
land in the same code path: the grid render. Plan to **refactor the
grid renderer once** to accept a per-cell `CellStyle` aggregate
(alignment + border + fill + font), rather than threading each new
attribute through ad-hoc parameters. The refactor is the “refactor”
step of the first Tier 1 feature; subsequent features slot in.

Suggested aggregate type in `l123-core/src/cell_render.rs` (extending
the existing module):

```rust
pub struct CellStyle {
    pub alignment: Alignment,
    pub border: Border,
    pub fill: Fill,
    pub font: FontStyle,
    pub text: TextStyle,    // bold/italic/underline (existing)
    pub format: Format,      // number format (existing)
}
```

`CellStyle::default()` is the no-style cell. App-side state becomes a
single `cell_styles: HashMap<Address, CellStyle>` rather than five
parallel maps. Migration: lazy — build the aggregate at render time
from the existing maps until each one is migrated in turn.

### 7.2 Round-trip fidelity invariant

A workbook authored in Excel, opened in L123, modified by zero
operations, and saved back, must produce a file that Excel reads
identically (or with documented losses).

Add an integration test `tests/round_trip_no_op.rs` that:

1. Loads a fixture of every supported feature.
2. Saves to a temp path with no edits.
3. Reloads.
4. Asserts every accessor returns identical data.

Run this for every Tier 1 and Tier 2 feature added.

### 7.3 Color handling is shared infra

Don't reimplement RGB → ANSI quantization per feature. Settle on one
function in `l123-ui/src/color.rs` and reuse it for fills, fonts, and
sheet-tab color. Unit-test it against a fixed table of color pairs.

### 7.4 IronCalc API drift

When upgrading IronCalc, re-run all engine integration tests. The
adapter is the only place that touches IronCalc types — if a field
rename breaks a test, the fix is local.

---

## 8. Definition of done — XLSX import parity v1

Tier 1 and Tier 2 are considered done when:

1. All accessors and setters listed in §2 and §3 are implemented and
   covered by engine integration tests.
2. Every feature has an acceptance transcript under
   `tests/acceptance/xlsx_<feature>.tsv` that loads a fixture and
   asserts visible behavior.
3. The round-trip-no-op test from §7.2 passes for every supported
   feature.
4. `cargo clippy --workspace --all-targets -- -D warnings` is clean.
5. `docs/SPEC.md` is updated to enumerate which xlsx features L123
   preserves vs. drops.

Tier 3 is purely opportunistic — done when individual features ship.

Tier 4 is out of scope for this plan.

---

## 9. Open questions

Resolve before starting:

1. **Aggregate `CellStyle` migration timing.** Refactor the grid
   renderer up front (one big PR, no behavior change), or migrate
   incrementally as each feature lands? Recommend up front: the
   incremental path produces five HashMap fields and a tangled render
   loop.
2. **`FontStyle` vs. `TextStyle` separation.** Keep them separate (as
   recommended in §3.1) or fold into one struct? The 1-2-3
   `:Format Bold` command needs only the three bits; xlsx needs more.
   Recommend separate, fold into `CellStyle` aggregate.
3. **ANSI-256 quantization tolerance.** What's an acceptable color
   error? Pick a numeric threshold (e.g. CIE76 ΔE < 10) and assert it
   in the quantization unit tests.
4. **Comment UX surface.** F1-on-commented-cell, or a status-bar
   indicator + dedicated `/Range Note Show` command? Pick one before
   §2.7. Recommend status-bar indicator + show-on-cursor as MVP; full
   note management menu later.
5. **Row height units.** Excel points → terminal lines is lossy. What
   ratio? Pick by experiment with a fixture; document the conversion
   constant in `l123-core/src/cell_render.rs`.

---

## 10. References

- IronCalc source: `/Users/ddmoore/dev/IronCalc/base/src/`
  - `types.rs` — Alignment, Border, Fill, Font, Row, ArrayFormula
  - `worksheets.rs` — load_merge_cells, load_sheet_color, load_comments,
    pane parsing
  - `styles.rs` — Font, CellXfs, the styles table
  - `tables.rs` — Table struct
- L123 wiring points:
  - `crates/l123-engine/src/ironcalc_adapter.rs` — accessors live here
  - `crates/l123-engine/src/engine.rs` — trait additions
  - `crates/l123-ui/src/app.rs:128–170` — `WorkbookState` field
    declarations
  - `crates/l123-ui/src/app.rs:2700–2950` — xlsx load/save fan-out
- Existing reference impls to copy from:
  - `TextStyle` (per-cell scalar attribute)
  - `Format` (per-cell enum mapped to Excel string)
  - column widths (per-column scalar)

---

## 11. How to use this document in a fresh session

1. Pick a feature from §2, §3, or §4. Open this file and the
   referenced IronCalc source paths.
2. Re-verify the IronCalc API paths still match (struct field
   names, function locations).
3. Write the failing tests in this order: unit on the new core type,
   adapter integration test, acceptance transcript.
4. Run `cargo test -p l123-engine` and confirm the integration test
   fails for the right reason (missing accessor, not a compile error).
5. Implement bottom-up: core type → adapter accessor → trait method →
   UI state field → drain into state on load → render → push back on
   save.
6. Run `cargo test --workspace` and `cargo clippy --workspace
   --all-targets -- -D warnings`. Both green = ready to commit.
7. Update §6 of this doc with the actual effort vs. estimate so future
   estimates are calibrated.
