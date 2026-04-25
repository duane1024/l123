# L123 — Lotus 1-2-3 TUI Spreadsheet (Spec v0.3)

> v0.3 retargets the project from Lotus 1-2-3 Release 3.1 (1990) to
> **Release 3.4a for DOS** (1993) — the final DOS release, with the
> WYSIWYG add-in promoted to a standard, always-on feature and the R3.4
> icon panel. v0.2 was rewritten from v0.1 after deep review of the
> *Lotus 1-2-3 Release 3.1 Reference*, *Tutorial*, and *Quick Reference*
> manuals (1990). Corrections vs. v0.1 are called out inline as **Δ**.

---

## 1. What L123 is

A terminal spreadsheet whose **interaction model is indistinguishable from
Lotus 1-2-3 Release 3.4a for DOS** (1993), but whose compute and I/O layers are modern:
xlsx-compatible, Unicode-correct, cross-platform.

**Two promises — the project fails if either breaks:**

1. An experienced 1-2-3 R3.4a user sitting down cold can drive L123 without
   reading anything.
2. L123 files round-trip cleanly to and from `.xlsx` via a modern formula
   engine.

## 2. What L123 is NOT

- A visual homage. The goal is *functional* fidelity, not CRT nostalgia (that
  is a stretch goal, §18).
- A DOS emulator. No INT21h, no 8.3 filenames, no code pages. Strings are
  UTF-8 end-to-end. 1-2-3's LMBCS character set is a compatibility input only.
- A reimplementation of the 1-2-3 compute core. We delegate to IronCalc.
- A macro player for existing `.WK3` files (the macro language is in scope
  per §18; replaying macros embedded in legacy `.WK3` files is not). Note:
  read-only `.WK3` import — values, formulas, basic styles, column widths —
  *is* shipped via `ironcalc_lotus`; see §14.

---

## 3. Stack decision

| Layer | Choice | Why |
|---|---|---|
| Language | Rust (stable) | Ecosystem fit for both TUI and engine |
| TUI | **ratatui** + **crossterm** | De facto Rust TUI stack; custom grid widget required |
| Engine | **IronCalc 0.7.x** | Native `.xlsx` round-trip; row/col/move ops shipped in v0.6; active TUI reference in `TironCalc` |
| `.xls` read | `calamine` | Only practical option, read-only |
| `.wk3` read | `ironcalc_lotus` (sibling crate to `ironcalc`/`ironcalc_base`) | Read-only; values, formulas, basic styles, col widths |
| CSV | `csv` crate | Standard |

**Δ from v0.1:** IronCalc is confirmed as primary (not "or Formualizer");
Formualizer is a fallback kept on the bench only. `set_user_input` is the
setter we build around.

**Engine gap we must paper over ourselves:**

- IronCalc has no undo/redo → we maintain a command journal (also the
  foundation for `/Worksheet Learn` macro recording).
- `set_user_input` couples parse+mutation → we wrap it with a routing shim
  that respects 1-2-3's value-vs-label first-char rule and label prefixes.
- `.wk3` read lives in `ironcalc_lotus` (read-only; sibling to
  `ironcalc_base`/`ironcalc`). l123-engine's adapter wires it in via the
  `Engine::load_wk3` slot.

---

## 4. Screen anatomy

The screen is divided into five fixed zones. Matching the layout exactly is
load-bearing for authenticity.

```
┌──────────────────────────────────────────────────────────────────┐
│ A:B5: (C2) +B3*1.08                                    READY     │ ← control panel line 1
│ 12,960.00                                                         │ ← control panel line 2
│ Worksheet Range Copy Move File Print Graph Data System Add-In Q… │ ← control panel line 3
├──┬────────┬────────┬────────┬────────┬────────┬────────┬─────────┤
│  │   A    │   B    │   C    │   D    │   E    │   F    │   G     │ ← column frame
├──┼────────┼────────┼────────┼────────┼────────┼────────┼─────────┤
│ 1│Revenue │ Q1     │ Q2     │ Q3     │ Q4     │ Total  │         │
│ 2│────────│────────│────────│────────│────────│────────│         │
│ 3│Sales   │  12000 │  14500 │  16200 │  17800 │  60500 │         │
│ …│        │        │        │        │        │        │         │
├──┴────────┴────────┴────────┴────────┴────────┴────────┴─────────┤
│ INC3.WK3          23-Apr-26 14:02          CAPS  NUM  UNDO  CALC  │ ← status line
└──────────────────────────────────────────────────────────────────┘
```

### Control panel — three lines (load-bearing)

| Line | Contents |
|---|---|
| 1 | `<sheet>:<addr>: (<fmt>) <contents>` on the left; **mode indicator** right-justified |
| 2 | In READY: blank. In LABEL/VALUE/EDIT: the entry buffer with cursor. In MENU: the current menu's items, horizontally, highlighted item in reverse video. In FILES/NAMES: list row. |
| 3 | In READY: blank. In MENU on a highlighted item: either (a) the next submenu previewed, or (b) a one-line help string describing what that menu item does. In a command prompt: `Enter range to format: B5..B5`. In data entry: in-progress expression. |

**Δ from v0.1:** v0.1 collapsed this to a single "formula/input line". The
three-line split and line-3 preview are *the* defining Lotus UI affordance;
omitting them breaks the contract.

### Column/row frame

- Column letters above (`A`..`IV`, with `A:`..`IV:` sheet prefix shown only
  when multi-sheet or in POINT across sheets).
- Row numbers on the left, right-aligned. Rows go to 8192.
- Cell pointer is reverse-video over the cell.
- POINT-mode highlight extends in reverse video from the anchor to the pointer.

### Status line (bottom)

- Left: current filename, or `*` followed by default when unsaved.
- Centre: clock/date (configurable: `International`, `Standard`, `None`,
  `Filename`).
- Right: status indicators — stacked horizontally, only the active ones shown.

---

## 5. Mode model

Every mode is a first-class state with a visible indicator.

| Mode | Indicator | When |
|---|---|---|
| READY | `READY` | Accepting navigation, `/`, or first keystroke of entry |
| LABEL | `LABEL` | First char was a label-starter or had a label prefix |
| VALUE | `VALUE` | First char was `0-9 + - . ( @ # $` |
| EDIT | `EDIT` | F2 pressed, or a bad entry was refused |
| POINT | `POINT` | Pointing at a cell/range (during formula or command) |
| MENU | `MENU` | A slash menu is active |
| FILES | `FILES` | A file list is shown (File Retrieve, etc.) |
| NAMES | `NAMES` | A name list is shown (F3) |
| HELP | `HELP` | Help overlay |
| ERROR | `ERROR` | An error message is visible; waits for ESC/ENTER |
| WAIT | `WAIT` | Long operation (retrieve, save, recalc) |
| FIND | `FIND` | /Data Query Find is highlighting a record |
| STAT | `STAT` | /Worksheet Status or /WGD Status panel shown |

**Δ from v0.1:** v0.1 listed 4 modes; the real model has 13. These all must be
implemented for authenticity; FIND/STAT are late-stage but the others are
MVP-adjacent. v0.1's `Enter mode` is split into `LABEL` and `VALUE` per
first-char rule — this is not cosmetic; it controls whether label prefixes
apply.

### Mode transitions (MVP-relevant)

```
READY --/--> MENU --letter/arrow+Enter--> (submenu MENU | prompt)
READY --<value-char>--> VALUE --<Enter>--> READY
READY --<label-char>--> LABEL --<Enter>--> READY
READY --F2--> EDIT --<Enter>--> READY
READY --F5--> (prompt in line 3) --<Enter>--> READY  (goto)
VALUE --<cell-ref-char-while-at-operator>--> POINT --<.>--> POINT(anchored)
<any prompt expecting range> --> POINT (auto-anchored)
POINT --<Esc>--> (unanchored, cursor freely movable)
POINT --<Esc while unanchored>--> back to the command's previous step
<any> --<Ctrl-Break>--> READY
<any menu/prompt> --<Esc>--> one level up
```

---

## 6. Status indicators (bottom line)

Show only when active. Lowercase in this table; displayed in uppercase.

| Indicator | Meaning |
|---|---|
| CALC | Recalc pending (blue) or in progress (red) — F9 triggers |
| CAPS | CAPS LOCK |
| NUM | NUM LOCK |
| SCROLL | SCROLL LOCK: arrows scroll viewport, not pointer |
| OVR | INS toggled to overstrike in entry modes |
| END | END pressed; next arrow jumps to edge of data |
| FILE | Ctrl-End pressed; next arrow navigates between active files |
| CIRC | Workbook has a circular reference |
| MEM | Free memory low |
| CMD | A macro is running |
| PRT | Printing |
| RO | File read-only |
| GROUP | File is in GROUP mode (3D edits propagate) |
| LEARN | Recording keystrokes (Alt-F5) |
| STEP | Macro STEP mode (Alt-F2) |
| SST | Macro running in STEP |
| ZOOM | Current window enlarged (Alt-F6) |
| UNDO | Undo enabled (on when `/WGD Other Undo Enable`) |

MVP needs: CALC, CAPS, NUM, OVR, END, CIRC, UNDO.
Post-MVP adds the rest as their features land.

---

## 7. Function key map

Exactly matches R3.4a. Keys in *italic* are MVP.

| Key | Plain | Alt |
|---|---|---|
| F1 | *HELP* (context) | COMPOSE |
| F2 | *EDIT* | RECORD |
| F3 | *NAME* (list names at prompts) | RUN |
| F4 | *ABS* (cycle $ in reference at EDIT/entry) | *UNDO* |
| F5 | *GOTO* | LEARN |
| F6 | WINDOW (cycle split) | ZOOM |
| F7 | QUERY (repeat last /Data Query) | APP1 |
| F8 | TABLE (repeat last /Data Table) | APP2 |
| F9 | *CALC* (recalc; in EDIT: value-freeze) | APP3 |
| F10 | GRAPH (current graph full-screen) | ADDIN |

Shift-keys: Tab = right; Shift-Tab = left. Ctrl-←/→ = BigLeft/BigRight (one
screen). Ctrl-PgUp/PgDn = prev/next sheet. Ctrl-Home = A:A1 of current file.
End + Ctrl-Home = last non-blank cell. Ctrl-End then arrows/PgUp = file
navigation. Ctrl-Break = abort.

**Δ:** v0.1 listed `Ctrl+C: Quit`. 1-2-3 uses `/QY`, not Ctrl-C. Ctrl-C is
unused. Ctrl-Break is the abort key.

---

## 8. Input model

### First-character rule

The first character typed in READY decides LABEL vs VALUE:

- **Value starters**: `0-9  +  -  .  (  @  #  $` (when `$` is the currency)
- Everything else → LABEL; a `'` (or user-configured default) prefix is
  auto-inserted.

### Label prefixes

| Prefix | Behavior |
|---|---|
| `'` | Left-align (default) |
| `"` | Right-align |
| `^` | Center |
| `\` | Repeat the label across the cell's width (`\-` → dashes) |
| `\|` | Non-print row (first col only) — deferred to Print milestone |

### Commit keys

- `Enter` — commits, pointer stays put.
- Arrow/Tab — commits and moves one cell in that direction.

### Cursor conveniences

- Typing while LABEL/VALUE: the entry buffer is on line 2; cursor is a single
  char.
- F2 EDIT mode: full line editing; Ctrl-← / Ctrl-→ jump word; Home/End jump
  line-ends; F4 cycles absoluteness of the reference the cursor is on.

---

## 9. POINT mode semantics

POINT is a mode, not a state machine hack. Any input that expects a cell or
range enters POINT and auto-anchors at the current pointer. Rules:

1. The range highlight extends from anchor to pointer, inclusive.
2. `.` (period) rotates the free corner through the 4 corners of the range
   (cycling which corner moves with the arrows).
3. `Esc` unanchors — pointer becomes a single moving cell without growing the
   range. Second `Esc` cancels out to the command's previous prompt.
4. Typing a range address replaces the highlight: e.g. `c8..d12` in the buffer
   replaces the pointed range.
5. `F3` pops up a list of named ranges / named graphs / files (context).
6. `F5` goes to a cell; pointer moves, highlight follows.
7. `Enter` commits.

**Δ:** v0.1 had "Arrow keys expand selection" as the full spec. The anchor
rules above must be implemented exactly or range selection will feel wrong.

---

## 10. Menu tree

The full tree is in `docs/MENU.md` (to be generated from reference research).
This spec lists the **MVP slice** (fully implemented) vs. **Complete** vs.
**Stretch**.

### Top-level (unchanged from 1-2-3): single-letter accepted at all levels

```
/ Worksheet Range Copy Move File Print Graph Data System Add-In Quit
```

Quick mnemonic: **W R C M F P G D S A Q**.

### MVP menu slice

```
/Worksheet
  /Worksheet Global Format …        (all formats)
  /Worksheet Global Column-Width
  /Worksheet Global Recalc Natural | Manual | Automatic
  /Worksheet Global Default Other Undo Enable | Disable
  /Worksheet Insert  Row | Column
  /Worksheet Delete  Row | Column
  /Worksheet Column  Set-Width | Reset-Width | Hide | Display
  /Worksheet Erase   Yes | No
  /Worksheet Titles  Both | Horizontal | Vertical | Clear
  /Worksheet Status
/Range
  /Range Format …                   (all formats)
  /Range Label    Left | Right | Center
  /Range Erase
  /Range Name     Create | Delete | Labels | Reset | Table
  /Range Justify
  /Range Prot | Unprot
  /Range Value
  /Range Trans
  /Range Search   Formulas | Labels | Both → Find | Replace
/Copy
/Move
/File
  /File Retrieve
  /File Save  Cancel | Replace | Backup
  /File Combine Copy | Add | Subtract
  /File Xtract  Formulas | Values
  /File List    Worksheet | Active
  /File Import  Text | Numbers
  /File Dir
  /File New     Before | After
  /File Open    Before | After
/Print
  /Print Printer | File …  Range | Options (Header/Footer/Margins/Setup/Other)
                            Go | Align | Quit
/Quit No | Yes
```

**Δ:** v0.1's top-level was missing `Print Graph Data System Add-In`. Those
exist at the menu level from day one (descending into an unimplemented leaf
shows a "Not yet implemented" message in line 3 rather than hiding the
menu) — keeping the top-level correct from MVP onward protects muscle
memory.

### Complete menu slice (post-MVP)

- `/Graph` tree (7 types, named graphs, A-F ranges)
- `/Data` tree (Fill, Sort, Query, Table 1/2, Distribution, Regression, Parse)
- `/Print` advanced (Borders, Fonts, Image, Encoded)
- `/File Admin` (Reservation, Seal)
- `/Worksheet Window Perspective | Map | Graph`
- `/Worksheet Insert Sheet`, `/Worksheet Global Group`
- `/Range Input`

### Stretch menu slice

- `/Data Matrix`, `/Data External`
- `/Add-In` loader
- `/Data Table 3` (3D), Labeled variants
- Macros, `{…}` advanced commands
- Lotus-era print drivers

### Menu rendering rules (critical)

- Every menu item starts with a unique letter; the **first letter is the
  accelerator** and descends immediately without `Enter`.
- Arrow keys move the highlight horizontally; wrap at ends; Home/End jump to
  first/last.
- When an item is highlighted, line 3 shows either (a) that item's submenu
  items previewed, or (b) a one-line help string. Both conventions exist in
  1-2-3; we follow the manual: submenu preview when the highlighted item is a
  parent, help text when it's a leaf.
- `Esc` = back one level. Ctrl-Break = all the way out.

---

## 11. Formula and reference semantics

### Reference forms

| Form | Example | Meaning |
|---|---|---|
| Relative | `B3` | Current sheet |
| Absolute | `$B$3` | Current sheet |
| Mixed | `$B3`, `B$3` | Current sheet |
| Cross-sheet | `A:B3`, `$A:$B$3`, etc. | Explicit sheet letter |
| 3D range | `A:B3..C:D5` | Spans sheets A-C |
| File ref | `<<wages.wk3>>A:B2` | Value pulled from another file |
| Range | `A1..B5` | `..` separator (`A1.B5` is accepted, normalized) |
| Named range | `SALES` | ≤15 chars, letter-first; `$NAME` forces absolute |

### Operator precedence (high → low)

```
^               (exponent)
+ -             (unary)
* /             (mul/div)
+ -             (add/sub)
= <> < > <= >=  (comparison)
#NOT#           (logical not)
#AND# #OR# &    (logical and/or, string concatenation)
```

### Value starters — reiterated

`0-9 + - . ( @ # $`. **`=` is NOT a value starter in 1-2-3**; typing `=A1`
enters `=A1` as a *label*. `+A1` is the correct canonical reference formula.

### F4 ABS cycle

For a reference at the cursor, F4 cycles through the 8 possible `$`
combinations (2 for sheet × 2 for col × 2 for row).

### ERR / NA propagation

- ERR is produced by: divide-by-zero, out-of-range number (>1E+99 or <1E-99),
  undefined range name, bad type coercion, bad file reference, out-of-memory,
  deleted cell that a named-range boundary pointed to.
- Any cell referencing an ERR or NA cell itself becomes ERR/NA.
- IronCalc uses Excel-style `#DIV/0!`, `#REF!`, `#NAME?` etc. The UI layer
  translates these to `ERR` or `NA` for display (and the control panel may
  tag with the original code as a disclosure). Round-trip xlsx preserves the
  Excel error; display is 1-2-3.

---

## 12. Cell format tags

The tag in parentheses on control-panel line 1. **Every format has a tag; the
tag is 1-2-3's in-situ type inspector.**

| Tag | Format |
|---|---|
| `(F0)`..`(F15)` | Fixed, N decimals |
| `(S0)`..`(S15)` | Scientific |
| `(C0)`..`(C15)` | Currency |
| `(,0)`..`(,15)` | Comma |
| `(G)` | General |
| `(+)` | +/- bar |
| `(P0)`..`(P15)` | Percent |
| `(D1)`..`(D5)` | Date (DD-MMM-YY, DD-MMM, MMM-YY, Long Intn'l, Short Intn'l) |
| `(D6)`..`(D9)` | Time (with/without seconds; 12h/24h Intn'l) |
| `(T)` | Text — show formula, not value |
| `(H)` | Hidden |
| `(A)` | Automatic (type-sniffs) |
| `(L)` | Label-only cell |

Overflow: `*********` across the cell when the value can't fit at the current
format and column width. General format falls back to scientific before
asterisks.

Column-width tag `[Wn]` appears after the format tag when width is non-default.

---

## 13. 3D worksheet model

**Δ:** v0.1 didn't mention 3D at all. 1-2-3 R3.x is natively 3D and L123
cannot skip it.

- A "workbook" (what 1-2-3 calls an *active file*) contains 1..256 worksheets
  lettered `A..IV`.
- Cell `A:B3` = sheet A, col B, row 3. The sheet prefix is always present in
  the control panel readout.
- Ctrl-PgUp / Ctrl-PgDn = prev/next sheet.
- `/Worksheet Insert Sheet Before|After` inserts sheets.
- `/Worksheet Global Group Enable` turns on GROUP mode: Format/Label/Column/
  Row operations propagate across all sheets. GROUP indicator lights up.
- 3D ranges: `A:B3..C:D5`. Supported in formulas, /Copy, /Print, /Graph, etc.
- `/Worksheet Window Perspective` — stacked oblique view of 3 sheets.
- Ctrl-End then Ctrl-PgUp/PgDn / Home/End = navigate between active files
  (multi-file).

Backing model: a 1-2-3 "active file" maps to one IronCalc workbook. Multiple
active files = multiple in-memory IronCalc workbooks. Cross-file refs
(`<<foo.wk3>>A:B2`) need a resolver layer above IronCalc.

---

## 14. File formats

| Command | What it does |
|---|---|
| `/File Retrieve` | Replace all active files with one from disk. `.WK3` source files load read-only and retarget the save path to `<orig>.WK3.xlsx` so `/File Save` writes xlsx without overwriting the legacy file. |
| `/File Open` | Add file to active set (Before/After), enabling multi-file |
| `/File Save` | Save all active files |
| `/File Combine Copy/Add/Subtract` | Merge disk file into current cell, element-wise |
| `/File Xtract Formulas/Values` | Save a range to a new `.WK3` |
| `/File Import Text/Numbers` | Import ASCII into cells |
| `/File List` | Overlay list of files with metadata |
| `/File Erase` | Delete file on disk |

**Δ:** v0.1 said `Save As`. 1-2-3 does not have "Save As". To save under a
different name: `/File Save`, then edit the prompted filename before Enter.
To save a *subset*: `/File Xtract`. L123 must match this model.

### Extensions (native)

- `.WK3` — worksheet (default save)
- `.WK1` — Release 2 worksheet (compat)
- `.BAK` — backup on save
- `.PRN` — ASCII print output
- `.ENC` — encoded print output
- `.CGM` / `.PIC` — graph

### Extensions (L123's modern path)

- `.xlsx` — primary format. Offer `/File Save` default extension switch
  (`/Worksheet Global Default Ext Save .xlsx`); save to `.xlsx` by default for
  modern workflow.
- `.csv` — /File Import Text / Numbers target
- `.wk3` — read-only import (values, formulas, basic styles, col
  widths). Save converts to `.xlsx` (`<orig>.WK3.xlsx`); writing `.wk3`
  is a non-goal.

---

## 15. @Function surface

All of R3.4a's ~120 @functions, grouped by category. Full list in
`docs/AT_FUNCTIONS.md`. MVP subset (must ship in M2):

- Math: `@ABS @INT @MOD @ROUND @SQRT @EXP @LN @LOG @RAND @PI`
- Trig: `@SIN @COS @TAN @ASIN @ACOS @ATAN @ATAN2`
- Stats: `@SUM @AVG @COUNT @MAX @MIN @STD @STDS @VAR @VARS`
- Logical: `@IF @TRUE @FALSE @NA @ERR @ISERR @ISNA @ISNUMBER @ISSTRING`
- String: `@LEN(@LENGTH) @LEFT @RIGHT @MID @UPPER @LOWER @PROPER @TRIM
  @REPEAT @FIND @EXACT @STRING @VALUE @CHAR @CODE`
- Date/Time: `@DATE @DATEVALUE @DAY @MONTH @YEAR @NOW @TODAY @TIME
  @TIMEVALUE @HOUR @MINUTE @SECOND`
- Financial: `@PMT @PV @FV @NPV @IRR @RATE @CTERM @TERM @SLN @SYD @DDB`
- Lookup: `@VLOOKUP @HLOOKUP @INDEX @CHOOSE`
- Reference: `@CELL @CELLPOINTER @@ @ROWS @COLS @SHEETS @COORD`

Post-MVP: the `@D...` database family, `@SUMPRODUCT`, `@VDB`, `@INFO`,
`@N @S`, `@DQUERY @ISRANGE`, full financial set.

### @-name → engine mapping

Most 1-2-3 functions map 1:1 to Excel/IronCalc names. A small translation
table handles the rest (e.g. `@AVG` → `AVERAGE`, `@STDS` → `STDEV.S`,
`@PROPER` = `PROPER`). `@@` is 1-2-3's indirect and maps to `INDIRECT`.
Where 1-2-3 has no Excel equivalent (e.g. `@CELLPOINTER` in some forms), we
emulate at the interpreter level.

---

## 16. Data model

```rust
enum Value {
    Number(f64),
    Text(String),
    Bool(bool),                 // internal; 1-2-3 represents TRUE/FALSE as 1/0
    Error(ErrKind),             // Err, Na, DivZero, Ref, Name, Value, Num
}

enum CellContents {
    Empty,
    Label { prefix: LabelPrefix, text: String },
    Constant(Value),
    Formula { expr: String, cached_value: Option<Value>, cached_format_override: Option<Format> },
}

enum LabelPrefix { Apostrophe, Quote, Caret, Backslash, Pipe }

struct Cell {
    contents: CellContents,
    format: Format,       // None means inherit from sheet global
    protected: bool,
    note: Option<String>,
}

struct Address { file: FileId, sheet: u16, col: u16, row: u32 }
struct Range { start: Address, end: Address }   // inclusive, may span sheets within one file

struct Workbook {          // one active file
    path: Option<PathBuf>,
    sheets: Vec<Sheet>,    // 1..=256
    names: BTreeMap<String, RangeRef>,       // named ranges
    print_settings: Vec<NamedPrintSettings>,
    graphs: BTreeMap<String, GraphDef>,
    group_mode: bool,
    ironcalc: ironcalc::Model,               // backing engine
}

struct Session {
    active_files: Vec<Workbook>,             // multi-file mode
    current: usize,
    mode: Mode,
    pointer: Address,
    point_state: Option<PointState>,
    menu_stack: Vec<MenuNode>,
    command_journal: Vec<JournalEntry>,      // for Undo + Learn
    learn_range: Option<Range>,
    indicators: StatusBits,
}
```

Cell storage is sparse (BTreeMap or hash-based); IronCalc already does this,
we don't duplicate the cell store but do track label prefixes and notes
on top.

---

## 17. Engine abstraction

We define a trait so Formualizer remains a swappable fallback.

```rust
trait Engine {
    fn set_user_input(&mut self, addr: Address, input: &str) -> Result<()>;
    fn set_value(&mut self, addr: Address, v: Value) -> Result<()>;
    fn set_formula(&mut self, addr: Address, expr: &str) -> Result<()>;
    fn get_cell(&self, addr: Address) -> CellView;

    fn insert_rows(&mut self, sheet: u16, at: u32, n: u32) -> Result<()>;
    fn delete_rows(&mut self, sheet: u16, at: u32, n: u32) -> Result<()>;
    fn insert_cols(&mut self, sheet: u16, at: u16, n: u16) -> Result<()>;
    fn delete_cols(&mut self, sheet: u16, at: u16, n: u16) -> Result<()>;
    fn insert_sheet(&mut self, at: u16) -> Result<()>;
    fn delete_sheet(&mut self, at: u16) -> Result<()>;

    fn copy_range(&mut self, src: Range, dst: Address) -> Result<()>;
    fn move_range(&mut self, src: Range, dst: Address) -> Result<()>;

    fn define_name(&mut self, name: &str, range: Range) -> Result<()>;
    fn delete_name(&mut self, name: &str) -> Result<()>;
    fn names(&self) -> Vec<(String, Range)>;

    fn recalc(&mut self);
    fn recalc_mode(&self) -> RecalcMode;
    fn set_recalc_mode(&mut self, mode: RecalcMode);

    fn load_xlsx(&mut self, path: &Path) -> Result<()>;
    fn save_xlsx(&self, path: &Path) -> Result<()>;
}
```

`set_user_input` bypasses our label/value disambiguation; it's used when the
UI already did the classification. Raw `set_formula` is needed when writing
label prefixes verbatim (since IronCalc would re-parse).

---

## 18. Scope tiers

### MVP (must ship v1.0)

- Grid rendering with column letters, row numbers, cell highlight
- Three-line control panel with mode indicator, line-3 menu-hint/preview
- Status line with CALC/CAPS/NUM/OVR/END/CIRC/UNDO
- Modes: READY, LABEL, VALUE, EDIT, POINT, MENU, FILES, NAMES, HELP, ERROR, WAIT
- First-char rule, label prefixes, commit-on-arrow
- F1 help (context-aware), F2 edit, F3 names, F4 abs, F5 goto, F9 calc, Alt-F4 undo
- POINT with anchor, `.` corner cycle, Esc unanchor
- Menu tree: full top-level; MVP leaves implemented; other leaves show "Not
  implemented yet" in line 3 but preserve the menu path
- `/Worksheet`: Insert/Delete Row/Column/Sheet, Global Format, Column, Erase,
  Titles, Status, Global Group
- `/Range`: Format, Label, Erase, Name Create/Delete/Labels/Table, Justify,
  Prot, Unprot
- `/Copy`, `/Move` (with POINT for both source and dest)
- `/File`: Retrieve, Save, Xtract, List, Import (CSV as Numbers), New, Open
- `/Print File` (ASCII output only, no printer drivers)
- `/Quit`
- MVP @function set (§15)
- Undo (journal-based; honors /WGD Other Undo Enable)
- 3D model: multi-sheet, Ctrl-PgUp/Dn, GROUP, 3D ranges in formulas
- `.xlsx` round-trip via IronCalc; `.csv` import/export; `.wk3` read-only
  import via `ironcalc_lotus` (save converts to `.xlsx`)

### Complete (v1.x)

- `/Graph` full tree (line, bar, xy, stack, pie, hlco, mixed); graph pane;
  F10 full-screen graph
- `/Data`: Fill, Sort, Query, Table 1/2, Distribution, Regression, Parse
- `/Print Printer` and `/Print Encoded` with Options (headers, footers,
  borders, margins, page breaks)
- `/Worksheet Window` Perspective, Map, Graph pane
- `/Range Input` form mode
- `/File Admin` (Reservation, Seal)
- Full @function catalog
- Macros: `/X` commands, `{BRANCH}`, `{IF}`, `{LET}`, `{GETLABEL}`,
  `{GETNUMBER}`, `{MENUBRANCH}`, `{QUIT}`, `{RETURN}`, subroutine calls,
  `\A..\Z` naming, `\0` autoexec, `/Worksheet Learn`

### Stretch (nice to have)

- `/Data Matrix`, `/Data External` (DataLens-style)
- Advanced macro commands (`{FOR}`, `{OPEN}`, `{READ}`, `{FORM}`, `{DEFINE}`)
- Wysiwyg-style add-in layer (fonts/shading); limited in terminal, could
  target Kitty image protocol
- Lotus-era print drivers → replaced with `lp` / PDF export
- CRT themes (green, amber)
- `.wk3` write
- Multi-user `/File Admin Reservation` over a shared filesystem

---

## 19. Non-goals (won't do)

- Lotus Release 1/2 menu set (we target 3.1 specifically)
- SmartSuite or 1-2-3 for Windows look
- Real DOS code page support; we are UTF-8 with LMBCS input translation
- Binary compatibility with `.WK3` byte layout on write
- In-terminal cell charts beyond what Unicode + Kitty image protocol allows

---

## 20. Authenticity contract

These behaviors are **not optional**. If any are missing, L123 fails promise 1
(§1) regardless of how much else is correct.

1. **Three-line control panel** with line 3's dynamic submenu preview / help
   text / prompt.
2. **Mode indicator** top-right, updating in real time across all 13 modes.
3. **Menu accelerators**: single letter at any level descends without Enter.
4. **POINT mode** with automatic anchor on command prompts, `.` to cycle
   corners, Esc to unanchor.
5. **First-char rule** for LABEL vs VALUE, including `'` auto-prefix.
6. **`@` as formula sigil** (not `=`); `..` as range separator (not `:`);
   `#AND#` / `#OR#` / `#NOT#` as logical ops.
7. **Format tag** `(C2)` etc. in the control panel's cell readout.
8. **Commit-on-arrow** during data entry.
9. **`\-` fills a cell with dashes**; values overflow as `*********`.
10. **`/File Retrieve` wipes memory; `/File Save` does not prompt "Save As"**
    — the filename field is pre-filled and user-editable at the prompt.
11. **R3.4a WYSIWYG icon panel** visible on screen, with the 17-icon R3.4a
    layout and mouse activation — the most immediate visual marker that
    distinguishes R3.4a from R3.1.

These items are the acceptance checklist for MVP authenticity review.

---

## 21. Glossary

- **Active file**: a file currently loaded into memory. 1-2-3 supports
  multiple active files simultaneously.
- **Worksheet**: a single 256-col × 8192-row grid within an active file. A
  file has 1..256 worksheets lettered A..IV.
- **Range**: a rectangular block of cells within one file; may span multiple
  sheets of that file (3D range).
- **Point**: to select a cell/range by moving the pointer rather than typing
  its address.
- **Anchor**: the fixed corner of a POINT-mode range; `.` cycles which corner
  is anchored.

---

## 22. Change log

- **v0.2** (2026-04-23) — Rewrite after Reference + Tutorial manual review.
  Corrections: full mode model (13, not 4); three-line control panel
  explicit; complete top-level menu; `=` is not a value starter; no
  "Save As"; 3D model; authenticity contract §20.
- **v0.1** — Initial draft.
