# L123 Implementation Plan

Companion to `SPEC.md v0.4`. This document is the execution path.

---

## 1. Guiding principles

1. **Ship the skeleton early, add muscle later.** A minimum *visibly
   correct* 1-2-3 shell (modes, panel, menu navigation, grid, POINT) should
   run within the first two weeks even without a working compute engine.
2. **Authenticity is tested, not asserted.** Each milestone has a
   keystroke-transcript test suite (§7) that exercises the Authenticity
   Contract (SPEC §20) at its current scope.
3. **Engine is behind a trait from day one.** We use IronCalc, but we do not
   let IronCalc types leak above `engine/`. This keeps Formualizer a ~2-day
   swap if IronCalc hits a wall.
4. **Every /menu/ path works from day one**, even if the leaf says "Not
   implemented yet." Menu muscle memory is the feature.
5. **No feature flags for half-done work.** If a command is shown in the
   menu, it is either fully implemented or explicitly guarded with a single
   message in line 3.

---

## 2. Repository layout

```
l123/
├── Cargo.toml                # workspace
├── crates/
│   ├── l123-core/            # Value, Address, Range, Format, Mode
│   ├── l123-engine/          # Engine trait + IronCalc adapter
│   ├── l123-cmd/             # Command enum, interpreter, journal, undo
│   ├── l123-menu/            # Menu tree (compile-time), accelerator dispatch
│   ├── l123-parse/           # 1-2-3 formula parser/printer (@ ↔ excel)
│   ├── l123-ui/              # ratatui widgets: ControlPanel, Grid, StatusLine
│   ├── l123-io/              # xlsx/csv/wk3 adapters
│   ├── l123-macro/           # /X + {} macro system (post-MVP)
│   └── l123/                 # binary; main.rs + app loop
├── docs/
│   ├── SPEC.md
│   ├── PLAN.md               # this file
│   ├── MENU.md               # auto-generated full menu tree
│   └── AT_FUNCTIONS.md       # @function catalog
└── tests/
    ├── acceptance/           # keystroke transcripts (authenticity checks)
    └── fixtures/
```

Workspace crates, not one mega-crate — this keeps compile times tolerable and
enforces the layering.

---

## 3. Milestones

Each milestone ends with a shippable state and an acceptance test set.

### M0 — Bring-up (week 1)

- Cargo workspace, CI with fmt + clippy + test
- ratatui + crossterm hello-world: renders a fixed grid at 80×25, cell
  pointer navigation (arrow keys, PgUp/PgDn, Home, End)
- Status line renders with a fixed `READY` and a clock
- `Session`, `Address`, `Range`, `Mode` enums from SPEC §16 compile

**Acceptance:** `cargo run` opens a screen, arrows move the pointer, `/QY`
quits.

### M1 — Control panel, modes, first-char input (week 2-3)

- Three-line control panel widget. Line 1 shows `A:A1: `; line 2 blank; line
  3 blank. Mode indicator right-justified on line 1.
- Mode engine: READY ↔ LABEL ↔ VALUE ↔ EDIT transitions on first-char rule.
  `'`, `"`, `^`, `\` label prefixes recognised.
- Entry buffer rendering on line 2 with cursor.
- Enter commits; Arrow/Tab commit-and-move. Esc cancels to READY.
- F2 opens EDIT; F4 abs-cycle is stubbed (needs formula parser).

**Acceptance transcripts:**

```
Type: hello<Enter>
  → A1 cell shows "hello"; control panel line 1 reads `A:A1: 'hello`
Type: 123<Enter>
  → A2 shows "123" right-aligned; line 1 reads `A:A2: 123`
Type: "right<Enter>
  → line 1 reads `A:A3: "right`; cell displays right-aligned
Type: \-<Enter>
  → line 1 reads `A:A4: \-`; cell displays "---------" (dashes to column width)
```

No compute engine wired yet — entries are stored as strings in a stub
backend. This milestone is purely UX.

### M2 — Engine wire-up, values, formulas, recalc (week 4-5)

- `l123-engine` crate: `Engine` trait + IronCalc adapter.
- `set_user_input` router: respects first-char rule, LABEL → stores as
  label, VALUE → forwards to IronCalc.
- F9 triggers recalc; CALC indicator lights when pending.
- `l123-parse` translates `@SUM(A1..A5)` ↔ IronCalc's `SUM(A1:A5)`. `..` ↔
  `:`; `@` sigil handling; `#AND#`/`#OR#`/`#NOT#` ↔ `AND()`/`OR()`/`NOT()`;
  `&` string concat preserved.
- MVP @function set (SPEC §15) works end-to-end.
- Formula display in control panel line 1 (as typed) vs. value in the cell
  (computed). Format tag `(G)` by default.
- F4 ABS cycle working on current reference.

**Acceptance:**

```
A1..A5: 10,20,30,40,50
C1: @sum(<POINT A1..A5>)<Enter>
  → C1 displays 150; line 1 reads `A:C1: (G) @SUM(A1..A5)`
F9 → recalc; no CALC indicator
Edit A1 to 100: C1 recomputes to 240
```

### M3 — Menu system and slash commands (week 5-7)

- `l123-menu` crate: static compile-time menu tree. Each node carries:
  letter, name, help text, children or action.
- Menu rendering on line 2 (items horizontally) + line 3 (preview/help).
- `/` → MENU mode; single-letter accelerator descends; arrow keys highlight;
  Esc backs out; Ctrl-Break aborts.
- Command interpreter (`l123-cmd`): each leaf builds a `Command` variant,
  then the engine executes it.
- MVP leaves implemented: `/Worksheet Insert/Delete Row|Column`,
  `/Worksheet Column Set-Width`, `/Worksheet Erase`, `/Range Format`,
  `/Range Label`, `/Range Erase`, `/Range Name Create|Delete`, `/Copy`,
  `/Move`, `/Quit`. Every other leaf displays "Not implemented yet" in line
  3 and refuses to act.
- POINT mode engaged from `/Copy`, `/Move`, `/Range` prompts.
- `.` corner cycle; Esc unanchor; typed address replaces highlight; F3
  pops up range-name list.

**Acceptance (the keystroke suite is the spec):**

```
/RFC2~<POINT B3..B8>~
  → B3..B8 display as (C2); currency format applied
/C~<POINT A1..A5>~<POINT C1>~
  → A1..A5 copied to C1..C5 with relative references adjusted
/RNC sales~<POINT B3..B8>~
  → named range "sales" defined
/WIR~
  → row inserted above current pointer
```

### M4 — Files: xlsx and csv (week 7-8)

- `/File Retrieve`: FILES mode, directory listing on line 3 (arrow to scroll,
  Enter to load, Esc to cancel).
- `/File Save`: prompts with current path; Cancel/Replace/Backup sub-menu.
- `/File Xtract`: POINT for range, save as new file. Formulas vs Values.
- `/File Import Numbers`: CSV import into range starting at pointer.
- `/File New`, `/File Open` (multi-file), `/File Dir`, `/File List`.

**Acceptance:**

```
Open a known-good xlsx (with formulas) via /FR; values recompute correctly.
Save back out; reopen in Excel — no formula loss.
Import a CSV; modify; export via /File Save (default extension .xlsx).
```

### M5 — 3D, GROUP, named ranges, undo (week 9-10)

- `/Worksheet Insert Sheet Before|After`; Ctrl-PgUp/PgDn navigation.
- `/Worksheet Global Group Enable|Disable`; GROUP indicator; propagation
  across sheets for format/column/row commands.
- 3D range references in formulas (`A:B3..C:D5`).
- Multi-file support: `/File Open Before|After`; Ctrl-End + PgUp/PgDn for
  file navigation; FILE indicator.
- Command journal for Undo; Alt-F4 reverts to previous READY state.
- UNDO indicator respects `/WGD Other Undo Enable`.

**Acceptance:**

```
/WIS A~ creates sheet B; Ctrl-PgDn moves to B.
In GROUP: /RFP on sheet A applies to all sheets.
@SUM(A:B3..C:B3) correctly sums across sheets.
/FON <file> loads second file; Ctrl-End Ctrl-PgUp switches.
Delete a row, press Alt-F4 — row is restored.
```

### M6 — Printing (ASCII only) and Range Search (week 11)

- `/Print File`: Range, Options (Header, Footer, Margins, Page-Length,
  Other: Formatted/Unformatted/As-Displayed/Cell-Formulas), Go, Align, Clear.
- `|` in first col hides row from print output.
- `/Range Search Formulas|Labels|Both → Find|Replace`: prompts, POINTs
  matches.

**Acceptance:**

```
/PF inc.prn~R<POINT A1..F20>~OHi|Center|Rq~G
  → inc.prn is a plain-text dump of that range with "i|Center|R" formatted
    header and page footer.
/RSF find sum<Enter>replace total<Enter>
  → all "sum" in formulas replaced with "total" in selected range.
```

### M7 — Graphs (week 12-14)

- `/Graph` tree; X and A..F ranges; Type Line|Bar|XY|Stack|Pie|HLCO|Mixed;
  Options (Titles, Legend, Scale, Grid, Color/B&W, Format, Data-Labels).
- F10 renders current graph full-screen. Use Unicode braille/block chars
  for line/bar; fall back to Kitty image protocol on capable terminals.
- `/Graph Name Create|Use|Delete|Reset|Table`.
- `/Graph Save` to `.CGM` — write to a simple SVG with the same filename
  (with toggle).

**Acceptance:** open INC3-equivalent workbook; `/GTB /GA..F` select
quarterly ranges; F10 shows a bar chart.

### M8 — Data commands (week 15-17)

- `/Data Fill`: sequence generator (numbers, dates, times).
- `/Data Sort`: Data-Range + Primary/Secondary/Extra keys; Ascending /
  Descending; Go.
- `/Data Query Find|Extract|Unique|Del|Modify`; F7 repeats.
- `/Data Table 1|2`; F8 repeats.
- `/Data Distribution`, `/Data Regression`, `/Data Parse`.
  - Distribution output gets an extra Unicode-bar histogram column
    (`▁▂▃▄▅▆▇█`) right of the count column, scaled to the largest bin.
    The histogram is plain `Label` cells (apostrophe prefix), so xlsx
    round-trip is clean and the user can erase or restyle it like any
    other range. Acceptance: `M8_data_distribution.tsv` asserts the
    bars render at the right positions.

### M9 — Macros and Learn (week 18-20)

- Command journal already captures every keystroke/command. `/Worksheet
  Learn Range` writes the journal into a range as macro text.
- `\A..\Z` naming convention: range name on a cell starts macro from there.
- Macro interpreter: `{BRANCH}`, `{IF}`, `{LET}`, `{PUT}`, `{GETLABEL}`,
  `{GETNUMBER}`, `{GET}`, `{LOOK}`, `{MENUBRANCH}`, `{MENUCALL}`, `{QUIT}`,
  `{RETURN}`, `{SUBROUTINE}`, `{ONERROR}`, `{WAIT}`, `{?}`, `~` = Enter,
  `{xxx}` = special keys per SPEC §11 ref.
- `/X` legacy commands on top of the above.
- Alt-F3 RUN to invoke by name; Alt-F5 LEARN toggle; Alt-F2 STEP.
- `\0` autoexec runs on file retrieve if `/WGD Autoexec Yes`.
- **`.l123log` sidecar (v0.4):** while LEARN is on, the same journal
  that backs Undo (§4.3) is also streamed line-by-line to a sidecar
  `<workbook>.l123log` file as one JSON record per command. The
  sidecar is human-readable and replayable via `l123 --replay
  <file>.l123log`, which is the regression-test harness for
  acceptance transcripts that exceed `.tsv`'s expressiveness. The
  sidecar is *additive* to the in-range macro that `/Worksheet Learn
  Range` already produces — picking one over the other is a workflow
  choice. Acceptance: `M9_learn_sidecar_replay.tsv` records a
  session, replays it on a fresh workbook, asserts identical end
  state.

### M10 — Polish, help, themes (week 21-22)

- F1 context help: in every mode, F1 shows a help panel relevant to the
  current screen (MENU → help for the highlighted item; READY → general
  help; ERROR → help for the last error).
- Compose key (Alt-F1) for LMBCS characters.
- CRT themes (green, amber, classic blue-on-black) as `--theme` flag.
- Documentation pass on `docs/AT_FUNCTIONS.md`, `docs/MENU.md`, README.

(Read-only `.wk3` import landed earlier via `ironcalc_lotus`; saving
converts to `<orig>.WK3.xlsx`.)

### M11 — Modern data import & Range Compare (week 23)

Adds three `/File Import` verbs for modern formats. Reuses the
FILES-mode infrastructure from M4 and the typed-loader pattern that
already handles xlsx/csv. SPEC §14 v0.4.

- `l123-io` adds `json_loader.rs`, `parquet_loader.rs`,
  `sqlite_loader.rs`. Each implements a common `RecordLoader` trait
  (open path → iterator of typed rows + header row).
- Type widening rule: numbers → `Value::Number`; bools → 1/0;
  null/None → `CellContents::Empty`; strings → `Label{prefix:
  Apostrophe}`; dates → `Number` with format tag `(D1)`.
- `l123-menu` adds three new leaves under `/File Import`. Each leaf
  enters FILES mode at the cwd, filtered by extension
  (`.json`/`.jsonl`, `.parquet`, `.sqlite`/`.db`).
- For sqlite: after picking the file, a NAMES-mode picker over the
  tables in that file appears (a single sqlite file can hold many).
- Errors (malformed JSON, parquet schema mismatch, sqlite locked) map
  to ERROR mode with a one-line cause on line 3.

**Acceptance transcripts** (new under `tests/acceptance/`):

```
M11_import_json.tsv        — array-of-objects → header row + records
M11_import_jsonl.tsv       — JSON-Lines streaming variant
M11_import_parquet.tsv     — typed columns preserved (Number, D1 date)
M11_import_sqlite.tsv      — picks file, then table from NAMES list
M11_import_json_error.tsv  — malformed JSON → ERROR mode, no partial load
```

```
/FIJ <FILES nav to fixtures/sales.json><Enter>
  → header row at A1 (id name qty); records at A2..D11
  → control panel reads `A:A1: 'id`
/FIS <FILES nav to fixtures/inventory.db><Enter>
  → NAMES list: items, suppliers, orders
  <select items><Enter>
  → table loaded at pointer; numeric cols right-aligned, string cols
    left-aligned with apostrophe prefix
```

**`/Range Compare` (v0.4):** new leaf under `/Range`. Prompts for two
ranges (POINT for both, identical to `/Copy`); produces a third
range, anchored at a third user-pointed cell, with one row per
differing cell: `(addr, left_value, right_value, diff_kind)` where
`diff_kind ∈ {only-left, only-right, both-different, type-mismatch}`.
Equal cells produce no output row. Useful for diffing xlsx
revisions, reconciliation, and regression checking. Implementation
sits in `l123-cmd` as a pure pass over two `RangeView`s; no engine
changes.

```
M11_range_compare_basic.tsv      — diff two ranges, 3 differing cells
M11_range_compare_type_mismatch.tsv — number vs label flagged distinctly
M11_range_compare_size_mismatch.tsv — different shapes → ERROR mode
M11_range_compare_clean.tsv      — equal ranges → empty output, info on line 3
```

```
/RC<POINT A1..C5>~<POINT E1..G5>~<POINT A20>~
  → output at A20..D? lists each differing cell with both values
```

### M12 — `/Data External` (week 24-25)

Live SQL source. Honors 1-2-3 R3.4a's DataLens architecture: a range
that is *backed by* an external query, refreshable on demand.
SPEC §18 (Complete tier, v0.4).

- New crate `l123-extdata` (or module under `l123-io`) with two
  drivers: `sqlite` (file path) and `postgres` (libpq URL). Both
  behind a `DataSource` trait so MySQL/DuckDB can land later.
- `/Data External` submenu: `Connect`, `Use`, `Refresh`, `List`,
  `Reset`, `Disconnect`.
  - `Connect` — prompts for a name (≤15 chars, named-range rules) and
    a connection string. Tests connectivity; ERROR mode on failure.
  - `Use <name> <query>` — prompts for SQL; runs it; populates a
    pointer-anchored range with the result. The range is marked
    *external-bound*; cells are display-protected (PROT visible).
  - `Refresh` — re-runs the query in WAIT mode; replaces values
    in-place (preserving column widths and adjacent formulas).
  - `List` — overlay listing all external connections + last-refresh
    timestamps, like `/File List`.
- Persistence: external-bound ranges are saved as xlsx custom
  document properties so `/File Save` round-trips the binding.
  Connection strings are stored without credentials; passwords come
  from a `~/.l123/credentials` keyfile or env var on reconnect.
- Recalc interaction: when an external-bound range's cells are
  referenced by formulas, recalc uses the cached values (no implicit
  query refresh on F9 — explicit `/DER` only).
- READ-only this milestone; write-back to the external source is a
  non-goal.

**Acceptance transcripts:**

```
M12_external_sqlite_connect.tsv   — Connect/Use/Refresh round-trip
M12_external_protected.tsv        — bound cells show PROT; direct edit refused
M12_external_xlsx_roundtrip.tsv   — save, close, reopen → binding restored
M12_external_refresh_wait.tsv     — long query → WAIT mode → CALC clears
M12_external_error_disconnect.tsv — db unreachable → ERROR; range frozen
```

```
/DEC sales sqlite:tests/fixtures/inventory.db<Enter>
  → connection 'sales' established
/DEU sales SELECT * FROM items WHERE qty > 0<Enter><POINT C5>
  → C5..F? populated; range marked external-bound (PROT visible on
    pointer enter); status shows last-refresh timestamp
/DER sales<Enter>
  → WAIT mode during query; cells refreshed in place
```

### M13 — ADDIN key + Data Workbench (week 26-29)

The native plug-in surface. SPEC §22 (v0.4). The workbench is a
distinct overlay surface, not part of the §20 authenticity contract;
its in-overlay UX is permitted to break 1-2-3 conventions.

Sub-milestones:

**M13a — overlay framing (week 26)**

- New `WORKBENCH` mode + `WORKBNCH` indicator. Mode model in
  `l123-core` extends with `Mode::Workbench { input: Range, view:
  WorkbenchState }`.
- `l123-ui` adds `WorkbenchOverlay` widget that takes over the screen,
  swapping out `Grid`/`ControlPanel` for a workbench-native layout.
- Alt-F10 in READY → POINT for input range → overlay opens. Esc-Esc
  or second Alt-F10 → POINT for optional write-back range → return to
  READY.
- 1-2-3 keys are *intentionally non-functional* inside the overlay;
  pressing `/` shows a one-line "1-2-3 keys disabled in Workbench"
  hint on the workbench status bar. Esc still backs out (the only
  shared key).

**M13b — transform set (week 27-28)**

- New crate `l123-workbench`. Reads a typed view of the input range
  (uses `l123-engine`'s `CellView` to extract type-tagged columns).
- v0.4 transforms: `Sort` (multi-key), `RegexFilter`, `Frequency`
  (with Unicode-bar histogram column), `Describe` (count, null,
  mean, median, stdev, min, max, q1, q3).
- Transform stack rendered in the workbench status bar; each
  transform is non-destructive over the previous layer.
- Vim-style movement (hjkl + g/G + /) inside the overlay. Type
  inference per column. Sort/filter wired through the transform stack.

**M13c — write-back (week 29)**

- `W` (write) inside overlay → exit-then-POINT for target range →
  current top-of-stack transform output committed as plain values
  starting at the target. No formulas, no label prefixes (apostrophe
  is auto-applied by the standard label-prefix rule on commit).
- Workbench state is in-memory only; no persistence into xlsx.
- Undo journal records the write-back as a single
  `Command::WorkbenchCommit` so Alt-F4 reverses the entire write.

**M13d — Stretch (post-v0.4)**

- Melt, Transpose, Unfurl, Join, TypeAudit transforms.
- APP1/APP2/APP3 binding via `~/.l123/plugins.toml` (Rust `cdylib`
  load).

**Acceptance transcripts (M13a-c):**

```
M13_addin_open_close.tsv         — Alt-F10 → POINT → overlay → Esc-Esc → READY
M13_workbench_mode_indicator.tsv — WORKBNCH visible; READY hidden
M13_workbench_disables_slash.tsv — `/` inside overlay → hint, no menu
M13_workbench_sort.tsv           — Sort by col 2 desc, top row matches
M13_workbench_frequency.tsv      — histogram column renders ▁..█
M13_workbench_describe.tsv       — 9 stats per typed column
M13_workbench_writeback.tsv      — W → POINT F1 → cells F1..H? populated
M13_workbench_writeback_undo.tsv — Alt-F4 reverses full write-back
M13_workbench_xlsx_clean.tsv     — workbench use leaves xlsx round-trip clean
```

Each Workbench acceptance test is structured: 1-2-3 keys outside, then
Alt-F10, then a workbench keystroke sub-block, then exit + outside
verification. The test harness needs a flag for "the next N keystrokes
are workbench keystrokes" so tests don't leak workbench keys back into
the 1-2-3 dispatch. Add this to `tests/acceptance/README.md` as part of
M13a.

---

## 4. Key design decisions

### 4.1 Menu tree: data or code?

**Decision: data.** Menu is a compile-time constant tree (`static MENU:
MenuNode = ...`), with leaves carrying an `Action` variant. This makes the
menu trivially enumerable (for `docs/MENU.md` generation and for the
"Not implemented yet" safety net).

### 4.2 Formula parser: ours or borrowed?

**Decision: ours, thin.** IronCalc has an internal parser but wants Excel
syntax. We keep our own thin 1-2-3 parser whose only job is to translate to
an Excel-shaped string, then hand to `set_user_input`. Names we can't map
(e.g. `@CELLPOINTER`, `@@`, 3D ranges IronCalc doesn't accept in that form)
are rewritten or computed at the interpreter level.

### 4.3 Undo

**Decision: command journal, not snapshots.** Each mutating `Command`
records a reverse-Command before execution. Undo pops and executes reverse.
Works for all cell ops, row/col ops, name ops. File ops (retrieve, save)
are NOT undone, consistent with 1-2-3. Macro replay uses the same journal
format — this is the Learn feature for free.

### 4.4 Multi-file

**Decision: multiple IronCalc `Model`s.** A 1-2-3 "active file" = one
IronCalc Model. Cross-file refs resolve at formula-translation time by
looking up the file in the session's active-file list and substituting a
value cache if present; if the referenced file isn't open and
`/File Admin Link-Refresh` is off, we preserve the last cached value.

### 4.5 3D ranges

**Decision: expand in translation.** `A:B3..C:B3` as an argument to `@SUM`
translates to `SUM(Sheet1!B3, Sheet2!B3, Sheet3!B3)` for IronCalc. This
keeps IronCalc happy; we reverse on round-trip.

### 4.6 Character set

**Decision: UTF-8 internally.** LMBCS input accepted via `@CHAR` and paste
events; emitted as UTF-8 codepoints. Compose key (Alt-F1) maps Lotus's
multi-keystroke sequences to Unicode.

### 4.7 Long operations and WAIT mode

**Decision: every potentially-slow op runs on a Tokio task; WAIT mode
is the user-visible contract.** SPEC §5 already lists `WAIT` as a
first-class mode. We honor it for: `/File Retrieve`, `/File Save`,
`/File Import` (all formats incl. M11 JSON/Parquet/Sqlite),
`/Data External Refresh` (M12), recalc on workbooks > 50k cells, and
xlsx round-trip. Pattern:

1. Op begins → mode switches to `WAIT` → indicator visible top-right
   → control-panel line 3 shows "Loading foo.parquet…" (or
   equivalent).
2. Op runs on a background task; the UI thread continues to render
   at 60Hz, drawing only spinner/progress updates (no input handled
   except Ctrl-Break).
3. Progress: where the loader can report it (file bytes read,
   sqlite rows yielded), a simple `[████████░░] 80%` bar renders on
   line 3. Where it can't, an animated spinner is sufficient.
4. Ctrl-Break interrupts, returns to READY, leaves no partial state.
5. Op completes → mode returns to caller's previous mode (usually
   READY) → results are committed in one journal entry (so Alt-F4
   undoes the whole load).

This is a cross-cutting infra item, not a milestone. Lands as part
of M4 (`l123-io::AsyncOp` trait + WAIT-mode plumbing in `l123-ui`)
and is reused by every later milestone that needs it.

---

## 5. Risks and mitigations

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| IronCalc's `set_user_input` swallows `@SUM` as text | Med | High | Parser layer always translates `@` → Excel names before the call. Covered by parse tests. |
| IronCalc row/col insertion breaks references across sheets | Med | High | Regression suite of multi-sheet formulas before/after every move/insert/delete op. Fall back to Formualizer if broken. |
| ratatui diffing drops frames on rapid cursor movement | Low | Med | Use crossterm synchronized output escape; profile at ~30 Hz key-repeat. |
| Kitty keyboard protocol unavailable in user's terminal | Med | Low | Detect at startup; degrade gracefully (F11/F12 combos, document limitation). |
| 1-2-3 menu accelerators clash with terminal keybindings (Ctrl-C, Ctrl-S) | Med | Med | Raw mode captures them all; Ctrl-C is intentionally inert (per SPEC §7 Δ — `/QY` is the only quit path); document Ctrl-Break for abort once implemented. |
| `.wk3` format underdocumented for writing | High | Low | Stretch goal; values-only read is the MVP-plus target; write deferred. |
| Author (you) burns out on completeness | Med | High | MVP is sufficient as a real product. Complete / Stretch tiers are separable releases. |

---

## 6. Test strategy

### 6.1 Unit tests

- `l123-core`: Address arithmetic, range intersection, reference parsing.
- `l123-parse`: 1-2-3 → Excel round-trip for the full MVP @function set;
  operator precedence fixtures; label-prefix parsing; 3D range handling.
- `l123-cmd`: Command → reverse-Command generation for every mutating op.
- `l123-menu`: every leaf dispatches to an Action; menu paths resolve.

### 6.2 Engine integration tests

- Drive IronCalc through the `Engine` trait with ~50 fixture workbooks:
  simple formulas, named ranges, row/col inserts, copy/move, 3D SUMs,
  cross-file refs. Assert values at well-known cells.
- xlsx round-trip: load, save, reload, assert bit-identical values (not
  byte-identical files).

### 6.3 Acceptance transcripts

The *authenticity* tests. Each is a text file under
`tests/acceptance/*.tsv` that describes keystrokes and expected screen
state deltas. Format:

```
# M1_label_entry.tsv
KEY    hello
ENTER
ASSERT_PANEL_L1  A:A1: 'hello
ASSERT_CELL      A1  hello
KEY    "right
ENTER
ASSERT_PANEL_L1  A:A2: "right
ASSERT_CELL      A2  right    # right-aligned
```

A test harness runs each transcript against a headless `Session`, asserting
panel text, cell display, mode indicator, and status bits. One file per
Authenticity-Contract item (SPEC §20).

### 6.4 Property tests

- Arbitrary @function expressions: generate, translate to Excel, translate
  back, compare ASTs.
- Arbitrary journal sequences: apply, then apply reverse sequence, assert
  equivalence to starting state.

### 6.5 Human-in-the-loop review

After each milestone, record a terminal session with `asciinema` and review
it against the Reference manual's corresponding section screenshots.
Capture in `tests/acceptance/sessions/Mn-*.cast`.

---

## 7. Definition of done (MVP v1.0)

1. All Authenticity Contract items (SPEC §20) pass their acceptance
   transcripts.
2. MVP menu slice (SPEC §10) fully functional; non-MVP leaves show
   "Not implemented yet" in line 3 and refuse to act.
3. `.xlsx` round-trip: a 100-row/10-formula fixture workbook loads, edits,
   saves, and reopens in Microsoft Excel with no formula loss.
4. CSV import and export work.
5. Undo reverses the previous READY-to-READY batch.
6. 3D worksheets work: insert, delete, navigate, GROUP format, 3D `@SUM`.
7. Clean build on macOS, Linux, and Windows (crossterm).
8. `docs/SPEC.md`, `docs/MENU.md`, `docs/AT_FUNCTIONS.md` are current.
9. README with installation + first-session walkthrough.

---

## 8. Commit and branching hygiene

- One topic per PR. Each PR maps to a sub-item of a milestone.
- Every mutating PR adds or updates at least one acceptance transcript.
- Conventional commits: `feat(ui): ...`, `fix(engine): ...`,
  `docs(spec): ...`, `test(acceptance): ...`.
- `main` always builds and passes acceptance tests. `dev/*` for in-flight.

---

## 9. Open questions (resolve before M2)

1. **Sheet letter vs. name.** IronCalc sheets have names (strings). 1-2-3
   sheets have letters A..IV derived from position. Decision: keep a
   letter↔name map inside `Workbook`; users see letters, IronCalc sees
   names. Default names `Sheet A`, `Sheet B`, …
2. **Column widths in xlsx.** 1-2-3 widths are character counts; xlsx
   widths are Calibri-11-ish units. Preserve 1-2-3 widths in a file-level
   property; map to xlsx "best match" on save.
3. **Error type surfaces.** Display as `ERR`/`NA` (1-2-3) or `#REF!` (Excel)?
   Decision per SPEC §11: show `ERR`/`NA` in the grid and control panel;
   disclose the Excel code via F1 help on the erroring cell.
4. **Key-repeat rate for POINT expansion.** Too slow → painful; too fast →
   overshoots. Default to terminal repeat; offer a `--point-repeat-ms`
   flag.
5. **Kitty keyboard protocol.** Enable by default; offer `--no-kitty-kbd`
   flag for users on terminals where it misbehaves.

---

## 10. Immediate next steps

1. Initialize the cargo workspace (crates listed in §2).
2. Wire IronCalc 0.7.x as a dependency behind the `Engine` trait with a
   single passing test: `create workbook → set A1=1, A2=2, A3=+A1+A2 →
   recalc → assert A3 = 3`.
3. Stand up a ratatui hello-world with a cell grid and arrow-key
   navigation (M0).
4. Draft `docs/MENU.md` from the Reference research notes; commit as the
   source of truth for M3.
5. Write the first acceptance transcript (`M1_label_entry.tsv`) and the
   harness that runs it.

When those five are done, we're at end-of-M0 and can start M1's control
panel in earnest.
