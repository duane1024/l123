# L123 Implementation Plan

Companion to `SPEC.md v0.2`. This document is the execution path.

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

### M10 — Polish, help, themes (week 21-22)

- F1 context help: in every mode, F1 shows a help panel relevant to the
  current screen (MENU → help for the highlighted item; READY → general
  help; ERROR → help for the last error).
- Compose key (Alt-F1) for LMBCS characters.
- CRT themes (green, amber, classic blue-on-black) as `--theme` flag.
- `.wk3` read-only import (values): port of libwps's WK3 parser.
- Documentation pass on `docs/AT_FUNCTIONS.md`, `docs/MENU.md`, README.

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
