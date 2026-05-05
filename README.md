# l123

**A Lotus 1-2-3–style terminal spreadsheet with modern Excel compatibility.**

l123 recreates the classic DOS-era spreadsheet experience — slash menus,
three-line control panel, keyboard-first workflows, WYSIWYG icon panel,
and all — on top of a modern formula engine with native `.xlsx`
round-trip.

Its interaction model targets **Lotus 1-2-3 Release 3.4a for DOS**
(1993). Its compute and I/O layers are Rust, IronCalc, and UTF-8.

![l123 running in iTerm](docs/iterm-screenshot.png)

---

## ✦ Status

Currently `v1.1.1`, with all MVP milestones (M0–M10) shipped.
Tracking the milestone plan in [`docs/PLAN.md`](docs/PLAN.md):

| Milestone | Scope | State |
|---|---|---|
| M0 | Grid, pointer nav, workspace bring-up | ✅ done |
| M1 | Control panel, modes, first-char input | ✅ done |
| M2 | Engine wire-up, formulas, recalc | ✅ done |
| M3 | Menu system and MVP slash commands | ✅ done |
| M4 | `.xlsx` and CSV round-trip | ✅ done |
| M5 | 3D sheets, GROUP, named ranges, undo | ✅ done |
| M6 | Printing (ASCII, PDF, line-printer) and Range Search | ✅ done |
| M7 | Graphs: 7 chart types, F10 view, SVG/PNG save | ✅ done |
| M8 | `/Data`: Fill, Sort, Query, Table, Distribution, Regression, Parse | ✅ done |
| M9 | Macros: `/X`, `{BRANCH}`, `{IF}`, `{MENUBRANCH}`, Learn | ✅ done |
| M10 | Polish: R3.4a WYSIWYG icon panel, startup splash, F1 context help, DOS theme | ✅ done |
| M11 | Modern data import (JSON, Parquet, SQLite) and `/Range Compare` | 🚧 in progress |
| M12 | `/Data External` (live SQL: SQLite, Postgres) | 📋 planned |
| M13 | Alt-F10 ADDIN: Data Workbench transform overlay | 📋 planned |

The MVP slice is feature-complete and stable. Currently 232 acceptance
transcripts under `tests/acceptance/` lock the authenticity contract
(SPEC §20) plus a tutorial-derived suite (T01–T12) tracking the 1-2-3
R3.1 Tutorial chapters end-to-end.

---

## ✦ Install

### Homebrew (macOS / Linux)

[Homebrew](https://brew.sh) is the standard package manager for macOS
(and works on Linux too). If you don't already have it, install it
first from [brew.sh](https://brew.sh), then:

```bash
brew install duane1024/l123/l123
```

Or tap once and install by name:

```bash
brew tap duane1024/l123
brew install l123
```

Upgrade later with `brew upgrade l123`.

### From source

Requires Rust stable (pinned via `rust-toolchain.toml`).

```bash
git clone git@github.com:duane1024/l123.git
cd l123
cargo build --release
./target/release/l123            # or: cargo run -p l123
```

Open an existing workbook:

```bash
l123 financials.xlsx
```

Inspect or initialize configuration:

```bash
l123 config            # show effective settings and their sources
l123 config --init     # write a sample ~/.l123/L123.CNF
```

See [`docs/CONFIG.md`](docs/CONFIG.md) for the full list of keys and
environment variables.

---

## ✦ Keyboard, the short version

The keyboard *is* the product.

| Key | What it does |
|---|---|
| `/` | Open the slash menu |
| `:` | Open the `:` (WYSIWYG) menu |
| First letter | Descend into a menu item (no `Enter` needed) |
| Arrows / `Tab` | Move pointer; during entry, commit-and-move |
| `Enter` | Commit cell entry |
| `Esc` | Back out one level (menu, prompt, POINT anchor) |
| `Ctrl-Break` | Abort to READY from anywhere; cancel a WAIT op |
| `.` (in POINT) | Cycle which corner of the range is anchored |
| `F1` | Context help (824-page R3.1 manual, navigable) |
| `F2` | Edit current cell |
| `F3` | List named ranges (overlay; works in GOTO and prompts) |
| `F4` | Cycle `$` absoluteness in a reference |
| `F5` | GOTO cell or named range |
| `F7` | Repeat last `/Data Query` |
| `F8` | Repeat last `/Data Table` |
| `F9` | Recalculate |
| `F10` | Full-screen graph view |
| `Alt-F2` | STEP — single-step macro execution |
| `Alt-F3` | RUN — invoke a macro by name |
| `Alt-F4` | Undo |
| `Alt-F5` | LEARN — toggle keystroke recording |
| `Alt-`*letter* | Run macro `\`*letter* (e.g. `Alt-A` → `\A`) |
| `Ctrl-PgUp` / `Ctrl-PgDn` | Previous / next sheet |
| `Ctrl-End` then `Ctrl-PgUp/PgDn` | Previous / next open file |

Mouse is supported for the WYSIWYG icon panel (17 icons, R3.4a layout)
plus cell click, drag-select, scroll wheel, and POINT extend.

Formulas use 1-2-3 syntax: `@SUM(A1..A5)`, not `=SUM(A1:A5)`. The `@`
sigil and `..` separator are required. `#AND#`, `#OR#`, `#NOT#` are the
logical operators.

First character typed in READY decides label vs. value: digits and
`+ - . ( @ # $` start a value; anything else starts a label (with an
auto-inserted `'` prefix). `"` = right-align, `^` = center, `\-` fills the
cell with dashes.

---

## ✦ What works today

**Core UX**

- Three-line control panel with live mode indicator
- 13 modes (READY, LABEL, VALUE, EDIT, POINT, MENU, FILES, NAMES, HELP,
  ERROR, WAIT, FIND, STAT)
- Full slash-menu tree: every path in `docs/MENU.md` is reachable; MVP
  leaves execute, non-MVP leaves show "Not implemented yet" in line 3
- POINT mode with `.` corner cycle, F3 named-range picker, typed
  addresses, and mouse drag-select
- Async WAIT mode for long ops (file load/save, recalc, large imports)
  with progress bar and Ctrl-Break cancellation

**Worksheet, Range, File**

- `/Worksheet`, `/Range`, `/Copy`, `/Move`, `/File`, `/Quit` —
  full menus implemented
- `.xlsx` round-trip through IronCalc (formulas, alignment, borders,
  comments, fills, fonts, frozen panes, hidden sheets, merges,
  sheet color, tables); `.csv` import and export
- `.WK3` read (values, formulas, basic styles, column widths) via the
  optional `wk3` cargo feature, gated behind a local `ironcalc_lotus`
  fork — see CLAUDE.md
- 3D workbooks: `A..IV` sheets, `A:B3..C:D5` ranges, GROUP mode
- Named ranges (Create/Delete/Reset/Labels/Note/Table/Undefine);
  `@` function set including the legacy `@D360`, `@DGET`, `@REPLACE`
- Command-journal undo (Alt-F4), toggleable via `/WGD Other Undo`
- Multi-file sessions (`/File Open Before|After`, Ctrl-End navigation,
  `/File List`)

**Print, Graph, Data**

- `/Print File` to ASCII, PDF, or line-printer output; headers, footers,
  margins, page-length, formatted / unformatted / as-displayed /
  cell-formulas modes; column page breaks; `|` in first column hides
  rows from print
- `/Range Search Formulas|Labels|Both` Find and Replace
- `/Graph` tree: Line, Bar, XY, Stack, Pie, HLCO, Mixed; Titles, Legend,
  Scale, Grid, Color/B&W, Data-Labels; Name Create/Use/Delete/Reset/Table
- F10 / `/Graph View` full-screen rendering with Unicode bar + line
  output; Kitty / iTerm2 / Sixel image support via `ratatui-image`
- `/Graph Save` to SVG (and plotters PNG output for all chart types)
- `/Data` tree (full): Fill, Sort (with extra keys), Query
  (Find/Extract/Unique/Delete/Modify, F7 repeat), Table 1/2 (F8 repeat),
  Distribution (with Unicode-bar histogram), Regression, Parse, Matrix

**WYSIWYG, Macros, Help**

- R3.4a WYSIWYG icon panel: all 17 icons, mouse-wired, hover hints
- `:` (WYSIWYG) menu: Format (Bold, Color, Alignment), Display (Mode,
  Grid), Special (Copy, Move, Color), Worksheet status; xlsx round-trip
- Macros: `/X` legacy commands; `{BRANCH}`, `{IF}`, `{LET}`, `{PUT}`,
  `{GETLABEL}`, `{GETNUMBER}`, `{MENUBRANCH}`, `{QUIT}`, `{?}` pause,
  `{BEEP}`, special-key tokens, subroutines, `\0` autoexec
- `/Worksheet Learn Range` keystroke recording (Alt-F5 toggle, Alt-F2
  STEP, Alt-F3 RUN)
- F1 context help: full 1-2-3 R3.1 reference manual (824 pages),
  context-aware per mode
- Startup splash screen
- DOS theme (black-on-cyan headers) via `--theme dos`, `L123_THEME`,
  or `theme=` in `~/.l123/L123.CNF`
- Auto-contrast for xlsx cells with light fills on dark terminals

## ✦ Coming

- **M11** `/File Import Json|Jsonl|Parquet|Sqlite` — modern data
  loaders with type inference; `/Range Compare` for diffing two ranges
- **M12** `/Data External` — live SQL backing ranges (SQLite, Postgres)
  with `/DEC Connect`, `/DEU Use`, `/DER Refresh`; xlsx persistence of
  the binding
- **M13** Alt-F10 ADDIN — built-in Data Workbench overlay with a
  VisiData-inspired transform stack (Sort, RegexFilter, Frequency,
  Describe) and value write-back
- Stretch: LMBCS compose key (Alt-F1), additional CRT themes (green,
  amber, classic blue-on-black), APP1/APP2/APP3 plug-in slots

---

## ✦ Architecture

Rust workspace, strict layering:

```
l123-core  ← types only, zero external deps
  ↑
l123-parse, l123-menu
  ↑
l123-engine           (wraps IronCalc behind a trait)
  ↑
l123-cmd, l123-io, l123-graph, l123-print
  ↑
l123-ui               (ratatui + crossterm; engine-agnostic)
  ↑
l123                  (binary)
```

IronCalc is behind the `Engine` trait so it can be swapped. 1-2-3 formula
syntax (`@SUM`, `..`, `#AND#`) is translated to Excel syntax in
`l123-parse` before it reaches the engine. The UI never sees IronCalc
types.

---

## ✦ Authenticity contract

l123 makes two promises (`docs/SPEC.md` §1):

1. An experienced 1-2-3 R3.4a user can drive l123 cold, without reading
   anything.
2. Files round-trip cleanly to and from `.xlsx`.

SPEC §20 enumerates the behaviors — three-line control panel, menu
accelerators, POINT anchor semantics, first-char rule, `@` sigil, format
tags, commit-on-arrow, WYSIWYG icon panel, and so on — that the project
fails if it misses. Every item in the contract has at least one
acceptance transcript under `tests/acceptance/`.

---

## ✦ Development

Strict red / green / refactor. Conventions live in [`CLAUDE.md`](CLAUDE.md).

```bash
cargo test --workspace                                  # all tests
cargo test -p l123-ui --test acceptance                 # keystroke transcripts
cargo clippy --workspace --all-targets -- -D warnings   # lint
cargo fmt --all                                         # format
```

Acceptance transcripts are `.tsv` files describing keystrokes in and
screen state out; see `tests/acceptance/README.md` for the directive
syntax. Every UI-visible change lands with a transcript.

Canonical docs (treat as sources of truth):

- [`docs/SPEC.md`](docs/SPEC.md) — what l123 *is*
- [`docs/PLAN.md`](docs/PLAN.md) — milestones and risk register
- [`docs/MENU.md`](docs/MENU.md) — the complete menu tree

If code and doc disagree, fix the doc first.

---

## ✦ Non-goals

- Not a DOS emulator. No INT21h, no code pages. Strings are UTF-8.
- Not a visual homage. Functional fidelity, not CRT nostalgia. (Green /
  amber themes are a stretch goal, not the point.)
- Not a macro player for existing `.WK3` files. Read-only `.WK3` import
  is a stretch goal; write is not planned.
- Not a reimplementation of the 1-2-3 compute core. IronCalc does that.
- Not aimed at Lotus 1-2-3 for Windows or SmartSuite. Release 3.4a for
  DOS only.

---

## ✦ Philosophy

Spreadsheets didn't get worse — they just got heavier.

l123 brings back the speed, clarity, and keyboard-driven precision of
early spreadsheet software, without sacrificing compatibility with modern
workflows.

---

## ✦ Why?

Because the `/` key was never the problem.

---

## ✦ License

Licensed under the MIT License. See [`LICENSE`](LICENSE).
