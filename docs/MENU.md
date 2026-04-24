# L123 — Complete Menu Tree

Source of truth for `l123-menu`. Derived from *Lotus 1-2-3 Release 3.4a
Reference* (1993), which retains the R3.1 menu tree and adds the
always-on WYSIWYG commands and R3.4 icon panel.

Legend: **[MVP]** = in MVP slice; **[CPL]** = Complete-tier; **[STR]** =
Stretch. Every top-level entry and every leaf is reachable from day one —
non-MVP leaves display "Not implemented yet" in control-panel line 3.

---

## Top level

```
/ Worksheet  Range  Copy  Move  File  Print  Graph  Data  System  Add-In  Quit
  W          R      C     M     F     P      G      D     S       A       Q
```

Accelerators are the capitalized first letter. Arrow keys highlight; first
letter descends immediately; `Esc` backs out one level; `Ctrl-Break` aborts
to READY.

---

## /Worksheet  (W)

```
Global       Insert   Delete   Column   Erase   Titles   Window   Status   Page   Hide   Learn
```

### /Worksheet Global  (G)

```
Format  Label  Col-Width  Prot  Zero  Recalc  Default  Group  Quit
```

- **Format** → `Fixed 0-15 | Sci 0-15 | Currency 0-15 | , (Comma) 0-15 | General | +/- | Percent 0-15 | Date (1..5, Time 1..4) | Text | Hidden | Other → Automatic | Color → Negative/Reset | Label | Parentheses Yes/No | Reset`   **[MVP]**
- **Label** → Left | Right | Center    **[MVP]**
- **Col-Width** (1-240)                 **[MVP]**
- **Prot** → Enable | Disable           **[CPL]**
- **Zero** → No | Yes | Label           **[CPL]**
- **Recalc** → Natural | Columnwise | Rowwise | Automatic | Manual | Iteration (1-50)   **[MVP]**
- **Default**                           **[CPL]** (see below)
- **Group** → Enable | Disable          **[MVP]** (3D GROUP mode)

#### /Worksheet Global Default  **[CPL]**

```
Printer  Dir  Status  Update  Other  Autoexec  Ext  Graph  Temp  Quit
```

- **Printer** → Interface 1-9 | AutoLf | Left | Right | Top | Bottom | Pg-Length | Wait | Setup | Name | Quit
- **Dir** (default directory)
- **Status** (display STAT screen)
- **Update** (write `123R31.CNF`; L123 writes `~/.config/l123/l123.toml`)
- **Other** → International (Punctuation A-H, Currency Prefix/Suffix, Date A-D, Time A-D, Negative, Release-2 LICS/LMBCS, File-Translation, Quit) | Help Instant/Removable | Clock Standard/International/None/Filename | **Undo Enable/Disable** **[MVP-critical]** | Beep | Add-In | Expanded-Memory
- **Autoexec** → Yes | No    (run `\0` macro on retrieve)
- **Ext** → Save (default ext) | List (filter for /File List/Retrieve)
- **Graph** → Columnwise/Rowwise auto-graph; CGM | PIC default type
- **Temp** (temp dir)

### /Worksheet Insert  (I)  **[MVP]**

```
Column   Row   Sheet
```

Prompts Before/After and count.

### /Worksheet Delete  (D)  **[MVP]**

```
Column   Row   Sheet   File
```

`File` removes the active file from memory.

### /Worksheet Column  (C)  **[MVP]**

```
Set-Width   Reset-Width   Hide   Display   Column-Range
```

- Column-Range → Set-Width | Reset-Width (apply width across columns)

### /Worksheet Erase  (E)  **[MVP]**

```
No   Yes
```

Clears all active files from memory; leaves one blank.

### /Worksheet Titles  (T)  **[MVP]**

```
Both   Horizontal   Vertical   Clear
```

Freeze panes at pointer.

### /Worksheet Window  (W)  **[CPL]**

```
Horizontal  Vertical  Sync  Unsync  Clear  Perspective  Map  Graph
```

- **Perspective** — stacked oblique view of 3 sheets
- **Map** — glyph display (`"` label, `#` number, `+` formula)
- **Graph** — live graph pane

### /Worksheet Status  (S)  **[MVP]**

Display STAT screen: memory, recalc mode, circular refs, coprocessor, formats.

### /Worksheet Page  (P)  **[CPL]**

Insert manual page break at pointer.

### /Worksheet Hide  (H)  **[CPL]**

```
Enable   Disable
```

Hide entire sheets.

### /Worksheet Learn  (L)  **[CPL]**

```
Range   Cancel   Erase
```

Assign/clear the Learn range (where Alt-F5 recorded keystrokes go).

---

## /Range  (R)

```
Format  Label  Erase  Name  Justify  Prot  Unprot  Input  Value  Trans  Search
```

- **Format** → same format list as /Worksheet Global Format, plus **Reset**  **[MVP]**
- **Label** → Left | Right | Center (change prefix on existing labels)  **[MVP]**
- **Erase**  **[MVP]**
- **Name** → Create | Delete | Labels (Right/Down/Left/Up) | Reset | Table | Undefine | Note (Create/Delete/Reset/Table/Quit)  **[MVP: Create, Delete, Labels, Reset, Table]**
- **Justify** (word-wrap long label into block)  **[MVP]**
- **Prot** | **Unprot**  **[MVP]**
- **Input** (form-style input limited to unprotected cells)  **[CPL]**
- **Value** (copy formulas → values)  **[CPL]**
- **Trans** (transpose rows↔cols↔sheets; can convert formulas→values)  **[CPL]**
- **Search** → Formulas | Labels | Both → Find | Replace  **[CPL]**

---

## /Copy  (C)  **[MVP]**

Two-step POINT: FROM range, then TO anchor. Relative refs adjust.

---

## /Move  (M)  **[MVP]**

Two-step POINT: FROM range, then TO anchor. Formulas that reference moved
cells are updated.

---

## /File  (F)

```
Retrieve  Save  Combine  Xtract  Erase  List  Import  Dir  New  Open  Admin
```

- **Retrieve**  **[MVP]** — wipes memory; loads one file
- **Save** → (for first save: prompt filename) → Cancel | Replace | Backup  **[MVP]**
- **Combine** → Copy | Add | Subtract → Entire-File | Named/Specified-Range  **[CPL]**
- **Xtract** → Formulas | Values → Cancel | Replace  **[MVP]**
- **Erase** → Worksheet | Print | Graph | Other  **[CPL]**
- **List** → Worksheet | Print | Graph | Other | Active | Linked  **[MVP: Worksheet, Active]**
- **Import** → Text | Numbers  **[MVP: Numbers (CSV)]**
- **Dir** (change session directory)  **[MVP]**
- **New** → Before | After  **[MVP]**
- **Open** → Before | After  **[MVP]**
- **Admin** → Reservation (Get/Release) | Seal (File/Reservation-Setting/Disable) | Table (W/P/G/O/Active/Linked) | Link-Refresh  **[STR]**

---

## /Print  (P)

```
Printer  File  Encoded  Cancel  Hold  Resume  Suspend
```

- **Printer** → …
- **File** → (same submenu)  **[MVP]**
- **Encoded** → (same submenu)  **[STR]**

Each of Printer/File/Encoded shares the submenu:

```
Range  Line  Page  Options  Clear  Align  Go  Quit
```

- **Range**  **[MVP]** (comma-sep list; `*GRAPHNAME` to embed a graph)
- **Line**, **Page**  **[MVP]**
- **Options**:
  - Header, Footer (`|` splits L/C/R; `#` page; `@` date; `\name`)
  - Margins (Left 0-1000, Right, Top, Bottom)
  - Pg-Length (1-1000, default 66)
  - Borders → Columns | Rows | Frame | No-Frame | All | Range | Clear
  - Setup (printer escape sequence)
  - Other → As-Displayed | Cell-Formulas | Formatted | Unformatted | Blank-Header (Print/Suppress)
  - Name → Create | Use | Delete | Reset | Table
  - Advanced → AutoLf | Color | Device | Fonts | Images | Layout | Priority | Wait
  - Quit
- **Clear** → All | Range | Borders | Format | Image | Device
- **Align** (page counter ← 1)
- **Go**

MVP Print scope: `/Print File Range … Options Margins Pg-Length Header Footer Other As-Displayed/Formatted Go` and the surrounding structure. `Printer` and `Encoded` are menu-level placeholders.

---

## /Graph  (G)  **[CPL]**

```
Type  X  A  B  C  D  E  F  Reset  View  Save  Options  Name  Group  Quit
```

- **Type** → Line | Bar | XY | Stack-Bar | Pie | HLCO | Mixed | Features (Vertical/Horizontal, Stacked, 100%, 2Y-Ranges A-F, Y-Ranges A-F)
- **X**, **A**..**F** (data ranges)
- **Reset** → Graph | X | A-F | Ranges | Options | Quit
- **View** (full-screen; same as F10)
- **Save** (write `.CGM` or `.PIC`)
- **Options** → Legend, Format (Lines/Symbols/Both/Neither/Area), Titles, Grid, Scale, Color, B&W, Data-Labels, Advanced
- **Name** → Use | Create | Delete | Reset | Table
- **Group** → Columnwise | Rowwise

---

## /Data  (D)  **[CPL]**

```
Fill  Table  Sort  Query  Distribution  Matrix  Regression  Parse  External
```

- **Fill**  **[CPL]** (numbers, dates, times)
- **Table** → 1 | 2 | 3 | Labeled | Reset  **[CPL: 1, 2]**
- **Sort** → Data-Range | Primary-Key | Secondary-Key | Extra-Key | Reset | Go | Quit  **[CPL]**
- **Query** → Input | Criteria | Output | Find | Extract | Unique | Del | Modify | Reset | Quit  **[CPL]**
- **Distribution**  **[CPL]**
- **Matrix** → Invert | Multiply  **[STR]**
- **Regression**  **[CPL]**
- **Parse**  **[CPL]**
- **External** → Use | List | Create | Delete | Other | Reset | Quit  **[STR]**

---

## /System  (S)  **[MVP]**

Suspend 1-2-3, shell out (`$SHELL` or `cmd.exe`); `exit` returns. On
modern systems this is a proper shell with the alt-screen stashed.

---

## /Add-In  (A)  **[STR]**

```
Load  Remove  Invoke  Clear  Table  Settings  Quit
```

- **Load** — read `.PLC`/`.ADN` into memory; optional assign to APP1/APP2/APP3/No-Key
- **Remove** | **Invoke** | **Clear**
- **Table** → @Functions | Macros | Applications
- **Settings** → File (Set/Cancel/Quit) | System (Set/Cancel/Directory/Update/Quit)

Open question: L123 can repurpose /Add-In for native Rust plug-ins loaded
as dylibs, or Wasm components. Stretch-goal decision.

---

## /Quit  (Q)  **[MVP]**

```
No   Yes
```

Exit with confirmation.

---

## Implementation notes

- The tree is encoded as a static `&'static MenuNode` in `l123-menu`.
- Each node: `letter: char`, `name: &'static str`, `help: &'static str`,
  body: either `Submenu(&'static [MenuNode])` or `Leaf(Action)`.
- Unimplemented leaves carry `Leaf(Action::NotYet(&'static str))`; the
  interpreter displays the string in control-panel line 3 and refuses to
  mutate.
- Every node's letters must be unique within its parent.

### Letter-uniqueness quirks

Watch out — 1-2-3 kept letter-uniqueness rigorously but some siblings
collide under casual reading:

- `/Worksheet Insert`: Column | Row | Sheet → **C** | **R** | **S**
- `/Worksheet Delete`: Column | Row | Sheet | File → **C** | **R** | **S** | **F**
- `/Worksheet Global`: Format | Label | Col-Width | **P**rot | **Z**ero | **R**ecalc | **D**efault | **G**roup | Quit — all unique
- `/Worksheet Global Default Other`: International | Help | Clock | **U**ndo | **B**eep | **A**dd-In | **E**xpanded-Memory — all unique

Verify the full tree at test time with a walk-and-assert in `l123-menu::tests`.

### Numbered leaves

Where a submenu accepts a numeric argument (Fixed 0-15, Currency 0-15,
Iteration 1-50, Interface 1-9, Printer Setup), the node is still a single
node; it prompts on entry rather than branching per digit.

### The Release-2.x-only commands

Release 3.x retained everything from Release 2 and added these that did
not exist in R2: `/Worksheet Insert Sheet`, `/Worksheet Delete Sheet`,
`/Worksheet Global Group`, `/Worksheet Window Perspective`, `/File Open`,
`/File Admin`, `/Data Table 3`, `/Data Table Labeled`, `/Data External`,
and the @ functions marked **♦** in `AT_FUNCTIONS.md`. R3.4a additionally
promotes the WYSIWYG add-in to always-on and ships the 17-icon R3.4 icon
panel.
