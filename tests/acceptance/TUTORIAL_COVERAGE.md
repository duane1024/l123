# Tutorial-Derived Acceptance Coverage

These transcripts exercise the Lotus 1-2-3 Release 3.1 Tutorial flows
that L123 currently implements, while keeping the suite executable.
The project targets 1-2-3 R3.4a for DOS, but the R3.1 tutorial remains
useful because the core worksheet, menu, graph, print, file, sort,
query, and macro workflows carry forward.

## Executable Transcripts

- `T01_tutorial_labels_and_fast_entry.tsv` covers Lesson 3 label entry,
  long-label storage, auto apostrophe prefix, F5 GOTO, F2
  cursor-position editing, and pointer-key commit.
- `T02_tutorial_values_erase_and_repeating_label.tsv` covers Lesson 4
  value entry, typed range erase, repeating labels, single-cell-to-range
  copy, named ranges as command range input, and F3 NAMES selection from
  a command range prompt.
- `T03_tutorial_calculation_and_named_ranges.tsv` covers Lesson 5
  arithmetic formulas, `@SUM`, recalculation, relative-reference formula
  copy, typed range-name creation, F3 GOTO by name, and named ranges in
  formulas.
- `T04_tutorial_formatting_and_printing.tsv` covers Lesson 6 range
  formatting by typed range, automatic currency/comma format inference,
  global column width, centered labels, row insertion, and multi-range
  print-to-file output.
- `T05_tutorial_graph_setup_view_save.tsv` covers Lesson 7 graph data
  ranges and X labels selected by F3 NAMES, graph type switching, graph
  view, and graph save.
- `T06_tutorial_multiple_sheets_group_and_3d.tsv` covers Lessons 10-11
  sheet insertion/navigation, 3D copy destinations, GROUP formatting,
  and 3D `@SUM`.
- `T07_tutorial_file_retrieve_and_open.tsv` covers Lesson 12 file
  save, retrieve, open, and inter-file navigation using L123's xlsx
  persistence path. Both directions of Ctrl-End Ctrl-PgUp / Ctrl-PgDn
  are exercised, including the round-trip back to the prior file.
- `T08_tutorial_macros.tsv` covers Chapter 5 macro authoring as a
  label cell, `\letter` /Range Name binding to Alt-letter, Alt-run,
  Lesson 15's missing-tilde debug flow (run a macro stranded in
  LABEL mode, F2 EDIT the macro cell to restore the closing tilde,
  re-run successfully), and Alt-F2 STEP single-step debugging with
  SST status, SPACE advance, and Esc abort.
- `T09_tutorial_learn_record.tsv` covers /Worksheet Learn Range plus
  Alt-F5 record/disarm, including the LEARN status indicator and the
  recorded source materializing in the learn range as a label.
- `T10_tutorial_data_sort.tsv` covers Lesson 13 /Data Sort: Data-Range,
  Primary-Key with Asc/Desc submenu, Secondary-Key tie-breaking, and
  Go. The Sort menu's stickiness (each leaf returns to it) is
  asserted on every step.
- `T11_tutorial_data_query.tsv` covers Lesson 14 /Data Query: Input,
  Criteria (single-field equality), Output, Find (pointer jumps to
  first matching record), and Extract (matching records copied below
  the output header in input order).
- `T12_tutorial_print_macro.tsv` covers Lesson 16 print macro
  patterns: `'/`-prefixed slash-command labels (which the macro pump
  replays as live menu keystrokes), multi-character macro names like
  `print_data` (vs. the `\letter` Alt-letter form), and Alt-F3 RUN
  to launch them via the NAMES overlay. The transcript writes via
  /Print File rather than the tutorial's /Print Printer so the
  output is verifiable on disk; the slash-in-macro mechanic is
  identical.

The tutorial's "drop one of multiple active files" flow is covered by
the existing `M5_delete_file.tsv` transcript (`/Worksheet Delete File`
removes the foreground file; the survivor takes focus; deleting the
last active file leaves a single blank workbook).

## Tutorial Areas Excluded Until Implemented

- Worksheet perspective/window views and physical printer background
  behavior (`/Worksheet Window` is still a `NotImplemented` leaf).
  This blocks the Lesson 8 what-if /Worksheet Window Graph flow and
  the Lesson 14 ZOOM (Alt-F6) full-screen toggle (T11 stays
  single-sheet rather than splitting the criteria/output ranges
  across worksheets B and C as the tutorial shows).
- Formula entry by *keyboard* pointer-splicing range arguments. Mouse
  splicing during VALUE/EDIT is implemented (see
  `M10_mouse_click_splice.tsv`), but pressing arrow keys mid-formula
  to enter POINT — the tutorial's primary idiom — still commits and
  moves the pointer.
- 3D print ranges. `l123-print` renders a single sheet per page; the
  range's `start.sheet` is the only sheet consulted.
- /Graph titles, legends, options, group assignment, and named
  graphs. Only Type, X/A..F ranges, Reset, View, and Save are wired,
  so the Lesson 7 title/legend/Group flows and all of Lessons 8-9
  (named graphs and graph printing) stay out of scope for tutorial
  transcripts.
- Lesson 14 advanced criteria. The /Data Query matcher is
  equality-only and case-insensitive; logical-formula criteria
  (`+YEARS_EMPLOYED>=3`, `+SALARY>20000`) and the tilde-prefix
  exclusion idiom (`~atlanta` for "not Atlanta") are not implemented,
  so T11 demonstrates only the single-field equality criterion.
- Cross-file formula links and `/File Admin Link-Refresh` (the
  `/File Admin` leaves are wired as named actions that close the
  menu without effect).
