# Tutorial-Derived Acceptance Coverage

These transcripts exercise the Lotus 1-2-3 Release 3.1 Tutorial flows
that L123 currently implements, while keeping the suite executable.
The project targets 1-2-3 R3.4a for DOS, but the R3.1 tutorial remains
useful because the core worksheet, menu, graph, print, and file
workflows carry forward.

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
- `T07_tutorial_file_retrieve_and_open.tsv` covers file save, retrieve,
  open, and active-file navigation using L123's xlsx persistence path.
- `T08_tutorial_macros.tsv` covers Chapter 4 macro authoring as a label
  cell, `\letter` /Range Name binding to Alt-letter, Alt-run, and
  Alt-F2 STEP single-step debugging with SST status, SPACE advance,
  and Esc abort.
- `T09_tutorial_learn_record.tsv` covers /Worksheet Learn Range plus
  Alt-F5 record/disarm, including the LEARN status indicator and the
  recorded source materializing in the learn range as a label.

The tutorial's "drop one of multiple active files" flow is covered by
the existing `M5_delete_file.tsv` transcript (`/Worksheet Delete File`
removes the foreground file; the survivor takes focus; deleting the
last active file leaves a single blank workbook).

## Tutorial Areas Excluded Until Implemented

- Worksheet perspective/window views and physical printer background
  behavior (`/Worksheet Window` is still a `NotImplemented` leaf).
- Formula entry by *keyboard* pointer-splicing range arguments. Mouse
  splicing during VALUE/EDIT is implemented (see
  `M10_mouse_click_splice.tsv`), but pressing arrow keys mid-formula
  to enter POINT — the tutorial's primary idiom — still commits and
  moves the pointer.
- 3D print ranges. `l123-print` renders a single sheet per page; the
  range's `start.sheet` is the only sheet consulted.
- /Graph titles, legends, options, group assignment, and named
  graphs. Only Type, X/A..F ranges, Reset, View, and Save are wired.
- Data Sort, Query, Fill, Table, Distribution, Regression, and Parse
  (the entire `/Data` submenu is `NotImplemented` leaves).
- Cross-file formula links and `/File Admin Link-Refresh` (the
  `/File Admin` leaves are wired as named actions that close the
  menu without effect).
