# Tutorial-Derived Acceptance Coverage

These transcripts exercise the Lotus 1-2-3 Release 3.1 Tutorial flows
that L123 currently implements, while keeping the suite executable.
The project targets 1-2-3 R3.4a for DOS, but the R3.1 tutorial remains
useful because the core worksheet, menu, graph, print, and file
workflows carry forward.

## Executable Transcripts

- `T01_tutorial_labels_and_fast_entry.tsv` covers Lesson 3 label entry,
  long-label storage, auto apostrophe prefix, and pointer-key commit.
- `T02_tutorial_values_erase_and_repeating_label.tsv` covers Lesson 4
  value entry, range erase by POINT highlighting, and repeating labels.
- `T03_tutorial_calculation_and_named_ranges.tsv` covers Lesson 5
  arithmetic formulas, `@SUM`, recalculation, and named ranges in
  formulas.
- `T04_tutorial_formatting_and_printing.tsv` covers Lesson 6 range
  formatting, global column width, centered labels, row insertion, and
  print-to-file output.
- `T05_tutorial_graph_setup_view_save.tsv` covers Lesson 7 graph data
  ranges, X labels, graph type switching, graph view, and graph save.
- `T06_tutorial_multiple_sheets_group_and_3d.tsv` covers Lessons 10-11
  sheet insertion/navigation, GROUP formatting, and 3D `@SUM`.
- `T07_tutorial_file_retrieve_and_open.tsv` covers file save, retrieve,
  open, and active-file navigation using L123's xlsx persistence path.

## Tutorial Areas Excluded Until Implemented

- F5 GOTO prompts and cursor-position editing inside EDIT prompts.
- Typed range addresses, comma-separated ranges, and F3 name-list
  selection inside POINT prompts.
- Copying one source cell across a larger destination range, 3D copy
  destinations, and relative-reference adjustment when copying formulas.
- Worksheet perspective/window views and physical printer background
  behavior.
- Automatic formatting inferred from typed currency/comma/percent
  literals.
- Graph titles, legends, options, group assignment, and named graphs.
- Data Sort, Query, Fill, Table, Distribution, Regression, and Parse.
- Cross-file formula links, File Admin Link-Refresh, and selective
  active-file deletion.
- Macro entry, naming, Alt-run, STEP debugging, and Record/Learn.
