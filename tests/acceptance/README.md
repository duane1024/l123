# L123 Acceptance Transcripts

Each `.tsv` file is a keystroke transcript + assertions that exercise a
slice of the Authenticity Contract (SPEC §20). The harness lives in
`crates/l123-ui/tests/acceptance.rs`.

## Directive format

One directive per line. Whitespace is insignificant between the directive
and its argument (separated by tabs or spaces). `#` starts a comment.

### Keystroke directives

| Directive | Meaning |
|---|---|
| `KEY <char>` | Press a single character (no modifiers) |
| `KEYS <text>` | Press each character of `<text>` in sequence |
| `ENTER` | Press Enter |
| `ESC` | Press Escape |
| `TAB` | Press Tab |
| `BACKSPACE` | Press Backspace |
| `UP` / `DOWN` / `LEFT` / `RIGHT` | Arrow keys |
| `HOME` / `END` | Home / End |
| `PGUP` / `PGDN` | PageUp / PageDown |
| `F <n>` | Function key Fn (1..10) |
| `CTRL <char>` | Ctrl + character |
| `ALT <char>` | Alt + character |
| `ALT_F <n>` | Alt + function key |

### Assertion directives

| Directive | Meaning |
|---|---|
| `ASSERT_POINTER <addr>` | Assert cell pointer is at `<addr>` (e.g. `A:A1`) |
| `ASSERT_MODE <mode>` | Assert mode indicator (`READY`, `LABEL`, …) |
| `ASSERT_PANEL_L1 <substr>` | Assert control panel line 1 contains `<substr>` |
| `ASSERT_PANEL_L2 <substr>` | Assert control panel line 2 contains `<substr>` |
| `ASSERT_PANEL_L3 <substr>` | Assert control panel line 3 contains `<substr>` |
| `ASSERT_STATUS <substr>` | Assert status line contains `<substr>` |
| `ASSERT_SCREEN <substr>` | Assert some row of the rendered buffer contains `<substr>` (useful for overlays like `/File List`) |
| `ASSERT_RUNNING <true\|false>` | Assert app running state |
| `ASSERT_CELL_STYLE <addr> <marker\|plain>` | Assert the cell's WYSIWYG text style. Marker is the space-joined form used on line 1 (`Bold`, `Bold Italic`, `Bold Italic Underline`). Use `plain` or an empty trailer for cells with no style override. |
| `ASSERT_BEEP_COUNT <n>` | Assert the cumulative error-beep count (bumped once per edge-collision or equivalent invalid move; unaffected while beep is disabled). |

### Misc

| Directive | Meaning |
|---|---|
| `SIZE <w> <h>` | Set render buffer size (default 80×25) |
| `RM_FILE <path>` | Delete a file on disk (ignored if missing). Use at the top of a transcript to start from a known state. |
| `ASSERT_FILE_CONTAINS <path>  <substr>` | Assert the named file's text contents contain `<substr>`. Supports `\n`, `\t`, `\f`, `\\` escapes in the substring. |
| `ASSERT_FILE_NOT_CONTAINS <path>  <substr>` | Negation of the above. |
| `HOVER_ICON <panel> <slot>` | Pin the icon-bar hover state as if the mouse were over `(panel, slot)`, where panel is `1..=7` and slot is `0..=16`. Drives control-panel line 3's hover-description tooltip. Transcripts use this instead of synthetic mouse-move events because the headless render buffer doesn't materialize a real icon panel to hit-test against. |
| `HOVER_CLEAR` | Clear the hover state set by a prior `HOVER_ICON`. |
| `MOUSE_CLICK <col> <row>` | Synthesize a left-button mouse-down at the given terminal coordinates. Drives grid click-to-move-pointer (READY mode) and — if you pre-stash the icon panel rect elsewhere — icon panel clicks. The directive renders a frame first so `last_grid_area` is populated, mirroring what the real event loop does between frames. |

### Filesystem sandbox

Each transcript runs inside a scratch directory at
`std::env::temp_dir().join("l123_accept_<transcript_stem>")`. Transcripts
that read or write files (save, retrieve, xtract, import, print) refer
to this directory via the literal placeholder `$TMPDIR`, which the
harness substitutes into every directive argument before the directive
runs. Example:

```
RM_FILE $TMPDIR/out.prn
…
KEYS $TMPDIR/out.prn
ENTER
…
ASSERT_FILE_CONTAINS  $TMPDIR/out.prn  expected text
```

The harness creates the directory fresh at transcript start and cleans
it up on successful completion. Tests that need a wider panel to fit a
long temp-dir path in `ASSERT_PANEL_L*` assertions should include
`SIZE 200 25` (or similar) at the top.

### Comments

A `#` starts a line comment when it is either at the start of the line
or preceded by whitespace — so directives carrying literal `#` in
their arguments (`KEYS P#`) work unescaped.
