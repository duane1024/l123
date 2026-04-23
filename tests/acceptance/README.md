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
| `ASSERT_RUNNING <true\|false>` | Assert app running state |

### Misc

| Directive | Meaning |
|---|---|
| `SIZE <w> <h>` | Set render buffer size (default 80×25) |
