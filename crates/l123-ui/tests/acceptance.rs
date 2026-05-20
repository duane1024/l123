//! Acceptance transcript harness.
//!
//! Reads `tests/acceptance/*.tsv` from the workspace root and drives an
//! `App` through the directives. Format documented in
//! `tests/acceptance/README.md`.

use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers, MouseButton, MouseEvent, MouseEventKind};
use l123_ui::App;

/// Transcripts share process CWD (set per transcript) and write to
/// `target/`. `/FD` mutates CWD; parallel tests racing on CWD or on
/// file-existence checks (`/FS` Cancel/Replace branching) are the
/// cause of historical flakes. Serializing the transcripts is cheap
/// (tests take <200ms total) and removes the race outright.
static ACCEPTANCE_LOCK: Mutex<()> = Mutex::new(());

fn parse_hex_rgb(s: &str) -> Option<(u8, u8, u8)> {
    let s = s.trim().strip_prefix('#').unwrap_or(s.trim());
    if s.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&s[0..2], 16).ok()?;
    let g = u8::from_str_radix(&s[2..4], 16).ok()?;
    let b = u8::from_str_radix(&s[4..6], 16).ok()?;
    Some((r, g, b))
}

fn workspace_root() -> PathBuf {
    // CARGO_MANIFEST_DIR is crates/l123-ui; workspace root is two up.
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf()
}

fn run_transcript(path: &Path) {
    // Serialize across transcripts — see ACCEPTANCE_LOCK. Poison is
    // ignored: a panicked test elsewhere shouldn't block later runs.
    let _guard = ACCEPTANCE_LOCK.lock().unwrap_or_else(|p| p.into_inner());
    // Run each transcript with CWD set to the workspace root. /FD
    // mutates CWD so the lock above is what actually keeps
    // concurrent writes correct.
    let _ = std::env::set_current_dir(workspace_root());

    // Per-transcript scratch dir under std::env::temp_dir(). Tests
    // that need to write files use the `$TMPDIR` placeholder, which
    // the harness substitutes into directive arguments (filenames,
    // ASSERT_FILE_* paths). Avoids leaving artifacts under `target/`
    // on every run.
    let test_name = path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("transcript");
    let tmp = std::env::temp_dir().join(format!("l123_accept_{test_name}"));
    let _ = std::fs::remove_dir_all(&tmp);
    std::fs::create_dir_all(&tmp).unwrap_or_else(|e| panic!("mkdir {}: {e}", tmp.display()));
    let tmp_str = tmp.to_string_lossy().into_owned();

    let body = fs::read_to_string(path).unwrap_or_else(|e| panic!("read {}: {e}", path.display()));

    let mut app = App::new();
    let mut width: u16 = 80;
    let mut height: u16 = 25;

    for (ln, raw) in body.lines().enumerate() {
        let line_no = ln + 1;
        let stripped = strip_comment(raw);
        let expanded: String = stripped.replace("$TMPDIR", &tmp_str);
        let line = expanded.trim();
        if line.is_empty() {
            continue;
        }
        let (directive, rest) = split_directive(line);
        // §4.7 — drain any queued async op before the next directive
        // runs, blocking the harness until the worker completes so
        // sequential KEY-then-ASSERT_CELL flows see the post-op
        // state. Gated by `block_next_async_op` so transcripts that
        // assert mid-flight WAIT state stay deterministic.
        app.drain_async_op_blocking();
        match directive {
            // ---- keystrokes ----
            "KEY" => press_char(&mut app, rest, line_no, path),
            "KEYS" => {
                for c in rest.chars() {
                    app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
                }
            }
            "ENTER" => app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE)),
            "ESC" => app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE)),
            "SPACE" => app.handle_key(KeyEvent::new(KeyCode::Char(' '), KeyModifiers::NONE)),
            "TAB" => app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::NONE)),
            "BACKSPACE" => app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE)),
            "DEL" => app.handle_key(KeyEvent::new(KeyCode::Delete, KeyModifiers::NONE)),
            "UP" => app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE)),
            "DOWN" => app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE)),
            "LEFT" => app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE)),
            "RIGHT" => app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE)),
            "HOME" => app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE)),
            "END" => app.handle_key(KeyEvent::new(KeyCode::End, KeyModifiers::NONE)),
            "PGUP" => app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::NONE)),
            "PGDN" => app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::NONE)),
            "CTRL_PGUP" => app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::CONTROL)),
            "CTRL_PGDN" => app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::CONTROL)),
            "CTRL_END" => app.handle_key(KeyEvent::new(KeyCode::End, KeyModifiers::CONTROL)),
            "F" => {
                let n: u8 = rest.parse().expect("F directive needs number");
                app.handle_key(KeyEvent::new(KeyCode::F(n), KeyModifiers::NONE));
            }
            "CTRL" => {
                let c = single_char(rest, line_no, path);
                app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::CONTROL));
            }
            "ALT" => {
                let c = single_char(rest, line_no, path);
                app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::ALT));
            }
            "ALT_F" => {
                let n: u8 = rest.parse().expect("ALT_F directive needs number");
                app.handle_key(KeyEvent::new(KeyCode::F(n), KeyModifiers::ALT));
            }
            // SPEC §7: Ctrl-Break is the canonical abort key. Crossterm
            // models the Pause/Break key as KeyCode::Pause; the CONTROL
            // modifier disambiguates Break from a plain Pause.
            "CTRL_BREAK" => {
                app.handle_key(KeyEvent::new(KeyCode::Pause, KeyModifiers::CONTROL));
            }
            // §4.7 test hooks. The next async file op the App spawns
            // will park itself until RESUME_OP fires, letting the
            // transcript observe mid-flight WAIT state. No-op outside
            // tests; production ops drain on their own schedule.
            "BLOCK_NEXT_OP" => app.test_block_next_async_op(),
            "RESUME_OP" => app.test_resume_async_op(),
            // §4.7 — lower (or raise) the F9-recalc cell-count
            // threshold. Lets a small transcript exercise the
            // recalc-on-50k-cells WAIT path without actually
            // populating 50k cells.
            "RECALC_WAIT_THRESHOLD" => {
                let n: usize = rest.parse().expect("RECALC_WAIT_THRESHOLD needs number");
                app.test_set_recalc_wait_threshold(n);
            }
            // §4.7 — write synthetic done/total progress numbers
            // onto the currently-pending async op so transcripts
            // can assert the `[████░░] N%` bar shape without
            // timing the worker. Format: "SEED_PROGRESS <done> <total>".
            "SEED_PROGRESS" => {
                let mut parts = rest.split_whitespace();
                let done: u64 = parts
                    .next()
                    .and_then(|s| s.parse().ok())
                    .expect("SEED_PROGRESS needs <done>");
                let total: u64 = parts
                    .next()
                    .and_then(|s| s.parse().ok())
                    .expect("SEED_PROGRESS needs <total>");
                app.test_seed_async_progress(done, total);
            }
            "MACRO" => app.run_macro_text(rest),
            // M9 v0.4 — pin the .l123log sidecar path for the next
            // LEARN session so the transcript can assert / replay it
            // without going through /File Save.
            "SET_LEARN_SIDECAR_PATH" => {
                let path = std::path::PathBuf::from(rest);
                app.test_set_learn_sidecar_path(Some(path));
            }
            // M9 v0.4 — read the .l123log sidecar at `path` and
            // dispatch each `{"keys": "..."}` token through the macro
            // interpreter, exactly as `l123 --replay <path>` would.
            "REPLAY" => {
                let path = std::path::PathBuf::from(rest);
                app.replay_sidecar(&path)
                    .unwrap_or_else(|e| panic!("{}:{line_no}: replay: {e}", path.display()));
            }

            // ---- assertions ----
            "ASSERT_POINTER" => {
                let got = app.pointer().display_full();
                assert_eq!(
                    got,
                    rest,
                    "{}:{line_no}: pointer expected {rest} got {got}",
                    path.display()
                );
            }
            "ASSERT_MODE" => {
                let got = app.mode().indicator();
                assert_eq!(
                    got,
                    rest,
                    "{}:{line_no}: mode expected {rest} got {got}",
                    path.display()
                );
            }
            "ASSERT_ENTRY_CURSOR" => {
                let want: usize = rest.parse().unwrap_or_else(|e| {
                    panic!(
                        "{}:{line_no}: ASSERT_ENTRY_CURSOR needs a number, got {rest:?} ({e})",
                        path.display()
                    )
                });
                let got = app.entry_cursor().unwrap_or_else(|| {
                    panic!(
                        "{}:{line_no}: ASSERT_ENTRY_CURSOR with no active entry",
                        path.display()
                    )
                });
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: entry cursor expected {want} got {got}",
                    path.display()
                );
            }
            "ASSERT_PANEL_L1" => assert_panel_line(&app, width, height, 0, rest, path, line_no),
            "ASSERT_PANEL_L2" => assert_panel_line(&app, width, height, 1, rest, path, line_no),
            "ASSERT_PANEL_L3" => assert_panel_line(&app, width, height, 2, rest, path, line_no),
            "ASSERT_STATUS" => {
                let buf = app.render_to_buffer(width, height);
                let text = App::line_text(&buf, height - 1);
                assert!(
                    text.contains(rest),
                    "{}:{line_no}: status line {text:?} does not contain {rest:?}",
                    path.display()
                );
            }
            // Substring search across every row of the rendered buffer.
            // Useful for overlays (e.g. /File List) that render outside
            // the fixed panel lines.
            "ASSERT_SCREEN" => {
                let buf = app.render_to_buffer(width, height);
                let found = (0..height).any(|y| App::line_text(&buf, y).contains(rest));
                assert!(
                    found,
                    "{}:{line_no}: screen does not contain {rest:?}",
                    path.display()
                );
            }
            // "ASSERT_SCREEN_COL <x> <substring>" — substring search
            // down a single buffer column. Used for vertical text such
            // as the Y-Axis title (one character per row).
            "ASSERT_SCREEN_COL" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let x_tok = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let x: u16 = x_tok.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_SCREEN_COL bad x coordinate {x_tok:?}",
                        path.display()
                    )
                });
                let buf = app.render_to_buffer(width, height);
                let col = App::column_text(&buf, x);
                assert!(
                    col.contains(want),
                    "{}:{line_no}: column {x} does not contain {want:?} (got {col:?})",
                    path.display()
                );
            }
            "ASSERT_SCREEN_NOT_CONTAINS" => {
                let buf = app.render_to_buffer(width, height);
                let hit = (0..height).find(|y| App::line_text(&buf, *y).contains(rest));
                assert!(
                    hit.is_none(),
                    "{}:{line_no}: screen unexpectedly contains {rest:?} on row {hit:?}",
                    path.display(),
                );
            }
            "ASSERT_CELL" => {
                // "A:A1  hello" — address then expected trimmed-cell text.
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_rendered_text(&buf, addr).unwrap_or_else(|| {
                    panic!(
                        "{}:{line_no}: address {addr:?} not in viewport",
                        path.display()
                    )
                });
                assert_eq!(
                    got.trim(),
                    want,
                    "{}:{line_no}: cell {addr} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // Assert a cell is not visible in the grid (hidden column or
            // off-viewport). `cell_rendered_text` returns None only for
            // cells that the visible-column layout skipped or that are
            // outside the rendered area.
            "ASSERT_CELL_HIDDEN" => {
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_rendered_text(&buf, rest);
                assert!(
                    got.is_none(),
                    "{}:{line_no}: expected {rest} hidden, got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_TABLES A Table1,Sales" — assert the comma-joined
            // list of table names on a sheet (single letter argument).
            // Use `none` to assert the sheet has no tables.
            "ASSERT_TABLES" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let letter = parts.next().unwrap_or("").trim();
                let want_raw = parts.next().unwrap_or("").trim();
                let want: &str = if want_raw.eq_ignore_ascii_case("none") {
                    ""
                } else {
                    want_raw
                };
                let ch = letter.chars().next().unwrap_or('?');
                let got = app.table_names(ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: tables on sheet {letter} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_STATUS_SHEET_FG FF0000" — assert the fg color of
            // the sheet-letter character in the status-line indicator.
            // Use `none` when the sheet has no tab color (the letter
            // renders DarkGray).
            "ASSERT_STATUS_SHEET_FG" => {
                let want = rest.trim();
                let buf = app.render_to_buffer(width, height);
                let got = app.status_sheet_letter_fg(&buf);
                if want.eq_ignore_ascii_case("none") {
                    assert!(
                        got.is_none(),
                        "{}:{line_no}: expected no sheet-letter tint, got {got:?}",
                        path.display()
                    );
                } else {
                    let want_rgb = parse_hex_rgb(want).unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: ASSERT_STATUS_SHEET_FG bad hex {want:?}",
                            path.display()
                        )
                    });
                    assert_eq!(
                        got,
                        Some(want_rgb),
                        "{}:{line_no}: sheet-letter fg expected {want:?} got {got:?}",
                        path.display()
                    );
                }
            }
            // "ASSERT_CELL_RIGHT_GLYPH A:A1 │" — assert the character
            // painted at the rightmost column of a cell's slot.  Used
            // for xlsx-imported right-border glyphs.  Use the literal
            // token `SPACE` to assert a blank (the parser trims, so
            // a literal " " can't be carried here).
            "ASSERT_CELL_RIGHT_GLYPH" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let raw = parts.next().unwrap_or("").trim();
                let want = if raw == "SPACE" { " " } else { raw };
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_right_edge_char(&buf, addr).unwrap_or_else(|| {
                    panic!("{}:{line_no}: cell {addr} not in viewport", path.display())
                });
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: cell {addr} right-edge expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_CELL_FG A:A1 FF0000" — assert the rendered
            // buffer's foreground color at the cell.  Use `none` for
            // "no explicit RGB fg" (terminal default).
            "ASSERT_CELL_FG" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_fg_rendered(&buf, addr);
                if want.eq_ignore_ascii_case("none") {
                    assert!(
                        got.is_none(),
                        "{}:{line_no}: cell {addr} expected no fg, got {got:?}",
                        path.display()
                    );
                } else {
                    let want_rgb = parse_hex_rgb(want).unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: ASSERT_CELL_FG bad hex {want:?}",
                            path.display()
                        )
                    });
                    assert_eq!(
                        got,
                        Some(want_rgb),
                        "{}:{line_no}: cell {addr} fg expected {want:?} got {got:?}",
                        path.display()
                    );
                }
            }
            // "ASSERT_CELL_STRIKE A:A1 true" — assert whether the cell
            // renders with the CROSSED_OUT modifier.  Argument is
            // "true" / "false".
            "ASSERT_CELL_STRIKE" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let want_b: bool = want.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_CELL_STRIKE expected true/false, got {want:?}",
                        path.display()
                    )
                });
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_strike_rendered(&buf, addr);
                assert_eq!(
                    got,
                    want_b,
                    "{}:{line_no}: cell {addr} strike expected {want_b} got {got}",
                    path.display()
                );
            }
            // "ASSERT_CELL_BG A:A1 FF0000" — assert the rendered
            // buffer's background color at the cell.  Color is an
            // uppercase 6-char hex RGB.  Use `none` when the cell
            // should render with the terminal default (no explicit
            // background).
            "ASSERT_CELL_BG" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let buf = app.render_to_buffer(width, height);
                let got = app.cell_bg_rendered(&buf, addr);
                if want.eq_ignore_ascii_case("none") {
                    assert!(
                        got.is_none(),
                        "{}:{line_no}: cell {addr} expected no bg, got {got:?}",
                        path.display()
                    );
                } else {
                    let want_rgb = parse_hex_rgb(want).unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: ASSERT_CELL_BG bad hex {want:?}",
                            path.display()
                        )
                    });
                    assert_eq!(
                        got,
                        Some(want_rgb),
                        "{}:{line_no}: cell {addr} bg expected {want:?} got {got:?}",
                        path.display()
                    );
                }
            }
            // "ASSERT_CELL_STYLE A:A1  Bold Italic" — assert the cell's
            // WYSIWYG text-style override. Use the literal word `plain`
            // (or an empty trailer) when the cell should have no style
            // entry at all.  Non-plain expectations use the marker
            // names exactly as they appear on control-panel line 1
            // (`Bold`, `Italic`, `Underline`, space-joined in that order).
            "ASSERT_CELL_STYLE" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want_raw = parts.next().unwrap_or("").trim();
                let got = app.cell_text_style(addr);
                if want_raw.is_empty() || want_raw.eq_ignore_ascii_case("plain") {
                    assert!(
                        got.is_none(),
                        "{}:{line_no}: cell {addr} expected plain, got {got:?}",
                        path.display()
                    );
                } else {
                    let got_marker = got
                        .and_then(|s| s.marker())
                        .unwrap_or_else(|| "plain".to_string());
                    assert_eq!(
                        got_marker,
                        want_raw,
                        "{}:{line_no}: cell {addr} style expected {want_raw:?} got {got_marker:?}",
                        path.display()
                    );
                }
            }
            // Current (unnamed) graph's type. Use an ASCII all-caps
            // token: LINE | BAR | XY | STACK | PIE | HLCO | MIXED.
            "ASSERT_GRAPH_TYPE" => {
                let got = app.graph_type_str();
                assert_eq!(
                    got,
                    rest,
                    "{}:{line_no}: graph type expected {rest} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SERIES X  A:A1..A:A3" — slot letter, then
            // expected range text. Use the literal word `none` (or an
            // empty trailer) to assert the slot is unset.
            "ASSERT_GRAPH_SERIES" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let slot = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want_raw = parts.next().unwrap_or("").trim();
                let want = if want_raw.eq_ignore_ascii_case("none") {
                    ""
                } else {
                    want_raw
                };
                let got = app.graph_series_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph series {slot} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_NAMES_COUNT 2" — number of named graphs
            // stored on the current workbook (Workbook::graphs).
            "ASSERT_GRAPH_NAMES_COUNT" => {
                let want: usize = rest.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_GRAPH_NAMES_COUNT expects an integer, got {rest:?}",
                        path.display()
                    )
                });
                let got = app.graph_names_count();
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph names count expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_MODE Y AUTO" — axis token (Y, X, or 2)
            // and expected mode (AUTO | MANUAL).
            "ASSERT_GRAPH_SCALE_MODE" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want = parts.next().unwrap_or("").trim();
                let got = app.graph_scale_mode_str(axis_ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale mode {axis_ch} expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_BOUND Y LOWER 100" — axis token
            // (Y, X, or 2), bound kind (LOWER | UPPER), expected
            // numeric value (or `none` for unset).
            "ASSERT_GRAPH_SCALE_BOUND" => {
                let mut parts = rest.split_whitespace();
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let kind_tok = parts.next().unwrap_or("");
                let want_raw = parts.next().unwrap_or("");
                let want: Option<f64> = if want_raw.eq_ignore_ascii_case("none") {
                    None
                } else {
                    Some(want_raw.parse().unwrap_or_else(|_| {
                        panic!(
                            "{}:{line_no}: ASSERT_GRAPH_SCALE_BOUND expected number or 'none', got {want_raw:?}",
                            path.display()
                        )
                    }))
                };
                let upper = matches!(kind_tok.to_ascii_uppercase().as_str(), "UPPER");
                let got = app.graph_scale_bound(axis_ch, upper);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale bound {axis_ch} {kind_tok} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_SKIP 5" — current skip factor.
            "ASSERT_GRAPH_SCALE_SKIP" => {
                let want: u32 = rest.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_GRAPH_SCALE_SKIP expects an integer, got {rest:?}",
                        path.display()
                    )
                });
                let got = app.graph_scale_skip();
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph skip expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_DATA_LABELS A  A:A1..A:A5" — slot letter
            // A..F then expected range string. Use `none` (or empty
            // trailer) for an unset slot.
            "ASSERT_GRAPH_DATA_LABELS" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let slot_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want_raw = parts.next().unwrap_or("").trim();
                let want = if want_raw.eq_ignore_ascii_case("none") {
                    ""
                } else {
                    want_raw
                };
                let slot = match slot_ch.to_ascii_uppercase() {
                    'A' => 0,
                    'B' => 1,
                    'C' => 2,
                    'D' => 3,
                    'E' => 4,
                    'F' => 5,
                    _ => panic!(
                        "{}:{line_no}: ASSERT_GRAPH_DATA_LABELS bad slot {slot_ch:?}",
                        path.display()
                    ),
                };
                let got = app.graph_data_labels_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph data-labels {slot_ch} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_INDICATOR Y Yes" — axis token (Y,
            // X, or 2) and expected magnitude indicator setting.
            "ASSERT_GRAPH_SCALE_INDICATOR" => {
                let mut parts = rest.split_whitespace();
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want = parts.next().unwrap_or("");
                let got = app.graph_scale_indicator_str(axis_ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale indicator {axis_ch} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_EXPONENT Y 3" — axis token (Y, X,
            // or 2) and expected order-of-magnitude shift.
            "ASSERT_GRAPH_SCALE_EXPONENT" => {
                let mut parts = rest.split_whitespace();
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want_raw = parts.next().unwrap_or("");
                let want: i8 = want_raw.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_GRAPH_SCALE_EXPONENT expects an integer, got {want_raw:?}",
                        path.display()
                    )
                });
                let got = app.graph_scale_exponent(axis_ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale exponent {axis_ch} expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_WIDTH Y 8" — axis token (Y, X, or
            // 2) and expected scale-label max width.
            "ASSERT_GRAPH_SCALE_WIDTH" => {
                let mut parts = rest.split_whitespace();
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want_raw = parts.next().unwrap_or("");
                let want: u8 = want_raw.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_GRAPH_SCALE_WIDTH expects an integer, got {want_raw:?}",
                        path.display()
                    )
                });
                let got = app.graph_scale_width(axis_ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale width {axis_ch} expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_SCALE_TYPE Y Linear" — axis token (Y, X,
            // or 2) and expected scale type ("Linear" or
            // "Logarithmic").
            "ASSERT_GRAPH_SCALE_TYPE" => {
                let mut parts = rest.split_whitespace();
                let axis_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want = parts.next().unwrap_or("");
                let got = app.graph_scale_type_str(axis_ch);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph scale type {axis_ch} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_GRID_Y_AXIS Y" — current y-axis grid
            // origin (one of `none`, `Y`, `2Y`, `Both`).
            "ASSERT_GRAPH_GRID_Y_AXIS" => {
                let want = rest.trim();
                let got = app.graph_grid_y_axis_str();
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: grid y-axis expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_DATA_LABELS_PLACEMENT A  Center" — slot
            // letter A..F, then expected placement (Center | Left |
            // Above | Right | Below).
            "ASSERT_GRAPH_DATA_LABELS_PLACEMENT" => {
                let mut parts = rest.split_whitespace();
                let slot_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want = parts.next().unwrap_or("");
                let slot = match slot_ch {
                    'A' | 'a' => 0,
                    'B' | 'b' => 1,
                    'C' | 'c' => 2,
                    'D' | 'd' => 3,
                    'E' | 'e' => 4,
                    'F' | 'f' => 5,
                    _ => panic!(
                        "{}:{line_no}: ASSERT_GRAPH_DATA_LABELS_PLACEMENT bad slot {slot_ch:?}",
                        path.display()
                    ),
                };
                let got = app.graph_data_labels_placement_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: data-labels placement {slot_ch} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_LEGEND A  Net Sales" — slot letter A..F
            // then expected text. Use `none` (or empty trailer) for
            // an unset legend.
            "ASSERT_GRAPH_LEGEND" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let slot_ch = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want_raw = parts.next().unwrap_or("").trim();
                let want = if want_raw.eq_ignore_ascii_case("none") {
                    None
                } else {
                    Some(want_raw)
                };
                let slot = match slot_ch.to_ascii_uppercase() {
                    'A' => 0,
                    'B' => 1,
                    'C' => 2,
                    'D' => 3,
                    'E' => 4,
                    'F' => 5,
                    _ => panic!(
                        "{}:{line_no}: ASSERT_GRAPH_LEGEND bad slot {slot_ch:?}",
                        path.display()
                    ),
                };
                let got = app.graph_legend_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph legend {slot_ch} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_TITLE First  Sales 1991" — slot token then
            // the expected text. Tokens: First, Second, X, Y, 2Y, Note,
            // Other. Use the literal `none` (or empty trailer) for unset.
            "ASSERT_GRAPH_TITLE" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let slot_tok = parts.next().unwrap_or("");
                let want_raw = parts.next().unwrap_or("").trim();
                let want = if want_raw.eq_ignore_ascii_case("none") {
                    None
                } else {
                    Some(want_raw)
                };
                use l123_ui::GraphTitleSlot;
                let slot = match slot_tok {
                    "First" => GraphTitleSlot::First,
                    "Second" => GraphTitleSlot::Second,
                    "X" => GraphTitleSlot::XAxis,
                    "Y" => GraphTitleSlot::YAxis,
                    "2Y" => GraphTitleSlot::TwoYAxis,
                    "Note" => GraphTitleSlot::Note,
                    "Other" => GraphTitleSlot::OtherNote,
                    _ => panic!(
                        "{}:{line_no}: ASSERT_GRAPH_TITLE bad slot {slot_tok:?}",
                        path.display()
                    ),
                };
                let got = app.graph_title_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph title {slot_tok} expected {want:?} got {got:?}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_FORMAT A  BOTH" — slot letter A..F, then
            // expected format token (LINES, SYMBOLS, BOTH, NEITHER, AREA).
            "ASSERT_GRAPH_FORMAT" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let slot = parts.next().unwrap_or("").chars().next().unwrap_or(' ');
                let want = parts.next().unwrap_or("").trim();
                let got = app.graph_format_str(slot);
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: graph format {slot} expected {want} got {got}",
                    path.display()
                );
            }
            // "ASSERT_GRAPH_GRID hv" — assert the current graph's /Graph
            // Options Grid mask. Token is the lowercase letters of the
            // active flags (h, v, y) in that order, or `none`.
            "ASSERT_GRAPH_GRID" => {
                let got = app.graph_grid_str();
                assert_eq!(
                    got,
                    rest,
                    "{}:{line_no}: graph grid expected {rest} got {got}",
                    path.display()
                );
            }
            "ASSERT_BEEP_COUNT" => {
                let want: u64 = rest.parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ASSERT_BEEP_COUNT expects an integer, got {rest:?}",
                        path.display()
                    )
                });
                let got = app.beep_count();
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: beep count expected {want} got {got}",
                    path.display()
                );
            }
            "ASSERT_RUNNING" => {
                let want = match rest {
                    "true" => true,
                    "false" => false,
                    other => panic!(
                        "{}:{line_no}: ASSERT_RUNNING expects true|false, got {other}",
                        path.display()
                    ),
                };
                let got = app.is_running();
                assert_eq!(
                    got,
                    want,
                    "{}:{line_no}: running expected {want}, got {got}",
                    path.display()
                );
            }
            "SIZE" => {
                let mut parts = rest.split_ascii_whitespace();
                width = parts.next().unwrap().parse().unwrap();
                height = parts.next().unwrap().parse().unwrap();
            }
            // "SPLASH <user>|<organization>" — flip the startup splash on
            // with test-controlled identity text. The pipe lets user
            // names carry whitespace without a tabs-vs-spaces footgun.
            "SPLASH" => {
                let mut parts = rest.splitn(2, '|');
                let user = parts.next().unwrap_or("").trim().to_string();
                let org = parts.next().unwrap_or("").trim().to_string();
                app.show_splash(user, org);
            }
            // Pre-clean a file on disk so the transcript starts from a
            // known state. Errors (e.g. not-present) are ignored.
            "RM_FILE" => {
                let _ = std::fs::remove_file(rest);
            }
            // Drop the current `App` and rebuild from `App::new()` to
            // simulate a process quit + relaunch. Required for any
            // claim that exercises only-on-load behavior (autoexec,
            // named-range repopulation) — otherwise stale UI maps from
            // the same session let the test pass without exercising
            // the load path.
            "RESET_APP" => {
                app = App::new();
            }
            // "COPY_FILE <src>  <dst>" — copy a binary fixture into the
            // transcript sandbox. Two args separated by ≥2 spaces or a
            // tab so paths with single spaces still parse.
            "COPY_FILE" => {
                let (src, dst) = rest
                    .split_once('\t')
                    .or_else(|| rest.split_once("  "))
                    .unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: COPY_FILE expects `<src>  <dst>` (tab- or 2+space-separated), got {rest:?}",
                            path.display()
                        )
                    });
                let src = src.trim();
                let dst = dst.trim();
                if let Some(parent) = std::path::Path::new(dst).parent() {
                    if !parent.as_os_str().is_empty() {
                        let _ = std::fs::create_dir_all(parent);
                    }
                }
                std::fs::copy(src, dst).unwrap_or_else(|e| {
                    panic!(
                        "{}:{line_no}: COPY_FILE {src} -> {dst} failed: {e}",
                        path.display()
                    )
                });
            }
            // "HOVER_ICON <panel> <slot>" — pin the icon-hover state
            // as if the mouse were over (`panel`, `slot`). The headless
            // render buffer has no real icon panel to hit-test against,
            // so transcripts short-circuit the mouse path and poke the
            // App's hover state directly.
            "HOVER_ICON" => {
                let mut parts = rest.split_ascii_whitespace();
                let panel_n: u8 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: HOVER_ICON expects `<panel 1..7> <slot 0..16>`, got {rest:?}",
                        path.display()
                    )
                });
                let slot: usize = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: HOVER_ICON expects `<panel 1..7> <slot 0..16>`, got {rest:?}",
                        path.display()
                    )
                });
                let panel = l123_graph::Panel::ORDER
                    .get(panel_n.saturating_sub(1) as usize)
                    .copied()
                    .unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: HOVER_ICON panel number must be 1..=7, got {panel_n}",
                            path.display()
                        )
                    });
                app.set_hovered_icon(panel, slot);
            }
            // Clear the hover state set by a prior `HOVER_ICON`.
            "HOVER_CLEAR" => {
                app.clear_hovered_icon();
            }
            // "ICON_CLICK <panel> <slot>" — dispatch the SmartIcon at
            // `(panel, slot)` directly, mirroring a real click. The
            // headless render buffer has no real icon panel to hit-test
            // against, so the harness short-circuits the mouse-coord
            // path and pokes the App's dispatcher.
            "ICON_CLICK" => {
                let mut parts = rest.split_ascii_whitespace();
                let panel_n: u8 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ICON_CLICK expects `<panel 1..7> <slot 0..16>`, got {rest:?}",
                        path.display()
                    )
                });
                let slot: usize = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: ICON_CLICK expects `<panel 1..7> <slot 0..16>`, got {rest:?}",
                        path.display()
                    )
                });
                let panel = l123_graph::Panel::ORDER
                    .get(panel_n.saturating_sub(1) as usize)
                    .copied()
                    .unwrap_or_else(|| {
                        panic!(
                            "{}:{line_no}: ICON_CLICK panel number must be 1..=7, got {panel_n}",
                            path.display()
                        )
                    });
                app.dispatch_icon_for_test(panel, slot);
            }
            // "MOUSE_CLICK <col> <row>" — synthesize a left-button
            // mouse-down at the given terminal coordinates. Grid-click
            // hit-testing needs the last_grid_area cache, which is set
            // by the preceding render pass — so a transcript must have
            // produced at least one ASSERT_* (which triggers a render)
            // before clicking, or the click will be a no-op.
            "MOUSE_CLICK" => {
                let mut parts = rest.split_ascii_whitespace();
                let col: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_CLICK expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                let row: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_CLICK expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                // Prime the grid-area cache by rendering first — this
                // mirrors what the real event loop does between frames.
                let _ = app.render_to_buffer(width, height);
                app.handle_mouse(MouseEvent {
                    kind: MouseEventKind::Down(MouseButton::Left),
                    column: col,
                    row,
                    modifiers: KeyModifiers::NONE,
                });
            }
            // "MOUSE_DRAG <col> <row>" — synthesize a left-button mouse
            // drag (button held, cursor moved). The harness assumes a
            // prior MOUSE_CLICK has primed the grid-area cache; we render
            // again here defensively so a transcript that drags without a
            // preceding click still has geometry to hit-test against.
            "MOUSE_DRAG" => {
                let mut parts = rest.split_ascii_whitespace();
                let col: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_DRAG expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                let row: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_DRAG expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                let _ = app.render_to_buffer(width, height);
                app.handle_mouse(MouseEvent {
                    kind: MouseEventKind::Drag(MouseButton::Left),
                    column: col,
                    row,
                    modifiers: KeyModifiers::NONE,
                });
            }
            // "MOUSE_UP <col> <row>" — synthesize the left-button release
            // ending a drag. Pairs with MOUSE_CLICK / MOUSE_DRAG to
            // exercise the full press → drag → release lifecycle.
            "MOUSE_UP" => {
                let mut parts = rest.split_ascii_whitespace();
                let col: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_UP expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                let row: u16 = parts.next().unwrap_or("").parse().unwrap_or_else(|_| {
                    panic!(
                        "{}:{line_no}: MOUSE_UP expects `<col> <row>`, got {rest:?}",
                        path.display()
                    )
                });
                app.handle_mouse(MouseEvent {
                    kind: MouseEventKind::Up(MouseButton::Left),
                    column: col,
                    row,
                    modifiers: KeyModifiers::NONE,
                });
            }
            // "SCROLL_DOWN" / "SCROLL_UP" — synthesize a scroll-wheel
            // tick. No coordinates: scroll affects the viewport
            // regardless of cursor position. Position 10,10 is fed to
            // the event for completeness; the handler ignores it.
            "SCROLL_DOWN" => {
                app.handle_mouse(MouseEvent {
                    kind: MouseEventKind::ScrollDown,
                    column: 10,
                    row: 10,
                    modifiers: KeyModifiers::NONE,
                });
            }
            "SCROLL_UP" => {
                app.handle_mouse(MouseEvent {
                    kind: MouseEventKind::ScrollUp,
                    column: 10,
                    row: 10,
                    modifiers: KeyModifiers::NONE,
                });
            }
            // "ASSERT_FILE_CONTAINS <path>  <substr>" — split on the
            // first whitespace run. The remainder is matched as a
            // substring inside the file's text contents. `\n`, `\t`,
            // and `\\` in the substring are unescaped so transcripts
            // can express specific whitespace layouts.
            "ASSERT_FILE_CONTAINS" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let fpath = parts.next().unwrap_or("");
                let raw = parts.next().unwrap_or("").trim();
                let want = unescape(raw);
                let body = std::fs::read_to_string(fpath)
                    .unwrap_or_else(|e| panic!("{}:{line_no}: read {fpath}: {e}", path.display()));
                assert!(
                    body.contains(&want),
                    "{}:{line_no}: file {fpath:?} does not contain {want:?}; got {body:?}",
                    path.display()
                );
            }
            "ASSERT_FILE_NOT_EXISTS" => {
                let fpath = rest.trim();
                let exists = std::fs::metadata(fpath).is_ok();
                assert!(
                    !exists,
                    "{}:{line_no}: file {fpath:?} unexpectedly exists",
                    path.display()
                );
            }
            "ASSERT_FILE_EXISTS" => {
                let fpath = rest.trim();
                let exists = std::fs::metadata(fpath).is_ok();
                assert!(
                    exists,
                    "{}:{line_no}: file {fpath:?} expected to exist",
                    path.display()
                );
            }
            "ASSERT_FILE_NOT_CONTAINS" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let fpath = parts.next().unwrap_or("");
                let raw = parts.next().unwrap_or("").trim();
                let want = unescape(raw);
                let body = std::fs::read_to_string(fpath)
                    .unwrap_or_else(|e| panic!("{}:{line_no}: read {fpath}: {e}", path.display()));
                assert!(
                    !body.contains(&want),
                    "{}:{line_no}: file {fpath:?} unexpectedly contains {want:?}; got {body:?}",
                    path.display()
                );
            }
            // Bytes-level variant for binary outputs (PDF). Same
            // substring semantics as ASSERT_FILE_CONTAINS, but reads
            // raw bytes and matches on byte slices — necessary because
            // PDF's magic header is not valid UTF-8 by spec.
            "ASSERT_FILE_BYTES_CONTAIN" => {
                let mut parts = rest.splitn(2, char::is_whitespace);
                let fpath = parts.next().unwrap_or("");
                let raw = parts.next().unwrap_or("").trim();
                let want = unescape(raw);
                let want_bytes = want.as_bytes();
                let body = std::fs::read(fpath)
                    .unwrap_or_else(|e| panic!("{}:{line_no}: read {fpath}: {e}", path.display()));
                let found = want_bytes.len() <= body.len()
                    && body.windows(want_bytes.len()).any(|w| w == want_bytes);
                assert!(
                    found,
                    "{}:{line_no}: file {fpath:?} bytes do not contain {want:?}",
                    path.display()
                );
            }
            other => {
                panic!("{}:{line_no}: unknown directive {other:?}", path.display());
            }
        }
    }

    // Clean up the per-transcript temp dir on successful completion.
    // A panic will skip this; the OS reaps temp dirs eventually.
    let _ = std::fs::remove_dir_all(&tmp);
}

/// Interpret `\n`, `\t`, and `\\` escape sequences inside a directive
/// argument. Used by the `ASSERT_FILE_*` directives to match specific
/// whitespace layouts.
/// Strip a trailing `#`-comment from a transcript line. A `#` is a
/// comment only when it's at the very start of the line or preceded
/// by whitespace — this lets data-bearing directives (KEYS, ASSERT_*)
/// carry literal `#` characters through.
fn strip_comment(raw: &str) -> &str {
    let bytes = raw.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'#' && (i == 0 || bytes[i - 1].is_ascii_whitespace()) {
            return &raw[..i];
        }
        i += 1;
    }
    raw
}

fn unescape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.next() {
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('f') => out.push('\x0c'),
                Some('\\') => out.push('\\'),
                Some(other) => {
                    out.push('\\');
                    out.push(other);
                }
                None => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    out
}

fn split_directive(line: &str) -> (&str, &str) {
    match line.find(|c: char| c.is_ascii_whitespace()) {
        Some(idx) => {
            let (a, b) = line.split_at(idx);
            (a, b.trim_start())
        }
        None => (line, ""),
    }
}

fn press_char(app: &mut App, rest: &str, line_no: usize, path: &Path) {
    let c = single_char(rest, line_no, path);
    app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
}

fn single_char(rest: &str, line_no: usize, path: &Path) -> char {
    let mut chars = rest.chars();
    let c = chars
        .next()
        .unwrap_or_else(|| panic!("{}:{line_no}: missing char argument", path.display()));
    assert!(
        chars.next().is_none(),
        "{}:{line_no}: expected single char, got {rest:?}",
        path.display()
    );
    c
}

fn assert_panel_line(
    app: &App,
    width: u16,
    height: u16,
    y: u16,
    needle: &str,
    path: &Path,
    line_no: usize,
) {
    let buf = app.render_to_buffer(width, height);
    let text = App::line_text(&buf, y);
    assert!(
        text.contains(needle),
        "{}:{line_no}: panel line {y} {text:?} does not contain {needle:?}",
        path.display()
    );
}

// Declare one #[test] per transcript file found at compile time.
macro_rules! transcripts {
    ( $( $name:ident => $file:literal ),* $(,)? ) => {
        $(
            #[test]
            fn $name() {
                let p = workspace_root().join("tests/acceptance").join($file);
                run_transcript(&p);
            }
        )*
    };
}

transcripts! {
    m0_arrow_nav    => "M0_arrow_nav.tsv",
    m0_scroll_down  => "M0_scroll_down.tsv",
    m0_scroll_right => "M0_scroll_right.tsv",
    m0_quit         => "M0_quit.tsv",
    m1_label_entry  => "M1_label_entry.tsv",
    m1_value_entry  => "M1_value_entry.tsv",
    m1_label_prefixes => "M1_label_prefixes.tsv",
    m1_entry_cancel   => "M1_entry_cancel.tsv",
    m1_commit_on_arrow => "M1_commit_on_arrow.tsv",
    m1_edit_f2         => "M1_edit_f2.tsv",
    m1_goto_f5         => "m1_goto_f5.tsv",
    m1_label_mid_buffer_edit => "m1_label_mid_buffer_edit.tsv",
    m1_value_delete_key      => "m1_value_delete_key.tsv",
    m2_formula_entry   => "M2_formula_entry.tsv",
    m2_value_currency               => "m2_value_currency.tsv",
    m2_value_percent                => "m2_value_percent.tsv",
    m2_value_comma                  => "m2_value_comma.tsv",
    m2_value_paren_negate           => "m2_value_paren_negate.tsv",
    m2_value_dollar_paren_negate    => "m2_value_dollar_paren_negate.tsv",
    m2_value_plain_preserves_format => "m2_value_plain_preserves_format.tsv",
    m2_edit_cursor_movement          => "m2_edit_cursor_movement.tsv",
    m2_edit_cursor_backspace_delete  => "m2_edit_cursor_backspace_delete.tsv",
    m2_edit_cursor_insert            => "m2_edit_cursor_insert.tsv",
    m2_f2_during_label               => "m2_f2_during_label.tsv",
    m2_f9_calc         => "M2_f9_calc.tsv",
    m2_format_tag      => "M2_format_tag.tsv",
    m3_menu_navigation => "M3_menu_navigation.tsv",
    m3_quit            => "M3_quit.tsv",
    m3_quit_dirty      => "m3_quit_dirty.tsv",
    m3_insert_delete_row_col => "M3_insert_delete_row_col.tsv",
    m3_range_erase     => "M3_range_erase.tsv",
    m3_range_erase_multi => "m3_range_erase_multi.tsv",
    m3_range_format_multi => "m3_range_format_multi.tsv",
    m3_copy            => "M3_copy.tsv",
    m3_copy_lotus_tutorial => "m3_copy_lotus_tutorial.tsv",
    m3_move            => "M3_move.tsv",
    m3_range_label     => "M3_range_label.tsv",
    m3_range_format    => "M3_range_format.tsv",
    m3_point_typed_range => "m3_point_typed_range.tsv",
    m3_point_named_range => "m3_point_named_range.tsv",
    m3_wg_recalc       => "M3_wg_recalc.tsv",
    m3_ws_col_width    => "M3_ws_col_width.tsv",
    m3_ws_col_reset_width       => "M3_ws_col_reset_width.tsv",
    m3_ws_col_range_set_width   => "M3_ws_col_range_set_width.tsv",
    m3_ws_col_range_reset_width => "M3_ws_col_range_reset_width.tsv",
    m3_ws_col_hide     => "M3_ws_col_hide.tsv",
    m3_ws_col_display  => "M3_ws_col_display.tsv",
    m3_ws_erase_confirm => "M3_ws_erase_confirm.tsv",
    m3_range_format_date => "M3_range_format_date.tsv",
    m3_wg_col_width    => "M3_wg_col_width.tsv",
    m3_wg_label        => "M3_wg_label.tsv",
    m3_range_name      => "M3_range_name.tsv",
    m3_f3_names_in_point        => "m3_f3_names_in_point.tsv",
    m3_f3_names_in_goto         => "m3_f3_names_in_goto.tsv",
    m3_f3_names_in_name_delete  => "m3_f3_names_in_name_delete.tsv",
    m3_f3_names_empty           => "m3_f3_names_empty.tsv",
    m3_range_name_reset         => "m3_range_name_reset.tsv",
    m3_range_name_labels        => "m3_range_name_labels.tsv",
    m3_range_name_table         => "m3_range_name_table.tsv",
    m3_range_name_undefine      => "m3_range_name_undefine.tsv",
    m3_range_name_note          => "m3_range_name_note.tsv",
    m3_range_prot               => "m3_range_prot.tsv",
    m3_range_input              => "m3_range_input.tsv",
    m3_range_value              => "m3_range_value.tsv",
    m3_range_trans              => "m3_range_trans.tsv",
    m3_range_justify            => "m3_range_justify.tsv",
    m3_range_format_hidden      => "m3_range_format_hidden.tsv",
    m3_range_format_time        => "m3_range_format_time.tsv",
    m3_range_format_plus_minus  => "m3_range_format_plus_minus.tsv",
    m3_range_format_other_automatic => "m3_range_format_other_automatic.tsv",
    m3_range_format_other_label     => "m3_range_format_other_label.tsv",
    m3_range_format_parens          => "m3_range_format_parens.tsv",
    m3_range_format_neg_color       => "m3_range_format_neg_color.tsv",
    m3_range_format_text_shows_formula => "m3_range_format_text_shows_formula.tsv",
    m3_range_format_label_only      => "m3_range_format_label_only.tsv",
    m3_range_format_label_only_input => "m3_range_format_label_only_input.tsv",
    m3_xlsx_phase2_round_trip       => "m3_xlsx_phase2_round_trip.tsv",
    m3_beep_edge       => "M3_beep_edge.tsv",
    m4_file_save       => "M4_file_save.tsv",
    m4_file_save_replace => "M4_file_save_replace.tsv",
    m4_file_retrieve   => "M4_file_retrieve.tsv",
    m4_file_retrieve_error => "M4_file_retrieve_error.tsv",
    m4_file_retrieve_csv   => "M4_file_retrieve_csv.tsv",
    m4_file_xtract     => "M4_file_xtract.tsv",
    m4_file_import_numbers => "M4_file_import_numbers.tsv",
    m4_file_import_text    => "M4_file_import_text.tsv",
    m4_file_erase          => "M4_file_erase.tsv",
    m4_file_combine        => "M4_file_combine.tsv",
    m4_file_new        => "M4_file_new.tsv",
    m4_file_new_clears_formatting => "M4_file_new_clears_formatting.tsv",
    m4_file_dir        => "M4_file_dir.tsv",
    m4_file_list_active => "M4_file_list_active.tsv",
    m4_file_list_other => "M4_file_list_other.tsv",
    m4_file_list_worksheet => "M4_file_list_worksheet.tsv",
    m4_wait_progress       => "M4_wait_progress.tsv",
    m4_wait_ctrl_break     => "M4_wait_ctrl_break.tsv",
    m4_wait_save           => "M4_wait_save.tsv",
    m4_wait_import_numbers => "M4_wait_import_numbers.tsv",
    m4_wait_import_text    => "M4_wait_import_text.tsv",
    m4_wait_recalc         => "M4_wait_recalc.tsv",
    m4_wait_progress_bar   => "M4_wait_progress_bar.tsv",
    m5_insert_sheet    => "M5_insert_sheet.tsv",
    m5_delete_sheet    => "M5_delete_sheet.tsv",
    m5_delete_file     => "M5_delete_file.tsv",
    m5_group_format    => "M5_group_format.tsv",
    m5_3d_sum          => "M5_3d_sum.tsv",
    m5_file_open       => "M5_file_open.tsv",
    m5_undo            => "M5_undo.tsv",
    m5_undo_toggle     => "M5_undo_toggle.tsv",
    m5_undo_coverage   => "M5_undo_coverage.tsv",
    wgdo_clock         => "wgdo_clock.tsv",
    m5_wg_format       => "M5_wg_format.tsv",
    m5_wg_format_date  => "M5_wg_format_date.tsv",
    m5_wg_format_date_time => "m5_wg_format_date_time.tsv",
    m5_wg_format_hidden    => "m5_wg_format_hidden.tsv",
    m5_wg_format_undo  => "M5_wg_format_undo.tsv",
    m5_wg_format_plus_minus => "m5_wg_format_plus_minus.tsv",
    m5_wg_format_other_automatic => "m5_wg_format_other_automatic.tsv",
    m5_wg_format_other_label     => "m5_wg_format_other_label.tsv",
    m5_ws_titles       => "M5_ws_titles.tsv",
    m5_ws_hide         => "M5_ws_hide.tsv",
    m5_wgd_status            => "M5_wgd_status.tsv",
    m5_wgd_dir_temp_ext      => "M5_wgd_dir_temp_ext.tsv",
    m5_wgd_autoexec_graph    => "M5_wgd_autoexec_graph.tsv",
    m5_wgd_printer           => "M5_wgd_printer.tsv",
    m5_wgdo_intl_punct     => "M5_wgdo_intl_punct.tsv",
    m5_wgdo_intl_currency  => "M5_wgdo_intl_currency.tsv",
    m5_wgdo_intl_negative  => "M5_wgdo_intl_negative.tsv",
    m5_wgdo_intl_date      => "M5_wgdo_intl_date.tsv",
    // Time Intl (D8/D9) rendering is unit-tested in l123-core/format;
    // exposing D8/D9 via /RF Date Time is a separate menu-wiring task.
    m5_wgdo_intl_undo      => "M5_wgdo_intl_undo.tsv",
    m6_print_file      => "M6_print_file.tsv",
    m6_print_multi_range => "m6_print_multi_range.tsv",
    m6_print_options_header => "M6_print_options_header.tsv",
    m6_print_options_setup  => "M6_print_options_setup.tsv",
    m6_print_encoded        => "M6_print_encoded.tsv",
    m6_print_cancel         => "M6_print_cancel.tsv",
    m6_print_advanced_device => "M6_print_advanced_device.tsv",
    m6_print_pipe_row  => "M6_print_pipe_row.tsv",
    m6_ws_page         => "M6_ws_page.tsv",
    m6_ws_page_column  => "M6_ws_page_column.tsv",
    m6_range_search_replace => "M6_range_search_replace.tsv",
    m6_print_cell_formulas  => "M6_print_cell_formulas.tsv",
    m6_print_margins        => "M6_print_margins.tsv",
    m6_print_pagination     => "M6_print_pagination.tsv",
    m6_print_header_tokens  => "M6_print_header_tokens.tsv",
    m6_print_align_clear    => "M6_print_align_clear.tsv",
    m6_print_printer_menu   => "M6_print_printer_menu.tsv",
    m6_print_pdf            => "M6_print_pdf.tsv",
    m6_range_search_find    => "M6_range_search_find.tsv",
    m7_graph_type       => "M7_graph_type.tsv",
    m7_graph_series     => "M7_graph_series.tsv",
    m7_graph_reset      => "M7_graph_reset.tsv",
    m7_graph_view_f10   => "M7_graph_view_f10.tsv",
    graph_view_full_screen => "graph_view_full_screen.tsv",
    m7_graph_save       => "M7_graph_save.tsv",
    graph_settings_visible => "graph_settings_visible.tsv",
    graph_features_toggles => "graph_features_toggles.tsv",
    graph_features_y_axis_and_frame => "graph_features_y_axis_and_frame.tsv",
    graph_options_color => "graph_options_color.tsv",
    graph_options_grid => "graph_options_grid.tsv",
    graph_options_format => "graph_options_format.tsv",
    graph_options_titles => "graph_options_titles.tsv",
    graph_options_legend => "graph_options_legend.tsv",
    graph_options_legend_range => "graph_options_legend_range.tsv",
    graph_reset_leaves => "graph_reset_leaves.tsv",
    graph_options_scale_bounds => "graph_options_scale_bounds.tsv",
    graph_frame_y_axis => "graph_frame_y_axis.tsv",
    graph_name_table => "graph_name_table.tsv",
    graph_data_labels_placement => "graph_data_labels_placement.tsv",
    graph_options_grid_y_axis => "graph_options_grid_y_axis.tsv",
    graph_options_scale_type => "graph_options_scale_type.tsv",
    graph_options_scale_width => "graph_options_scale_width.tsv",
    graph_options_scale_exponent => "graph_options_scale_exponent.tsv",
    graph_options_scale_indicator => "graph_options_scale_indicator.tsv",
    graph_render_data_labels => "graph_render_data_labels.tsv",
    graph_render_data_labels_bar => "graph_render_data_labels_bar.tsv",
    graph_render_data_labels_stack => "graph_render_data_labels_stack.tsv",
    graph_render_data_labels_xy => "graph_render_data_labels_xy.tsv",
    graph_render_data_labels_hlco => "graph_render_data_labels_hlco.tsv",
    graph_render_data_labels_bar_horizontal => "graph_render_data_labels_bar_horizontal.tsv",
    graph_options_data_labels => "graph_options_data_labels.tsv",
    graph_options_scale => "graph_options_scale.tsv",
    graph_options_advanced_shell => "graph_options_advanced_shell.tsv",
    graph_name => "graph_name.tsv",
    graph_group => "graph_group.tsv",
    graph_render_titles => "graph_render_titles.tsv",
    graph_render_legend => "graph_render_legend.tsv",
    graph_render_grid => "graph_render_grid.tsv",
    graph_render_frame => "graph_render_frame.tsv",
    graph_render_orientation => "graph_render_orientation.tsv",
    graph_render_stacked => "graph_render_stacked.tsv",
    graph_render_percent => "graph_render_percent.tsv",
    graph_render_table => "graph_render_table.tsv",
    graph_render_clustered_bar => "graph_render_clustered_bar.tsv",
    graph_render_drop_shadow => "graph_render_drop_shadow.tsv",
    graph_render_notes => "graph_render_notes.tsv",
    graph_render_y_axis_title => "graph_render_y_axis_title.tsv",
    graph_render_pie_labels => "graph_render_pie_labels.tsv",
    graph_render_pie_unicode => "graph_render_pie_unicode.tsv",
    graph_render_xy_unicode => "graph_render_xy_unicode.tsv",
    graph_render_mixed_unicode => "graph_render_mixed_unicode.tsv",
    graph_render_hlco_unicode => "graph_render_hlco_unicode.tsv",
    m10_startup_splash  => "M10_startup_splash.tsv",
    m11_f1_help_open_close => "m11_f1_help_open_close.tsv",
    m11_f1_help_menu_context => "m11_f1_help_menu_context.tsv",
    m10_status_line_filename => "M10_status_line_filename.tsv",
    m10_status_line_dirty    => "m10_status_line_dirty.tsv",
    m10_worksheet_status     => "M10_worksheet_status.tsv",
    m10_label_spill          => "M10_label_spill.tsv",
    m10_wysiwyg_bold         => "M10_wysiwyg_bold.tsv",
    m10_wysiwyg_lines        => "m10_wysiwyg_lines.tsv",
    m10_wysiwyg_compound     => "M10_wysiwyg_compound.tsv",
    m10_wysiwyg_clear        => "M10_wysiwyg_clear.tsv",
    m10_wysiwyg_undo         => "M10_wysiwyg_undo.tsv",
    m10_wysiwyg_panel_marker => "M10_wysiwyg_panel_marker.tsv",
    m10_wysiwyg_alignment    => "M10_wysiwyg_alignment.tsv",
    m10_wysiwyg_color        => "M10_wysiwyg_color.tsv",
    m10_wysiwyg_xlsx_round_trip => "M10_wysiwyg_xlsx_round_trip.tsv",
    m10_wysiwyg_col_width    => "M10_wysiwyg_col_width.tsv",
    m10_wysiwyg_display_mode => "M10_wysiwyg_display_mode.tsv",
    m10_wysiwyg_display_grid => "M10_wysiwyg_display_grid.tsv",
    m10_wysiwyg_special_copy  => "M10_wysiwyg_special_copy.tsv",
    m10_wysiwyg_special_move  => "M10_wysiwyg_special_move.tsv",
    m10_wysiwyg_special_color => "M10_wysiwyg_special_color.tsv",
    m10_xlsx_format_round_trip  => "M10_xlsx_format_round_trip.tsv",
    m10_status_line_sheet       => "M10_status_line_sheet.tsv",
    m10_icon_hover              => "M10_icon_hover.tsv",
    m10_mouse_click_cell        => "M10_mouse_click_cell.tsv",
    m10_mouse_click_point_extend => "M10_mouse_click_point_extend.tsv",
    m10_mouse_click_splice       => "M10_mouse_click_splice.tsv",
    m10_mouse_drag_select        => "M10_mouse_drag_select.tsv",
    m10_mouse_scroll_wheel       => "M10_mouse_scroll_wheel.tsv",
    xlsx_alignment               => "xlsx_alignment.tsv",
    xlsx_fill                    => "xlsx_fill.tsv",
    xlsx_sheet_color             => "xlsx_sheet_color.tsv",
    xlsx_font                    => "xlsx_font.tsv",
    xlsx_auto_contrast           => "auto_contrast.tsv",
    xlsx_borders                 => "xlsx_borders.tsv",
    xlsx_comments                => "xlsx_comments.tsv",
    xlsx_merges                  => "xlsx_merges.tsv",
    xlsx_frozen                  => "xlsx_frozen.tsv",
    xlsx_hidden_sheets           => "xlsx_hidden_sheets.tsv",
    xlsx_tables                  => "xlsx_tables.tsv",
    xlsx_date_formats            => "xlsx_date_formats.tsv",
    t01_tutorial_labels_and_fast_entry => "T01_tutorial_labels_and_fast_entry.tsv",
    t02_tutorial_values_erase_and_repeating_label => "T02_tutorial_values_erase_and_repeating_label.tsv",
    t03_tutorial_calculation_and_named_ranges => "T03_tutorial_calculation_and_named_ranges.tsv",
    t04_tutorial_formatting_and_printing => "T04_tutorial_formatting_and_printing.tsv",
    t05_tutorial_graph_setup_view_save => "T05_tutorial_graph_setup_view_save.tsv",
    t06_tutorial_multiple_sheets_group_and_3d => "T06_tutorial_multiple_sheets_group_and_3d.tsv",
    t07_tutorial_file_retrieve_and_open => "T07_tutorial_file_retrieve_and_open.tsv",
    t08_tutorial_macros               => "T08_tutorial_macros.tsv",
    t09_tutorial_learn_record         => "T09_tutorial_learn_record.tsv",
    t10_tutorial_data_sort            => "T10_tutorial_data_sort.tsv",
    t11_tutorial_data_query           => "T11_tutorial_data_query.tsv",
    t12_tutorial_print_macro          => "T12_tutorial_print_macro.tsv",
    m9_macro_basic_keystrokes => "M9_macro_basic_keystrokes.tsv",
    m9_macro_special_keys     => "M9_macro_special_keys.tsv",
    m9_macro_alt_letter       => "M9_macro_alt_letter.tsv",
    m9_macro_alt_f3_run       => "M9_macro_alt_f3_run.tsv",
    m9_macro_autoexec         => "M9_macro_autoexec.tsv",
    m9_macro_autoexec_restart => "M9_macro_autoexec_restart.tsv",
    m9_macro_branch           => "M9_macro_branch.tsv",
    m9_macro_quit             => "M9_macro_quit.tsv",
    m9_macro_subroutine       => "M9_macro_subroutine.tsv",
    m9_macro_if               => "M9_macro_if.tsv",
    m9_macro_let              => "M9_macro_let.tsv",
    m9_macro_blank            => "M9_macro_blank.tsv",
    m9_macro_question_pause   => "M9_macro_question_pause.tsv",
    m9_macro_getlabel         => "M9_macro_getlabel.tsv",
    m9_macro_menubranch       => "M9_macro_menubranch.tsv",
    m9_macro_beep             => "M9_macro_beep.tsv",
    m9_macro_x_commands       => "M9_macro_x_commands.tsv",
    m9_learn_record           => "M9_learn_record.tsv",
    m9_macro_step             => "M9_macro_step.tsv",
    m8_data_fill              => "M8_data_fill.tsv",
    m8_data_sort              => "M8_data_sort.tsv",
    m8_data_distribution      => "M8_data_distribution.tsv",
    m8_data_regression        => "M8_data_regression.tsv",
    m8_data_matrix            => "M8_data_matrix.tsv",
    m8_data_parse             => "M8_data_parse.tsv",
    m8_data_table_1           => "M8_data_table_1.tsv",
    m8_data_table_2           => "M8_data_table_2.tsv",
    m8_data_sort_extra        => "M8_data_sort_extra.tsv",
    m10_icon_sort             => "M10_icon_sort.tsv",
    m8_data_parse_format_line => "M8_data_parse_format_line.tsv",
    m8_data_query             => "M8_data_query.tsv",
    function_renames    => "function_renames.tsv",
    function_argfix     => "function_argfix.tsv",
    function_emulations => "function_emulations.tsv",
    function_reload_round_trip => "function_reload_round_trip.tsv",
    function_sidecar_round_trip => "function_sidecar_round_trip.tsv",
    m9_learn_sidecar_replay          => "M9_learn_sidecar_replay.tsv",
    m11_import_json                  => "M11_import_json.tsv",
    m11_import_jsonl                 => "M11_import_jsonl.tsv",
    m11_import_json_error            => "M11_import_json_error.tsv",
    m11_import_parquet               => "M11_import_parquet.tsv",
    m11_import_sqlite                => "M11_import_sqlite.tsv",
    m11_range_compare_basic          => "M11_range_compare_basic.tsv",
    m11_range_compare_type_mismatch  => "M11_range_compare_type_mismatch.tsv",
    m11_range_compare_size_mismatch  => "M11_range_compare_size_mismatch.tsv",
    m11_range_compare_clean          => "M11_range_compare_clean.tsv",
    m12_external_sqlite_connect      => "M12_external_sqlite_connect.tsv",
    m12_external_refresh             => "M12_external_refresh.tsv",
    m12_external_list                => "M12_external_list.tsv",
    m12_external_xlsx_roundtrip      => "M12_external_xlsx_roundtrip.tsv",
    m12_external_refresh_wait        => "M12_external_refresh_wait.tsv",
    m12_external_disconnect          => "M12_external_disconnect.tsv",
    m12_external_reset               => "M12_external_reset.tsv",
    m12_external_protected           => "M12_external_protected.tsv",
}

#[cfg(feature = "wk3")]
#[test]
fn wk3_retrieve_saves_as_xlsx() {
    let p = workspace_root().join("tests/acceptance/wk3_retrieve_saves_as_xlsx.tsv");
    run_transcript(&p);
}
