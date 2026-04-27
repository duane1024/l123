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
    m3_beep_edge       => "M3_beep_edge.tsv",
    m4_file_save       => "M4_file_save.tsv",
    m4_file_save_replace => "M4_file_save_replace.tsv",
    m4_file_retrieve   => "M4_file_retrieve.tsv",
    m4_file_retrieve_error => "M4_file_retrieve_error.tsv",
    m4_file_retrieve_csv   => "M4_file_retrieve_csv.tsv",
    m4_file_xtract     => "M4_file_xtract.tsv",
    m4_file_import_numbers => "M4_file_import_numbers.tsv",
    m4_file_new        => "M4_file_new.tsv",
    m4_file_dir        => "M4_file_dir.tsv",
    m4_file_list_active => "M4_file_list_active.tsv",
    m4_file_list_worksheet => "M4_file_list_worksheet.tsv",
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
    m5_wg_format_undo  => "M5_wg_format_undo.tsv",
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
    m5_wgdo_intl_stat      => "M5_wgdo_intl_stat.tsv",
    m6_print_file      => "M6_print_file.tsv",
    m6_print_multi_range => "m6_print_multi_range.tsv",
    m6_print_options_header => "M6_print_options_header.tsv",
    m6_print_options_setup  => "M6_print_options_setup.tsv",
    m6_print_encoded        => "M6_print_encoded.tsv",
    m6_print_cancel         => "M6_print_cancel.tsv",
    m6_print_advanced_device => "M6_print_advanced_device.tsv",
    m6_print_pipe_row  => "M6_print_pipe_row.tsv",
    m6_ws_page         => "M6_ws_page.tsv",
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
    m7_graph_save       => "M7_graph_save.tsv",
    m10_startup_splash  => "M10_startup_splash.tsv",
    m11_f1_help_open_close => "m11_f1_help_open_close.tsv",
    m10_status_line_filename => "M10_status_line_filename.tsv",
    m10_worksheet_status     => "M10_worksheet_status.tsv",
    m10_label_spill          => "M10_label_spill.tsv",
    m10_wysiwyg_bold         => "M10_wysiwyg_bold.tsv",
    m10_wysiwyg_compound     => "M10_wysiwyg_compound.tsv",
    m10_wysiwyg_clear        => "M10_wysiwyg_clear.tsv",
    m10_wysiwyg_undo         => "M10_wysiwyg_undo.tsv",
    m10_wysiwyg_panel_marker => "M10_wysiwyg_panel_marker.tsv",
    m10_wysiwyg_xlsx_round_trip => "M10_wysiwyg_xlsx_round_trip.tsv",
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
    xlsx_borders                 => "xlsx_borders.tsv",
    xlsx_comments                => "xlsx_comments.tsv",
    xlsx_merges                  => "xlsx_merges.tsv",
    xlsx_frozen                  => "xlsx_frozen.tsv",
    xlsx_hidden_sheets           => "xlsx_hidden_sheets.tsv",
    xlsx_tables                  => "xlsx_tables.tsv",
    t01_tutorial_labels_and_fast_entry => "T01_tutorial_labels_and_fast_entry.tsv",
    t02_tutorial_values_erase_and_repeating_label => "T02_tutorial_values_erase_and_repeating_label.tsv",
    t03_tutorial_calculation_and_named_ranges => "T03_tutorial_calculation_and_named_ranges.tsv",
    t04_tutorial_formatting_and_printing => "T04_tutorial_formatting_and_printing.tsv",
    t05_tutorial_graph_setup_view_save => "T05_tutorial_graph_setup_view_save.tsv",
    t06_tutorial_multiple_sheets_group_and_3d => "T06_tutorial_multiple_sheets_group_and_3d.tsv",
    t07_tutorial_file_retrieve_and_open => "T07_tutorial_file_retrieve_and_open.tsv",
}

#[cfg(feature = "wk3")]
#[test]
fn wk3_retrieve_saves_as_xlsx() {
    let p = workspace_root().join("tests/acceptance/wk3_retrieve_saves_as_xlsx.tsv");
    run_transcript(&p);
}
