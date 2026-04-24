//! Acceptance transcript harness.
//!
//! Reads `tests/acceptance/*.tsv` from the workspace root and drives an
//! `App` through the directives. Format documented in
//! `tests/acceptance/README.md`.

use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use l123_ui::App;

/// Transcripts share process CWD (set per transcript) and write to
/// `target/`. `/FD` mutates CWD; parallel tests racing on CWD or on
/// file-existence checks (`/FS` Cancel/Replace branching) are the
/// cause of historical flakes. Serializing the transcripts is cheap
/// (tests take <200ms total) and removes the race outright.
static ACCEPTANCE_LOCK: Mutex<()> = Mutex::new(());

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
    m2_formula_entry   => "M2_formula_entry.tsv",
    m2_f9_calc         => "M2_f9_calc.tsv",
    m2_format_tag      => "M2_format_tag.tsv",
    m3_menu_navigation => "M3_menu_navigation.tsv",
    m3_quit            => "M3_quit.tsv",
    m3_insert_delete_row_col => "M3_insert_delete_row_col.tsv",
    m3_range_erase     => "M3_range_erase.tsv",
    m3_copy            => "M3_copy.tsv",
    m3_move            => "M3_move.tsv",
    m3_range_label     => "M3_range_label.tsv",
    m3_range_format    => "M3_range_format.tsv",
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
    m4_file_save       => "M4_file_save.tsv",
    m4_file_save_replace => "M4_file_save_replace.tsv",
    m4_file_retrieve   => "M4_file_retrieve.tsv",
    m4_file_xtract     => "M4_file_xtract.tsv",
    m4_file_import_numbers => "M4_file_import_numbers.tsv",
    m4_file_new        => "M4_file_new.tsv",
    m4_file_dir        => "M4_file_dir.tsv",
    m4_file_list_active => "M4_file_list_active.tsv",
    m4_file_list_worksheet => "M4_file_list_worksheet.tsv",
    m5_insert_sheet    => "M5_insert_sheet.tsv",
    m5_group_format    => "M5_group_format.tsv",
    m5_3d_sum          => "M5_3d_sum.tsv",
    m5_file_open       => "M5_file_open.tsv",
    m5_undo            => "M5_undo.tsv",
    m5_undo_toggle     => "M5_undo_toggle.tsv",
    m5_undo_coverage   => "M5_undo_coverage.tsv",
    m6_print_file      => "M6_print_file.tsv",
    m6_print_options_header => "M6_print_options_header.tsv",
    m6_print_pipe_row  => "M6_print_pipe_row.tsv",
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
    m10_status_line_filename => "M10_status_line_filename.tsv",
    m10_worksheet_status     => "M10_worksheet_status.tsv",
}
