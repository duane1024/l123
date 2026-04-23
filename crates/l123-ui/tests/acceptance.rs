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

    let body = fs::read_to_string(path)
        .unwrap_or_else(|e| panic!("read {}: {e}", path.display()));

    let mut app = App::new();
    let mut width: u16 = 80;
    let mut height: u16 = 25;

    for (ln, raw) in body.lines().enumerate() {
        let line_no = ln + 1;
        let line = raw.split('#').next().unwrap().trim();
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
            "BACKSPACE" => {
                app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE))
            }
            "UP" => app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE)),
            "DOWN" => app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE)),
            "LEFT" => app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE)),
            "RIGHT" => app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE)),
            "HOME" => app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE)),
            "END" => app.handle_key(KeyEvent::new(KeyCode::End, KeyModifiers::NONE)),
            "PGUP" => {
                app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::NONE))
            }
            "PGDN" => {
                app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::NONE))
            }
            "CTRL_PGUP" => {
                app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::CONTROL))
            }
            "CTRL_PGDN" => {
                app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::CONTROL))
            }
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
                    got, rest,
                    "{}:{line_no}: pointer expected {rest} got {got}",
                    path.display()
                );
            }
            "ASSERT_MODE" => {
                let got = app.mode().indicator();
                assert_eq!(
                    got, rest,
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
            "ASSERT_CELL" => {
                // "A:A1  hello" — address then expected trimmed-cell text.
                let mut parts = rest.splitn(2, char::is_whitespace);
                let addr = parts.next().unwrap_or("");
                let want = parts.next().unwrap_or("").trim();
                let buf = app.render_to_buffer(width, height);
                let got = app
                    .cell_rendered_text(&buf, addr)
                    .unwrap_or_else(|| panic!(
                        "{}:{line_no}: address {addr:?} not in viewport",
                        path.display()
                    ));
                assert_eq!(
                    got.trim(), want,
                    "{}:{line_no}: cell {addr} expected {want:?} got {got:?}",
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
                    got, want,
                    "{}:{line_no}: running expected {want}, got {got}",
                    path.display()
                );
            }
            "SIZE" => {
                let mut parts = rest.split_ascii_whitespace();
                width = parts.next().unwrap().parse().unwrap();
                height = parts.next().unwrap().parse().unwrap();
            }
            // Pre-clean a file on disk so the transcript starts from a
            // known state. Errors (e.g. not-present) are ignored.
            "RM_FILE" => {
                let _ = std::fs::remove_file(rest);
            }
            other => {
                panic!(
                    "{}:{line_no}: unknown directive {other:?}",
                    path.display()
                );
            }
        }
    }
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
}
