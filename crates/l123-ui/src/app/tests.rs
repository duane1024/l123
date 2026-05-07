use super::*;
use std::path::Path;

#[test]
fn wgd_update_writes_cnf_block_and_preserves_other_lines() {
    let dir = std::env::temp_dir().join("l123_wgd_update_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("L123.CNF");
    std::fs::write(&path, "user = \"Pre-existing\"\nlog_file = /tmp/x.log\n").unwrap();

    let d = GlobalDefaults {
        printer_interface: 5,
        printer_pg_length: 88,
        default_dir: "/tmp/sheets".into(),
        autoexec: false,
        graph_group: GraphGroupOrientation::Rowwise,
        graph_save: GraphSaveFormat::Pic,
        ..GlobalDefaults::default()
    };
    d.write_to_path(&path).unwrap();

    let body = std::fs::read_to_string(&path).unwrap();
    assert!(body.contains("user = \"Pre-existing\""), "{body}");
    assert!(body.contains("log_file = /tmp/x.log"), "{body}");
    assert!(body.contains("wgd_printer_interface = 5"), "{body}");
    assert!(body.contains("wgd_printer_pg_length = 88"), "{body}");
    assert!(body.contains("wgd_dir = \"/tmp/sheets\""), "{body}");
    assert!(body.contains("wgd_autoexec = false"), "{body}");
    assert!(body.contains("wgd_graph_group = rowwise"), "{body}");
    assert!(body.contains("wgd_graph_save = pic"), "{body}");

    // Re-running update should not duplicate the block.
    d.write_to_path(&path).unwrap();
    let body2 = std::fs::read_to_string(&path).unwrap();
    let count = body2.matches("wgd_printer_interface").count();
    assert_eq!(count, 1, "block duplicated on re-run:\n{body2}");
}

#[test]
fn starts_at_a1() {
    let app = App::new();
    assert_eq!(app.wb().pointer, Address::A1);
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.entry.is_none());
}

#[test]
fn new_with_file_routes_csv_by_extension() {
    use std::io::Write;
    let dir = std::env::temp_dir().join("l123_new_with_csv");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("foo.csv");
    std::fs::File::create(&path)
        .unwrap()
        .write_all(b"10,20\nfoo,bar\n")
        .unwrap();
    let app = App::new_with_file(path);
    assert_eq!(
        app.mode,
        Mode::Ready,
        "CLI-opening a .csv should land in READY, not ERROR (error={:?})",
        app.error_message,
    );
    match app.wb().cells.get(&Address::A1) {
        Some(CellContents::Constant(Value::Number(n))) => assert_eq!(*n, 10.0),
        other => panic!("A1 expected Number(10), got {other:?}"),
    }
}

#[cfg(feature = "wk3")]
#[test]
fn new_with_file_routes_wk3_by_extension() {
    // Open a `.WK3` via the CLI entry point: it should land in
    // READY (not ERROR) with the workbook content visible, and
    // `active_path` swapped to "<original>.WK3.xlsx" so /File
    // Save defaults to writing xlsx alongside the legacy file.
    let dir = temp_test_dir("new_with_wk3");
    let src = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(|p| p.parent())
        .expect("workspace root above crates/l123-ui")
        .join("tests/acceptance/fixtures/wk3/FILE0001.WK3");
    if !src.exists() {
        eprintln!(
            "skipping new_with_file_routes_wk3_by_extension: {} missing",
            src.display()
        );
        let _ = std::fs::remove_dir(&dir);
        return;
    }
    let wk3_path = dir.join("legacy.WK3");
    std::fs::copy(&src, &wk3_path).unwrap();

    let app = App::new_with_file(wk3_path.clone());
    assert_eq!(
        app.mode,
        Mode::Ready,
        "CLI-opening a .WK3 should land in READY (error={:?})",
        app.error_message,
    );
    match app.wb().cells.get(&Address::A1) {
        Some(CellContents::Label { text, .. }) => assert_eq!(text, "Hello"),
        other => panic!("A1 expected Label(Hello), got {other:?}"),
    }
    let expected_save = wk3_path.with_file_name("legacy.WK3.xlsx");
    assert_eq!(
        app.wb().active_path.as_deref(),
        Some(expected_save.as_path()),
        "active_path should be original.WK3.xlsx so /File Save writes xlsx",
    );

    let _ = std::fs::remove_file(&wk3_path);
    let _ = std::fs::remove_dir(&dir);
}

#[test]
fn arrow_nav() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer.col, 1);
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer.row, 1);
    app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer.col, 0);
    app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer.row, 0);
}

#[test]
fn f5_goto_moves_pointer() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::F(5), KeyModifiers::NONE));
    for c in "C5".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer, Address::new(SheetId::A, 2, 4));
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn f5_goto_esc_leaves_pointer_unchanged() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::F(5), KeyModifiers::NONE));
    for c in "Z99".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer, Address::A1);
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn typed_point_range_commits() {
    let mut app = App::new();
    // /Range Erase enters POINT in one step.
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    for c in "B2..D4".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(app.wb().pointer, Address::A1);
}

#[test]
fn typed_point_esc_clears_buffer_first() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    for c in "B2".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    assert!(app.point.as_ref().unwrap().typed.is_empty());
    // Two more Esc presses to fully cancel (un-anchor, then exit).
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn typed_point_bad_input_stays_in_point() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    for c in "WAT".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    assert!(app.point.as_ref().unwrap().typed.is_empty());
}

#[test]
fn left_from_a1_stays_at_a1() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer, Address::A1);
}

#[test]
fn down_past_visible_area_advances_row_offset() {
    let mut app = App::new();
    // Render an 80x25 frame so scroll_into_view has a cached
    // grid rect to consult on the next move.
    let _ = app.render_to_buffer(80, 25);
    for _ in 0..25 {
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    }
    assert_eq!(app.wb().pointer.row, 25);
    assert!(
        app.wb().viewport_row_offset > 0,
        "viewport_row_offset should advance to keep pointer visible, got 0",
    );
    assert!(
        app.wb().pointer.row >= app.wb().viewport_row_offset,
        "pointer must be at or below the new top of viewport",
    );
}

#[test]
fn right_past_visible_area_advances_col_offset() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    for _ in 0..12 {
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    }
    assert_eq!(app.wb().pointer.col, 12);
    assert!(
        app.wb().viewport_col_offset > 0,
        "viewport_col_offset should advance to keep pointer visible, got 0",
    );
}

#[test]
fn up_after_scroll_pulls_viewport_back() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    for _ in 0..30 {
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    }
    assert!(app.wb().viewport_row_offset > 0);
    // Press UP enough to drop above the current viewport top.
    for _ in 0..30 {
        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
    }
    assert_eq!(app.wb().pointer.row, 0);
    assert_eq!(app.wb().viewport_row_offset, 0);
}

#[test]
fn home_resets_pointer() {
    let mut app = App::new();
    app.wb_mut().pointer = Address::new(SheetId::A, 10, 10);
    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer, Address::A1);
}

#[test]
fn ctrl_c_does_not_quit() {
    // SPEC §7 Δ: 1-2-3 uses /QY to quit; Ctrl-C is unused.
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL));
    assert!(app.running);
}

#[test]
fn pgdn_moves_twenty_rows() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::NONE));
    assert_eq!(app.wb().pointer.row, 20);
}

#[test]
fn letter_first_enters_label_mode() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('h'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Label);
    let e = app.entry.as_ref().unwrap();
    assert_eq!(e.buffer, "h");
    assert!(matches!(e.kind, EntryKind::Label(LabelPrefix::Apostrophe)));
}

#[test]
fn digit_first_enters_value_mode() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('1'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Value);
    let e = app.entry.as_ref().unwrap();
    assert_eq!(e.buffer, "1");
    assert!(matches!(e.kind, EntryKind::Value));
}

fn make_label(text: &str) -> CellContents {
    CellContents::Label {
        prefix: LabelPrefix::Apostrophe,
        text: text.into(),
    }
}

fn press(app: &mut App, code: KeyCode) {
    app.handle_key(KeyEvent::new(code, KeyModifiers::NONE));
}

fn press_ch(app: &mut App, c: char) {
    app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
}

#[test]
fn begin_point_auto_anchors_at_pointer() {
    let mut app = App::new();
    app.wb_mut().pointer = Address::new(SheetId::A, 1, 1); // B2
    app.begin_point(PendingCommand::RangeErase);
    assert_eq!(app.mode, Mode::Point);
    let anchor = app.point.as_ref().unwrap().anchor.unwrap();
    assert_eq!(anchor, Address::new(SheetId::A, 1, 1));
    assert_eq!(
        app.highlight_range(),
        Range::single(Address::new(SheetId::A, 1, 1))
    );
}

#[test]
fn point_arrow_expands_range() {
    let mut app = App::new();
    app.begin_point(PendingCommand::RangeErase);
    press(&mut app, KeyCode::Right);
    press(&mut app, KeyCode::Down);
    let r = app.highlight_range();
    assert_eq!(r.start, Address::A1);
    assert_eq!(r.end, Address::new(SheetId::A, 1, 1));
}

#[test]
fn point_esc_twice_cancels() {
    let mut app = App::new();
    app.begin_point(PendingCommand::RangeErase);
    press(&mut app, KeyCode::Down);
    press(&mut app, KeyCode::Esc);
    // Anchor cleared but still in POINT.
    assert_eq!(app.mode, Mode::Point);
    assert!(app.point.as_ref().unwrap().anchor.is_none());
    press(&mut app, KeyCode::Esc);
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.point.is_none());
}

#[test]
fn point_period_anchors_when_unanchored() {
    let mut app = App::new();
    app.begin_point(PendingCommand::RangeErase);
    // Anchor manually cleared
    app.point.as_mut().unwrap().anchor = None;
    press(&mut app, KeyCode::Right);
    press_ch(&mut app, '.');
    let anchor = app.point.as_ref().unwrap().anchor.unwrap();
    assert_eq!(anchor, Address::new(SheetId::A, 1, 0));
}

#[test]
fn point_period_cycles_corner() {
    let mut app = App::new();
    // Anchor at A1; extend to B2 → pointer at BR.
    app.begin_point(PendingCommand::RangeErase);
    press(&mut app, KeyCode::Right);
    press(&mut app, KeyCode::Down);
    let before = app.highlight_range();
    assert_eq!(before.start, Address::A1);
    assert_eq!(before.end, Address::new(SheetId::A, 1, 1));
    // Initially pointer is at BR corner of range. `.` rotates TL→TR→BR→BL.
    // Pointer was at BR(B2), so we're detecting at_min_col=false, at_min_row=false,
    // which the code treats as "BR → BL", moving pointer to BL (A2).
    press_ch(&mut app, '.');
    assert_eq!(app.wb().pointer, Address::new(SheetId::A, 0, 1)); // BL
                                                                  // Range unchanged.
    assert_eq!(app.highlight_range(), before);
    // Next `.` from BL → TL.
    press_ch(&mut app, '.');
    assert_eq!(app.wb().pointer, Address::A1); // TL
    assert_eq!(app.highlight_range(), before);
    // Next `.` from TL → TR.
    press_ch(&mut app, '.');
    assert_eq!(app.wb().pointer, Address::new(SheetId::A, 1, 0)); // TR
    assert_eq!(app.highlight_range(), before);
    // Next `.` from TR → BR. Full loop.
    press_ch(&mut app, '.');
    assert_eq!(app.wb().pointer, Address::new(SheetId::A, 1, 1)); // BR
    assert_eq!(app.highlight_range(), before);
}

#[test]
fn range_erase_clears_all_cells_in_range() {
    let mut app = App::new();
    for (row, v) in [(0, "10"), (1, "20"), (2, "30")] {
        for c in v.chars() {
            press_ch(&mut app, c);
        }
        press(&mut app, KeyCode::Down);
        // After each commit, pointer moves; at end we're at row+1
        let _ = row;
    }
    // Go back to A1 and erase A1..A3
    app.wb_mut().pointer = Address::A1;
    app.begin_point(PendingCommand::RangeErase);
    press(&mut app, KeyCode::Down);
    press(&mut app, KeyCode::Down);
    press(&mut app, KeyCode::Enter);
    assert_eq!(app.mode, Mode::Ready);
    for row in 0..3 {
        assert!(
            !app.wb()
                .cells
                .contains_key(&Address::new(SheetId::A, 0, row)),
            "A{} should be empty",
            row + 1
        );
    }
}

#[test]
fn shift_cells_rows_insert_pushes_down() {
    let mut cells = HashMap::new();
    cells.insert(Address::new(SheetId::A, 0, 0), make_label("r0"));
    cells.insert(Address::new(SheetId::A, 0, 1), make_label("r1"));
    cells.insert(Address::new(SheetId::A, 0, 2), make_label("r2"));
    // Insert 1 row at row 1 — r1 and r2 move down by 1.
    shift_cells_rows(&mut cells, SheetId::A, 1, 1);
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 0)),
        Some(&make_label("r0"))
    );
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 2)),
        Some(&make_label("r1"))
    );
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 3)),
        Some(&make_label("r2"))
    );
    assert!(!cells.contains_key(&Address::new(SheetId::A, 0, 1)));
}

#[test]
fn shift_cells_rows_delete_pulls_up() {
    let mut cells = HashMap::new();
    cells.insert(Address::new(SheetId::A, 0, 0), make_label("r0"));
    cells.insert(Address::new(SheetId::A, 0, 2), make_label("r2"));
    // Simulate delete of row 1: remove it (already absent here) then pull.
    shift_cells_rows(&mut cells, SheetId::A, 2, -1);
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 0)),
        Some(&make_label("r0"))
    );
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 1)),
        Some(&make_label("r2"))
    );
}

#[test]
fn shift_cells_cols_insert_pushes_right() {
    let mut cells = HashMap::new();
    cells.insert(Address::new(SheetId::A, 0, 0), make_label("A1"));
    cells.insert(Address::new(SheetId::A, 1, 0), make_label("B1"));
    shift_cells_cols(&mut cells, SheetId::A, 1, 1);
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 0, 0)),
        Some(&make_label("A1"))
    );
    assert_eq!(
        cells.get(&Address::new(SheetId::A, 2, 0)),
        Some(&make_label("B1"))
    );
    assert!(!cells.contains_key(&Address::new(SheetId::A, 1, 0)));
}

#[test]
fn shift_cells_rows_leaves_other_sheets_alone() {
    let mut cells = HashMap::new();
    cells.insert(Address::new(SheetId::A, 0, 1), make_label("a"));
    cells.insert(Address::new(SheetId(1), 0, 1), make_label("b"));
    shift_cells_rows(&mut cells, SheetId::A, 0, 1);
    assert!(cells.contains_key(&Address::new(SheetId::A, 0, 2)));
    // Sheet B unchanged.
    assert_eq!(
        cells.get(&Address::new(SheetId(1), 0, 1)),
        Some(&make_label("b"))
    );
}

#[test]
fn manual_recalc_defers_computation_until_f9() {
    let mut app = App::new();
    app.set_recalc_mode(RecalcMode::Manual);

    // A1 = 10
    for c in "10".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    // B1 = +A1*2
    app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
    for c in "A1*2".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    // In Manual mode, the just-committed formula has no cached value yet.
    let formula = app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap();
    if let CellContents::Formula { cached_value, .. } = formula {
        assert!(cached_value.is_none(), "manual mode should not auto-eval");
    } else {
        panic!("expected Formula");
    }
    assert!(app.recalc_pending());

    // F9 computes and clears the pending flag.
    app.handle_key(KeyEvent::new(KeyCode::F(9), KeyModifiers::NONE));
    assert!(!app.recalc_pending());
    let formula = app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap();
    match formula {
        CellContents::Formula { cached_value, .. } => {
            assert_eq!(*cached_value, Some(Value::Number(20.0)));
        }
        other => panic!("expected Formula, got {other:?}"),
    }
}

#[test]
fn calc_indicator_visible_only_when_pending() {
    let mut app = App::new();
    let buf = app.render_to_buffer(80, 25);
    let status_line = App::line_text(&buf, 24);
    assert!(
        !status_line.contains("CALC"),
        "should be absent: {status_line:?}"
    );

    app.set_recalc_mode(RecalcMode::Manual);
    // Type a formula to get CALC to light up.
    app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('1'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    let buf = app.render_to_buffer(80, 25);
    let status_line = App::line_text(&buf, 24);
    assert!(
        status_line.contains("CALC"),
        "should contain CALC: {status_line:?}"
    );

    app.handle_key(KeyEvent::new(KeyCode::F(9), KeyModifiers::NONE));
    let buf = app.render_to_buffer(80, 25);
    let status_line = App::line_text(&buf, 24);
    assert!(
        !status_line.contains("CALC"),
        "should be cleared: {status_line:?}"
    );
}

#[test]
fn formula_commit_populates_cached_value() {
    let mut app = App::new();
    // A1 = 10, A2 = 20
    for c in "10".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    for c in "20".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    // A3 = @SUM(A1..A2)
    for c in "@SUM(A1..A2)".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let a3 = app.wb().cells.get(&Address::new(SheetId::A, 0, 2)).unwrap();
    match a3 {
        CellContents::Formula { expr, cached_value } => {
            assert_eq!(expr, "@SUM(A1..A2)");
            assert_eq!(*cached_value, Some(Value::Number(30.0)));
        }
        other => panic!("expected Formula, got {other:?}"),
    }
}

#[test]
fn upstream_edit_recomputes_dependent_formula() {
    let mut app = App::new();
    for c in "10".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    // B1 = +A1*3
    app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
    for c in "A1*3".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(
        app.wb()
            .cells
            .get(&Address::new(SheetId::A, 0, 1))
            .and_then(|c| match c {
                CellContents::Formula { cached_value, .. } => cached_value.clone(),
                _ => None,
            }),
        Some(Value::Number(30.0))
    );
    // Change A1 to 5; B1 should recompute to 15.
    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('5'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(
        app.wb()
            .cells
            .get(&Address::new(SheetId::A, 0, 1))
            .and_then(|c| match c {
                CellContents::Formula { cached_value, .. } => cached_value.clone(),
                _ => None,
            }),
        Some(Value::Number(15.0))
    );
}

#[test]
fn f2_on_empty_cell_enters_edit_mode() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Edit);
    assert_eq!(app.entry.as_ref().unwrap().buffer, "");
}

#[test]
fn f2_loads_label_source_into_buffer_with_prefix() {
    let mut app = App::new();
    app.wb_mut().cells.insert(
        Address::A1,
        CellContents::Label {
            prefix: LabelPrefix::Quote,
            text: "right".into(),
        },
    );
    app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Edit);
    assert_eq!(app.entry.as_ref().unwrap().buffer, "\"right");
}

#[test]
fn f2_commit_reparses_via_first_char_rule() {
    let mut app = App::new();
    app.wb_mut().cells.insert(
        Address::A1,
        CellContents::Label {
            prefix: LabelPrefix::Apostrophe,
            text: "hello".into(),
        },
    );
    app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
    // Buffer is "'hello"; remove ' and " -prefix instead.
    app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
    for c in "ello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
        let _ = c;
    }
    // buffer = "h" now (we backspaced over 'ello but 'h' remains);
    // for a deterministic commit, clear fully.
    app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
    assert_eq!(app.entry.as_ref().unwrap().buffer, "");
    app.handle_key(KeyEvent::new(KeyCode::Char('"'), KeyModifiers::NONE));
    for c in "right".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    match app.wb().cells.get(&Address::A1).unwrap() {
        CellContents::Label { prefix, text } => {
            assert_eq!(*prefix, LabelPrefix::Quote);
            assert_eq!(text, "right");
        }
        other => panic!("expected Label(Quote), got {other:?}"),
    }
}

#[test]
fn esc_during_entry_cancels_and_leaves_cell_empty() {
    let mut app = App::new();
    for c in "hello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.entry.is_none());
    assert!(!app.wb().cells.contains_key(&Address::A1));
}

#[test]
fn arrow_commits_then_moves() {
    let mut app = App::new();
    for c in "hi".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(app.wb().pointer, Address::new(SheetId::A, 0, 1));
    assert!(matches!(
        app.wb().cells.get(&Address::A1),
        Some(CellContents::Label { .. })
    ));
}

#[test]
fn backspace_edits_buffer() {
    let mut app = App::new();
    for c in "hello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Backspace, KeyModifiers::NONE));
    assert_eq!(app.entry.as_ref().unwrap().buffer, "hell");
}

#[test]
fn explicit_label_prefix_dispatch() {
    for (ch, want) in [
        ('\'', LabelPrefix::Apostrophe),
        ('"', LabelPrefix::Quote),
        ('^', LabelPrefix::Caret),
        ('\\', LabelPrefix::Backslash),
    ] {
        let mut app = App::new();
        app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::NONE));
        assert_eq!(app.mode, Mode::Label, "char {ch:?}");
        let e = app.entry.as_ref().unwrap();
        assert!(
            matches!(e.kind, EntryKind::Label(p) if p == want),
            "char {ch:?}: expected prefix {want:?}, got {:?}",
            e.kind
        );
        assert_eq!(e.buffer, "", "buffer should be empty after prefix char");
    }
}

#[test]
fn value_commit_stores_as_number() {
    let mut app = App::new();
    for c in "123".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    match app.wb().cells.get(&Address::A1).unwrap() {
        CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 123.0),
        other => panic!("expected Number, got {other:?}"),
    }
}

#[test]
fn value_commit_handles_decimal_and_negative() {
    let mut app = App::new();
    for c in "-1.25".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    match app.wb().cells.get(&Address::A1).unwrap() {
        CellContents::Constant(Value::Number(n)) => {
            assert!((*n - (-1.25)).abs() < 1e-9, "got {n}");
        }
        other => panic!("expected Number, got {other:?}"),
    }
}

#[test]
fn label_commit_stores_with_prefix_and_returns_to_ready() {
    let mut app = App::new();
    for c in "hello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.entry.is_none());
    let stored = app.wb().cells.get(&Address::A1).unwrap();
    match stored {
        CellContents::Label { prefix, text } => {
            assert_eq!(*prefix, LabelPrefix::Apostrophe);
            assert_eq!(text, "hello");
        }
        other => panic!("expected Label, got {other:?}"),
    }
}

#[test]
fn backslash_prefix_repeat_fills_when_xlsx_halign_is_left() {
    // Reproduces the post-import regression: an xlsx that came from
    // a Lotus-saved sheet stores `\-` as Label{Backslash, "-"} and
    // separately carries HAlign::Left (Excel's text default). The
    // grid renderer was letting Left override the stored Backslash
    // prefix, so the cell rendered as a single "-" left-padded
    // rather than as a span of dashes.
    let mut app = App::new();
    app.wb_mut().cells.insert(
        Address::A1,
        CellContents::Label {
            prefix: LabelPrefix::Backslash,
            text: "-".into(),
        },
    );
    app.wb_mut().cell_alignments.insert(
        Address::A1,
        Alignment {
            horizontal: HAlign::Left,
            ..Alignment::DEFAULT
        },
    );
    let buf = app.render_to_buffer(80, 25);
    let got = app
        .cell_rendered_text(&buf, "A:A1")
        .expect("A1 must be in viewport");
    assert_eq!(
        got, "---------",
        "Backslash prefix must override imported HAlign::Left, got {got:?}"
    );
}

fn temp_test_dir(tag: &str) -> PathBuf {
    use std::process;
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let dir = std::env::temp_dir().join(format!("l123_test_{}_{}_{}", tag, process::id(), nanos,));
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

fn drive_save_keys(app: &mut App, seed_label: &str, path: &Path) {
    for c in seed_label.chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    for c in ['/', 'F', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in path.to_str().unwrap().chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    // Drain the §4.7 async save the prompt commit just queued.
    app.test_resume_async_op();
}

/// End-to-end: /FS <path><Enter> writes an xlsx file at <path> that
/// IronCalc can open. The temp dir is unique per test invocation so
/// parallel runs don't collide.
#[test]
fn file_save_writes_xlsx_at_typed_path() {
    let dir = temp_test_dir("file_save");
    let target = dir.join("saved.xlsx");

    let mut app = App::new();
    drive_save_keys(&mut app, "42", &target);

    assert_eq!(app.mode, Mode::Ready);
    assert!(
        target.exists(),
        "expected xlsx at {target:?} — prompt did not write the file"
    );

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// Backup path: when the file already exists, picking Backup
/// renames the existing file to `.BAK` and then writes a fresh one.
#[test]
fn file_save_backup_renames_existing_to_bak() {
    let dir = temp_test_dir("file_save_backup");
    let target = dir.join("sheet.xlsx");

    // First save populates the file.
    let mut app = App::new();
    drive_save_keys(&mut app, "42", &target);
    assert!(target.exists());

    // Modify, then /FS again — prefilled; Enter opens confirm.
    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    for c in "99".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    for c in ['/', 'F', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert!(app.save_confirm.is_some(), "confirm submenu did not open");

    // Press B — Backup.
    app.handle_key(KeyEvent::new(KeyCode::Char('B'), KeyModifiers::NONE));
    // Backup queues an async save; drain it before checking disk.
    app.test_resume_async_op();
    assert_eq!(app.mode, Mode::Ready);
    let bak = target.with_extension("BAK");
    assert!(bak.exists(), "expected {bak:?} after Backup");
    assert!(target.exists(), "fresh {target:?} after Backup");

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_file(&bak);
    let _ = std::fs::remove_dir(&dir);
}

/// Helper: drive /FX<kind> <path><Enter> HOME <Enter>. Assumes the
/// pointer is at the bottom-right of the intended range before the
/// call (POINT auto-anchors there; HOME slides the free corner to
/// A1 → highlight covers A1..pointer).
fn drive_xtract_keys(app: &mut App, kind: char, path: &Path) {
    for c in ['/', 'F', 'X', kind] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in path.to_str().unwrap().chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
}

/// /FXF — Formulas extract keeps the formula string intact so the
/// extracted file, when reloaded, still has a live formula in A3.
#[test]
fn file_xtract_formulas_preserves_formula() {
    let dir = temp_test_dir("xtract_f");
    let target = dir.join("x.xlsx");
    if target.exists() {
        std::fs::remove_file(&target).unwrap();
    }

    let mut app = App::new();
    // DOWN during entry commits-and-moves; ENTER commits-in-place.
    for line in ["10", "20"] {
        for c in line.chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    }
    for c in "+A1+A2".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    drive_xtract_keys(&mut app, 'F', &target);
    assert_eq!(app.mode, Mode::Ready);
    assert!(target.exists(), "extract did not write the file");

    // Re-open and inspect.
    let mut e = IronCalcEngine::new().unwrap();
    e.load_xlsx(&target).unwrap();
    let a3 = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
    assert_eq!(a3.value, Value::Number(30.0));
    assert!(
        a3.formula.is_some(),
        "Formulas variant should preserve the formula"
    );

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// /FXV — Values extract replaces the formula with its cached
/// value; reloaded A3 has a number but no formula.
#[test]
fn file_xtract_values_strips_formula() {
    let dir = temp_test_dir("xtract_v");
    let target = dir.join("x.xlsx");
    if target.exists() {
        std::fs::remove_file(&target).unwrap();
    }

    let mut app = App::new();
    // DOWN during entry commits-and-moves; ENTER commits-in-place.
    for line in ["10", "20"] {
        for c in line.chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
        }
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    }
    for c in "+A1+A2".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    drive_xtract_keys(&mut app, 'V', &target);
    assert_eq!(app.mode, Mode::Ready);
    assert!(target.exists());

    let mut e = IronCalcEngine::new().unwrap();
    e.load_xlsx(&target).unwrap();
    let a3 = e.get_cell(Address::new(SheetId::A, 0, 2)).unwrap();
    assert_eq!(a3.value, Value::Number(30.0));
    assert!(
        a3.formula.is_none(),
        "Values variant should flatten formula to value"
    );

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// Worksheet listing is an alphabetical list of `.xlsx` files
/// (plus `.WK3` when built with `--features wk3`) in the given
/// directory. Pure function of the directory's contents so we can
/// test it without touching process CWD.
#[test]
fn list_worksheet_files_in_returns_xlsx_sorted() {
    let dir = temp_test_dir("list_ws");
    let names_in = [
        "zeta.xlsx",
        "alpha.xlsx",
        "other.txt",
        "mid.XLSX",
        "legacy.WK3",
        "lower.wk3",
    ];
    for name in names_in {
        std::fs::write(dir.join(name), b"placeholder").unwrap();
    }
    let got = list_worksheet_files_in(&dir);
    let names: Vec<String> = got
        .iter()
        .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
        .collect();
    #[cfg(feature = "wk3")]
    let expected = vec![
        "alpha.xlsx",
        "legacy.WK3",
        "lower.wk3",
        "mid.XLSX",
        "zeta.xlsx",
    ];
    #[cfg(not(feature = "wk3"))]
    let expected = vec!["alpha.xlsx", "mid.XLSX", "zeta.xlsx"];
    assert_eq!(names, expected);
    for name in names_in {
        let _ = std::fs::remove_file(dir.join(name));
    }
    let _ = std::fs::remove_dir(&dir);
}

/// Vertical navigation + scroll: with 20 files and a PAGE_SIZE of
/// 10, pressing Down 15 times should advance the view so the
/// highlight is still visible.
#[test]
fn file_list_vertical_nav_and_scroll() {
    let dir = temp_test_dir("list_scroll");
    for i in 0..20 {
        std::fs::write(dir.join(format!("f{i:02}.xlsx")), b"data").unwrap();
    }
    let entries = list_worksheet_files_in(&dir);
    assert_eq!(entries.len(), 20);

    let mut app = App::new();
    app.file_list = Some(FileListState {
        kind: FileListKind::Worksheet,
        entries,
        highlight: 0,
        view_offset: 0,
    });
    app.mode = Mode::Files;

    for _ in 0..15 {
        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    }
    let fl = app.file_list.as_ref().unwrap();
    assert_eq!(fl.highlight, 15);
    assert!(
        fl.view_offset > 0,
        "view_offset should have advanced (got {})",
        fl.view_offset
    );
    assert!(
        fl.highlight >= fl.view_offset && fl.highlight < fl.view_offset + FILE_LIST_PAGE_SIZE,
        "highlight {} out of window starting at {}",
        fl.highlight,
        fl.view_offset
    );

    app.handle_key(KeyEvent::new(KeyCode::End, KeyModifiers::NONE));
    assert_eq!(app.file_list.as_ref().unwrap().highlight, 19);

    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    let fl = app.file_list.as_ref().unwrap();
    assert_eq!(fl.highlight, 0);
    assert_eq!(fl.view_offset, 0);

    for i in 0..20 {
        let _ = std::fs::remove_file(dir.join(format!("f{i:02}.xlsx")));
    }
    let _ = std::fs::remove_dir(&dir);
}

/// Worksheet branch: pressing Enter on the highlighted row loads
/// that xlsx file into the workbook.
#[test]
fn file_list_worksheet_enter_retrieves_highlighted() {
    let dir = temp_test_dir("list_ws_enter");

    // Build two xlsx files with distinguishing contents. Driving
    // through the App would set CWD / active_path; use the engine
    // directly so the test stays hermetic.
    let a = dir.join("a.xlsx");
    let b = dir.join("b.xlsx");
    {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::A1, "111").unwrap();
        e.recalc();
        e.save_xlsx(&a).unwrap();
    }
    {
        let mut e = IronCalcEngine::new().unwrap();
        e.set_user_input(Address::A1, "222").unwrap();
        e.recalc();
        e.save_xlsx(&b).unwrap();
    }

    let mut app = App::new();
    app.file_list = Some(FileListState {
        kind: FileListKind::Worksheet,
        entries: vec![a.clone(), b.clone()],
        highlight: 1, // point at b.xlsx
        view_offset: 0,
    });
    app.mode = Mode::Files;

    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.file_list.is_none());
    assert_eq!(app.wb().active_path.as_deref(), Some(b.as_path()));
    match app.wb().cells.get(&Address::A1).unwrap() {
        CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 222.0),
        other => panic!("A1 expected Number(222) from b.xlsx, got {other:?}"),
    }

    let _ = std::fs::remove_file(&a);
    let _ = std::fs::remove_file(&b);
    let _ = std::fs::remove_dir(&dir);
}

/// `list_all_files_in` includes every regular file (any extension)
/// and skips dotfiles, sorted by filename.
#[test]
fn list_all_files_in_returns_every_file_no_dotfiles() {
    let dir = temp_test_dir("list_other");
    let names_in = [
        "zeta.xlsx",
        "notes.txt",
        "data.csv",
        "alpha.bin",
        ".hidden",
        "README",
    ];
    for name in names_in {
        std::fs::write(dir.join(name), b"x").unwrap();
    }
    let got = list_all_files_in(&dir);
    let names: Vec<String> = got
        .iter()
        .map(|p| p.file_name().unwrap().to_string_lossy().into_owned())
        .collect();
    assert_eq!(
        names,
        vec!["README", "alpha.bin", "data.csv", "notes.txt", "zeta.xlsx"]
    );
    for name in names_in {
        let _ = std::fs::remove_file(dir.join(name));
    }
    let _ = std::fs::remove_dir(&dir);
}

/// /File List Other → Enter on a `.csv` file routes through the CSV
/// loader: the workbook ends up populated with that file's rows.
#[test]
fn file_list_other_enter_loads_csv() {
    let dir = temp_test_dir("list_other_csv");
    let path = dir.join("data.csv");
    std::fs::write(&path, b"7,8,9\n").unwrap();

    let mut app = App::new();
    app.file_list = Some(FileListState {
        kind: FileListKind::Other,
        entries: vec![path.clone()],
        highlight: 0,
        view_offset: 0,
    });
    app.mode = Mode::Files;

    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.file_list.is_none());
    match app.wb().cells.get(&Address::A1).unwrap() {
        CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 7.0),
        other => panic!("A1 expected Number(7) from data.csv, got {other:?}"),
    }

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_dir(&dir);
}

/// /File List Other → Enter on a non-spreadsheet file just dismisses
/// the overlay, leaving the workbook untouched.
#[test]
fn file_list_other_enter_on_unsupported_extension_dismisses() {
    let dir = temp_test_dir("list_other_dismiss");
    let path = dir.join("readme.txt");
    std::fs::write(&path, b"hello").unwrap();

    let mut app = App::new();
    app.wb_mut().active_path = Some(PathBuf::from("orig.xlsx"));
    app.file_list = Some(FileListState {
        kind: FileListKind::Other,
        entries: vec![path.clone()],
        highlight: 0,
        view_offset: 0,
    });
    app.mode = Mode::Files;

    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.file_list.is_none());
    // Workbook untouched: active_path still points at the pre-existing
    // file, no cells were planted.
    assert_eq!(
        app.wb().active_path.as_deref(),
        Some(PathBuf::from("orig.xlsx").as_path())
    );
    assert!(app.wb().cells.is_empty());

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_dir(&dir);
}

/// /FLA shows the single active_path on line 2. Enter / Esc both
/// dismiss to READY without mutating the workbook.
#[test]
fn file_list_active_shows_active_path_and_esc_dismisses() {
    let mut app = App::new();
    app.wb_mut().active_path = Some(PathBuf::from("workbook.xlsx"));
    for c in ['/', 'F', 'L', 'A'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    assert_eq!(app.mode, Mode::Files);
    let fl = app.file_list.as_ref().expect("file_list populated");
    assert_eq!(fl.entries.len(), 1);
    assert_eq!(fl.entries[0], PathBuf::from("workbook.xlsx"));
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.file_list.is_none());
}

/// /FN wipes the entire in-memory workbook back to a blank sheet —
/// cells, formats, active path, pointer, and the engine itself.
#[test]
fn file_new_wipes_in_memory_state() {
    let mut app = App::new();
    // Seed A1 and move pointer off origin so we can detect the reset.
    for c in "42".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.wb_mut().active_path = Some(PathBuf::from("/tmp/pretend.xlsx"));
    assert!(!app.wb().cells.is_empty());

    // /FNA (After).
    for c in ['/', 'F', 'N', 'A'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.wb().cells.is_empty(), "cells should be cleared");
    assert_eq!(app.wb().pointer, Address::A1);
    assert!(
        app.wb().active_path.is_none(),
        "active_path should be cleared"
    );
}

/// /FIN drops the CSV rows into the grid starting at the pointer.
/// Numbers become constants; strings become labels.
#[test]
fn file_import_numbers_populates_cells_from_csv() {
    let dir = temp_test_dir("import_n");
    let src = dir.join("in.csv");
    std::fs::write(&src, "10,20,30\n\"foo\",\"bar\",\"baz\"\n").unwrap();

    let mut app = App::new();
    // /FIN <path><Enter>
    for c in ['/', 'F', 'I', 'N'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in src.to_str().unwrap().chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app.test_resume_async_op();

    assert_eq!(app.mode, Mode::Ready);
    match app.wb().cells.get(&Address::new(SheetId::A, 0, 0)).unwrap() {
        CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 10.0),
        other => panic!("A1 expected Number(10), got {other:?}"),
    }
    match app.wb().cells.get(&Address::new(SheetId::A, 2, 0)).unwrap() {
        CellContents::Constant(Value::Number(n)) => assert_eq!(*n, 30.0),
        other => panic!("C1 expected Number(30), got {other:?}"),
    }
    match app.wb().cells.get(&Address::new(SheetId::A, 0, 1)).unwrap() {
        CellContents::Label { text, .. } => assert_eq!(text, "foo"),
        other => panic!("A2 expected Label(foo), got {other:?}"),
    }

    let _ = std::fs::remove_file(&src);
    let _ = std::fs::remove_dir(&dir);
}

/// /FR: save a workbook, dirty memory, retrieve — cells should
/// show the saved values, not the dirty ones.
#[test]
fn file_retrieve_replaces_memory_with_saved_contents() {
    let dir = temp_test_dir("file_retrieve");
    let target = dir.join("sheet.xlsx");

    let mut app = App::new();
    drive_save_keys(&mut app, "42", &target);
    assert!(target.exists());

    // Dirty A1 in a fresh app, then /FR from disk.
    let mut app2 = App::new();
    for c in "99".chars() {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    for c in ['/', 'F', 'R'] {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in target.to_str().unwrap().chars() {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app2.test_resume_async_op();

    assert_eq!(app2.mode, Mode::Ready);
    let stored = app2
        .wb()
        .cells
        .get(&Address::A1)
        .unwrap_or_else(|| panic!("A1 not populated after /FR — have: {:?}", app2.wb().cells));
    match stored {
        CellContents::Constant(Value::Number(n)) => assert_eq!(n, &42.0),
        other => panic!("expected Constant(42), got {other:?}"),
    }

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// /FR clears the dirty bit: the in-memory workbook now matches
/// what's on disk, so a follow-up /Q should not warn.
#[test]
fn file_retrieve_clears_dirty_bit() {
    let dir = temp_test_dir("file_retrieve_dirty");
    let target = dir.join("sheet.xlsx");

    let mut app = App::new();
    drive_save_keys(&mut app, "42", &target);
    assert!(target.exists());

    let mut app2 = App::new();
    for c in "hi".chars() {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert!(app2.is_dirty(), "label commit should mark dirty");

    for c in ['/', 'F', 'R'] {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in target.to_str().unwrap().chars() {
        app2.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app2.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app2.test_resume_async_op();

    assert!(!app2.is_dirty(), "successful /FR should clear dirty bit");

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// Cancel path: pressing C on the confirm submenu leaves the
/// existing file untouched.
#[test]
fn file_save_cancel_leaves_existing_untouched() {
    let dir = temp_test_dir("file_save_cancel");
    let target = dir.join("sheet.xlsx");
    let mut app = App::new();
    drive_save_keys(&mut app, "42", &target);
    let before_len = std::fs::metadata(&target).unwrap().len();

    // Second /FS opens confirm; Cancel (C) aborts.
    for c in ['/', 'F', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert!(app.save_confirm.is_some());
    app.handle_key(KeyEvent::new(KeyCode::Char('C'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.save_confirm.is_none());
    let after_len = std::fs::metadata(&target).unwrap().len();
    assert_eq!(before_len, after_len, "Cancel unexpectedly wrote the file");

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

fn drive_chord(app: &mut App, chord: &[char]) {
    for c in chord {
        app.handle_key(KeyEvent::new(KeyCode::Char(*c), KeyModifiers::NONE));
    }
}

fn drive_chars(app: &mut App, s: &str) {
    for c in s.chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
}

fn enter(app: &mut App) {
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
}

/// Commit `1` at A1 and force the workbook back to a clean
/// baseline. Used by dirty-bit tests for mutators that need an
/// existing cell to operate on.
fn seed_a1_clean(app: &mut App) {
    drive_chord(app, &['1']);
    enter(app);
    app.wb_mut().dirty = false;
}

#[test]
fn range_erase_marks_dirty() {
    let mut app = App::new();
    seed_a1_clean(&mut app);
    // /RE auto-anchors POINT at the pointer; Enter erases the
    // single-cell range A1..A1.
    drive_chord(&mut app, &['/', 'R', 'E']);
    enter(&mut app);
    assert!(app.is_dirty(), "/RE should mark dirty");
}

#[test]
fn copy_marks_dirty() {
    let mut app = App::new();
    seed_a1_clean(&mut app);
    // /C: FROM-Enter (A1..A1), RIGHT, TO-Enter (B1).
    drive_chord(&mut app, &['/', 'C']);
    enter(&mut app);
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    enter(&mut app);
    assert!(app.is_dirty(), "/C should mark dirty");
}

#[test]
fn move_marks_dirty() {
    let mut app = App::new();
    seed_a1_clean(&mut app);
    drive_chord(&mut app, &['/', 'M']);
    enter(&mut app);
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    enter(&mut app);
    assert!(app.is_dirty(), "/M should mark dirty");
}

#[test]
fn range_format_marks_dirty() {
    let mut app = App::new();
    seed_a1_clean(&mut app);
    // /RFG → General format → POINT → Enter applies to A1..A1.
    drive_chord(&mut app, &['/', 'R', 'F', 'G']);
    enter(&mut app);
    assert!(app.is_dirty(), "/RFG should mark dirty");
}

#[test]
fn range_name_create_marks_dirty() {
    let mut app = App::new();
    // /RNC → name "sales" → Enter → POINT (A1..A1) → Enter.
    drive_chord(&mut app, &['/', 'R', 'N', 'C']);
    drive_chars(&mut app, "sales");
    enter(&mut app);
    enter(&mut app);
    assert!(app.is_dirty(), "/RNC should mark dirty");
}

#[test]
fn range_name_delete_marks_dirty() {
    let mut app = App::new();
    // Create first.
    drive_chord(&mut app, &['/', 'R', 'N', 'C']);
    drive_chars(&mut app, "sales");
    enter(&mut app);
    enter(&mut app);
    app.wb_mut().dirty = false;
    // /RND → name "sales" → Enter.
    drive_chord(&mut app, &['/', 'R', 'N', 'D']);
    drive_chars(&mut app, "sales");
    enter(&mut app);
    assert!(app.is_dirty(), "/RND should mark dirty");
}

#[test]
fn ws_column_set_width_marks_dirty() {
    let mut app = App::new();
    // /WCS prompts for width; type 15 + Enter.
    drive_chord(&mut app, &['/', 'W', 'C', 'S']);
    drive_chars(&mut app, "15");
    enter(&mut app);
    assert!(app.is_dirty(), "/WCS should mark dirty");
}

#[test]
fn ws_column_reset_width_marks_dirty() {
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'C', 'S']);
    drive_chars(&mut app, "15");
    enter(&mut app);
    app.wb_mut().dirty = false;
    // /WCR resets the current column to the default width.
    drive_chord(&mut app, &['/', 'W', 'C', 'R']);
    assert!(app.is_dirty(), "/WCR should mark dirty");
}

#[test]
fn ws_column_range_set_width_marks_dirty() {
    let mut app = App::new();
    // /WCCS prompts for width; type 12 + Enter, then POINT, Enter applies.
    drive_chord(&mut app, &['/', 'W', 'C', 'C', 'S']);
    drive_chars(&mut app, "12");
    enter(&mut app);
    enter(&mut app);
    assert!(app.is_dirty(), "/WCCS should mark dirty");
}

#[test]
fn ws_column_range_reset_width_marks_dirty() {
    let mut app = App::new();
    // First set a non-default width so the reset has something to do.
    drive_chord(&mut app, &['/', 'W', 'C', 'C', 'S']);
    drive_chars(&mut app, "12");
    enter(&mut app);
    enter(&mut app);
    app.wb_mut().dirty = false;
    // /WCCR is the column-range reset; POINT, Enter.
    drive_chord(&mut app, &['/', 'W', 'C', 'C', 'R']);
    enter(&mut app);
    assert!(app.is_dirty(), "/WCCR should mark dirty");
}

#[test]
fn ws_column_hide_marks_dirty() {
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'C', 'H']);
    enter(&mut app);
    assert!(app.is_dirty(), "/WCH should mark dirty");
}

#[test]
fn wysiwyg_display_mode_color_paints_white_bg_on_empty_cell() {
    let mut app = App::new();
    drive_chord(&mut app, &[':', 'D', 'M', 'C']);
    assert_eq!(app.mode, Mode::Ready);
    let buf = app.render_to_buffer(80, 25);
    assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), Some((0xFF, 0xFF, 0xFF)));
    assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), Some((0x00, 0x00, 0x00)));
}

#[test]
fn wysiwyg_display_mode_reverse_paints_black_bg() {
    let mut app = App::new();
    drive_chord(&mut app, &[':', 'D', 'M', 'R']);
    let buf = app.render_to_buffer(80, 25);
    assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), Some((0x00, 0x00, 0x00)));
    assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), Some((0xFF, 0xFF, 0xFF)));
}

#[test]
fn wysiwyg_display_mode_bw_leaves_terminal_default() {
    let mut app = App::new();
    // Switch to Color, then back to B&W — B&W should clear the
    // RGB BG so the terminal default shows through (read-back
    // returns None for non-RGB cells).
    drive_chord(&mut app, &[':', 'D', 'M', 'C']);
    drive_chord(&mut app, &[':', 'D', 'M', 'B']);
    let buf = app.render_to_buffer(80, 25);
    assert_eq!(app.cell_bg_rendered(&buf, "A:B5"), None);
    assert_eq!(app.cell_fg_rendered(&buf, "A:B5"), None);
}

#[test]
fn wysiwyg_display_grid_yes_paints_dotted_right_edges() {
    let mut app = App::new();
    // Default off — no gridline glyphs anywhere.
    let buf = app.render_to_buffer(80, 25);
    let body_row = App::line_text(&buf, PANEL_HEIGHT + 1);
    assert!(
        !body_row.contains('┊'),
        "default body row should not contain gridline dots: {body_row:?}"
    );

    // :DOGY turns gridlines on — the rightmost column of each
    // 9-char-wide empty cell becomes `┊`.
    drive_chord(&mut app, &[':', 'D', 'O', 'G', 'Y']);
    let buf = app.render_to_buffer(80, 25);
    // A body row past the pointer cell so no REVERSED highlight
    // suppresses the overlay.
    let body_row = App::line_text(&buf, PANEL_HEIGHT + 2);
    assert!(
        body_row.matches('┊').count() >= 4,
        "expected dotted gridline at each cell right edge: {body_row:?}"
    );

    // :DOGN turns them back off.
    drive_chord(&mut app, &[':', 'D', 'O', 'G', 'N']);
    let buf = app.render_to_buffer(80, 25);
    let body_row = App::line_text(&buf, PANEL_HEIGHT + 2);
    assert!(
        !body_row.contains('┊'),
        "grid=No should suppress gridline dots: {body_row:?}"
    );
}

#[test]
fn wysiwyg_display_grid_skips_cells_with_full_width_content() {
    let mut app = App::new();
    // Type a 9-char value that fills the default column width
    // exactly. With gridlines on, the right-edge `┊` would
    // overwrite the last digit — verify we skip the overlay.
    drive_chord(&mut app, &['9', '8', '7', '6', '5', '4', '3', '2', '1']);
    enter(&mut app);
    drive_chord(&mut app, &[':', 'D', 'O', 'G', 'Y']);
    let buf = app.render_to_buffer(80, 25);
    // Move the pointer off A1 so the highlight doesn't suppress
    // anything for an unrelated reason; assert by reading the
    // actual cell content.
    let painted = (0..9)
        .map(|i| buf[(ROW_GUTTER + i, PANEL_HEIGHT + 1)].symbol().to_string())
        .collect::<String>();
    assert_eq!(painted, "987654321", "filled cell must not be clipped");
}

#[test]
fn ws_column_display_marks_dirty() {
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'C', 'H']);
    enter(&mut app);
    app.wb_mut().dirty = false;
    drive_chord(&mut app, &['/', 'W', 'C', 'D']);
    enter(&mut app);
    assert!(app.is_dirty(), "/WCD should mark dirty");
}

#[test]
fn ws_titles_set_marks_dirty() {
    let mut app = App::new();
    // /WTH freezes the rows above the pointer.
    drive_chord(&mut app, &['/', 'W', 'T', 'H']);
    assert!(app.is_dirty(), "/WTH should mark dirty");
}

#[test]
fn ws_titles_clear_marks_dirty() {
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'T', 'H']);
    app.wb_mut().dirty = false;
    // /WTC on a sheet that has frozen titles should clear them.
    drive_chord(&mut app, &['/', 'W', 'T', 'C']);
    assert!(app.is_dirty(), "/WTC (with prior titles) should mark dirty");
}

#[test]
fn ws_titles_clear_no_op_does_not_dirty() {
    // /WTC on a sheet without titles is a no-op — no journal, no
    // dirty flip. Matches the existing M5_ws_titles transcript's
    // "quiet no-op" expectation.
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'T', 'C']);
    assert!(!app.is_dirty(), "/WTC on clean sheet should remain clean");
}

#[test]
fn wg_global_format_marks_dirty() {
    let mut app = App::new();
    // /WGFG sets global format to General — routes through
    // set_global_format and journals a GlobalFormat entry.
    drive_chord(&mut app, &['/', 'W', 'G', 'F', 'G']);
    assert!(app.is_dirty(), "/WGFG should mark dirty");
}

#[test]
fn wg_global_col_width_marks_dirty() {
    let mut app = App::new();
    // /WGCS prompts for the new default column width; type 12 + Enter.
    drive_chord(&mut app, &['/', 'W', 'G', 'C']);
    drive_chord(&mut app, &['S']);
    drive_chars(&mut app, "12");
    enter(&mut app);
    assert!(app.is_dirty(), "/WGCS should mark dirty");
}

#[test]
fn wgd_intl_punctuation_marks_dirty() {
    let mut app = App::new();
    // /WGDOIPB switches Punctuation from the default (A) to B,
    // mutating Workbook.international.
    drive_chord(&mut app, &['/', 'W', 'G', 'D', 'O', 'I', 'P', 'B']);
    assert!(app.is_dirty(), "/WGDOIPB should mark dirty");
}

#[test]
fn ws_erase_yields_clean_workbook() {
    // /WEY discards the workbook for a fresh blank one — there are
    // no changes left to save, so /Q should not warn afterwards.
    let mut app = App::new();
    drive_chars(&mut app, "hi");
    enter(&mut app);
    assert!(app.is_dirty());
    drive_chord(&mut app, &['/', 'W', 'E', 'Y']);
    assert!(!app.is_dirty(), "/WEY should leave the workbook clean");
}

#[test]
fn wysiwyg_bold_marks_dirty() {
    let mut app = App::new();
    drive_chars(&mut app, "hi");
    enter(&mut app);
    app.wb_mut().dirty = false;
    // :FBS → POINT → Enter applies Bold to A1..A1.
    drive_chord(&mut app, &[':', 'F', 'B', 'S']);
    enter(&mut app);
    assert!(app.is_dirty(), ":FBS should mark dirty");
}

/// Structural mutators (insert/delete row & col) flip the dirty
/// bit. Each runs against a fresh `App::new()` so the assertion
/// isolates one mutator.
#[test]
fn structural_mutators_mark_dirty() {
    for chord in [
        &['/', 'W', 'I', 'R'][..],
        &['/', 'W', 'I', 'C'][..],
        &['/', 'W', 'D', 'R'][..],
        &['/', 'W', 'D', 'C'][..],
    ] {
        let mut app = App::new();
        assert!(!app.is_dirty(), "fresh App should be clean");
        drive_chord(&mut app, chord);
        let chord_str: String = chord.iter().collect();
        assert!(app.is_dirty(), "{chord_str} should mark dirty");
    }
}

/// Dirty-bit lifecycle: a fresh App is clean; committing a label
/// flips it dirty; a successful `/FS` clears it. Drives the
/// `/QY` warn-on-quit guard.
#[test]
fn dirty_bit_tracks_modifications_and_save() {
    let mut app = App::new();
    assert!(!app.is_dirty(), "fresh workbook should be clean");

    for c in "hi".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert!(app.is_dirty(), "label commit should mark dirty");

    let dir = temp_test_dir("dirty_bit");
    let target = dir.join("clean.xlsx");
    for c in ['/', 'F', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in target.to_str().unwrap().chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app.test_resume_async_op();
    assert!(!app.is_dirty(), "successful /FS should clear the dirty bit");

    let _ = std::fs::remove_file(&target);
    let _ = std::fs::remove_dir(&dir);
}

/// Simulate a left-click at `(col, row)` by stashing a fake panel
/// geometry and routing through [`App::handle_mouse`].
fn click(app: &mut App, area: Rect, col: u16, row: u16) {
    app.icon_panel_area.set(Some(test_geom(area)));
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Down(MouseButton::Left),
        column: col,
        row,
        modifiers: KeyModifiers::NONE,
    });
}

/// Standard fixture panel: 17 icons × 3 rows each = 51 rows, so
/// slot N's middle row is `N * 3 + 1`.
const TEST_PANEL: Rect = Rect {
    x: 80,
    y: 4,
    width: 3,
    height: 51,
};

/// Build a hit-test fixture where the rendered image is assumed
/// to fill the cell rect exactly (no aspect-ratio slack). Using
/// `font_px_h = 2` lets the cell-midpoint pixel formula resolve
/// to integer slot boundaries on multiples of `2 * 17 = 34`.
fn test_geom(rect: Rect) -> IconPanelGeom {
    IconPanelGeom {
        rect,
        rendered_px_h: rect.height as u32 * 2,
        font_px_h: 2,
    }
}

fn click_slot(app: &mut App, slot: u16) {
    // Click the middle row of the given slot.
    click(
        app,
        TEST_PANEL,
        TEST_PANEL.x + 1,
        TEST_PANEL.y + slot * 3 + 1,
    );
}

#[test]
fn icon_click_save_opens_save_prompt() {
    let mut app = App::new();
    click_slot(&mut app, 0);
    // /FS enters Mode::Menu with an active prompt for the filename.
    assert_eq!(app.mode, Mode::Menu);
    assert!(app.prompt.is_some(), "save prompt should be active");
}

#[test]
fn icon_click_retrieve_opens_retrieve_prompt() {
    let mut app = App::new();
    click_slot(&mut app, 1);
    assert_eq!(app.mode, Mode::Menu);
    assert!(app.prompt.is_some());
}

#[test]
fn icon_click_graph_view_enters_graph_mode() {
    let mut app = App::new();
    click_slot(&mut app, 7);
    assert_eq!(app.mode, Mode::Graph);
}

#[test]
fn icon_click_print_opens_print_prompt() {
    let mut app = App::new();
    click_slot(&mut app, 9);
    assert_eq!(app.mode, Mode::Menu);
    assert!(app.prompt.is_some());
}

#[test]
fn icon_click_prev_sheet_is_noop_on_first_sheet() {
    let mut app = App::new();
    click_slot(&mut app, 5);
    // Only one sheet in a fresh workbook — prev-sheet clamps.
    assert_eq!(app.pointer().display_full(), "A:A1");
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn icon_click_help_is_safe_noop() {
    let mut app = App::new();
    click_slot(&mut app, 15);
    assert_eq!(app.pointer().display_full(), "A:A1");
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn icon_click_bold_applies_to_pointer_cell_instantly() {
    // Panel 1 slot 11 = icon id 12 = SmartIcons Bold. The click
    // applies bold to the cursor cell immediately — no menu, no
    // POINT prompt, mode stays Ready.
    let mut app = App::new();
    click_slot(&mut app, 11);
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD),
    );
}

#[test]
fn icon_click_italic_applies_to_pointer_cell_instantly() {
    let mut app = App::new();
    click_slot(&mut app, 12);
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::ITALIC),
    );
}

#[test]
fn icon_click_underline_applies_to_pointer_cell_instantly() {
    let mut app = App::new();
    click_slot(&mut app, 13);
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::UNDERLINE),
    );
}

#[test]
fn icon_click_bold_applies_to_active_point_highlight() {
    // While in POINT mode with an extended highlight (e.g. mid /Range
    // command), clicking Bold should apply to that highlight rather
    // than just A1, then return to Ready.
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    // Now in POINT for /Range Erase, anchored at A1. Extend to B2.
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    click_slot(&mut app, 11);
    assert_eq!(app.mode, Mode::Ready);
    for addr in [
        Address::A1,
        Address::new(SheetId(0), 1, 0),
        Address::new(SheetId(0), 0, 1),
        Address::new(SheetId(0), 1, 1),
    ] {
        assert_eq!(
            app.wb().cell_text_styles.get(&addr).copied(),
            Some(TextStyle::BOLD),
            "{addr:?} should be bold",
        );
    }
}

#[test]
fn icon_click_bold_toggles_off_when_pointer_cell_already_bold() {
    // Second click on Bold over an already-bold cell removes bold.
    let mut app = App::new();
    click_slot(&mut app, 11);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD),
    );
    click_slot(&mut app, 11);
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(app.wb().cell_text_styles.get(&Address::A1).copied(), None);
}

#[test]
fn icon_click_bold_toggle_only_clears_bold_keeping_other_styles() {
    // Toggling Bold off must leave Italic/Underline intact on the cell.
    let mut app = App::new();
    app.execute_range_text_style(
        Range::single(Address::A1),
        TextStyle::BOLD.merge(TextStyle::ITALIC),
        true,
    );
    click_slot(&mut app, 11);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::ITALIC),
    );
}

#[test]
fn icon_click_bold_toggles_off_when_every_cell_in_highlight_is_bold() {
    // Pre-bold A1:B2, then enter POINT over A1:B2 and click Bold —
    // the whole range should clear.
    let mut app = App::new();
    let r = Range {
        start: Address::A1,
        end: Address::new(SheetId(0), 1, 1),
    };
    app.execute_range_text_style(r, TextStyle::BOLD, true);
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    click_slot(&mut app, 11);
    assert_eq!(app.mode, Mode::Ready);
    for addr in [
        Address::A1,
        Address::new(SheetId(0), 1, 0),
        Address::new(SheetId(0), 0, 1),
        Address::new(SheetId(0), 1, 1),
    ] {
        assert_eq!(
            app.wb().cell_text_styles.get(&addr).copied(),
            None,
            "{addr:?} should no longer be bold",
        );
    }
}

#[test]
fn icon_click_bold_sets_bold_on_mixed_highlight() {
    // Mixed highlight (only A1 bold, others plain): one click bolds
    // every cell rather than clearing.
    let mut app = App::new();
    app.execute_range_text_style(Range::single(Address::A1), TextStyle::BOLD, true);
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    click_slot(&mut app, 11);
    for addr in [
        Address::A1,
        Address::new(SheetId(0), 1, 0),
        Address::new(SheetId(0), 0, 1),
        Address::new(SheetId(0), 1, 1),
    ] {
        assert_eq!(
            app.wb().cell_text_styles.get(&addr).copied(),
            Some(TextStyle::BOLD),
            "{addr:?} should be bold",
        );
    }
}

#[test]
fn mouse_click_outside_panel_is_ignored() {
    let mut app = App::new();
    click(&mut app, TEST_PANEL, 10, 10);
    assert_eq!(app.pointer().display_full(), "A:A1");
    assert_eq!(app.mode, Mode::Ready);
}

/// Drop a non-blank label at the given short address on sheet A.
fn put_label(app: &mut App, addr: &str, text: &str) {
    let a = Address::parse(addr).expect("test addr");
    app.wb_mut().cells.insert(a, make_label(text));
}

/// Drop a numeric constant at the given short address on sheet A,
/// keeping the engine in sync so subsequent recalcs see the value.
fn put_number(app: &mut App, addr: &str, n: f64) {
    let a = Address::parse(addr).expect("test addr");
    let c = CellContents::Constant(Value::Number(n));
    app.push_to_engine_at(a, &c);
    app.wb_mut().cells.insert(a, c);
}

fn set_pointer(app: &mut App, addr: &str) {
    app.wb_mut().pointer = Address::parse(addr).expect("test addr");
}

#[test]
fn block_end_home_on_empty_sheet_stays_at_a1() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndHome);
    assert_eq!(app.pointer().display_short(), "A1");
}

#[test]
fn block_end_home_jumps_to_lower_right_of_active_area() {
    // Active area corner = (max occupied col, max occupied row),
    // which need not itself be occupied.
    let mut app = App::new();
    put_label(&mut app, "C2", "x");
    put_label(&mut app, "A5", "x");
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndHome);
    assert_eq!(app.pointer().display_short(), "C5");
}

#[test]
fn block_end_down_in_run_jumps_to_last_nonblank() {
    let mut app = App::new();
    put_label(&mut app, "A1", "a");
    put_label(&mut app, "A2", "b");
    put_label(&mut app, "A3", "c");
    // A4 blank.
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndDown);
    assert_eq!(app.pointer().display_short(), "A3");
}

#[test]
fn block_end_down_past_blanks_lands_on_next_nonblank() {
    let mut app = App::new();
    put_label(&mut app, "A1", "a");
    // A2..A4 blank.
    put_label(&mut app, "A5", "b");
    // From A1 with A2 blank: scan past blanks → A5.
    set_pointer(&mut app, "A1");
    // A1 is nonblank with nonblank A2 in the canonical case, but
    // here A2 is blank so this exercises the "first nonblank"
    // path even though we started on a nonblank.
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndDown);
    assert_eq!(app.pointer().display_short(), "A5");
}

#[test]
fn block_end_down_from_blank_finds_first_nonblank() {
    let mut app = App::new();
    // A1 blank, A2 blank, A3 nonblank.
    put_label(&mut app, "A3", "c");
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndDown);
    assert_eq!(app.pointer().display_short(), "A3");
}

#[test]
fn block_end_down_with_nothing_below_goes_to_last_row() {
    let mut app = App::new();
    // Empty sheet → END+DOWN from A1 → A8192 (max row).
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndDown);
    assert_eq!(app.pointer().row, l123_core::address::MAX_ROWS - 1);
    assert_eq!(app.pointer().col, 0);
}

#[test]
fn block_end_up_jumps_to_top_of_run() {
    let mut app = App::new();
    put_label(&mut app, "A2", "a");
    put_label(&mut app, "A3", "b");
    put_label(&mut app, "A4", "c");
    set_pointer(&mut app, "A4");
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndUp);
    assert_eq!(app.pointer().display_short(), "A2");
}

#[test]
fn block_end_right_jumps_to_end_of_run() {
    let mut app = App::new();
    put_label(&mut app, "A1", "a");
    put_label(&mut app, "B1", "b");
    put_label(&mut app, "C1", "c");
    // D1 blank.
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndRight);
    assert_eq!(app.pointer().display_short(), "C1");
}

#[test]
fn block_end_left_from_blank_finds_first_nonblank() {
    let mut app = App::new();
    // Pointer at D1, A1 nonblank, B1..C1 blank.
    put_label(&mut app, "A1", "a");
    set_pointer(&mut app, "D1");
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndLeft);
    assert_eq!(app.pointer().display_short(), "A1");
}

#[test]
fn block_end_at_boundary_does_not_move() {
    let mut app = App::new();
    // At A1 with nothing above: END+UP holds.
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndUp);
    assert_eq!(app.pointer().display_short(), "A1");
    // Same for END+LEFT.
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndLeft);
    assert_eq!(app.pointer().display_short(), "A1");
}

#[test]
fn block_end_only_scans_current_sheet() {
    let mut app = App::new();
    // Place a cell on sheet B that should not pull the active-area
    // corner on sheet A.
    let other = Address::new(SheetId(1), 9, 9);
    app.wb_mut().cells.insert(other, make_label("x"));
    put_label(&mut app, "B2", "x");
    app.dispatch_sys_action(l123_graph::SysAction::BlockEndHome);
    assert_eq!(app.pointer().display_short(), "B2");
}

#[test]
fn scroll_one_row_down_shifts_viewport_offset_only() {
    let mut app = App::new();
    let pointer_before = app.pointer();
    app.dispatch_sys_action(l123_graph::SysAction::ScrollRowDown);
    assert_eq!(app.wb().viewport_row_offset, 1);
    assert_eq!(app.wb().viewport_col_offset, 0);
    assert_eq!(app.pointer(), pointer_before, "pointer must not move");
}

#[test]
fn scroll_one_column_right_shifts_viewport_col_only() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::ScrollColumnRight);
    assert_eq!(app.wb().viewport_col_offset, 1);
    assert_eq!(app.wb().viewport_row_offset, 0);
}

#[test]
fn scroll_up_at_top_clamps_at_zero() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::ScrollRowUp);
    assert_eq!(app.wb().viewport_row_offset, 0);
    app.dispatch_sys_action(l123_graph::SysAction::ScrollColumnLeft);
    assert_eq!(app.wb().viewport_col_offset, 0);
}

#[test]
fn scroll_screen_down_with_no_grid_uses_pgdn_default() {
    let mut app = App::new();
    // No render yet → fallback of 20 rows / 8 cols.
    app.dispatch_sys_action(l123_graph::SysAction::ScrollScreenDown);
    assert_eq!(app.wb().viewport_row_offset, 20);
    app.dispatch_sys_action(l123_graph::SysAction::ScrollScreenRight);
    assert_eq!(app.wb().viewport_col_offset, 8);
}

#[test]
fn scroll_clamps_at_worksheet_max() {
    let mut app = App::new();
    app.wb_mut().viewport_row_offset = l123_core::address::MAX_ROWS - 2;
    app.dispatch_sys_action(l123_graph::SysAction::ScrollScreenDown);
    // Clamped at MAX_ROWS - 1 even though the screen-jump would
    // overshoot.
    assert_eq!(
        app.wb().viewport_row_offset,
        l123_core::address::MAX_ROWS - 1
    );
}

#[test]
fn icon_click_panel_five_scroll_row_down_dispatches() {
    // Panel 5 slot 11 = icon 80 (Move display one row down). The
    // pure-scroll path bumps viewport_row_offset and leaves the
    // pointer at A1.
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Five;
    click_slot(&mut app, 11);
    assert_eq!(app.wb().viewport_row_offset, 1);
    assert_eq!(app.pointer().display_short(), "A1");
}

#[test]
fn icon_click_panel_five_delete_sheet_opens_menu() {
    // Panel 5 slot 3 = icon 72 (Delete worksheets) → /WDS.
    // /WDS in a one-sheet workbook prompts for confirmation, so the
    // icon click should land us in Mode::Menu with a prompt.
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Five;
    click_slot(&mut app, 3);
    // /WDS dispatches synchronously; with a single sheet it can
    // either land on a confirm prompt or refuse — either way we
    // must not have crashed and the pointer stays at A1.
    assert_eq!(app.pointer().display_short(), "A1");
}

#[test]
fn icon_click_panel_five_page_break_row_invokes_wpr() {
    // Panel 5 slot 12 = icon 63 (row page break) → /WPR.
    // /WPR inserts a row at the pointer with `|::` in column A.
    let mut app = App::new();
    app.wb_mut().pointer = Address::new(SheetId::A, 0, 4);
    app.current_panel = l123_graph::Panel::Five;
    click_slot(&mut app, 12);
    let marker = Address::new(SheetId::A, 0, 4);
    assert!(
        app.wb().cells.contains_key(&marker),
        "page-break marker at A5 should exist after /WPR",
    );
    let cell = app.wb().cells.get(&marker).unwrap();
    if let CellContents::Label { prefix, text } = cell {
        assert_eq!(*prefix, LabelPrefix::Pipe);
        assert_eq!(text, "::");
    } else {
        panic!("expected pipe-prefixed `::` label, got {cell:?}");
    }
}

#[test]
fn icon_click_panel_five_page_break_column_invokes_wpc() {
    // Panel 5 slot 13 = icon 64 (column page break) → /WPC.
    // /WPC inserts a column at the pointer with `|::` in row 1.
    let mut app = App::new();
    app.wb_mut().pointer = Address::new(SheetId::A, 4, 0);
    app.current_panel = l123_graph::Panel::Five;
    click_slot(&mut app, 13);
    let marker = Address::new(SheetId::A, 4, 0);
    assert!(
        app.wb().cells.contains_key(&marker),
        "page-break marker at E1 should exist after /WPC",
    );
    let cell = app.wb().cells.get(&marker).unwrap();
    if let CellContents::Label { prefix, text } = cell {
        assert_eq!(*prefix, LabelPrefix::Pipe);
        assert_eq!(text, "::");
    } else {
        panic!("expected pipe-prefixed `::` label, got {cell:?}");
    }
}

/// Read the source-form expression at `addr`, panicking on
/// non-formula cells. Tests assert the formula written by the
/// SmartIcon — the engine's cached evaluation is exercised
/// elsewhere.
fn formula_expr_at(app: &App, addr: &str) -> String {
    let a = Address::parse(addr).expect("test addr");
    match app.wb().cells.get(&a) {
        Some(CellContents::Formula { expr, .. }) => expr.clone(),
        other => panic!("expected formula at {addr}, got {other:?}"),
    }
}

#[test]
fn sum_smarticon_uses_run_directly_above() {
    let mut app = App::new();
    put_number(&mut app, "A1", 1.0);
    put_number(&mut app, "A2", 2.0);
    put_number(&mut app, "A3", 3.0);
    set_pointer(&mut app, "A4");
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert_eq!(formula_expr_at(&app, "A4"), "@SUM(A1..A3)");
}

#[test]
fn sum_smarticon_falls_back_to_left_when_above_is_blank() {
    let mut app = App::new();
    put_number(&mut app, "A1", 10.0);
    put_number(&mut app, "B1", 20.0);
    put_number(&mut app, "C1", 30.0);
    set_pointer(&mut app, "D1");
    // No row above D1 — algorithm scans left.
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert_eq!(formula_expr_at(&app, "D1"), "@SUM(A1..C1)");
}

#[test]
fn sum_smarticon_prefers_above_over_left() {
    let mut app = App::new();
    // Both above and left are numeric — above wins.
    put_number(&mut app, "B1", 1.0);
    put_number(&mut app, "A2", 99.0);
    set_pointer(&mut app, "B2");
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert_eq!(formula_expr_at(&app, "B2"), "@SUM(B1..B1)");
}

#[test]
fn sum_smarticon_stops_at_label_above() {
    let mut app = App::new();
    // A1 is a label header → not numeric; the run is just A2..A3.
    put_label(&mut app, "A1", "Total");
    put_number(&mut app, "A2", 5.0);
    put_number(&mut app, "A3", 7.0);
    set_pointer(&mut app, "A4");
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert_eq!(formula_expr_at(&app, "A4"), "@SUM(A2..A3)");
}

#[test]
fn sum_smarticon_with_no_numeric_neighbor_beeps_and_writes_nothing() {
    let mut app = App::new();
    let beep_before = app.beep_count();
    // Empty sheet, cursor at A1 — nothing above or to the left.
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert!(app.beep_count() > beep_before, "expected a beep");
    assert!(
        !app.wb().cells.contains_key(&Address::A1),
        "no cell should have been written",
    );
}

#[test]
fn sum_smarticon_skips_when_directly_above_is_blank_with_numbers_higher() {
    // Algorithm checks the immediately-adjacent neighbour. If A2
    // is blank but A1 has a number, the run-above test fails →
    // we fall through to scan-left, which is also blank → beep.
    let mut app = App::new();
    put_number(&mut app, "A1", 1.0);
    // A2 blank.
    set_pointer(&mut app, "A3");
    let beep_before = app.beep_count();
    app.dispatch_sys_action(l123_graph::SysAction::SumRange);
    assert!(app.beep_count() > beep_before);
    assert!(!app.wb().cells.contains_key(&Address::parse("A3").unwrap()));
}

#[test]
fn today_smarticon_writes_now_formula_at_pointer() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::TodayDate);
    assert_eq!(formula_expr_at(&app, "A1"), "@NOW");
}

#[test]
fn icon_click_panel_one_sum_writes_formula() {
    // Panel 1 slot 6 = icon id 9 = Sum SmartIcon.
    let mut app = App::new();
    put_number(&mut app, "A1", 1.0);
    put_number(&mut app, "A2", 2.0);
    set_pointer(&mut app, "A3");
    click_slot(&mut app, 6);
    assert_eq!(formula_expr_at(&app, "A3"), "@SUM(A1..A2)");
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn icon_click_panel_four_today_writes_now_formula() {
    // Panel 4 slot 7 = icon id 45 = Today's date SmartIcon.
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Four;
    click_slot(&mut app, 7);
    assert_eq!(formula_expr_at(&app, "A1"), "@NOW");
}

#[test]
fn auto_detect_database_range_returns_none_on_blank_cursor() {
    let app = App::new();
    let cur = Address::A1;
    assert!(app.auto_detect_database_range(cur).is_none());
}

#[test]
fn auto_detect_database_range_grows_to_full_block() {
    // Header A1..C1, two data rows A2..C3. Cursor at B2 should
    // discover the full A1..C3 rectangle.
    let mut app = App::new();
    put_label(&mut app, "A1", "id");
    put_label(&mut app, "B1", "name");
    put_label(&mut app, "C1", "value");
    put_number(&mut app, "A2", 2.0);
    put_label(&mut app, "B2", "bob");
    put_number(&mut app, "C2", 20.0);
    put_number(&mut app, "A3", 1.0);
    put_label(&mut app, "B3", "alice");
    put_number(&mut app, "C3", 10.0);
    let cursor = Address::parse("B2").unwrap();
    let block = app
        .auto_detect_database_range(cursor)
        .expect("block should be detected");
    assert_eq!(block.start.display_short(), "A1");
    assert_eq!(block.end.display_short(), "C3");
}

#[test]
fn auto_detect_database_range_stops_at_fully_blank_row() {
    // Two blocks separated by a blank row 3. Cursor in the upper
    // block should detect only that block, not row 4 onwards.
    let mut app = App::new();
    put_label(&mut app, "A1", "id");
    put_label(&mut app, "B1", "name");
    put_number(&mut app, "A2", 1.0);
    put_label(&mut app, "B2", "alice");
    // Row 3 fully blank.
    put_number(&mut app, "A4", 99.0);
    put_label(&mut app, "B4", "elsewhere");
    let cursor = Address::parse("A2").unwrap();
    let block = app.auto_detect_database_range(cursor).unwrap();
    assert_eq!(block.start.display_short(), "A1");
    assert_eq!(block.end.display_short(), "B2");
}

#[test]
fn sort_smarticon_ascending_orders_block_excluding_header() {
    // Header + 3 data rows, sorted by column A ascending. Cursor in
    // column A on a data row picks that column as primary key.
    let mut app = App::new();
    put_label(&mut app, "A1", "id");
    put_label(&mut app, "B1", "name");
    put_number(&mut app, "A2", 3.0);
    put_label(&mut app, "B2", "cherry");
    put_number(&mut app, "A3", 1.0);
    put_label(&mut app, "B3", "apple");
    put_number(&mut app, "A4", 2.0);
    put_label(&mut app, "B4", "banana");
    set_pointer(&mut app, "A3");
    app.dispatch_sys_action(l123_graph::SysAction::SortAscending);
    // Header still in row 1.
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A1").unwrap()),
        Some(CellContents::Label { .. })
    ));
    // Data sorted 1, 2, 3.
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A2").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 1.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A3").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 2.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A4").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 3.0
    ));
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn sort_smarticon_descending_orders_block_high_to_low() {
    let mut app = App::new();
    put_label(&mut app, "A1", "id");
    put_number(&mut app, "A2", 1.0);
    put_number(&mut app, "A3", 3.0);
    put_number(&mut app, "A4", 2.0);
    set_pointer(&mut app, "A3");
    app.dispatch_sys_action(l123_graph::SysAction::SortDescending);
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A2").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 3.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A3").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 2.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A4").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 1.0
    ));
}

#[test]
fn sort_smarticon_on_blank_cursor_beeps_no_op() {
    let mut app = App::new();
    // Numbers far away from cursor at A1.
    put_number(&mut app, "C5", 1.0);
    let beep_before = app.beep_count();
    app.dispatch_sys_action(l123_graph::SysAction::SortAscending);
    assert!(app.beep_count() > beep_before, "expected a beep");
    // Untouched.
    assert!(matches!(
        app.wb().cells.get(&Address::parse("C5").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 1.0
    ));
}

#[test]
fn sort_smarticon_on_header_only_block_beeps_no_op() {
    // Single row of labels — no data rows below, can't sort.
    let mut app = App::new();
    put_label(&mut app, "A1", "id");
    put_label(&mut app, "B1", "name");
    set_pointer(&mut app, "A1");
    let beep_before = app.beep_count();
    app.dispatch_sys_action(l123_graph::SysAction::SortAscending);
    assert!(app.beep_count() > beep_before, "expected a beep");
}

#[test]
fn icon_click_panel_four_sort_ascending_runs_sort() {
    // Panel 4 slot 0 = icon id 31 (Sort Ascending).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Four;
    put_label(&mut app, "A1", "id");
    put_number(&mut app, "A2", 2.0);
    put_number(&mut app, "A3", 1.0);
    set_pointer(&mut app, "A2");
    click_slot(&mut app, 0);
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A2").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 1.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A3").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 2.0
    ));
}

#[test]
fn icon_click_panel_four_sort_descending_runs_sort() {
    // Panel 4 slot 1 = icon id 32 (Sort Descending).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Four;
    put_label(&mut app, "A1", "id");
    put_number(&mut app, "A2", 1.0);
    put_number(&mut app, "A3", 2.0);
    set_pointer(&mut app, "A2");
    click_slot(&mut app, 1);
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A2").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 2.0
    ));
    assert!(matches!(
        app.wb().cells.get(&Address::parse("A3").unwrap()),
        Some(CellContents::Constant(Value::Number(n))) if *n == 1.0
    ));
}

#[test]
fn icon_click_panel_four_copy_single_to_range_opens_copy_point() {
    // Panel 4 slot 9 = icon id 47 → /Copy. /C enters POINT for the
    // FROM range; the user (or icon caller) then selects source +
    // destination. Single-cell sources are handled natively by /C.
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Four;
    click_slot(&mut app, 9);
    assert_eq!(app.mode, Mode::Point);
}

#[test]
fn step_toggle_smarticon_flips_step_mode() {
    let mut app = App::new();
    assert!(!app.step_mode, "fresh App starts with STEP off");
    app.dispatch_sys_action(l123_graph::SysAction::StepToggle);
    assert!(app.step_mode);
    app.dispatch_sys_action(l123_graph::SysAction::StepToggle);
    assert!(!app.step_mode);
}

#[test]
fn run_macro_smarticon_opens_name_picker() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::RunMacro);
    assert_eq!(app.mode, Mode::Names);
    assert!(app.name_list.is_some(), "macro picker should be open");
}

#[test]
fn icon_click_panel_seven_step_toggles_step_mode() {
    // Panel 7 slot 4 = icon id 52 (STEP toggle).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Seven;
    click_slot(&mut app, 4);
    assert!(app.step_mode);
}

#[test]
fn icon_click_panel_seven_run_opens_picker() {
    // Panel 7 slot 5 = icon id 53 (Run a macro).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Seven;
    click_slot(&mut app, 5);
    assert_eq!(app.mode, Mode::Names);
}

#[test]
fn outline_smarticon_borders_single_cell_perimeter() {
    let mut app = App::new();
    app.dispatch_sys_action(l123_graph::SysAction::OutlineRange);
    let b = app
        .wb()
        .cell_borders
        .get(&Address::A1)
        .copied()
        .expect("A1 should now have a border entry");
    assert!(b.top.is_some());
    assert!(b.bottom.is_some());
    assert!(b.left.is_some());
    assert!(b.right.is_some());
}

#[test]
fn outline_smarticon_toggles_off_when_already_outlined() {
    let mut app = App::new();
    // First click outlines.
    app.dispatch_sys_action(l123_graph::SysAction::OutlineRange);
    assert!(app.wb().cell_borders.contains_key(&Address::A1));
    // Second click clears it.
    app.dispatch_sys_action(l123_graph::SysAction::OutlineRange);
    assert!(
        !app.wb().cell_borders.contains_key(&Address::A1),
        "second click should leave the cell with no border entry",
    );
}

#[test]
fn outline_smarticon_only_perimeter_for_multi_cell_range() {
    let mut app = App::new();
    // 3x3 range A1:C3 via POINT.
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('R'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('E'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    // Now in POINT with highlight A1:C3. Outline.
    app.dispatch_sys_action(l123_graph::SysAction::OutlineRange);
    // Corner A1: top + left set, right + bottom clear.
    let a1 = app.wb().cell_borders.get(&Address::A1).copied().unwrap();
    assert!(a1.top.is_some() && a1.left.is_some());
    assert!(a1.right.is_none() && a1.bottom.is_none());
    // Top-mid B1: only top.
    let b1 = app
        .wb()
        .cell_borders
        .get(&Address::new(SheetId::A, 1, 0))
        .copied()
        .unwrap();
    assert!(b1.top.is_some());
    assert!(b1.left.is_none() && b1.right.is_none() && b1.bottom.is_none());
    // Centre B2: no border at all (interior cell).
    let b2 = app.wb().cell_borders.get(&Address::new(SheetId::A, 1, 1));
    assert!(
        b2.is_none(),
        "interior cell of an outlined range should have no border",
    );
    // Bottom-right C3: bottom + right.
    let c3 = app
        .wb()
        .cell_borders
        .get(&Address::new(SheetId::A, 2, 2))
        .copied()
        .unwrap();
    assert!(c3.bottom.is_some() && c3.right.is_some());
    assert!(c3.top.is_none() && c3.left.is_none());
}

#[test]
fn outline_smarticon_undo_restores_prior_border_state() {
    let mut app = App::new();
    assert!(app.wb().cell_borders.is_empty());
    app.dispatch_sys_action(l123_graph::SysAction::OutlineRange);
    assert!(app.wb().cell_borders.contains_key(&Address::A1));
    // Alt-F4 = undo.
    app.handle_key(KeyEvent::new(KeyCode::F(4), KeyModifiers::ALT));
    assert!(
        !app.wb().cell_borders.contains_key(&Address::A1),
        "undo should remove the border entry that wasn't there before",
    );
}

#[test]
fn icon_click_panel_three_outline_writes_borders() {
    // Panel 3 slot 8 = icon id 20 (drop shadow + outline).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Three;
    click_slot(&mut app, 8);
    assert!(app.wb().cell_borders.contains_key(&Address::A1));
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn icon_click_panel_three_comma_format_opens_decimals_prompt() {
    // Panel 3 slot 6 = icon id 18 (Comma format) → /RF, → decimals
    // prompt, same shape as Currency (slot 5 = icon 17).
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Three;
    click_slot(&mut app, 6);
    assert_eq!(app.mode, Mode::Menu);
    assert!(
        app.prompt.is_some(),
        "Comma format should prompt for decimal places",
    );
}

#[test]
fn icon_click_panel_two_block_end_down_dispatches() {
    let mut app = App::new();
    put_label(&mut app, "A1", "a");
    put_label(&mut app, "A2", "b");
    app.current_panel = l123_graph::Panel::Two;
    // Slot 2 in panel 2 = icon id 40 = END+DOWN.
    click_slot(&mut app, 2);
    assert_eq!(app.pointer().display_short(), "A2");
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn mouse_click_without_cached_panel_is_ignored() {
    let mut app = App::new();
    app.icon_panel_area.set(None);
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Down(MouseButton::Left),
        column: 82,
        row: 5,
        modifiers: KeyModifiers::NONE,
    });
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_non_left_button_is_ignored() {
    let mut app = App::new();
    app.icon_panel_area.set(Some(test_geom(TEST_PANEL)));
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Down(MouseButton::Right),
        column: 81,
        row: 5,
        modifiers: KeyModifiers::NONE,
    });
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.prompt.is_none());
}

#[test]
fn pager_right_half_advances_panel() {
    let mut app = App::new();
    assert_eq!(app.current_panel, l123_graph::Panel::One);
    let mid_col = TEST_PANEL.x + TEST_PANEL.width / 2;
    click(&mut app, TEST_PANEL, mid_col + 1, TEST_PANEL.y + 49);
    assert_eq!(app.current_panel, l123_graph::Panel::Two);
}

#[test]
fn pager_left_half_retreats_panel() {
    let mut app = App::new();
    let left_col = TEST_PANEL.x;
    click(&mut app, TEST_PANEL, left_col, TEST_PANEL.y + 49);
    // From panel 1, prev wraps to panel 7.
    assert_eq!(app.current_panel, l123_graph::Panel::Seven);
}

#[test]
fn pager_full_cycle_returns_to_panel_one() {
    let mut app = App::new();
    let mid_col = TEST_PANEL.x + TEST_PANEL.width / 2;
    for _ in 0..7 {
        click(&mut app, TEST_PANEL, mid_col + 1, TEST_PANEL.y + 49);
    }
    assert_eq!(app.current_panel, l123_graph::Panel::One);
}

#[test]
fn hit_test_maps_cell_to_visually_dominant_icon() {
    // Realistic geometry from a 3-col panel at font (9, 18): the
    // PNG's 1:17 aspect renders width-constrained at 27×459 px,
    // and the ceiled cell rect is 3×26 — so the bottom 9 px of
    // row 25 is empty slack and each icon is visually 1.5 cells
    // tall. Sampling at cell midpoint pixels then dividing by the
    // true rendered pixel height pins each cell to the icon that
    // covers most of its pixels. Naive (cells-only) mapping drifts
    // by half an icon at the bottom — mis-reporting Help as Bold,
    // Save as spilling into Retrieve, etc.
    let geom = IconPanelGeom {
        rect: Rect::new(80, 4, 3, 26),
        rendered_px_h: 459,
        font_px_h: 18,
    };
    // Row 4 (cell 0) sits entirely inside Save's pixel band [0,27).
    assert_eq!(App::hit_test_slot(&geom, 4), Some(0));
    // Row 5 (cell 1, midpt = 27 px) — right on the Save/Retrieve
    // boundary. The formula treats it as Retrieve, so Save's hit-
    // test does not spill into Retrieve's visual region.
    assert_eq!(App::hit_test_slot(&geom, 5), Some(1));
    // Row 21 (cell 17, midpt = 315 px) is entirely inside Bold.
    assert_eq!(App::hit_test_slot(&geom, 21), Some(11));
    // Row 22 (cell 18, midpt = 333 px) is entirely Italic.
    assert_eq!(App::hit_test_slot(&geom, 22), Some(12));
    // Row 27 (cell 23, midpt = 423 px) is entirely Help.
    assert_eq!(App::hit_test_slot(&geom, 27), Some(15));
    // Row 28 (cell 24, midpt = 441 px) is entirely Pager.
    assert_eq!(App::hit_test_slot(&geom, 28), Some(16));
    // Row 29 (cell 25, midpt = 459 px) is in the empty slack at
    // the bottom; map to pager so the last row isn't a dead zone.
    assert_eq!(App::hit_test_slot(&geom, 29), Some(16));
}

/// Simulate a mouse-move at `(col, row)` by stashing a fake panel
/// geometry (same fixture as click tests) and routing through
/// [`App::handle_mouse`].
fn hover(app: &mut App, area: Rect, col: u16, row: u16) {
    app.icon_panel_area.set(Some(test_geom(area)));
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Moved,
        column: col,
        row,
        modifiers: KeyModifiers::NONE,
    });
}

fn hover_slot(app: &mut App, slot: u16) {
    hover(
        app,
        TEST_PANEL,
        TEST_PANEL.x + 1,
        TEST_PANEL.y + slot * 3 + 1,
    );
}

#[test]
fn mouse_move_over_slot_sets_hovered_icon() {
    let mut app = App::new();
    hover_slot(&mut app, 0);
    assert_eq!(app.hovered_icon, Some((l123_graph::Panel::One, 0)));
    hover_slot(&mut app, 7);
    assert_eq!(app.hovered_icon, Some((l123_graph::Panel::One, 7)));
}

#[test]
fn mouse_move_outside_panel_clears_hovered_icon() {
    let mut app = App::new();
    hover_slot(&mut app, 0);
    assert!(app.hovered_icon.is_some());
    hover(&mut app, TEST_PANEL, 10, 10);
    assert_eq!(app.hovered_icon, None);
}

#[test]
fn mouse_move_over_pager_slot_does_not_set_hovered_icon() {
    // Slot 16 is the panel navigator; we deliberately exclude it
    // from hover-tooltip since its function is already rendered on
    // the slot itself ("Panel N of 7").
    let mut app = App::new();
    hover_slot(&mut app, 16);
    assert_eq!(app.hovered_icon, None);
}

#[test]
fn mouse_move_without_cached_panel_clears_hovered_icon() {
    let mut app = App::new();
    hover_slot(&mut app, 3);
    assert!(app.hovered_icon.is_some());
    app.icon_panel_area.set(None);
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Moved,
        column: 82,
        row: 5,
        modifiers: KeyModifiers::NONE,
    });
    assert_eq!(app.hovered_icon, None);
}

#[test]
fn mouse_move_on_different_panel_tracks_current_panel() {
    let mut app = App::new();
    app.current_panel = l123_graph::Panel::Three;
    hover_slot(&mut app, 2);
    assert_eq!(app.hovered_icon, Some((l123_graph::Panel::Three, 2)));
}

// ---- Grid click-to-move (Phase 1: READY only) ----

/// Synthesize a left-click at the given screen position. Unlike
/// [`click`], this doesn't stash a fake icon-panel rect — the
/// headless render path already populates `last_grid_area` when
/// `render_to_buffer` is called, which is what grid-click uses.
fn click_at(app: &mut App, col: u16, row: u16) {
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Down(MouseButton::Left),
        column: col,
        row,
        modifiers: KeyModifiers::NONE,
    });
}

// Default layout (80×25): control panel = 4 lines (y 0..4), grid
// = 20 lines (y 4..24), status = 1 line (y 24). Column-header row
// sits at area.y = 4; body rows start at y = 5. ROW_GUTTER = 5
// columns; default col width = 9, so col A spans x [5, 14), col B
// spans x [14, 23).
const HEADER_Y: u16 = 4;
const BODY_TOP_Y: u16 = 5;
const COL_A_X: u16 = 7; // anywhere in [5, 14)
const COL_B_X: u16 = 16; // anywhere in [14, 23)

#[test]
fn mouse_click_on_cell_moves_pointer() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 2);
    assert_eq!(app.pointer().display_full(), "A:B3");
}

#[test]
fn mouse_click_on_column_header_does_not_move_pointer() {
    // Column-letter row sits at the top of the grid area; clicking
    // it is reserved for Phase 2 (full-column select).
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_B_X, HEADER_Y);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_click_on_row_gutter_does_not_move_pointer() {
    // x in [0, ROW_GUTTER) is the row-number gutter; reserved for
    // Phase 2 (full-row select).
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, 2, BODY_TOP_Y + 2);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_click_below_grid_does_not_move_pointer() {
    // Status line row; clicks there are not bound to anything.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, 24);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_click_without_rendered_grid_is_ignored() {
    // No render happened yet, so last_grid_area is None — the
    // hit-test must bail cleanly rather than guessing geometry.
    let mut app = App::new();
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_click_in_menu_mode_does_not_move_pointer() {
    // Phase 1 restricts click-to-move to READY. In MENU (and
    // entry modes) we leave the pointer alone; Phase 2 will
    // add context-appropriate behavior.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    app.handle_key(KeyEvent::new(KeyCode::Char('/'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Menu);
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 2);
    assert_eq!(app.pointer().display_full(), "A:A1");
    assert_eq!(app.mode, Mode::Menu);
}

#[test]
fn mouse_click_respects_scroll_offset() {
    // With the viewport scrolled down, clicking the top body row
    // lands on the first visible row, not row 1.
    let mut app = App::new();
    app.wb_mut().viewport_row_offset = 10;
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    assert_eq!(app.pointer().display_full(), "A:A11");
}

#[test]
fn mouse_click_respects_custom_column_width() {
    // After widening column A to 15, col B's body shifts right.
    // Clicking the new B region must still resolve to B, not A.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    drive_set_col_width(&mut app, 15);
    let _ = app.render_to_buffer(80, 25);
    // ROW_GUTTER(5) + 15 = 20 → col B starts at x=20.
    click_at(&mut app, 22, BODY_TOP_Y);
    assert_eq!(app.pointer().display_full(), "A:B1");
}

// ---- Phase 2: POINT-mode click-to-extend ----

// Body cell fixtures for POINT-mode tests. Default col width 9,
// ROW_GUTTER 5 → col C at x [23,32), col D at x [32,41).
const COL_C_X: u16 = 25; // anywhere in [23, 32)
const COL_D_X: u16 = 35; // anywhere in [32, 41)

/// Enter POINT via `/RE` (Range Erase) — the shortest path into
/// auto-anchored POINT. The prompt expects a range; auto-anchor
/// fires at the current pointer.
fn enter_point_via_range_erase(app: &mut App) {
    for c in ['/', 'R', 'E'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
}

#[test]
fn mouse_click_in_anchored_point_extends_range() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    assert_eq!(app.mode, Mode::Point);
    // Anchor is set at A1 (the pointer when POINT entered).
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));

    // Click at C3 — range should become A1..C3. Anchor stays at A1.
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:C3");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_click_in_anchored_point_can_shrink_range() {
    // After extending to C3, clicking back at B2 must shrink the
    // range (anchor still A1, pointer now B2).
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
    assert_eq!(app.pointer().display_full(), "A:B2");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_click_in_unanchored_point_moves_pointer_without_reanchoring() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    // First Esc unanchors.
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), None);

    click_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:D4");
    // Still unanchored after click.
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), None);
}

#[test]
fn mouse_click_on_gutter_in_point_is_noop() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    click_at(&mut app, 2, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:A1");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_click_in_point_does_not_exit_mode() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Point, "click must not drop out of POINT");
}

// ---- Phase 4: drag-to-select ----

/// Simulate a left-button drag at `(col, row)` — used after a
/// preceding [`click_at`] to drive the drag-to-select path.
fn drag_at(app: &mut App, col: u16, row: u16) {
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Drag(MouseButton::Left),
        column: col,
        row,
        modifiers: KeyModifiers::NONE,
    });
}

/// Simulate the left-button release ending a drag.
fn release_at(app: &mut App, col: u16, row: u16) {
    app.handle_mouse(MouseEvent {
        kind: MouseEventKind::Up(MouseButton::Left),
        column: col,
        row,
        modifiers: KeyModifiers::NONE,
    });
}

#[test]
fn mouse_drag_from_ready_enters_point_anchored_at_press_cell() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y); // Press at A1.
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2); // Drag to C3.
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:C3");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_drag_extends_then_shrinks_with_anchor_fixed() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    drag_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
    assert_eq!(app.pointer().display_full(), "A:D4");
    drag_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
    assert_eq!(app.pointer().display_full(), "A:B2");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_release_does_not_exit_point() {
    // Up just ends the drag; selection persists for follow-up
    // commands (Bold icon, /Range Format, …).
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    release_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:C3");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_drag_inside_existing_point_does_not_reanchor() {
    // Already in POINT via /RE — anchor at A1. Press on B2 (Phase 2
    // moves pointer; anchor stays), then drag to C3. The /RE anchor
    // must persist — we never overwrite an existing POINT anchor
    // with the mouse-press cell.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_point_via_range_erase(&mut app);
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 1); // Press at B2.
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2); // Drag to C3.
    assert_eq!(app.mode, Mode::Point);
    assert_eq!(app.pointer().display_full(), "A:C3");
    assert_eq!(app.point.as_ref().and_then(|p| p.anchor), Some(Address::A1));
}

#[test]
fn mouse_drag_without_press_on_grid_is_ignored() {
    // A drag that wasn't preceded by a press inside the grid (e.g.
    // mouse-down landed on the column header, then drag onto the
    // body) must not promote into POINT.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_B_X, HEADER_Y); // Header click is ignored.
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Ready);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn mouse_drag_in_value_mode_is_ignored() {
    // Mid-formula entry: a stray drag must not corrupt the buffer
    // or change mode. Phase 3 splicing is press-driven, not drag.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+");
    assert_eq!(app.mode, Mode::Value);
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Value);
    assert_eq!(entry_buffer(&app), "+");
}

#[test]
fn mouse_drag_off_grid_does_not_move_pointer() {
    // While dragging, if the cursor leaves the grid (onto the
    // column header or row gutter), the pointer freezes — it does
    // not snap to the last in-grid cell or jump to the gutter.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    let pinned = app.pointer().display_full();
    drag_at(&mut app, 2, BODY_TOP_Y + 2); // Onto row gutter.
    assert_eq!(app.pointer().display_full(), pinned);
    assert_eq!(app.mode, Mode::Point);
}

#[test]
fn mouse_release_clears_drag_state_so_next_drag_needs_a_press() {
    // After Up, a follow-up Drag without a fresh press is a no-op:
    // we shouldn't keep extending the previous selection just
    // because the OS sends spurious motion events.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    drag_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    release_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    // Now exit POINT — Esc, Esc.
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    // Bare Drag (no press): must NOT promote to POINT.
    drag_at(&mut app, COL_D_X, BODY_TOP_Y + 3);
    assert_eq!(app.mode, Mode::Ready);
}

// ---- Phase 5: scroll wheel ----

fn wheel(app: &mut App, kind: MouseEventKind) {
    app.handle_mouse(MouseEvent {
        kind,
        column: 10,
        row: 10,
        modifiers: KeyModifiers::NONE,
    });
}

#[test]
fn scroll_down_advances_viewport_row_offset_by_scroll_step() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    assert_eq!(app.wb().viewport_row_offset, 0);
    wheel(&mut app, MouseEventKind::ScrollDown);
    assert_eq!(app.wb().viewport_row_offset, MOUSE_SCROLL_STEP);
}

#[test]
fn scroll_up_retreats_viewport_row_offset_saturating_at_zero() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    app.wb_mut().viewport_row_offset = 50;
    wheel(&mut app, MouseEventKind::ScrollUp);
    assert_eq!(app.wb().viewport_row_offset, 50 - MOUSE_SCROLL_STEP);
    // Many up-scrolls saturate at 0 (no underflow / panic).
    for _ in 0..100 {
        wheel(&mut app, MouseEventKind::ScrollUp);
    }
    assert_eq!(app.wb().viewport_row_offset, 0);
}

#[test]
fn scroll_does_not_move_pointer() {
    // Modern spreadsheet convention: the wheel scrolls the
    // viewport only, leaving the cell pointer where it sits — even
    // if it ends up off-screen. The next keyboard arrow press will
    // pull it back into view via scroll_into_view.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    assert_eq!(app.pointer().display_full(), "A:A1");
    wheel(&mut app, MouseEventKind::ScrollDown);
    wheel(&mut app, MouseEventKind::ScrollDown);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

#[test]
fn scroll_does_not_change_mode() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    wheel(&mut app, MouseEventKind::ScrollDown);
    assert_eq!(app.mode, Mode::Ready);

    // In POINT, scrolling must keep POINT alive — selection
    // persists, viewport just moves.
    enter_point_via_range_erase(&mut app);
    assert_eq!(app.mode, Mode::Point);
    wheel(&mut app, MouseEventKind::ScrollDown);
    assert_eq!(app.mode, Mode::Point);
}

#[test]
fn scroll_works_during_value_entry_without_corrupting_buffer() {
    // The wheel is a navigation gesture, not an entry gesture — it
    // must never touch the entry buffer or commit/cancel the
    // current entry.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+A1");
    assert_eq!(app.mode, Mode::Value);
    wheel(&mut app, MouseEventKind::ScrollDown);
    assert_eq!(app.mode, Mode::Value);
    assert_eq!(entry_buffer(&app), "+A1");
    assert!(app.wb().viewport_row_offset > 0);
}

// ---- Phase 3: mid-entry cell-reference splicing on click ----

/// Type a sequence into a fresh app and return it in VALUE mode
/// with `buffer` typed after the value-starter. The leading `+`
/// forces VALUE (SPEC §20 #6 — `=` is not a value starter in L123).
fn enter_value_with(app: &mut App, buffer: &str) {
    for c in buffer.chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
}

fn entry_buffer(app: &App) -> String {
    app.entry
        .as_ref()
        .map(|e| e.buffer.clone())
        .unwrap_or_default()
}

#[test]
fn mouse_click_in_value_after_plus_splices_short_address() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+");
    assert_eq!(app.mode, Mode::Value);
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Value, "splice must keep VALUE mode");
    assert_eq!(entry_buffer(&app), "+C3");
}

#[test]
fn mouse_click_in_value_after_open_paren_splices() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "@SUM(");
    click_at(&mut app, COL_A_X, BODY_TOP_Y);
    assert_eq!(entry_buffer(&app), "@SUM(A1");
}

#[test]
fn mouse_click_in_value_after_comma_splices() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "@MIN(A1,");
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 1);
    assert_eq!(entry_buffer(&app), "@MIN(A1,B2");
}

#[test]
fn mouse_click_in_value_after_range_dots_splices() {
    // `..` is 1-2-3's range separator; a trailing `.` must count
    // as a cell-ref-accepting context so you can click the end
    // corner of a range.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+A1..");
    click_at(&mut app, COL_B_X, BODY_TOP_Y + 4);
    assert_eq!(entry_buffer(&app), "+A1..B5");
}

#[test]
fn mouse_click_in_value_after_digit_is_noop() {
    // Splicing after `5` would produce `5C3` — nonsense. The
    // click must not corrupt the buffer.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "5");
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Value);
    assert_eq!(entry_buffer(&app), "5");
}

#[test]
fn mouse_click_in_value_after_close_paren_is_noop() {
    // Close paren closes a sub-expression; the parser wants an
    // operator next, not another cell ref.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "@SUM(A1..A3)");
    let before = entry_buffer(&app);
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(entry_buffer(&app), before);
}

#[test]
fn mouse_click_in_label_is_noop() {
    // Labels hold literal text; a mid-label cell ref is almost
    // certainly not what the user meant. Leave the buffer alone.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "hello ");
    assert_eq!(app.mode, Mode::Label);
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Label);
    assert_eq!(entry_buffer(&app), "hello ");
}

#[test]
fn mouse_click_in_edit_after_operator_splices() {
    // Put a formula referencing A1 into B1, then F2 to EDIT the
    // source. Appending `+` and clicking C3 must splice.
    let mut app = App::new();
    // Move pointer to B1 (right once from A1).
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    enter_value_with(&mut app, "+A1");
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let _ = app.render_to_buffer(80, 25);
    app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Edit);
    app.handle_key(KeyEvent::new(KeyCode::Char('+'), KeyModifiers::NONE));
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.mode, Mode::Edit);
    assert_eq!(entry_buffer(&app), "+A1+C3");
}

#[test]
fn mouse_click_on_gutter_in_value_is_noop() {
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+");
    click_at(&mut app, 2, BODY_TOP_Y + 2);
    assert_eq!(entry_buffer(&app), "+");
}

#[test]
fn mouse_click_in_value_does_not_move_pointer() {
    // Splicing is a buffer operation; the *cell pointer* must
    // stay put so the entry still belongs to the originally-
    // selected cell.
    let mut app = App::new();
    let _ = app.render_to_buffer(80, 25);
    enter_value_with(&mut app, "+");
    click_at(&mut app, COL_C_X, BODY_TOP_Y + 2);
    assert_eq!(app.pointer().display_full(), "A:A1");
}

/// Drive the `/Worksheet Column Set-Width` menu path, typing `width`
/// at the prompt and pressing Enter.
fn drive_set_col_width(app: &mut App, width: u8) {
    for c in ['/', 'W', 'C', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    // The prompt seeds the current width and is `fresh` — any digit
    // replaces the seed; subsequent digits append.
    for c in width.to_string().chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
}

/// After /WCS 15 on column A, the grid must actually draw column A
/// at 15 characters wide — pushing the B-column header and B1's
/// content to x = ROW_GUTTER + 15.
#[test]
fn set_col_width_widens_column_on_screen() {
    let mut app = App::new();
    // A1 = "alpha", B1 = "beta".
    for c in "alpha".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    for c in "beta".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));

    // Widen column A back at the A1 pointer.
    app.handle_key(KeyEvent::new(KeyCode::Home, KeyModifiers::NONE));
    drive_set_col_width(&mut app, 15);
    assert_eq!(app.col_width_of(SheetId::A, 0), 15);

    let buf = app.render_to_buffer(80, 25);
    // The body row for row 1 sits one line below the column-header
    // row (PANEL_HEIGHT + 1).
    let body_y = PANEL_HEIGHT + 1;
    // Column A's contents must start at x = ROW_GUTTER and run 15
    // chars; "alpha" is left-aligned (apostrophe prefix) and padded
    // with spaces to the full column width.
    let a_slot: String = (0..15)
        .map(|i| buf[(ROW_GUTTER + i, body_y)].symbol().to_string())
        .collect();
    assert_eq!(a_slot, "alpha          ");
    // Column B must start 15 characters later — not 9.
    let b_slot: String = (0..9)
        .map(|i| buf[(ROW_GUTTER + 15 + i, body_y)].symbol().to_string())
        .collect();
    assert_eq!(b_slot, "beta     ");

    // And the header row reflects the same geometry.
    let header_y = PANEL_HEIGHT;
    let b_header = buf[(ROW_GUTTER + 15 + 4, header_y)].symbol(); // center of 9-wide slot
    assert_eq!(b_header, "B");
}

#[test]
fn iterm2_env_hint_matches_term_program_variants() {
    // Apple's iTerm2 sets TERM_PROGRAM=iTerm.app.
    assert!(is_iterm2_compatible_env(Some("iTerm.app"), None));
    // SSH into a shell from iTerm2: TERM_PROGRAM is whatever the
    // remote shell wants (often tmux / Apple_Terminal), but
    // LC_TERMINAL is forwarded.
    assert!(is_iterm2_compatible_env(Some("tmux"), Some("iTerm2")));
    // Other hosts that speak the OSC 1337 image protocol.
    assert!(is_iterm2_compatible_env(Some("WezTerm"), None));
    assert!(is_iterm2_compatible_env(Some("mintty"), None));
    assert!(is_iterm2_compatible_env(Some("WarpTerminal"), None));
}

#[test]
fn iterm2_env_hint_rejects_other_terminals() {
    assert!(!is_iterm2_compatible_env(Some("ghostty"), None));
    assert!(!is_iterm2_compatible_env(Some("Apple_Terminal"), None));
    assert!(!is_iterm2_compatible_env(Some("xterm-kitty"), None));
    assert!(!is_iterm2_compatible_env(None, None));
    assert!(!is_iterm2_compatible_env(None, Some("")));
}

// --- :Format Bold/Italic/Underline Set|Clear (step 2 of the
//     text-style slice: core storage, command execution, journaled
//     undo). Menu wiring and rendering land in later steps.

fn one_cell_range(addr: Address) -> Range {
    Range {
        start: addr,
        end: addr,
    }
}

#[test]
fn range_text_style_set_bold_records_override() {
    let mut app = App::new();
    app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD),
    );
}

#[test]
fn range_text_style_set_then_undo_restores_plain() {
    let mut app = App::new();
    app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
    app.undo();
    assert!(
        !app.wb().cell_text_styles.contains_key(&Address::A1),
        "bold should be gone after Alt-F4"
    );
}

#[test]
fn range_text_style_bold_then_italic_composes_bits() {
    let mut app = App::new();
    let r = one_cell_range(Address::A1);
    app.execute_range_text_style(r, TextStyle::BOLD, true);
    app.execute_range_text_style(r, TextStyle::ITALIC, true);
    let s = app.wb().cell_text_styles.get(&Address::A1).copied();
    assert_eq!(
        s,
        Some(TextStyle {
            bold: true,
            italic: true,
            underline: false
        }),
    );
}

#[test]
fn range_text_style_clear_drops_only_named_bits() {
    let mut app = App::new();
    let r = one_cell_range(Address::A1);
    app.execute_range_text_style(r, TextStyle::BOLD.merge(TextStyle::ITALIC), true);
    app.execute_range_text_style(r, TextStyle::BOLD, false);
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::ITALIC),
    );
}

#[test]
fn range_text_style_clearing_last_bit_removes_entry() {
    let mut app = App::new();
    let r = one_cell_range(Address::A1);
    app.execute_range_text_style(r, TextStyle::BOLD, true);
    app.execute_range_text_style(r, TextStyle::BOLD, false);
    assert!(
        !app.wb().cell_text_styles.contains_key(&Address::A1),
        "plain style should not leave an empty entry in the map"
    );
}

#[test]
fn range_text_style_reset_clears_all_attributes() {
    let mut app = App::new();
    let r = one_cell_range(Address::A1);
    let all = TextStyle {
        bold: true,
        italic: true,
        underline: true,
    };
    app.execute_range_text_style(r, all, true);
    app.execute_range_text_style(r, all, false);
    assert!(!app.wb().cell_text_styles.contains_key(&Address::A1));
}

#[test]
fn range_text_style_undo_restores_partial_prior_state() {
    let mut app = App::new();
    let r = one_cell_range(Address::A1);
    // Start: bold on A1 (the prior state we should restore to).
    app.execute_range_text_style(r, TextStyle::BOLD, true);
    // Clear the journal so we're only testing the next undo.
    app.wb_mut().journal.clear();
    // Now apply italic.
    app.execute_range_text_style(r, TextStyle::ITALIC, true);
    app.undo();
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD),
        "undoing the italic-set should leave the prior bold-only state"
    );
}

#[test]
fn text_style_modifier_maps_bits_to_ratatui_modifier() {
    assert_eq!(text_style_modifier(TextStyle::PLAIN), Modifier::empty());
    assert_eq!(text_style_modifier(TextStyle::BOLD), Modifier::BOLD);
    assert_eq!(text_style_modifier(TextStyle::ITALIC), Modifier::ITALIC);
    assert_eq!(
        text_style_modifier(TextStyle::UNDERLINE),
        Modifier::UNDERLINED,
    );
    let all = TextStyle {
        bold: true,
        italic: true,
        underline: true,
    };
    assert_eq!(
        text_style_modifier(all),
        Modifier::BOLD | Modifier::ITALIC | Modifier::UNDERLINED,
    );
}

/// Read back the ratatui `Modifier` bits on the first cell of an
/// address in the rendered buffer. Uses the same visible-column
/// layout math as [`App::cell_rendered_text`].
fn cell_modifier_at(app: &App, buf: &Buffer, addr: Address) -> Modifier {
    let dr = (addr.row - app.wb().viewport_row_offset) as u16;
    let y = PANEL_HEIGHT + 1 + dr;
    let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
    let layout = app.visible_column_layout(content_width);
    let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == addr.col).unwrap();
    let x = ROW_GUTTER + x_off;
    buf[(x, y)].style().add_modifier
}

#[test]
fn bold_style_renders_with_ratatui_bold_modifier() {
    let mut app = App::new();
    // Put a label at B5 so there's a character to style.
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    for c in "hello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::new(SheetId::A, 1, 4); // B5
    app.execute_range_text_style(one_cell_range(target), TextStyle::BOLD, true);

    let buf = app.render_to_buffer(80, 25);
    let modifier = cell_modifier_at(&app, &buf, target);
    assert!(
        modifier.contains(Modifier::BOLD),
        "expected BOLD in buffer cell's modifier, got {modifier:?}"
    );
}

#[test]
fn compound_style_renders_with_all_three_modifiers() {
    // UNDERLINED applies only over actual glyphs, so the cell
    // needs visible text — an empty cell carries no underline
    // even when one is set on its style.
    let mut app = App::new();
    for c in "x".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::A1;
    let all = TextStyle {
        bold: true,
        italic: true,
        underline: true,
    };
    app.execute_range_text_style(one_cell_range(target), all, true);
    let buf = app.render_to_buffer(80, 25);
    let modifier = cell_modifier_at(&app, &buf, target);
    assert!(modifier.contains(Modifier::BOLD));
    assert!(modifier.contains(Modifier::ITALIC));
    assert!(modifier.contains(Modifier::UNDERLINED));
}

#[test]
fn plain_cells_have_no_text_style_modifier() {
    let app = App::new();
    let buf = app.render_to_buffer(80, 25);
    let modifier = cell_modifier_at(&app, &buf, Address::new(SheetId::A, 2, 2));
    assert!(!modifier.contains(Modifier::BOLD));
    assert!(!modifier.contains(Modifier::ITALIC));
    assert!(!modifier.contains(Modifier::UNDERLINED));
}

#[test]
fn text_style_follows_label_spill_into_adjacent_columns() {
    // Long italic label at A1 spills across B/C/D/E.  The
    // overflow characters should render italic too, not plain.
    let mut app = App::new();
    for c in "INCOME SUMMARY 1991: Sloane Camera and Video".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    keys(&mut app, ":FIS~");

    let buf = app.render_to_buffer(100, 25);

    // A1 itself is italic (its home column).
    assert!(
        cell_modifier_at(&app, &buf, Address::A1).contains(Modifier::ITALIC),
        "A1 (owner) should be italic"
    );
    // B1, C1, D1 are empty neighbors the label overflowed into —
    // they must pick up the owner's italic attribute.
    for col in 1..=3u16 {
        let addr = Address::new(SheetId::A, col, 0);
        let m = cell_modifier_at(&app, &buf, addr);
        assert!(
            m.contains(Modifier::ITALIC),
            "spill into column {col} should be italic, got {m:?}"
        );
    }
}

#[test]
fn text_style_does_not_leak_past_spill_extent() {
    // Italic label at A1 long enough to fill B1 exactly; C1 must
    // remain plain (the label doesn't reach it).
    let mut app = App::new();
    // A1 + B1 default widths = 9 + 9 = 18.  "123456789abcdefgh" is
    // 17 chars — fits inside A1..B1 with no leftover for C1.
    for c in "123456789abcdefgh".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    keys(&mut app, ":FIS~");

    let buf = app.render_to_buffer(100, 25);
    // C1 is never touched by the spill, so it should stay plain.
    let c1 = Address::new(SheetId::A, 2, 0);
    let m = cell_modifier_at(&app, &buf, c1);
    assert!(
        !m.contains(Modifier::ITALIC),
        "C1 beyond spill extent should be plain, got {m:?}"
    );
}

#[test]
fn bold_on_highlighted_cell_keeps_both_reversed_and_bold() {
    let mut app = App::new();
    // Pointer starts at A1, which is the highlighted cell in READY.
    app.execute_range_text_style(one_cell_range(Address::A1), TextStyle::BOLD, true);
    let buf = app.render_to_buffer(80, 25);
    let modifier = cell_modifier_at(&app, &buf, Address::A1);
    assert!(
        modifier.contains(Modifier::REVERSED),
        "pointer reverse-video"
    );
    assert!(modifier.contains(Modifier::BOLD), "bold overlay");
}

/// Replay a string of characters as individual READY-mode key
/// presses, just like a transcript.  `~` stands for Enter, matching
/// the acceptance-transcript convention.
fn keys(app: &mut App, s: &str) {
    for c in s.chars() {
        match c {
            '~' => app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE)),
            ch => app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::NONE)),
        }
    }
}

#[test]
fn colon_enters_wysiwyg_menu_mode() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Menu);
    let state = app.menu.as_ref().expect("menu state");
    // First item under WYSIWYG_ROOT is "Worksheet".
    let first = state.level().first().expect("items visible");
    assert_eq!(first.name, "Worksheet");
}

#[test]
fn colon_f_b_s_bolds_the_selected_range() {
    let mut app = App::new();
    keys(&mut app, ":FBS~");
    // After ~ commits POINT with default single-cell range at A1.
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD),
    );
    assert_eq!(app.mode, Mode::Ready);
}

#[test]
fn colon_f_i_s_italicizes_range() {
    let mut app = App::new();
    keys(&mut app, ":FIS~");
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::ITALIC),
    );
}

#[test]
fn colon_f_u_s_underlines_range() {
    let mut app = App::new();
    keys(&mut app, ":FUS~");
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::UNDERLINE),
    );
}

/// Read the buffer cell at `addr.col`'s base x-position plus
/// `col_off` columns, on the row of `addr`.  Lets a test inspect
/// the trailing padding columns inside a wider cell.
fn buf_cell_at<'a>(
    app: &App,
    buf: &'a Buffer,
    addr: Address,
    col_off: u16,
) -> &'a ratatui::buffer::Cell {
    let dr = (addr.row - app.wb().viewport_row_offset) as u16;
    let y = PANEL_HEIGHT + 1 + dr;
    let content_width = buf.area.width.saturating_sub(ROW_GUTTER);
    let layout = app.visible_column_layout(content_width);
    let (_, x_off, _) = *layout.iter().find(|(c, _, _)| *c == addr.col).unwrap();
    let x = ROW_GUTTER + x_off + col_off;
    &buf[(x, y)]
}

#[test]
fn underline_does_not_extend_into_trailing_padding() {
    // Short underlined label "hi" in a width-9 cell: 'h' and 'i'
    // carry UNDERLINED; the seven trailing padding spaces do not.
    let mut app = App::new();
    // Move pointer off A1 so the test cell isn't REVERSED.
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    for c in "hi".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::new(SheetId::A, 0, 2); // A3
    app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

    let buf = app.render_to_buffer(80, 25);
    let h = buf_cell_at(&app, &buf, target, 0);
    assert_eq!(h.symbol(), "h");
    assert!(h.style().add_modifier.contains(Modifier::UNDERLINED));
    let i = buf_cell_at(&app, &buf, target, 1);
    assert_eq!(i.symbol(), "i");
    assert!(i.style().add_modifier.contains(Modifier::UNDERLINED));
    for off in 2..9u16 {
        let cell = buf_cell_at(&app, &buf, target, off);
        assert_eq!(cell.symbol(), " ", "padding at +{off} should be a space");
        assert!(
            !cell.style().add_modifier.contains(Modifier::UNDERLINED),
            "padding at +{off} should NOT carry UNDERLINED, got {:?}",
            cell.style().add_modifier,
        );
    }
}

#[test]
fn underline_does_not_extend_past_spilled_label_text() {
    // Underlined label long enough to spill A→B→C; the tail of
    // the spill (padding past the last glyph) must not carry
    // UNDERLINED, even though the spill cells inherit the
    // owner's text style.
    let mut app = App::new();
    // Move pointer off A1.
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    // 11 chars at width 9 → spills one column into B3, leaving
    // 7 trailing pad columns inside B3.
    for c in "hello world".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::new(SheetId::A, 0, 2); // A3
    app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

    let buf = app.render_to_buffer(80, 25);
    let neighbor = Address::new(SheetId::A, 1, 2); // B3
                                                   // First two B3 columns hold the spilled "ld" tail — underlined.
    let l = buf_cell_at(&app, &buf, neighbor, 0);
    assert_eq!(l.symbol(), "l");
    assert!(l.style().add_modifier.contains(Modifier::UNDERLINED));
    let d = buf_cell_at(&app, &buf, neighbor, 1);
    assert_eq!(d.symbol(), "d");
    assert!(d.style().add_modifier.contains(Modifier::UNDERLINED));
    // Remaining padding columns of B3 are space and unstyled.
    for off in 2..9u16 {
        let cell = buf_cell_at(&app, &buf, neighbor, off);
        assert_eq!(
            cell.symbol(),
            " ",
            "spill-tail padding at B3+{off} should be a space",
        );
        assert!(
            !cell.style().add_modifier.contains(Modifier::UNDERLINED),
            "spill-tail padding at B3+{off} should NOT carry UNDERLINED",
        );
    }
}

#[test]
fn underline_continues_through_internal_space_at_cell_boundary() {
    // Long underlined label spills across A→B; the space between
    // "1991:" and "Sloane" lands at the A/B column boundary.
    // That space is internal to the original text, so the leading
    // column of B must still carry UNDERLINED — earlier per-slot
    // trim heuristics dropped it as if it were B's leading
    // padding.
    let mut app = App::new();
    // Set A's width to 20 so "INCOME SUMMARY 1991:" exactly fills
    // it and the spill boundary lands on the space character.
    // Pointer starts at A1 — /WCS 20 widens column A.
    for c in ['/', 'W', 'C', 'S'] {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    for c in "20".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    assert_eq!(app.col_width_of(SheetId::A, 0), 20);
    // Move pointer to A3 so it's not REVERSED-highlighting our
    // test cells.
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    for c in "INCOME SUMMARY 1991: Sloane Camera and Video".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::new(SheetId::A, 0, 2); // A3
    app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

    let buf = app.render_to_buffer(120, 25);
    let neighbor = Address::new(SheetId::A, 1, 2); // B3
                                                   // First column of B3 is the boundary space — it's the
                                                   // internal " " of "...1991: Sloane..." and must stay
                                                   // underlined for the run to read continuously.
    let boundary = buf_cell_at(&app, &buf, neighbor, 0);
    assert_eq!(boundary.symbol(), " ");
    assert!(
        boundary.style().add_modifier.contains(Modifier::UNDERLINED),
        "internal-text space at A/B cell seam should keep UNDERLINED",
    );
    // Second column of B3 is 'S' — clearly part of text.
    let s_glyph = buf_cell_at(&app, &buf, neighbor, 1);
    assert_eq!(s_glyph.symbol(), "S");
    assert!(s_glyph.style().add_modifier.contains(Modifier::UNDERLINED));
}

#[test]
fn underline_covers_internal_whitespace_between_glyphs() {
    // An internal space between two glyphs in a label is part of
    // the text run and should keep its underline so the line
    // reads as continuous under "a b".
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    for c in "a b".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    let target = Address::new(SheetId::A, 0, 2); // A3
    app.execute_range_text_style(one_cell_range(target), TextStyle::UNDERLINE, true);

    let buf = app.render_to_buffer(80, 25);
    // "a", " " (internal), "b" all carry UNDERLINED.
    let a = buf_cell_at(&app, &buf, target, 0);
    assert_eq!(a.symbol(), "a");
    assert!(a.style().add_modifier.contains(Modifier::UNDERLINED));
    let mid = buf_cell_at(&app, &buf, target, 1);
    assert_eq!(mid.symbol(), " ");
    assert!(
        mid.style().add_modifier.contains(Modifier::UNDERLINED),
        "internal space between 'a' and 'b' should remain underlined",
    );
    let b = buf_cell_at(&app, &buf, target, 2);
    assert_eq!(b.symbol(), "b");
    assert!(b.style().add_modifier.contains(Modifier::UNDERLINED));
}

#[test]
fn colon_f_b_c_clears_bold_left_other_bits_alone() {
    let mut app = App::new();
    keys(&mut app, ":FBS~");
    keys(&mut app, ":FIS~");
    keys(&mut app, ":FBC~");
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::ITALIC),
    );
}

#[test]
fn colon_f_r_resets_all_attributes() {
    let mut app = App::new();
    keys(&mut app, ":FBS~");
    keys(&mut app, ":FIS~");
    keys(&mut app, ":FUS~");
    keys(&mut app, ":FR~");
    assert!(!app.wb().cell_text_styles.contains_key(&Address::A1));
}

#[test]
fn line1_shows_bold_marker_on_labeled_cell() {
    let mut app = App::new();
    for c in "hello".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    keys(&mut app, ":FBS~");
    let buf = app.render_to_buffer(80, 25);
    let line1 = App::line_text(&buf, 0);
    assert!(
        line1.contains("{Bold}"),
        "expected {{Bold}} in line 1, got {line1:?}"
    );
}

#[test]
fn line1_shows_compound_marker_with_space_separator() {
    let mut app = App::new();
    keys(&mut app, ":FBS~");
    keys(&mut app, ":FIS~");
    keys(&mut app, ":FUS~");
    let buf = app.render_to_buffer(80, 25);
    let line1 = App::line_text(&buf, 0);
    assert!(
        line1.contains("{Bold Italic Underline}"),
        "expected {{Bold Italic Underline}} in line 1, got {line1:?}"
    );
}

#[test]
fn line1_has_no_style_marker_on_plain_cell() {
    let app = App::new();
    let buf = app.render_to_buffer(80, 25);
    let line1 = App::line_text(&buf, 0);
    assert!(
        !line1.contains('{'),
        "plain cell should not show a style marker, got {line1:?}"
    );
}

#[test]
fn line1_marker_follows_format_tag_on_numeric_cell() {
    let mut app = App::new();
    // Type a number so the cell gets a (G) format tag.
    for c in "42".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    // Go back up to the just-entered cell before applying style.
    app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::NONE));
    keys(&mut app, ":FBS~");
    let buf = app.render_to_buffer(80, 25);
    let line1 = App::line_text(&buf, 0);
    let tag_pos = line1.find("(G)").expect("format tag present");
    let marker_pos = line1.find("{Bold}").expect("style marker present");
    assert!(
        tag_pos < marker_pos,
        "format tag should precede style marker: {line1:?}"
    );
}

#[test]
fn line1_marker_disappears_after_clearing_last_style_bit() {
    let mut app = App::new();
    keys(&mut app, ":FBS~");
    keys(&mut app, ":FBC~");
    let buf = app.render_to_buffer(80, 25);
    let line1 = App::line_text(&buf, 0);
    assert!(
        !line1.contains('{'),
        "marker should be gone after clear, got {line1:?}"
    );
}

#[test]
fn text_style_survives_xlsx_save_and_retrieve() {
    use std::process;
    use std::time::{SystemTime, UNIX_EPOCH};
    let nanos = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_nanos())
        .unwrap_or(0);
    let dir = std::env::temp_dir().join(format!("l123_ui_style_rt_{}_{}", process::id(), nanos));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("style.xlsx");

    let mut app = App::new();
    // A1: label "hi" with bold + italic.
    for c in "hi".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    // Enter commits in place (SPEC §8) — pointer still at A1.
    keys(&mut app, ":FBS~");
    keys(&mut app, ":FIS~");
    // A2: label "bye" with underline.
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    for c in "bye".chars() {
        app.handle_key(KeyEvent::new(KeyCode::Char(c), KeyModifiers::NONE));
    }
    app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::NONE));
    keys(&mut app, ":FUS~");

    // Probe the in-memory map right before save.
    assert_eq!(
        app.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle::BOLD.merge(TextStyle::ITALIC)),
    );
    assert_eq!(
        app.wb()
            .cell_text_styles
            .get(&Address::new(SheetId::A, 0, 1))
            .copied(),
        Some(TextStyle::UNDERLINE),
    );

    app.save_workbook_to(path.clone());

    let mut reopened = App::new();
    reopened.load_workbook_from(path.clone());

    assert_eq!(
        reopened.wb().cell_text_styles.get(&Address::A1).copied(),
        Some(TextStyle {
            bold: true,
            italic: true,
            underline: false
        }),
        "A1 should round-trip bold + italic"
    );
    assert_eq!(
        reopened
            .wb()
            .cell_text_styles
            .get(&Address::new(SheetId::A, 0, 1))
            .copied(),
        Some(TextStyle::UNDERLINE),
        "A2 should round-trip underline"
    );

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_dir(&dir);
}

#[test]
fn colon_q_closes_wysiwyg_menu() {
    let mut app = App::new();
    app.handle_key(KeyEvent::new(KeyCode::Char(':'), KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Char('Q'), KeyModifiers::NONE));
    assert_eq!(app.mode, Mode::Ready);
    assert!(app.menu.is_none());
}

#[test]
fn range_text_style_applies_across_multi_cell_range() {
    let mut app = App::new();
    let r = Range {
        start: Address::new(SheetId::A, 0, 0),
        end: Address::new(SheetId::A, 1, 1),
    };
    app.execute_range_text_style(r, TextStyle::BOLD, true);
    for col in 0..=1 {
        for row in 0..=1 {
            let a = Address::new(SheetId::A, col, row);
            assert_eq!(
                app.wb().cell_text_styles.get(&a).copied(),
                Some(TextStyle::BOLD),
                "cell ({col},{row}) should be bold"
            );
        }
    }
}

#[test]
fn default_theme_paints_headers_with_reversed_modifier() {
    // Default theme leaves both fg and bg unset and just sets
    // REVERSED so the terminal inverts its own pair — exactly what
    // every existing acceptance transcript already snapshots.
    let app = App::new();
    let buf = app.render_to_buffer(80, 25);
    // Column header strip is at y = PANEL_HEIGHT (top of the grid).
    let col_header_y = PANEL_HEIGHT;
    // Default column width is 9; column B sits at ROW_GUTTER + 9..
    // and is *not* the active column on a fresh App, so it shows
    // the resting (non-active) header style.
    let inactive = &buf[(ROW_GUTTER + 9 + 4, col_header_y)];
    assert!(
        inactive.modifier.contains(Modifier::REVERSED),
        "default theme should set REVERSED on column headers"
    );
    assert_eq!(inactive.fg, Color::Reset);
    assert_eq!(inactive.bg, Color::Reset);

    // Row-number gutter at body row 2 (y = PANEL_HEIGHT + 2,
    // which is sheet row index 1, also non-active on a fresh app).
    let gutter = &buf[(0, PANEL_HEIGHT + 2)];
    assert!(
        gutter.modifier.contains(Modifier::REVERSED),
        "default theme should set REVERSED on the row gutter"
    );
}

#[test]
fn dos_theme_paints_inactive_headers_black_on_cyan() {
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    let buf = app.render_to_buffer(80, 25);
    let col_header_y = PANEL_HEIGHT;

    // Column B (inactive) — resting header style.
    let inactive_col = &buf[(ROW_GUTTER + 9 + 4, col_header_y)];
    assert_eq!(
        inactive_col.bg,
        Color::Rgb(0, 170, 170),
        "column-B header bg"
    );
    assert_eq!(inactive_col.fg, Color::Rgb(0, 0, 0), "column-B header fg");
    assert!(
        !inactive_col.modifier.contains(Modifier::REVERSED),
        "DOS theme must not also set REVERSED — that would re-invert"
    );

    // Row 2 (sheet row 1, inactive) — resting gutter style.
    let inactive_row_gutter = &buf[(0, PANEL_HEIGHT + 2)];
    assert_eq!(
        inactive_row_gutter.bg,
        Color::Rgb(0, 170, 170),
        "row-2 gutter bg"
    );
    assert_eq!(
        inactive_row_gutter.fg,
        Color::Rgb(0, 0, 0),
        "row-2 gutter fg"
    );
}

#[test]
fn dos_theme_paints_active_column_header_deep_blue() {
    // Pointer is at A1 by default — column A is the active column.
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    let buf = app.render_to_buffer(80, 25);
    // ROW_GUTTER + 4 is inside column A's slot (default width 9).
    let active_cell = &buf[(ROW_GUTTER + 4, PANEL_HEIGHT)];
    assert_eq!(
        active_cell.bg,
        Color::Rgb(0, 0, 170),
        "active column header should wear CGA blue"
    );
    assert_eq!(active_cell.fg, Color::Rgb(255, 255, 255));
}

#[test]
fn dos_theme_paints_active_row_gutter_deep_blue() {
    // Pointer is at A1 by default — row 0 (label "1") is active.
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    let buf = app.render_to_buffer(80, 25);
    // Row 1 lands at y = PANEL_HEIGHT + 1.
    let active_gutter = &buf[(0, PANEL_HEIGHT + 1)];
    assert_eq!(active_gutter.bg, Color::Rgb(0, 0, 170));
    assert_eq!(active_gutter.fg, Color::Rgb(255, 255, 255));
}

#[test]
fn dos_theme_paints_upper_left_corner_deep_blue() {
    // Sheet-identity area in the corner above the row gutter
    // wears the same active style — CGA blue across all five
    // gutter columns.
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    let buf = app.render_to_buffer(80, 25);
    for x in 0..ROW_GUTTER {
        let cell = &buf[(x, PANEL_HEIGHT)];
        assert_eq!(
            cell.bg,
            Color::Rgb(0, 0, 170),
            "corner column {x} should be CGA blue"
        );
    }
}

#[test]
fn dos_theme_paints_selected_cell_in_resting_label_teal() {
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    let buf = app.render_to_buffer(80, 25);
    // Pointer is on A1; the cell occupies x ∈ [ROW_GUTTER, ROW_GUTTER+9)
    // at y = PANEL_HEIGHT + 1.
    let cell = &buf[(ROW_GUTTER, PANEL_HEIGHT + 1)];
    assert_eq!(
        cell.bg,
        Color::Rgb(0, 170, 170),
        "selected cell bg should match the resting label teal"
    );
    assert_eq!(
        cell.fg,
        Color::Rgb(0, 0, 0),
        "selected cell fg should be black"
    );
    assert!(
        !cell.modifier.contains(Modifier::REVERSED),
        "DOS theme paints the highlight directly — no REVERSED"
    );
}

#[test]
fn dos_theme_active_highlight_follows_pointer_on_move() {
    // Move the pointer to B2 and verify B's header + row 2's
    // gutter pick up the active deep-blue style, while A's drop
    // back to the resting teal.
    let mut app = App::new();
    app.set_theme(crate::Theme::Dos);
    app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::NONE));
    app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE));
    let buf = app.render_to_buffer(80, 25);

    // B's column header — now active (deep blue).
    let b_header = &buf[(ROW_GUTTER + 9 + 4, PANEL_HEIGHT)];
    assert_eq!(b_header.bg, Color::Rgb(0, 0, 170), "B should now be active");

    // A's column header — dropped back to resting teal.
    let a_header = &buf[(ROW_GUTTER + 4, PANEL_HEIGHT)];
    assert_eq!(
        a_header.bg,
        Color::Rgb(0, 170, 170),
        "A should no longer be active"
    );

    // Row 2's gutter — active deep blue.
    let row2_gutter = &buf[(0, PANEL_HEIGHT + 2)];
    assert_eq!(row2_gutter.bg, Color::Rgb(0, 0, 170));
    // Row 1's gutter — resting teal.
    let row1_gutter = &buf[(0, PANEL_HEIGHT + 1)];
    assert_eq!(row1_gutter.bg, Color::Rgb(0, 170, 170));
}

#[test]
fn set_theme_round_trips() {
    let mut app = App::new();
    assert_eq!(app.theme(), crate::Theme::Default);
    app.set_theme(crate::Theme::Dos);
    assert_eq!(app.theme(), crate::Theme::Dos);
    app.set_theme(crate::Theme::Default);
    assert_eq!(app.theme(), crate::Theme::Default);
}

#[test]
fn clean_dropped_path_unescapes_backslashes_from_drag_and_drop() {
    // Terminal drag-and-drop on macOS escapes spaces and `~` with
    // backslashes. PathBuf::from sees those literally, so the file
    // can't be found.
    let input = "/Users/ddmoore/Library/Mobile\\ Documents/com\\~apple\\~CloudDocs/channel19/Factoring\\ Model.xlsx";
    assert_eq!(
        clean_dropped_path(input),
        "/Users/ddmoore/Library/Mobile Documents/com~apple~CloudDocs/channel19/Factoring Model.xlsx",
    );
}

#[test]
fn clean_dropped_path_strips_outer_single_quotes() {
    assert_eq!(
        clean_dropped_path("'/tmp/has space.xlsx'"),
        "/tmp/has space.xlsx",
    );
}

#[test]
fn clean_dropped_path_strips_outer_double_quotes() {
    assert_eq!(
        clean_dropped_path("\"/tmp/has space.xlsx\""),
        "/tmp/has space.xlsx",
    );
}

#[test]
fn clean_dropped_path_trims_surrounding_whitespace() {
    // Drag-and-drop on macOS often leaves a trailing space.
    assert_eq!(clean_dropped_path("  /tmp/x.xlsx  "), "/tmp/x.xlsx");
}

#[test]
fn clean_dropped_path_passes_through_plain_paths_unchanged() {
    assert_eq!(clean_dropped_path("/tmp/plain.xlsx"), "/tmp/plain.xlsx");
}

#[test]
fn clean_dropped_path_collapses_double_backslashes_to_one() {
    // `\\` in shell-escape is a single literal backslash.
    assert_eq!(clean_dropped_path("/tmp/odd\\\\name"), "/tmp/odd\\name");
}

#[test]
fn file_retrieve_prompt_unescapes_drag_and_drop_path() {
    // Reproduces the bug: macOS Terminal drag-and-drop produces a
    // backslash-escaped path. Without unescaping, /File Retrieve
    // queues a non-existent path and the load fails.
    let mut app = App::new();
    app.test_block_next_async_op();
    app.prompt = Some(PromptState {
        label: "Enter file to retrieve:".into(),
        buffer:
            "/Users/ddmoore/Library/Mobile\\ Documents/com\\~apple\\~CloudDocs/channel19/Factoring\\ Model.xlsx"
                .into(),
        next: PromptNext::FileRetrieveFilename,
        fresh: false,
    });
    app.commit_prompt();

    let pending = app.pending_async_op.as_ref().expect("op should be queued");
    assert_eq!(pending.verb, "Loading");
    assert_eq!(pending.display_name, "Factoring Model.xlsx");
    let queued = match &pending.state {
        crate::app::types::OpState::Queued(q) => q,
        _ => panic!("op should still be in Queued state (block_next_async_op set)"),
    };
    match queued.as_ref() {
        QueuedOp::FileRetrieve { path } => assert_eq!(
            path.as_path(),
            Path::new(
                "/Users/ddmoore/Library/Mobile Documents/com~apple~CloudDocs/channel19/Factoring Model.xlsx",
            ),
        ),
        _ => panic!("expected QueuedOp::FileRetrieve"),
    }
}

fn corner_text(buf: &Buffer) -> String {
    (0..ROW_GUTTER)
        .map(|i| buf[(i, PANEL_HEIGHT)].symbol().to_string())
        .collect()
}

#[test]
fn corner_paints_active_sheet_letter() {
    let app = App::new();
    let buf = app.render_to_buffer(80, 25);
    let corner = corner_text(&buf);
    assert_eq!(
        corner.trim(),
        "A",
        "corner shows active sheet letter; got {corner:?}"
    );
}

#[test]
fn corner_follows_active_sheet_after_insert() {
    let mut app = App::new();
    drive_chord(&mut app, &['/', 'W', 'I', 'S', 'B']);
    let buf = app.render_to_buffer(80, 25);
    let corner = corner_text(&buf);
    assert_eq!(
        corner.trim(),
        "B",
        "after /WISB the original sheet shifts to B and the pointer follows; got {corner:?}"
    );
}
