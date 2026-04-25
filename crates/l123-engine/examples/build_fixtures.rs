//! Regenerate the xlsx fixtures under `tests/acceptance/fixtures/xlsx/`.
//!
//! Run from the workspace root:
//!
//!     cargo run -p l123-engine --example build_fixtures
//!
//! The fixtures are committed as binary blobs so `cargo test` does not
//! require running this script first; this entry point exists so the
//! files can be reproduced (and edited) without reaching for a separate
//! xlsx authoring tool.
//!
//! See `docs/XLSX_IMPORT_PLAN.md` §1 for the fixture-build strategy.
//! Caveat: we are building fixtures with the same IronCalc codepath we
//! later load them with, so a symmetric bug in IronCalc's read + write
//! will not surface here — it will only surface when we drop in an
//! Excel-authored fixture (provenance documented in the README).

use std::path::{Path, PathBuf};

use std::io::{Read, Write};

use ironcalc_xlsx::base::{
    types::{
        Alignment, Border, BorderItem, BorderStyle, Fill, HorizontalAlignment, VerticalAlignment,
    },
    Model,
};
use ironcalc_xlsx::export::save_to_xlsx;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let out_dir = fixtures_dir();
    std::fs::create_dir_all(&out_dir)?;
    regen(&out_dir.join("alignment.xlsx"), build_alignment)?;
    regen(&out_dir.join("fill.xlsx"), build_fill)?;
    regen(&out_dir.join("sheet_color.xlsx"), build_sheet_color)?;
    regen(&out_dir.join("font.xlsx"), build_font)?;
    regen(&out_dir.join("borders.xlsx"), build_borders)?;
    regen(&out_dir.join("comments.xlsx"), build_comments)?;
    regen(&out_dir.join("merges.xlsx"), build_merges)?;
    regen(&out_dir.join("frozen.xlsx"), build_frozen)?;
    regen(&out_dir.join("hidden_sheets.xlsx"), build_hidden_sheets)?;
    regen(&out_dir.join("tables.xlsx"), build_tables)?;
    println!("Wrote fixtures to {}", out_dir.display());
    Ok(())
}

/// Delete an existing fixture then re-emit it.  IronCalc's
/// `save_to_xlsx` refuses to overwrite, so callers must clear first.
fn regen(
    path: &Path,
    build: fn(&Path) -> Result<(), Box<dyn std::error::Error>>,
) -> Result<(), Box<dyn std::error::Error>> {
    if path.exists() {
        std::fs::remove_file(path)?;
    }
    build(path)
}

fn fixtures_dir() -> PathBuf {
    // `CARGO_MANIFEST_DIR` for this example is `crates/l123-engine`; the
    // workspace root is two levels up, and fixtures live under
    // `tests/acceptance/fixtures/xlsx/`.
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    manifest
        .parent()
        .and_then(|p| p.parent())
        .expect("workspace root above crates/l123-engine")
        .join("tests/acceptance/fixtures/xlsx")
}

/// `alignment.xlsx` — A1..D1 carry distinct horizontal alignments so a
/// read-side test can assert each value survives load.
///
///   A1 = 'left'    horizontal = left
///   B1 = 'center'  horizontal = center
///   C1 = 'right'   horizontal = right
///   D1 = 1234      horizontal = left     (number explicitly left-aligned)
///   E1 = 'wrap'    horizontal = left, vertical = top, wrap_text = true
fn build_alignment(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("alignment", "en", "UTC", "en")?;
    write_cell_with_alignment(&mut model, 1, 1, "'left", HorizontalAlignment::Left, false)?;
    write_cell_with_alignment(
        &mut model,
        1,
        2,
        "'center",
        HorizontalAlignment::Center,
        false,
    )?;
    write_cell_with_alignment(&mut model, 1, 3, "'right", HorizontalAlignment::Right, false)?;
    write_cell_with_alignment(&mut model, 1, 4, "1234", HorizontalAlignment::Left, false)?;
    write_cell_with_alignment(&mut model, 1, 5, "'wrap", HorizontalAlignment::Left, true)?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `fill.xlsx` — A1..C1 carry distinct solid background colors; D1 has
/// no fill.  The acceptance transcript loads this and asserts each
/// cell's rendered background in the terminal buffer.
///
///   A1 'red'    bg = #FF0000
///   B1 'green'  bg = #00C800  (a readable green — pure #00FF00 is harsh)
///   C1 'blue'   bg = #3366CC
///   D1 'plain'  no fill
fn build_fill(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("fill", "en", "UTC", "en")?;
    write_cell_with_fill(&mut model, 1, 1, "'red", "FF0000")?;
    write_cell_with_fill(&mut model, 1, 2, "'green", "00C800")?;
    write_cell_with_fill(&mut model, 1, 3, "'blue", "3366CC")?;
    // D1 'plain' — no fill applied; only the value is written.
    model.set_user_input(0, 1, 4, "'plain".to_string())?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `sheet_color.xlsx` — three sheets with distinct tab colors.
/// IronCalc 0.7's xlsx exporter doesn't write `<tabColor>` (see
/// `docs/XLSX_IMPORT_PLAN.md` §2.6), so we build the workbook
/// through IronCalc first, then hand-patch the sheet XML to inject a
/// `<sheetPr><tabColor rgb="FF..."/></sheetPr>` block on sheets 2 and
/// 3.  Sheet 1 stays untinted as the negative-control baseline.
fn build_sheet_color(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("sheet_color", "en", "UTC", "en")?;
    model.rename_sheet_by_index(0, "Overview")?;
    model.insert_sheet("Q1 Red", 1, None)?;
    model.insert_sheet("Q2 Blue", 2, None)?;
    model.set_user_input(0, 1, 1, "'overview".to_string())?;
    model.set_user_input(1, 1, 1, "'red tab".to_string())?;
    model.set_user_input(2, 1, 1, "'blue tab".to_string())?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    // Patch sheet2.xml + sheet3.xml with tabColor.  Sheet indexes in
    // xl/worksheets/ are 1-based: sheet1 = index 0, sheet2 = index 1,…
    patch_tab_color(path, "xl/worksheets/sheet2.xml", "FFDBBE29")?;
    patch_tab_color(path, "xl/worksheets/sheet3.xml", "FF3366CC")?;
    println!("  {}", path.display());
    Ok(())
}

/// Rewrite an xlsx zip, replacing `member` (a worksheet XML) with a
/// copy that has `<sheetPr><tabColor rgb="..."/></sheetPr>` inserted
/// as the first child of `<worksheet>`.  Every other member is
/// copied byte-for-byte so the patched file stays a valid xlsx.
fn patch_tab_color(
    path: &Path,
    member: &str,
    rgb: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read(path)?;
    let reader = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(reader)?;

    let mut out_buf: Vec<u8> = Vec::new();
    {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut out_buf));
        let opts = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();
            if name == member {
                let mut xml = String::new();
                entry.read_to_string(&mut xml)?;
                let patched = inject_tab_color(&xml, rgb)?;
                writer.start_file(name, opts)?;
                writer.write_all(patched.as_bytes())?;
            } else {
                let mut buf = Vec::with_capacity(entry.size() as usize);
                entry.read_to_end(&mut buf)?;
                writer.start_file(name, opts)?;
                writer.write_all(&buf)?;
            }
        }
        writer.finish()?;
    }
    std::fs::write(path, out_buf)?;
    Ok(())
}

/// Insert a `<sheetPr><tabColor rgb="FFxxxxxx"/></sheetPr>` block
/// immediately after the opening `<worksheet ...>` tag.  Fails loudly
/// if the XML shape doesn't match the IronCalc exporter's output —
/// catches the day IronCalc reorganises its XML and this helper needs
/// a rewrite.
fn inject_tab_color(xml: &str, rgb: &str) -> Result<String, String> {
    let insert = format!(r#"<sheetPr><tabColor rgb="{rgb}"/></sheetPr>"#);
    // Find the first `>` that closes the opening `<worksheet ...>` tag.
    let open_start = xml
        .find("<worksheet")
        .ok_or_else(|| "no <worksheet> open tag".to_string())?;
    let close_rel = xml[open_start..]
        .find('>')
        .ok_or_else(|| "unterminated <worksheet> open tag".to_string())?;
    let insert_pos = open_start + close_rel + 1;
    let mut out = String::with_capacity(xml.len() + insert.len());
    out.push_str(&xml[..insert_pos]);
    out.push_str(&insert);
    out.push_str(&xml[insert_pos..]);
    Ok(out)
}

/// `font.xlsx` — cells with distinct font fg colors + strikethrough.
/// IronCalc's exporter DOES emit font color / strike, so no xlsx
/// hand-patching is needed here.
///
///   A1 'red'     fg = #FF0000
///   B1 'blue'    fg = #3366CC
///   C1 'struck'  strikethrough, no color override
///   D1 'plain'   default font
fn build_font(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("font", "en", "UTC", "en")?;
    write_cell_with_font(&mut model, 1, 1, "'red", Some("#FF0000"), false)?;
    write_cell_with_font(&mut model, 1, 2, "'blue", Some("#3366CC"), false)?;
    write_cell_with_font(&mut model, 1, 3, "'struck", None, true)?;
    // D1 plain — no font override.
    model.set_user_input(0, 1, 4, "'plain".to_string())?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `borders.xlsx` — A1..D1 carry distinct right-side borders so the
/// transcript can verify the rendered glyph for each style.
///
///   A1 'thin'    right = thin
///   B1 'thick'   right = thick
///   C1 'double'  right = double
///   D1 'dashed'  right = mediumDashed (renders as Dashed in L123)
///   E1 'plain'   no border
///
/// `dotted` is intentionally not exercised here: IronCalc 0.7's xlsx
/// importer drops it back to `Thin` (see the
/// `dotted_border_degrades_on_xlsx_round_trip_upstream_gap` test in
/// the engine adapter).  When upstream closes the gap, add it.
fn build_borders(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("borders", "en", "UTC", "en")?;
    write_cell_with_right_border(&mut model, 1, 1, "'thin", BorderStyle::Thin)?;
    write_cell_with_right_border(&mut model, 1, 2, "'thick", BorderStyle::Thick)?;
    write_cell_with_right_border(&mut model, 1, 3, "'double", BorderStyle::Double)?;
    write_cell_with_right_border(&mut model, 1, 4, "'dashed", BorderStyle::MediumDashed)?;
    model.set_user_input(0, 1, 5, "'plain".to_string())?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `comments.xlsx` — a single sheet with two commented cells.
/// IronCalc 0.7's xlsx exporter doesn't write `xl/comments1.xml`,
/// so we build the base xlsx through IronCalc then inject the
/// comments file + the sheet→comments relationship by hand.
///
///   A1 'note 1' — comment "first note"
///   B2 'note 2' — comment "second note"
///   C3 'plain' — no comment
///
/// IronCalc's importer doesn't read author names — every loaded
/// comment surfaces with `author = ""`.  The fixture omits an author
/// for clarity rather than authoring one that would silently drop.
fn build_comments(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("comments", "en", "UTC", "en")?;
    model.set_user_input(0, 1, 1, "'note 1".to_string())?;
    model.set_user_input(0, 2, 2, "'note 2".to_string())?;
    model.set_user_input(0, 3, 3, "'plain".to_string())?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;

    // xl/comments1.xml — IronCalc's importer reads `commentList` and
    // pulls each `<comment ref="...">` with its concatenated `<t>` text.
    // It ignores authors entirely, so the `authors` block is here for
    // schema correctness only.  XML is one-line, no insignificant
    // whitespace — IronCalc's rels reader uses `first_child()` and is
    // sensitive to leading whitespace text nodes inside the root.
    let comments_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        r#"<authors><author>Alice</author></authors>"#,
        r#"<commentList>"#,
        r#"<comment ref="A1" authorId="0"><text><t>first note</t></text></comment>"#,
        r#"<comment ref="B2" authorId="0"><text><t>second note</t></text></comment>"#,
        r#"</commentList>"#,
        r#"</comments>"#,
    );
    let sheet_rels_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>"#,
        r#"</Relationships>"#,
    );
    add_to_xlsx(
        path,
        &[
            ("xl/comments1.xml", comments_xml),
            ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels_xml),
        ],
    )?;
    println!("  {}", path.display());
    Ok(())
}

/// `merges.xlsx` — merged-cell fixture exercising both horizontal
/// (one-row) and rectangular (multi-row) merges.  IronCalc's xlsx
/// exporter writes `<mergeCells>` faithfully so no hand-patching is
/// needed.
///
///   A1:C1 anchor=A1, "Header"  — horizontal three-cell merge
///   B3:C4 anchor=B3, "Box"     — 2x2 rectangular merge
///   A5    'plain'              — control cell, not merged
fn build_merges(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("merges", "en", "UTC", "en")?;
    model.set_user_input(0, 1, 1, "'Header".to_string())?;
    model.set_user_input(0, 3, 2, "'Box".to_string())?;
    model.set_user_input(0, 5, 1, "'plain".to_string())?;
    {
        let ws = model.workbook.worksheet_mut(0)?;
        ws.merge_cells.push("A1:C1".to_string());
        ws.merge_cells.push("B3:C4".to_string());
    }
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `frozen.xlsx` — exercises pinned rows + columns.
///
///   frozen_rows = 2, frozen_columns = 1
///   A1 'TL', A2 'L'              — frozen top-left corner cells
///   B1..AC1 'C{n}'               — long top frozen-row band
///   A3..A60 'R{n}'               — long left frozen-column band
///   B3..AC60 cells unimportant   — scrolling main region
///
/// IronCalc round-trips frozen panes natively, so no hand-patching.
fn build_frozen(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("frozen", "en", "UTC", "en")?;
    // Corner cells.
    model.set_user_input(0, 1, 1, "'TL".to_string())?;
    model.set_user_input(0, 2, 1, "'L".to_string())?;
    // Top frozen-row content: B1..AC1.
    for col in 2..=29 {
        let label = format!("'C{}", col - 1);
        model.set_user_input(0, 1, col, label)?;
    }
    // Left frozen-column content: A3..A60.
    for row in 3..=60 {
        let label = format!("'R{}", row);
        model.set_user_input(0, row, 1, label)?;
    }
    // A few cells in the scrolling main region for navigation cues.
    model.set_user_input(0, 5, 5, "'BODY".to_string())?;
    model.set_frozen_rows(0, 2)?;
    model.set_frozen_columns(0, 1)?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `hidden_sheets.xlsx` — exercises sheet visibility round-trip and
/// the L123 navigation skip behavior.
///
///   Sheet A (Visible)    "'Visible 1"
///   Sheet B (Hidden)     "'Hidden body"
///   Sheet C (VeryHidden) "'VeryHidden body"
///   Sheet D (Visible)    "'Visible 2"
///
/// IronCalc round-trips state natively via the workbook XML's
/// `<sheet state="..."/>` attribute.
fn build_hidden_sheets(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    use ironcalc_xlsx::base::types::SheetState;
    let mut model = Model::new_empty("hidden_sheets", "en", "UTC", "en")?;
    model.insert_sheet("Sheet2", 1, None)?;
    model.insert_sheet("Sheet3", 2, None)?;
    model.insert_sheet("Sheet4", 3, None)?;
    model.set_user_input(0, 1, 1, "'Visible 1".to_string())?;
    model.set_user_input(1, 1, 1, "'Hidden body".to_string())?;
    model.set_user_input(2, 1, 1, "'VeryHidden body".to_string())?;
    model.set_user_input(3, 1, 1, "'Visible 2".to_string())?;
    model.set_sheet_state(1, SheetState::Hidden)?;
    model.set_sheet_state(2, SheetState::VeryHidden)?;
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;
    println!("  {}", path.display());
    Ok(())
}

/// `tables.xlsx` — a single sheet with a 4×4 table named `Table1`.
/// IronCalc 0.7's xlsx exporter does NOT write `xl/tables/*.xml`,
/// so we build the base file with IronCalc and hand-inject the
/// table XML + sheet→table relationship.
///
/// Layout:
///   A1 'Year'   B1 'Q1'  C1 'Q2'  D1 'Q3'   — header row
///   A2 2024     B2 100   C2 110   D2 120
///   A3 2025     B3 130   C3 140   D3 150
///
/// One named table covers A1:D3 with autofilter enabled.
fn build_tables(path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut model = Model::new_empty("tables", "en", "UTC", "en")?;
    // Header row.
    model.set_user_input(0, 1, 1, "'Year".to_string())?;
    model.set_user_input(0, 1, 2, "'Q1".to_string())?;
    model.set_user_input(0, 1, 3, "'Q2".to_string())?;
    model.set_user_input(0, 1, 4, "'Q3".to_string())?;
    // Two data rows.
    for (row, year) in [(2, 2024), (3, 2025)] {
        model.set_user_input(0, row, 1, year.to_string())?;
    }
    for (row, vals) in [(2, [100, 110, 120]), (3, [130, 140, 150])] {
        for (i, v) in vals.iter().enumerate() {
            model.set_user_input(0, row, 2 + i as i32, v.to_string())?;
        }
    }
    let path_str = path
        .to_str()
        .ok_or_else(|| format!("non-UTF8 path: {}", path.display()))?;
    save_to_xlsx(&model, path_str)?;

    // xl/tables/table1.xml — the table definition itself.  One-line
    // XML to dodge the IronCalc-importer whitespace gotcha noted in
    // `build_comments` (its `first_child()` parser trips on leading
    // whitespace text nodes).
    let table_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
        r#"id="1" name="Table1" displayName="Table1" ref="A1:D3" totalsRowShown="0">"#,
        // IronCalc 0.7 sets `has_filters = true` only when
        // `<autoFilter>` has at least one child element — so we add
        // an empty `<filterColumn colId="0"/>` placeholder.  Without
        // it the flag round-trips as false even though the schema
        // intends "filters are present".
        r#"<autoFilter ref="A1:D3"><filterColumn colId="0"/></autoFilter>"#,
        r#"<tableColumns count="4">"#,
        r#"<tableColumn id="1" name="Year"/>"#,
        r#"<tableColumn id="2" name="Q1"/>"#,
        r#"<tableColumn id="3" name="Q2"/>"#,
        r#"<tableColumn id="4" name="Q3"/>"#,
        r#"</tableColumns>"#,
        r#"<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" "#,
        r#"showRowStripes="1" showColumnStripes="0"/>"#,
        r#"</table>"#,
    );
    // Sheet rels: point sheet1 at the new table part.  Same shape as
    // the comments fixture's relationship file.
    let sheet_rels_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>"#,
        r#"</Relationships>"#,
    );
    add_to_xlsx(
        path,
        &[
            ("xl/tables/table1.xml", table_xml),
            ("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels_xml),
        ],
    )?;
    println!("  {}", path.display());
    Ok(())
}

/// Append new zip members to an existing xlsx without disturbing the
/// existing parts.  Used when IronCalc's exporter doesn't emit a part
/// L123 needs (e.g. comments, sheet rels).  Members already present
/// in the archive are kept as-is — to *replace* a member, use
/// `patch_tab_color`'s style of inline rewrite.
fn add_to_xlsx(path: &Path, additions: &[(&str, &str)]) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read(path)?;
    let reader = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(reader)?;
    let existing: std::collections::HashSet<String> =
        (0..archive.len()).map(|i| archive.by_index(i).unwrap().name().to_string()).collect();

    let mut out_buf: Vec<u8> = Vec::new();
    {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(&mut out_buf));
        let opts = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);
        for i in 0..archive.len() {
            let mut entry = archive.by_index(i)?;
            let name = entry.name().to_string();
            let mut buf = Vec::with_capacity(entry.size() as usize);
            entry.read_to_end(&mut buf)?;
            writer.start_file(name, opts)?;
            writer.write_all(&buf)?;
        }
        for (name, body) in additions {
            if existing.contains(*name) {
                continue;
            }
            writer.start_file(*name, opts)?;
            writer.write_all(body.as_bytes())?;
        }
        writer.finish()?;
    }
    std::fs::write(path, out_buf)?;
    Ok(())
}

fn write_cell_with_right_border(
    model: &mut Model,
    row_1b: i32,
    col_1b: i32,
    input: &str,
    style: BorderStyle,
) -> Result<(), String> {
    model.set_user_input(0, row_1b, col_1b, input.to_string())?;
    let mut s = model.get_style_for_cell(0, row_1b, col_1b)?;
    s.border = Border {
        diagonal_up: false,
        diagonal_down: false,
        left: None,
        right: Some(BorderItem { style, color: None }),
        top: None,
        bottom: None,
        diagonal: None,
    };
    model.set_cell_style(0, row_1b, col_1b, &s)?;
    Ok(())
}

fn write_cell_with_font(
    model: &mut Model,
    row_1b: i32,
    col_1b: i32,
    input: &str,
    color_hex: Option<&str>,
    strike: bool,
) -> Result<(), String> {
    model.set_user_input(0, row_1b, col_1b, input.to_string())?;
    let mut style = model.get_style_for_cell(0, row_1b, col_1b)?;
    style.font.color = color_hex.map(str::to_string);
    style.font.strike = strike;
    model.set_cell_style(0, row_1b, col_1b, &style)?;
    Ok(())
}

fn write_cell_with_fill(
    model: &mut Model,
    row_1b: i32,
    col_1b: i32,
    input: &str,
    rgb_hex: &str,
) -> Result<(), String> {
    model.set_user_input(0, row_1b, col_1b, input.to_string())?;
    let mut style = model.get_style_for_cell(0, row_1b, col_1b)?;
    style.fill = Fill {
        pattern_type: "solid".to_string(),
        fg_color: None,
        // IronCalc's xlsx exporter prepends `FF` to this string, so we
        // hand it a 6-char RGB (no alpha) to get a well-formed ARGB
        // on disk.  Matches the shape used by `to_ic_fill` in the
        // L123 engine adapter.
        bg_color: Some(rgb_hex.to_string()),
    };
    model.set_cell_style(0, row_1b, col_1b, &style)?;
    Ok(())
}

fn write_cell_with_alignment(
    model: &mut Model,
    row_1b: i32,
    col_1b: i32,
    input: &str,
    horizontal: HorizontalAlignment,
    wrap_text: bool,
) -> Result<(), String> {
    model.set_user_input(0, row_1b, col_1b, input.to_string())?;
    let mut style = model.get_style_for_cell(0, row_1b, col_1b)?;
    style.alignment = Some(Alignment {
        horizontal,
        vertical: if wrap_text {
            VerticalAlignment::Top
        } else {
            VerticalAlignment::Bottom
        },
        wrap_text,
    });
    model.set_cell_style(0, row_1b, col_1b, &style)?;
    Ok(())
}
