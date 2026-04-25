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

use ironcalc::base::{
    types::{
        Alignment, Border, BorderItem, BorderStyle, Fill, HorizontalAlignment, VerticalAlignment,
    },
    Model,
};
use ironcalc::export::save_to_xlsx;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let out_dir = fixtures_dir();
    std::fs::create_dir_all(&out_dir)?;
    regen(&out_dir.join("alignment.xlsx"), build_alignment)?;
    regen(&out_dir.join("fill.xlsx"), build_fill)?;
    regen(&out_dir.join("sheet_color.xlsx"), build_sheet_color)?;
    regen(&out_dir.join("font.xlsx"), build_font)?;
    regen(&out_dir.join("borders.xlsx"), build_borders)?;
    regen(&out_dir.join("comments.xlsx"), build_comments)?;
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
