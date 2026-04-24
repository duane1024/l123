//! v3.4 WYSIWYG-style icon panel — 17 icons matching the default
//! layout from `ICONS3.CNF` (ids 5, 6, 67, 68, 69, 70, 9, 10, 57, 7,
//! 8, 12, 13, 15, 23, 4, 93).
//!
//! Icons are redrawn from scratch in plotters; we don't redistribute
//! the original bitmaps. See `examples/extract_icons.rs` and
//! `examples/dump_default_panel.rs` for the reverse-engineered format
//! if you need the originals for personal reference.

use std::io::Cursor;

use plotters::prelude::*;

pub const ICON_PANEL_WIDTH_PX: u32 = 56;
pub const ICON_HEIGHT_PX: u32 = 56;
pub const ICON_COUNT: u32 = 17;
pub const ICON_PANEL_HEIGHT_PX: u32 = ICON_COUNT * ICON_HEIGHT_PX;

/// Every slot in the default v3.4 icon panel, top to bottom.
///
/// The name describes the *function* the icon triggers, mirroring the
/// help text 1-2-3 shows when the mouse hovers over the icon.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum IconKind {
    SaveFile,
    RetrieveFile,
    FileOpenAfter,
    Perspective,
    NextSheet,
    PrevSheet,
    Sum,
    GraphView,
    AddGraph,
    Print,
    PrintPreview,
    Bold,
    Italic,
    UnderlineSingle,
    FontCycle,
    Help,
    UserDefined,
}

impl IconKind {
    pub const ORDER: [IconKind; 17] = [
        IconKind::SaveFile,
        IconKind::RetrieveFile,
        IconKind::FileOpenAfter,
        IconKind::Perspective,
        IconKind::NextSheet,
        IconKind::PrevSheet,
        IconKind::Sum,
        IconKind::GraphView,
        IconKind::AddGraph,
        IconKind::Print,
        IconKind::PrintPreview,
        IconKind::Bold,
        IconKind::Italic,
        IconKind::UnderlineSingle,
        IconKind::FontCycle,
        IconKind::Help,
        IconKind::UserDefined,
    ];

    /// One-line description for status-bar / tooltip text. Phrased
    /// like 1-2-3's own help strings.
    pub fn description(self) -> &'static str {
        match self {
            IconKind::SaveFile => "Save the current worksheet to a file",
            IconKind::RetrieveFile => "Replace the current file with one from disk",
            IconKind::FileOpenAfter => "Load a second file after the current file",
            IconKind::Perspective => "Show three worksheets in perspective view",
            IconKind::NextSheet => "Move to the next worksheet",
            IconKind::PrevSheet => "Move to the previous worksheet",
            IconKind::Sum => "Sum values in a range",
            IconKind::GraphView => "Display the current graph",
            IconKind::AddGraph => "Add the current graph to the worksheet",
            IconKind::Print => "Print the current print range",
            IconKind::PrintPreview => "Preview the current print range",
            IconKind::Bold => "Toggle bold on a range",
            IconKind::Italic => "Toggle italics on a range",
            IconKind::UnderlineSingle => "Toggle single underline on a range",
            IconKind::FontCycle => "Cycle through the fonts in the font set",
            IconKind::Help => "Start the 1-2-3 Help system",
            IconKind::UserDefined => "User-defined icon (currently unassigned)",
        }
    }
}

const BG: RGBColor = RGBColor(0xC0, 0xC0, 0xC0);
const INK: RGBColor = RGBColor(0x20, 0x20, 0x20);
const ACCENT: RGBColor = RGBColor(0x00, 0x80, 0x80); // 1-2-3 teal
const WARM: RGBColor = RGBColor(0xB0, 0x40, 0x40);
const PAPER: RGBColor = RGBColor(0xF4, 0xF4, 0xF0);

/// Render the full 17-icon strip to a PNG.
pub fn render_icon_panel_png() -> Vec<u8> {
    let w = ICON_PANEL_WIDTH_PX;
    let h = ICON_PANEL_HEIGHT_PX;
    let mut rgb = vec![0u8; (w as usize) * (h as usize) * 3];
    {
        let backend = BitMapBackend::with_buffer(&mut rgb, (w, h));
        let root = backend.into_drawing_area();
        let _ = root.fill(&BG);
        let cells = root.split_evenly((ICON_COUNT as usize, 1));
        for (i, kind) in IconKind::ORDER.iter().enumerate() {
            if let Some(cell) = cells.get(i) {
                let _ = cell.draw(&Rectangle::new(
                    [(0, 0), ((w as i32) - 1, (ICON_HEIGHT_PX as i32) - 1)],
                    INK.stroke_width(1),
                ));
                draw_icon(*kind, cell);
            }
        }
        let _ = root.present();
    }
    let img = match image::RgbImage::from_raw(w, h, rgb) {
        Some(i) => i,
        None => return Vec::new(),
    };
    let mut out = Cursor::new(Vec::new());
    match img.write_to(&mut out, image::ImageFormat::Png) {
        Ok(_) => out.into_inner(),
        Err(_) => Vec::new(),
    }
}

fn draw_icon<DB>(kind: IconKind, area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    match kind {
        IconKind::SaveFile => draw_save(area),
        IconKind::RetrieveFile => draw_retrieve(area),
        IconKind::FileOpenAfter => draw_file_open_after(area),
        IconKind::Perspective => draw_perspective(area),
        IconKind::NextSheet => draw_next_sheet(area),
        IconKind::PrevSheet => draw_prev_sheet(area),
        IconKind::Sum => draw_sum(area),
        IconKind::GraphView => draw_bar_chart(area),
        IconKind::AddGraph => draw_add_graph(area),
        IconKind::Print => draw_printer(area),
        IconKind::PrintPreview => draw_print_preview(area),
        IconKind::Bold => draw_letter(area, "B", true, false),
        IconKind::Italic => draw_letter(area, "I", false, false), // italic font metric handled visually
        IconKind::UnderlineSingle => draw_letter(area, "U", false, true),
        IconKind::FontCycle => draw_font_cycle(area),
        IconKind::Help => draw_letter(area, "?", true, false),
        IconKind::UserDefined => draw_user_defined(area),
    }
}

fn dims<DB>(area: &DrawingArea<DB, plotters::coord::Shift>) -> (i32, i32)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = area.dim_in_pixel();
    (w as i32, h as i32)
}

fn fill<DB>(area: &DrawingArea<DB, plotters::coord::Shift>, rect: [(i32, i32); 2], c: RGBColor)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let _ = area.draw(&Rectangle::new(rect, c.filled()));
}

fn stroke<DB>(area: &DrawingArea<DB, plotters::coord::Shift>, rect: [(i32, i32); 2], c: RGBColor, w: u32)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let _ = area.draw(&Rectangle::new(rect, c.stroke_width(w)));
}

fn text<DB>(
    area: &DrawingArea<DB, plotters::coord::Shift>,
    s: &str,
    x: i32,
    y: i32,
    size: i32,
    c: RGBColor,
) where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let font = ("sans-serif", size).into_font().color(&c);
    let _ = area.draw_text(s, &font, (x, y));
}

fn draw_save<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let p = w / 7;
    fill(area, [(p, p), (w - p, h - p)], ACCENT);
    // Metal shutter at top: a lighter horizontal band with a small cutout.
    fill(area, [(p + 2, p + 2), (w - p - 2, p + h / 7)], BG);
    fill(
        area,
        [((w / 2) - w / 14, p + 2), ((w / 2) + w / 14, p + h / 7)],
        INK,
    );
    // Label area at the bottom.
    fill(area, [(p + w / 7, h / 2), (w - p - w / 7, h - p - 2)], PAPER);
}

fn draw_retrieve<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let p = w / 6;
    // A piece of paper.
    fill(area, [(p, p), (w - p, h - p)], PAPER);
    stroke(area, [(p, p), (w - p, h - p)], INK, 1);
    // A few text lines.
    for i in 0..4 {
        let y = p + (h / 7) + i * (h / 9);
        fill(area, [(p + 3, y), (w - p - 3, y + 2)], INK);
    }
    // A downward arrow into the document.
    let cx = w / 2;
    let shaft_top = 2;
    let shaft_bottom = p - 1;
    fill(area, [(cx - 2, shaft_top), (cx + 2, shaft_bottom)], WARM);
    let _ = area.draw(&Polygon::new(
        vec![(cx - 6, p - 2), (cx + 6, p - 2), (cx, p + 4)],
        WARM.filled(),
    ));
}

fn draw_file_open_after<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let p = w / 6;
    // Front document.
    fill(area, [(p, p + 4), (w - p - 4, h - p)], PAPER);
    stroke(area, [(p, p + 4), (w - p - 4, h - p)], INK, 1);
    // Background document offset up-right (suggests "after current").
    stroke(area, [(p + 4, p), (w - p, h - p - 4)], INK, 1);
    // A "+" mark in the corner.
    let cx = w - p - 2;
    let cy = p + 2;
    fill(area, [(cx - 4, cy - 1), (cx + 4, cy + 1)], WARM);
    fill(area, [(cx - 1, cy - 4), (cx + 1, cy + 4)], WARM);
}

fn draw_perspective<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    // Three overlapping rectangles representing stacked sheets.
    for (i, offset) in [10, 5, 0].iter().enumerate() {
        let x0 = 6 + offset;
        let y0 = 6 + (2 - i as i32) * 4;
        let x1 = w - 6 - (10 - offset);
        let y1 = h - 6 - i as i32 * 4;
        fill(area, [(x0, y0), (x1, y1)], PAPER);
        stroke(area, [(x0, y0), (x1, y1)], INK, 1);
    }
}

fn draw_sheet_stack<DB>(area: &DrawingArea<DB, plotters::coord::Shift>, x0: i32, y0: i32, x1: i32, y1: i32)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    fill(area, [(x0, y0), (x1, y1)], PAPER);
    stroke(area, [(x0, y0), (x1, y1)], INK, 1);
    // Horizontal rule lines to suggest rows.
    let rows = 4;
    let dy = (y1 - y0) / (rows + 1);
    for i in 1..=rows {
        let y = y0 + i * dy;
        let _ = area.draw(&PathElement::new(vec![(x0 + 2, y), (x1 - 2, y)], INK.stroke_width(1)));
    }
}

fn draw_next_sheet<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    draw_sheet_stack(area, 6, h / 3, w - 6, h - 6);
    // Right-pointing arrow above.
    let y = h / 4;
    fill(area, [(6, y - 2), (w / 2 + 2, y + 2)], ACCENT);
    let _ = area.draw(&Polygon::new(
        vec![(w / 2 + 2, y - 6), (w / 2 + 2, y + 6), (w / 2 + 12, y)],
        ACCENT.filled(),
    ));
}

fn draw_prev_sheet<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    draw_sheet_stack(area, 6, h / 3, w - 6, h - 6);
    let y = h / 4;
    fill(area, [(w / 2 - 2, y - 2), (w - 6, y + 2)], ACCENT);
    let _ = area.draw(&Polygon::new(
        vec![(w / 2 - 2, y - 6), (w / 2 - 2, y + 6), (w / 2 - 12, y)],
        ACCENT.filled(),
    ));
}

fn draw_sum<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    // Render the Greek capital sigma via text (covered by sans-serif fallback).
    text(area, "\u{03A3}", w / 4, h / 6, h * 3 / 4, INK);
    // A tiny "123" caption.
    text(area, "123", w / 2, h * 2 / 3, h / 4, WARM);
}

fn draw_bar_chart<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let base = h - 8;
    let left = 6;
    let right = w - 6;
    // Axis baseline.
    let _ = area.draw(&PathElement::new(
        vec![(left, base), (right, base)],
        INK.stroke_width(1),
    ));
    let bars = [0.4, 0.7, 0.5, 0.9];
    let n = bars.len() as i32;
    let bar_w = (right - left - (n + 1) * 2) / n;
    for (i, &ratio) in bars.iter().enumerate() {
        let x0 = left + 2 + (i as i32) * (bar_w + 2);
        let top = base - ((base - 6) as f64 * ratio) as i32;
        fill(area, [(x0, top), (x0 + bar_w, base)], ACCENT);
    }
}

fn draw_add_graph<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    draw_sheet_stack(area, 4, 4, w - 4, h - 4);
    // Miniature bar chart inset on the right half.
    let base = h - 10;
    let left = w / 2 + 2;
    let right = w - 8;
    let bars = [0.3, 0.6, 0.9];
    let n = bars.len() as i32;
    let bw = (right - left) / n;
    for (i, &r) in bars.iter().enumerate() {
        let x0 = left + (i as i32) * bw;
        let top = base - ((base - h / 3) as f64 * r) as i32;
        fill(area, [(x0, top), (x0 + bw - 1, base)], WARM);
    }
}

fn draw_printer<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    // Paper feeding out the top.
    fill(area, [(w / 4, 4), (w - w / 4, h / 3)], PAPER);
    stroke(area, [(w / 4, 4), (w - w / 4, h / 3)], INK, 1);
    for i in 0..3 {
        let y = 8 + i * (h / 12);
        fill(area, [(w / 4 + 3, y), (w - w / 4 - 3, y + 1)], INK);
    }
    // Printer body.
    fill(area, [(6, h / 3), (w - 6, h - 10)], ACCENT);
    stroke(area, [(6, h / 3), (w - 6, h - 10)], INK, 1);
    // Output tray.
    fill(area, [(w / 5, h - 10), (w - w / 5, h - 6)], INK);
}

fn draw_print_preview<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    // A document.
    fill(area, [(6, 4), (w - 12, h - 12)], PAPER);
    stroke(area, [(6, 4), (w - 12, h - 12)], INK, 1);
    for i in 0..4 {
        let y = 8 + i * (h / 9);
        fill(area, [(9, y), (w - 16, y + 1)], INK);
    }
    // Magnifier in the lower-right corner.
    let cx = w - 10;
    let cy = h - 10;
    let r = h / 6;
    let _ = area.draw(&Circle::new((cx, cy), r, INK.stroke_width(2)));
    let _ = area.draw(&PathElement::new(
        vec![(cx + r - 1, cy + r - 1), (cx + r + 6, cy + r + 6)],
        INK.stroke_width(2),
    ));
}

fn draw_letter<DB>(
    area: &DrawingArea<DB, plotters::coord::Shift>,
    s: &str,
    bold: bool,
    underline: bool,
) where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let size = h * 3 / 4;
    // Plotters has no direct bold; draw twice with a 1-px offset when
    // we want a heavier stroke.
    let style = ("sans-serif", size).into_font().color(&INK);
    let x = w / 3;
    let y = h / 8;
    let _ = area.draw_text(s, &style, (x, y));
    if bold {
        let _ = area.draw_text(s, &style, (x + 1, y));
    }
    if underline {
        let uy = h - 8;
        fill(area, [(w / 4, uy), (w - w / 4, uy + 2)], INK);
    }
}

fn draw_font_cycle<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    let style = ("sans-serif", h * 3 / 5).into_font().color(&INK);
    let _ = area.draw_text("A", &style, (w / 5, h / 4));
    let small = ("sans-serif", h * 2 / 5).into_font().color(&INK);
    let _ = area.draw_text("a", &small, (w * 3 / 5, h / 2));
    // Small curved arrow suggesting "cycle".
    let cx = w / 2;
    let cy = h - 8;
    let _ = area.draw(&PathElement::new(
        vec![(cx - 6, cy), (cx + 6, cy)],
        ACCENT.stroke_width(2),
    ));
    let _ = area.draw(&Polygon::new(
        vec![(cx + 6, cy - 3), (cx + 6, cy + 3), (cx + 10, cy)],
        ACCENT.filled(),
    ));
}

fn draw_user_defined<DB>(area: &DrawingArea<DB, plotters::coord::Shift>)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, h) = dims(area);
    // A dashed border to suggest "empty slot / customize".
    let step = 4;
    for x in (6..w - 6).step_by(step as usize) {
        fill(area, [(x, 8), (x + 2, 9)], INK);
        fill(area, [(x, h - 9), (x + 2, h - 8)], INK);
    }
    for y in (8..h - 8).step_by(step as usize) {
        fill(area, [(6, y), (7, y + 2)], INK);
        fill(area, [(w - 7, y), (w - 6, y + 2)], INK);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn icon_order_is_seventeen() {
        assert_eq!(IconKind::ORDER.len(), 17);
        assert_eq!(IconKind::ORDER[0], IconKind::SaveFile);
        assert_eq!(IconKind::ORDER[16], IconKind::UserDefined);
    }

    #[test]
    fn icon_count_const_matches_order() {
        assert_eq!(ICON_COUNT as usize, IconKind::ORDER.len());
    }

    #[test]
    fn every_icon_has_a_description() {
        for k in IconKind::ORDER {
            let d = k.description();
            assert!(!d.is_empty(), "empty description for {k:?}");
        }
    }

    #[test]
    fn icon_panel_png_has_magic_and_decodes() {
        let bytes = render_icon_panel_png();
        assert!(bytes.len() > 8, "png too short");
        assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n");
        let img = image::load_from_memory(&bytes).expect("valid PNG");
        assert_eq!(img.width(), ICON_PANEL_WIDTH_PX);
        assert_eq!(img.height(), ICON_PANEL_HEIGHT_PX);
    }

    #[test]
    fn icon_panel_has_non_background_pixels() {
        let bytes = render_icon_panel_png();
        let img = image::load_from_memory(&bytes).unwrap().into_rgb8();
        let mut non_bg = 0usize;
        for p in img.pixels() {
            if p.0 != [0xC0, 0xC0, 0xC0] {
                non_bg += 1;
            }
        }
        assert!(non_bg > 2000, "expected a lot of drawn pixels, got {non_bg}");
    }
}
