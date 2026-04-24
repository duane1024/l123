//! v3.4-style WYSIWYG icon panel.
//!
//! Renders authentic 24×24 monochrome bitmaps from the catalog baked
//! into `icon_data.rs`. Seven panels are supported: panel 1's layout
//! is byte-authentic to the factory `ICONS3.CNF` (icon IDs 5, 6, 67,
//! 68, 69, 70, 9, 10, 57, 7, 8, 12, 13, 15, 23, 4 — with the 17th
//! slot reserved for the dynamic panel navigator). Panels 2-7 are
//! our thematic groupings from the same catalog; the factory 3.4
//! defaults for panels 2-7 are compiled into WYSIWYG.EXP and not
//! recoverable from disk.

use std::io::Cursor;

use plotters::prelude::*;

use crate::icon_data::{ICON_BITMAPS, ICON_DESCRIPTIONS, BITMAP_DIM};

/// Logical pixel size of one icon cell in the generated PNG. Ratatui-
/// image downscales this to fit the terminal area.
pub const CELL_SIZE_PX: u32 = 72;
pub const CELLS_PER_PANEL: u32 = 17;
pub const ICON_PANEL_WIDTH_PX: u32 = CELL_SIZE_PX;
pub const ICON_PANEL_HEIGHT_PX: u32 = CELLS_PER_PANEL * CELL_SIZE_PX;

/// One of seven selectable panels. Panel 1 is the factory default;
/// the rest are our thematic groupings.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum Panel {
    One,
    Two,
    Three,
    Four,
    Five,
    Six,
    Seven,
}

impl Panel {
    pub const ORDER: [Panel; 7] = [
        Panel::One, Panel::Two, Panel::Three, Panel::Four, Panel::Five, Panel::Six, Panel::Seven,
    ];

    /// 1-based panel number shown in the slot-16 navigator.
    pub fn number(self) -> u8 {
        match self {
            Panel::One => 1,
            Panel::Two => 2,
            Panel::Three => 3,
            Panel::Four => 4,
            Panel::Five => 5,
            Panel::Six => 6,
            Panel::Seven => 7,
        }
    }

    pub fn next(self) -> Panel {
        Panel::ORDER[(self.number() as usize) % Panel::ORDER.len()]
    }

    pub fn prev(self) -> Panel {
        let n = Panel::ORDER.len();
        Panel::ORDER[(self.number() as usize - 1 + n - 1) % n]
    }

    /// Icon IDs for the 16 non-pager slots of this panel. Slot 16 is
    /// always the dynamic panel navigator, not a catalog icon.
    pub fn icon_ids(self) -> [u8; 16] {
        match self {
            // Factory default — matches ICONS3.CNF bytes 4..19 from
            // the user's v3.4 install.
            Panel::One => [5, 6, 67, 68, 69, 70, 9, 10, 57, 7, 8, 12, 13, 15, 23, 4],

            // Cursor + editing.
            Panel::Two => [38, 39, 40, 41, 42, 43, 49, 50, 26, 27, 33, 34, 35, 36, 37, 4],

            // Format & WYSIWYG style.
            Panel::Three => [12, 13, 15, 16, 14, 17, 18, 19, 20, 21, 22, 23, 24, 25, 44, 4],

            // Data & alignment.
            Panel::Four => [31, 32, 28, 29, 30, 51, 9, 45, 46, 47, 48, 62, 60, 61, 49, 4],

            // View, scroll, window.
            Panel::Five => [65, 68, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 63, 64, 11, 4],

            // Graph-focused.
            Panel::Six => [10, 57, 58, 9, 45, 46, 7, 8, 5, 6, 50, 49, 65, 44, 11, 4],

            // File + macros / misc.
            Panel::Seven => [5, 6, 66, 67, 52, 53, 54, 55, 56, 59, 11, 44, 49, 50, 61, 4],
        }
    }
}

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum IconAction {
    /// Open the slash menu and descend via these chars as if the user
    /// typed "/" followed by each character in turn (case-insensitive).
    MenuPath(&'static str),
    /// Fire a non-menu key.
    SysKey(SysAction),
    /// Dynamic panel navigator. Slot 16 uses this; the UI decides
    /// prev vs next based on click x-coordinate inside the cell.
    PageNav,
    /// Safe no-op — feature not implemented yet.
    Noop,
}

#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum SysAction {
    /// F10 — enter full-screen graph view.
    GraphView,
    /// Alt-F4 — undo last command.
    Undo,
    /// Home — move cell pointer to A1 of the current sheet.
    Home,
    /// F9 — recalculate.
    Recalc,
    /// F2 — edit current cell.
    Edit,
    /// F5 — goto (address prompt).
    Goto,
    /// Ctrl-PgDn — next worksheet.
    NextSheet,
    /// Ctrl-PgUp — previous worksheet.
    PrevSheet,
}

/// Return the action a click on the given icon ID should fire.
///
/// Wired paths mirror the hover-help text from `ICON_DESCRIPTIONS`.
/// Unmapped IDs (mostly WYSIWYG formatting we don't implement and the
/// macro / custom-palette management that depends on features still
/// on the roadmap) are safe no-ops.
pub fn icon_action(id: u8) -> IconAction {
    use IconAction::*;
    use SysAction::*;
    match id {
        4 => Noop,                    // Help (context help not wired)
        5 => MenuPath("FS"),          // Save
        6 => MenuPath("FR"),          // Retrieve
        7 => MenuPath("PF"),          // Print file
        10 => SysKey(GraphView),
        11 => SysKey(Undo),
        14 => MenuPath("RFR"),        // Clear formats → Range Format Reset
        17 => MenuPath("RFC"),        // Currency
        19 => MenuPath("RFP"),        // Percent
        26 => MenuPath("C"),          // Copy
        27 => MenuPath("M"),          // Move
        28 => MenuPath("RLL"),        // Label Left
        29 => MenuPath("RLC"),        // Label Center
        30 => MenuPath("RLR"),        // Label Right
        33 => MenuPath("RE"),         // Range Erase
        34 => MenuPath("WIR"),        // Insert row
        35 => MenuPath("WIC"),        // Insert column
        36 => MenuPath("WDR"),        // Delete row
        37 => MenuPath("WDC"),        // Delete column
        38 => SysKey(Home),
        44 => SysKey(Recalc),
        49 => SysKey(Goto),
        50 => MenuPath("RS"),         // Range Search
        51 => MenuPath("DF"),         // Data Fill
        58 => SysKey(GraphView),
        61 => SysKey(Edit),
        66 => MenuPath("FN"),         // File New
        67 => MenuPath("FOA"),        // File Open After
        69 => SysKey(NextSheet),
        70 => SysKey(PrevSheet),
        71 => MenuPath("WISA"),       // Worksheet Insert Sheet After
        // Everything else: Noop until the underlying feature lands.
        _ => Noop,
    }
}

/// Description for a slot in a panel. Handles the special pager slot
/// separately since its text changes with the panel.
pub fn slot_description(panel: Panel, slot: usize) -> String {
    if slot == 16 {
        return format!(
            "Switch icon panel (currently panel {} of 7)",
            panel.number()
        );
    }
    let id = panel.icon_ids()[slot];
    let d = ICON_DESCRIPTIONS.get(id as usize).copied().unwrap_or("");
    if d.is_empty() {
        format!("Icon #{id}")
    } else {
        d.to_string()
    }
}

const BG: RGBColor = RGBColor(0xC0, 0xC0, 0xC0);
const INK: RGBColor = RGBColor(0x10, 0x10, 0x10);
const ACCENT: RGBColor = RGBColor(0x00, 0x80, 0x80);

pub fn render_panel_png(panel: Panel) -> Vec<u8> {
    let w = ICON_PANEL_WIDTH_PX;
    let h = ICON_PANEL_HEIGHT_PX;
    let mut rgb = vec![0u8; (w as usize) * (h as usize) * 3];
    {
        let backend = BitMapBackend::with_buffer(&mut rgb, (w, h));
        let root = backend.into_drawing_area();
        let _ = root.fill(&BG);
        let cells = root.split_evenly((CELLS_PER_PANEL as usize, 1));
        let ids = panel.icon_ids();
        for (slot, cell) in cells.iter().enumerate() {
            // Subtle frame around each cell.
            let (cw, ch) = cell.dim_in_pixel();
            let _ = cell.draw(&Rectangle::new(
                [(0, 0), ((cw as i32) - 1, (ch as i32) - 1)],
                INK.stroke_width(1),
            ));
            if slot < 16 {
                let id = ids[slot] as usize;
                paint_bitmap(cell, &ICON_BITMAPS[id]);
            } else {
                draw_pager(cell, panel);
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

/// Paint a 24×24 monochrome bitmap into `cell`, scaled to fit. The
/// bitmap stores one bit per pixel, MSB-first; bit value 0 = ink.
/// The last row of every catalog bitmap is a solid black separator
/// row that we crop out, rendering only rows 0..23 (top 23).
fn paint_bitmap<DB>(cell: &DrawingArea<DB, plotters::coord::Shift>, bits: &[u8; 72])
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (cw, ch) = cell.dim_in_pixel();
    let (cw, ch) = (cw as i32, ch as i32);
    // Leave 4-pixel padding around the icon so the frame doesn't
    // swallow the bitmap edges.
    let pad = 4;
    let usable_w = cw - 2 * pad;
    let usable_h = ch - 2 * pad;
    // Scale per source pixel, rounded down so we always fit. Use the
    // smaller scale so aspect ratio stays 1:1.
    let rows_used: i32 = (BITMAP_DIM as i32) - 1; // crop bottom separator row
    let cols_used: i32 = BITMAP_DIM as i32;
    let scale = (usable_w / cols_used).min(usable_h / rows_used).max(1);
    let draw_w = cols_used * scale;
    let draw_h = rows_used * scale;
    let off_x = pad + (usable_w - draw_w) / 2;
    let off_y = pad + (usable_h - draw_h) / 2;
    for y in 0..rows_used {
        for x in 0..cols_used {
            let byte = bits[(y * 3 + x / 8) as usize];
            let bit = 7 - (x % 8);
            if (byte >> bit) & 1 == 0 {
                let x0 = off_x + x * scale;
                let y0 = off_y + y * scale;
                let _ = cell.draw(&Rectangle::new(
                    [(x0, y0), (x0 + scale - 1, y0 + scale - 1)],
                    INK.filled(),
                ));
            }
        }
    }
}

/// The slot-16 pager: "◀ N ▶" where N is the current panel number.
fn draw_pager<DB>(cell: &DrawingArea<DB, plotters::coord::Shift>, panel: Panel)
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (cw, ch) = cell.dim_in_pixel();
    let (cw, ch) = (cw as i32, ch as i32);
    let cx = cw / 2;
    let cy = ch / 2;
    let tri_w = cw / 6;
    let tri_h = ch / 4;
    // Left arrow.
    let _ = cell.draw(&Polygon::new(
        vec![
            (cx - tri_w * 2, cy),
            (cx - tri_w, cy - tri_h / 2),
            (cx - tri_w, cy + tri_h / 2),
        ],
        ACCENT.filled(),
    ));
    // Right arrow.
    let _ = cell.draw(&Polygon::new(
        vec![
            (cx + tri_w * 2, cy),
            (cx + tri_w, cy - tri_h / 2),
            (cx + tri_w, cy + tri_h / 2),
        ],
        ACCENT.filled(),
    ));
    // Panel number in the middle.
    let label = format!("{}", panel.number());
    let size = ch * 2 / 5;
    let font = ("sans-serif", size).into_font().color(&INK);
    let _ = cell.draw_text(&label, &font, (cx - size / 4, cy - size * 3 / 5));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn panel_cycling_wraps_forward() {
        let mut p = Panel::One;
        for _ in 0..14 {
            p = p.next();
        }
        assert_eq!(p, Panel::One, "14 next() should wrap back to One");
    }

    #[test]
    fn panel_cycling_wraps_backward() {
        assert_eq!(Panel::One.prev(), Panel::Seven);
        assert_eq!(Panel::Seven.next(), Panel::One);
    }

    #[test]
    fn panel_numbers_are_1_through_7() {
        for (i, p) in Panel::ORDER.iter().enumerate() {
            assert_eq!(p.number() as usize, i + 1);
        }
    }

    #[test]
    fn panel_one_matches_icons3_cnf() {
        // The byte sequence taken directly from ICONS3.CNF (minus the
        // last slot, which we render as the dynamic pager).
        assert_eq!(
            Panel::One.icon_ids(),
            [5, 6, 67, 68, 69, 70, 9, 10, 57, 7, 8, 12, 13, 15, 23, 4]
        );
    }

    #[test]
    fn every_id_is_within_catalog() {
        for panel in Panel::ORDER {
            for &id in &panel.icon_ids() {
                assert!(
                    (id as usize) < ICON_BITMAPS.len(),
                    "{panel:?} slot uses out-of-range id {id}"
                );
            }
        }
    }

    #[test]
    fn icon_action_wires_save_retrieve_print() {
        assert_eq!(icon_action(5), IconAction::MenuPath("FS"));
        assert_eq!(icon_action(6), IconAction::MenuPath("FR"));
        assert_eq!(icon_action(7), IconAction::MenuPath("PF"));
    }

    #[test]
    fn icon_action_wires_sheet_nav() {
        assert_eq!(icon_action(69), IconAction::SysKey(SysAction::NextSheet));
        assert_eq!(icon_action(70), IconAction::SysKey(SysAction::PrevSheet));
    }

    #[test]
    fn icon_action_wires_graph_view() {
        assert_eq!(icon_action(10), IconAction::SysKey(SysAction::GraphView));
        assert_eq!(icon_action(58), IconAction::SysKey(SysAction::GraphView));
    }

    #[test]
    fn icon_action_unmapped_ids_are_noop() {
        assert_eq!(icon_action(4), IconAction::Noop); // help
        assert_eq!(icon_action(12), IconAction::Noop); // bold (wysiwyg)
        assert_eq!(icon_action(93), IconAction::Noop);
    }

    #[test]
    fn panel_png_decodes_for_each_panel() {
        for panel in Panel::ORDER {
            let bytes = render_panel_png(panel);
            assert!(bytes.len() > 8, "{panel:?}: png too short");
            assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n", "{panel:?}: bad magic");
            let img = image::load_from_memory(&bytes).expect("valid PNG");
            assert_eq!(img.width(), ICON_PANEL_WIDTH_PX);
            assert_eq!(img.height(), ICON_PANEL_HEIGHT_PX);
        }
    }

    #[test]
    fn pager_description_includes_current_number() {
        let d = slot_description(Panel::Three, 16);
        assert!(d.contains('3'), "pager description should mention number: {d}");
    }
}
