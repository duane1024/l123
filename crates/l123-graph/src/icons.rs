//! Icon-panel raster generator for the v3.1 WYSIWYG right-edge strip.
//!
//! Produces a single tall-narrow PNG containing the seven navigation
//! icons documented in the *Wysiwyg Publishing and Presentation* manual,
//! chapter 2, "The Icon Panel":
//!
//! | Icon | Action                                              |
//! |------|-----------------------------------------------------|
//! | ◀    | Move cell pointer left                              |
//! | ▼    | Move cell pointer down                              |
//! | ▲    | Move cell pointer up                                |
//! | ▶    | Move cell pointer right                             |
//! | ↑    | Move cell pointer forward through worksheets        |
//! | ↓    | Move cell pointer backward through worksheets       |
//! | ?    | Display Help                                        |
//!
//! Ratatui-image handles the downscale into the terminal's icon-panel
//! area (see `l123-ui`).

use std::io::Cursor;

use plotters::prelude::*;

/// Width in pixels of the generated icon strip. Chosen to comfortably
/// rasterize into 3 terminal columns (~24-30 px) at common font sizes.
pub const ICON_PANEL_WIDTH_PX: u32 = 80;
/// Per-icon height in pixels. Seven icons at 80 px → 560 px tall.
pub const ICON_HEIGHT_PX: u32 = 80;
/// Icons in display order, top to bottom.
pub const ICON_COUNT: u32 = 7;
/// Full generated image height.
pub const ICON_PANEL_HEIGHT_PX: u32 = ICON_COUNT * ICON_HEIGHT_PX;

/// The seven icons in the same vertical order as the rendered PNG.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum IconKind {
    Left,
    Down,
    Up,
    Right,
    SheetForward,
    SheetBackward,
    Help,
}

impl IconKind {
    pub const ORDER: [IconKind; 7] = [
        IconKind::Left,
        IconKind::Down,
        IconKind::Up,
        IconKind::Right,
        IconKind::SheetForward,
        IconKind::SheetBackward,
        IconKind::Help,
    ];
}

/// Render the icon panel to a PNG byte buffer.
pub fn render_icon_panel_png() -> Vec<u8> {
    let w = ICON_PANEL_WIDTH_PX;
    let h = ICON_PANEL_HEIGHT_PX;
    let mut rgb = vec![0u8; (w as usize) * (h as usize) * 3];
    {
        let backend = BitMapBackend::with_buffer(&mut rgb, (w, h));
        let root = backend.into_drawing_area();
        // Light-grey background, matching Lotus's muted panel tone.
        let bg = RGBColor(0xC0, 0xC0, 0xC0);
        let _ = root.fill(&bg);
        let cells = root.split_evenly((ICON_COUNT as usize, 1));
        for (i, kind) in IconKind::ORDER.iter().enumerate() {
            if let Some(cell) = cells.get(i) {
                draw_icon(*kind, cell);
                let _ = cell.draw(&Rectangle::new(
                    [(0, 0), ((w as i32) - 1, (ICON_HEIGHT_PX as i32) - 1)],
                    BLACK.stroke_width(1),
                ));
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
    let (w, h) = area.dim_in_pixel();
    let (w, h) = (w as i32, h as i32);
    let cx = w / 2;
    let cy = h / 2;
    let tri_color = RGBColor(0x20, 0x20, 0x20);
    let arrow_color = RGBColor(0x30, 0x30, 0x60);
    let pad = (w.min(h) / 5).max(4);
    match kind {
        IconKind::Left => {
            let poly = vec![
                (w - pad, pad),
                (w - pad, h - pad),
                (pad, cy),
            ];
            let _ = area.draw(&Polygon::new(poly, tri_color.filled()));
        }
        IconKind::Right => {
            let poly = vec![
                (pad, pad),
                (pad, h - pad),
                (w - pad, cy),
            ];
            let _ = area.draw(&Polygon::new(poly, tri_color.filled()));
        }
        IconKind::Down => {
            let poly = vec![
                (pad, pad),
                (w - pad, pad),
                (cx, h - pad),
            ];
            let _ = area.draw(&Polygon::new(poly, tri_color.filled()));
        }
        IconKind::Up => {
            let poly = vec![
                (pad, h - pad),
                (w - pad, h - pad),
                (cx, pad),
            ];
            let _ = area.draw(&Polygon::new(poly, tri_color.filled()));
        }
        IconKind::SheetForward => draw_vertical_arrow(area, cx, pad, h - pad, true, arrow_color),
        IconKind::SheetBackward => draw_vertical_arrow(area, cx, pad, h - pad, false, arrow_color),
        IconKind::Help => {
            // A centered question mark. Plotters' text drawing needs
            // TTF — the workspace already enables that feature.
            let size = (h - 2 * pad).max(16);
            let font = ("sans-serif", size).into_font().color(&tri_color);
            let _ = area.draw_text("?", &font, (cx - size / 3, pad));
        }
    }
}

/// A narrow rectangular shaft plus a triangular head, pointing up or down.
fn draw_vertical_arrow<DB>(
    area: &DrawingArea<DB, plotters::coord::Shift>,
    cx: i32,
    top: i32,
    bottom: i32,
    up: bool,
    color: RGBColor,
) where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let (w, _h) = area.dim_in_pixel();
    let shaft_halfwidth = ((w as i32) / 10).max(3);
    let head_halfwidth = ((w as i32) / 4).max(8);
    let head_height = (bottom - top) / 3;
    if up {
        // Shaft from near top down to the bottom.
        let shaft_top = top + head_height;
        let _ = area.draw(&Rectangle::new(
            [(cx - shaft_halfwidth, shaft_top), (cx + shaft_halfwidth, bottom)],
            color.filled(),
        ));
        // Triangle head at top.
        let poly = vec![
            (cx, top),
            (cx - head_halfwidth, shaft_top),
            (cx + head_halfwidth, shaft_top),
        ];
        let _ = area.draw(&Polygon::new(poly, color.filled()));
    } else {
        let shaft_bottom = bottom - head_height;
        let _ = area.draw(&Rectangle::new(
            [(cx - shaft_halfwidth, top), (cx + shaft_halfwidth, shaft_bottom)],
            color.filled(),
        ));
        let poly = vec![
            (cx, bottom),
            (cx - head_halfwidth, shaft_bottom),
            (cx + head_halfwidth, shaft_bottom),
        ];
        let _ = area.draw(&Polygon::new(poly, color.filled()));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn icon_order_has_seven() {
        assert_eq!(IconKind::ORDER.len(), 7);
        assert_eq!(IconKind::ORDER[0], IconKind::Left);
        assert_eq!(IconKind::ORDER[6], IconKind::Help);
    }

    #[test]
    fn icon_panel_png_has_magic_and_decodes() {
        let bytes = render_icon_panel_png();
        assert!(bytes.len() > 8, "png too short");
        assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n", "missing PNG magic");
        let img = image::load_from_memory(&bytes).expect("valid PNG");
        assert_eq!(img.width(), ICON_PANEL_WIDTH_PX);
        assert_eq!(img.height(), ICON_PANEL_HEIGHT_PX);
    }

    #[test]
    fn icon_panel_has_non_background_pixels() {
        // Make sure we're actually drawing icons, not just returning a
        // blank grey canvas.
        let bytes = render_icon_panel_png();
        let img = image::load_from_memory(&bytes).unwrap().into_rgb8();
        let mut non_bg = 0usize;
        for p in img.pixels() {
            if p.0 != [0xC0, 0xC0, 0xC0] {
                non_bg += 1;
            }
        }
        assert!(non_bg > 500, "expected at least 500 drawn pixels, got {non_bg}");
    }
}
