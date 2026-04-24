//! Encode a [`PageGrid`] to PDF.
//!
//! Uses [`pdf_writer`] (pure Rust, tiny dep graph) and the built-in
//! PDF type-1 Courier font — no font embedding, no external data.
//! One PDF page per `PageGrid` page; text is placed at fixed x/y
//! derived from the configured characters-per-inch and lines-per-inch,
//! matching classic line-printer geometry.

use pdf_writer::{Content, Name, Pdf, Rect, Ref, Str};

use crate::grid::PageGrid;

/// Physical page size for the PDF output.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageSize {
    /// 8.5" × 11" (612 × 792 pt).
    #[default]
    Letter,
    /// 210 mm × 297 mm (595 × 842 pt).
    A4,
}

impl PageSize {
    /// Page width in PDF points (1/72").
    pub fn width_pt(self) -> f32 {
        match self {
            PageSize::Letter => 612.0,
            PageSize::A4 => 595.0,
        }
    }
    /// Page height in PDF points.
    pub fn height_pt(self) -> f32 {
        match self {
            PageSize::Letter => 792.0,
            PageSize::A4 => 842.0,
        }
    }
}

#[derive(Debug, Clone)]
pub struct PdfOptions {
    pub page_size: PageSize,
    /// Characters per inch. 10 matches classic Courier / greenbar.
    pub cpi: u8,
    /// Lines per inch. 6 matches classic Courier / greenbar.
    pub lpi: u8,
}

impl Default for PdfOptions {
    fn default() -> Self {
        Self { page_size: PageSize::default(), cpi: 10, lpi: 6 }
    }
}

impl PdfOptions {
    /// Width of one column of monospace text, in PDF points.
    fn char_width(&self) -> f32 {
        72.0 / self.cpi.max(1) as f32
    }
    /// Height of one line, in PDF points (leading).
    fn line_height(&self) -> f32 {
        72.0 / self.lpi.max(1) as f32
    }
    /// Font size for Courier. Chosen so each glyph's 600/1000 em
    /// advance matches [`char_width`]: `char_width / 0.6`.
    fn font_size(&self) -> f32 {
        self.char_width() / 0.6
    }
}

/// Encode `grid` to a PDF byte buffer using Courier.
pub fn to_pdf(grid: &PageGrid, opts: &PdfOptions) -> Vec<u8> {
    let page_w = opts.page_size.width_pt();
    let page_h = opts.page_size.height_pt();
    let font_size = opts.font_size();
    let line_h = opts.line_height();
    let font_name = Name(b"F1");

    let mut pdf = Pdf::new();

    let catalog_id = Ref::new(1);
    let page_tree_id = Ref::new(2);
    let font_id = Ref::new(3);

    pdf.catalog(catalog_id).pages(page_tree_id);
    pdf.type1_font(font_id).base_font(Name(b"Courier"));

    // Allocate two consecutive refs per page (page object + its
    // content stream). Starting at 4 keeps us clear of the fixed
    // catalog/page-tree/font ids above.
    let page_refs: Vec<(Ref, Ref)> = (0..grid.pages.len())
        .map(|i| {
            let base = 4 + (2 * i) as i32;
            (Ref::new(base), Ref::new(base + 1))
        })
        .collect();

    pdf.pages(page_tree_id)
        .kids(page_refs.iter().map(|(p, _)| *p))
        .count(grid.pages.len() as i32);

    for (page, (page_id, content_id)) in grid.pages.iter().zip(&page_refs) {
        {
            let mut p = pdf.page(*page_id);
            p.parent(page_tree_id);
            p.media_box(Rect::new(0.0, 0.0, page_w, page_h));
            p.contents(*content_id);
            p.resources().fonts().pair(font_name, font_id);
        }

        // Lay out from the top of the page down. Text baselines sit
        // at y = top - (i+1) * line_h; the `+1` offset gives the
        // first line a full line's worth of headroom below the
        // `top_blank` gap.
        let top = page_h;
        let mut row_index: usize = page.top_blank as usize;

        let mut content = Content::new();
        content.begin_text();
        content.set_font(font_name, font_size);

        // Prime the text matrix at the first baseline.
        let first_baseline = top - (row_index as f32 + 1.0) * line_h;
        content.next_line(0.0, first_baseline);
        let mut last_baseline = first_baseline;

        let advance_to_row = |content: &mut Content, row_index: usize, last: &mut f32| {
            let target = top - (row_index as f32 + 1.0) * line_h;
            let dy = target - *last;
            if dy != 0.0 {
                content.next_line(0.0, dy);
                *last = target;
            }
        };

        if let Some(h) = &page.header {
            advance_to_row(&mut content, row_index, &mut last_baseline);
            content.show(Str(h.as_bytes()));
            row_index += 2; // header + blank line separator
        }
        for row in &page.rows {
            advance_to_row(&mut content, row_index, &mut last_baseline);
            content.show(Str(row.trim_end_matches('\n').as_bytes()));
            row_index += 1;
        }
        if let Some(f) = &page.footer {
            row_index += 1; // blank line separator before footer
            advance_to_row(&mut content, row_index, &mut last_baseline);
            content.show(Str(f.as_bytes()));
        }

        content.end_text();
        pdf.stream(*content_id, &content.finish());
    }

    pdf.finish()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::{Page, PageGrid};

    fn grid_with(rows: Vec<&str>) -> PageGrid {
        PageGrid {
            pages: vec![Page {
                number: 1,
                header: None,
                footer: None,
                rows: rows.into_iter().map(|s| format!("{s}\n")).collect(),
                top_blank: 0,
                bottom_blank: 0,
            }],
            page_width: 10,
        }
    }

    fn multi_page_grid(n: usize) -> PageGrid {
        PageGrid {
            pages: (0..n)
                .map(|i| Page {
                    number: (i + 1) as u32,
                    header: None,
                    footer: None,
                    rows: vec![format!("page{}\n", i + 1)],
                    top_blank: 0,
                    bottom_blank: 0,
                })
                .collect(),
            page_width: 10,
        }
    }

    #[test]
    fn starts_with_pdf_magic() {
        let bytes = to_pdf(&grid_with(vec!["hi"]), &PdfOptions::default());
        assert!(bytes.starts_with(b"%PDF-"), "got {:?}", &bytes[..8]);
    }

    #[test]
    fn embeds_row_text_verbatim() {
        let bytes = to_pdf(&grid_with(vec!["HelloWorld"]), &PdfOptions::default());
        assert!(
            bytes.windows(10).any(|w| w == b"HelloWorld"),
            "row text not found in PDF byte stream",
        );
    }

    #[test]
    fn emits_one_pdf_page_per_grid_page() {
        let bytes = to_pdf(&multi_page_grid(3), &PdfOptions::default());
        // pdf-writer emits `/Count 3` on the pages tree object for a
        // three-page document.
        assert!(
            bytes.windows(8).any(|w| w == b"/Count 3"),
            "expected `/Count 3` in pages tree, not found",
        );
    }

    #[test]
    fn references_courier_font() {
        let bytes = to_pdf(&grid_with(vec!["x"]), &PdfOptions::default());
        assert!(
            bytes.windows(7).any(|w| w == b"Courier"),
            "Courier base font not referenced",
        );
    }

    #[test]
    fn header_and_footer_are_placed() {
        let grid = PageGrid {
            pages: vec![Page {
                number: 1,
                header: Some("TOPBANNER".into()),
                footer: Some("BOTTOMLINE".into()),
                rows: vec!["middle\n".into()],
                top_blank: 0,
                bottom_blank: 0,
            }],
            page_width: 10,
        };
        let bytes = to_pdf(&grid, &PdfOptions::default());
        assert!(bytes.windows(9).any(|w| w == b"TOPBANNER"));
        assert!(bytes.windows(10).any(|w| w == b"BOTTOMLINE"));
        assert!(bytes.windows(6).any(|w| w == b"middle"));
    }

    #[test]
    fn page_size_points_are_correct() {
        assert_eq!(PageSize::Letter.width_pt(), 612.0);
        assert_eq!(PageSize::Letter.height_pt(), 792.0);
        assert_eq!(PageSize::A4.width_pt(), 595.0);
        assert_eq!(PageSize::A4.height_pt(), 842.0);
    }

    #[test]
    fn default_cpi_10_gives_12pt_courier() {
        // Classic greenbar: 10 CPI Courier at 12pt.
        let opts = PdfOptions::default();
        assert_eq!(opts.cpi, 10);
        assert!((opts.font_size() - 12.0).abs() < 0.01);
        assert!((opts.line_height() - 12.0).abs() < 0.01);
    }
}
