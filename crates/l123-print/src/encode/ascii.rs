//! Encode a [`PageGrid`] to plain ASCII — the `.prn` format that
//! `/Print File` has emitted since M6. Form-feed (`\x0c`) separates
//! pages; no trailing form-feed after the last page.

use crate::grid::PageGrid;

/// Concatenate `grid` to a single ASCII string.
pub fn to_ascii(grid: &PageGrid) -> String {
    let mut out = String::new();
    let last = grid.pages.len().saturating_sub(1);
    for (i, page) in grid.pages.iter().enumerate() {
        for _ in 0..page.top_blank {
            out.push('\n');
        }
        if let Some(h) = &page.header {
            out.push_str(h);
            out.push('\n');
            out.push('\n');
        }
        for row in &page.rows {
            out.push_str(row);
        }
        if let Some(f) = &page.footer {
            out.push('\n');
            out.push_str(f);
            out.push('\n');
        }
        for _ in 0..page.bottom_blank {
            out.push('\n');
        }
        if i < last {
            out.push('\x0c');
        }
    }
    out
}
