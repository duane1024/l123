//! The intermediate page model that encoders consume.
//!
//! A `PageGrid` holds already-rendered rows (label prefixes, numeric
//! alignment, overflow-to-asterisks, right-margin truncation, and
//! left-margin padding are all baked in) plus pre-formatted header /
//! footer strings. Encoders walk pages and emit bytes; they do not
//! re-layout.

/// Complete print job split into pages.
#[derive(Debug, Clone)]
pub struct PageGrid {
    pub pages: Vec<Page>,
    /// Sum of printed column widths, used for three-part header/footer
    /// centering. Capped at 1 so centering math never divides by zero.
    pub page_width: u16,
}

/// One page worth of output.
#[derive(Debug, Clone)]
pub struct Page {
    /// Logical page number (starts at `settings.start_page`).
    pub number: u32,
    /// Post-substitute, post-three-part, post-left-pad header line
    /// (without trailing newline). `None` when unformatted mode or no
    /// header was configured.
    pub header: Option<String>,
    /// As above for the footer.
    pub footer: Option<String>,
    /// Each row already has left-margin padding applied and a trailing
    /// `\n`. Ready to concatenate.
    pub rows: Vec<String>,
    /// Blank lines before the header (i.e. `margin_top`).
    pub top_blank: u16,
    /// Blank lines after the footer (i.e. `margin_bottom`).
    pub bottom_blank: u16,
}
