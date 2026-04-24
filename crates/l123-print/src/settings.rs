//! `/Print` options as a plain-data struct, consumed by
//! [`crate::render::render`].

/// How cell contents are rendered into the print file.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PrintContentMode {
    /// Render each cell as it appears on the screen.
    #[default]
    AsDisplayed,
    /// Render formula cells as their source expression; labels and
    /// constants unchanged.
    CellFormulas,
}

/// Whether page decorations (headers, footers, …) wrap the output.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum PrintFormatMode {
    #[default]
    Formatted,
    Unformatted,
}

/// Pre-resolved snapshot of the `/Print` session fields the renderer
/// needs. Built at Go-time so the renderer doesn't have to re-borrow
/// UI state.
#[derive(Debug, Clone)]
pub struct PrintSettings {
    pub header: String,
    pub footer: String,
    pub content_mode: PrintContentMode,
    pub format_mode: PrintFormatMode,
    pub margin_left: u16,
    pub margin_right: u16,
    pub margin_top: u16,
    pub margin_bottom: u16,
    pub pg_length: u16,
    /// Starting page number for the first page of this Go.
    pub start_page: u32,
}

impl Default for PrintSettings {
    fn default() -> Self {
        Self {
            header: String::new(),
            footer: String::new(),
            content_mode: PrintContentMode::AsDisplayed,
            format_mode: PrintFormatMode::Formatted,
            margin_left: 0,
            margin_right: 0,
            margin_top: 0,
            margin_bottom: 0,
            pg_length: 0,
            start_page: 1,
        }
    }
}
