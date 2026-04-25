//! Per-sheet visibility from xlsx imports.  See
//! `docs/XLSX_IMPORT_PLAN.md` §3.3.
//!
//! Excel distinguishes:
//! * `Visible`   — the normal state; the sheet appears in the tab bar.
//! * `Hidden`    — the sheet exists, but its tab is hidden.  Users
//!   can unhide it via Excel's right-click menu.
//! * `VeryHidden` — the sheet exists, but it can only be unhidden via
//!   VBA / macros (Excel's UI does not show an unhide path).
//!
//! L123 doesn't yet ship a `/Worksheet Hide` command, so this type
//! is currently round-trip-only: workbooks loaded from xlsx preserve
//! their visibility settings, and the UI skips non-visible sheets in
//! `Ctrl-PgUp/PgDn` navigation.

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq, Hash)]
pub enum SheetState {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

impl SheetState {
    /// True when the sheet should appear in tab navigation (Ctrl-PgUp /
    /// Ctrl-PgDn).  Hidden and VeryHidden are both skipped.
    pub fn is_visible(self) -> bool {
        matches!(self, SheetState::Visible)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_visible() {
        assert_eq!(SheetState::default(), SheetState::Visible);
        assert!(SheetState::Visible.is_visible());
    }

    #[test]
    fn hidden_and_very_hidden_skip_navigation() {
        assert!(!SheetState::Hidden.is_visible());
        assert!(!SheetState::VeryHidden.is_visible());
    }
}
