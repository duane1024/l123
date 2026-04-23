//! Static 1-2-3 menu tree.
//!
//! The full menu tree (SPEC §10, docs/MENU.md) is encoded as a
//! compile-time constant. Each node carries a single-letter accelerator,
//! a display name, a one-line help string (shown on control-panel line 3
//! when the item is highlighted at a leaf), and a body.
//!
//! ## Actions
//! [`Action`] names each distinct terminal command. Leaves that are not
//! yet implemented use [`MenuBody::NotImplemented`]; the menu can still
//! be navigated to them, but committing produces a status message rather
//! than a state change. The MVP slice (SPEC §10) enumerates which leaves
//! should graduate from `NotImplemented` to `Action(...)`.

#![allow(clippy::needless_lifetimes)]

/// Single menu item.
#[derive(Debug, Clone, Copy)]
pub struct MenuItem {
    pub letter: char,
    pub name: &'static str,
    pub help: &'static str,
    pub body: MenuBody,
}

#[derive(Debug, Clone, Copy)]
pub enum MenuBody {
    /// A submenu — descending selects one of its items.
    Submenu(&'static [MenuItem]),
    /// A terminal command the interpreter should execute.
    Action(Action),
    /// Placeholder for a leaf that is not yet wired up. The string is a
    /// short identifier shown on line 3 so the user knows the leaf was
    /// reached but is unimplemented.
    NotImplemented(&'static str),
}

/// Commands the interpreter can execute. One variant per distinct leaf;
/// leaves that take arguments (e.g. /Range Format's decimals prompt) build
/// their args via follow-up prompts driven by the UI.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Action {
    /// Close the menu and return to READY. Used for explicit "No" / "Cancel"
    /// leaves in confirm submenus.
    Cancel,
    Quit,
    WorksheetEraseConfirm,
    WorksheetInsertRow,
    WorksheetInsertColumn,
    WorksheetDeleteRow,
    WorksheetDeleteColumn,
    WorksheetColumnSetWidth,
    WorksheetGlobalRecalcAutomatic,
    WorksheetGlobalRecalcManual,
    RangeErase,
    RangeLabelLeft,
    RangeLabelRight,
    RangeLabelCenter,
    RangeNameCreate,
    RangeNameDelete,
    RangeFormatFixed,
    RangeFormatScientific,
    RangeFormatCurrency,
    RangeFormatComma,
    RangeFormatGeneral,
    RangeFormatPercent,
    RangeFormatDate,
    RangeFormatText,
    RangeFormatReset,
    Copy,
    Move,
    FileSave,
    FileRetrieve,
    FileXtractFormulas,
    FileXtractValues,
    FileImportNumbers,
    FileNew,
    FileDir,
    FileListWorksheet,
    FileListActive,
}

/// Resolve a path of letter accelerators from the root menu.  Returns
/// `None` if any letter fails to match. Letters are case-insensitive.
pub fn resolve(path: &[char]) -> Option<&'static MenuItem> {
    let mut items: &[MenuItem] = ROOT;
    let mut last: Option<&MenuItem> = None;
    for &letter in path {
        let item = items
            .iter()
            .find(|m| m.letter.eq_ignore_ascii_case(&letter))?;
        last = Some(item);
        items = match item.body {
            MenuBody::Submenu(sub) => sub,
            _ => return Some(item), // path ended at a leaf
        };
    }
    last
}

/// Children of a submenu item (or empty slice if not a submenu / terminal).
pub fn children(item: &MenuItem) -> &'static [MenuItem] {
    match item.body {
        MenuBody::Submenu(s) => s,
        _ => &[],
    }
}

/// Items visible on the control-panel menu bar when the user has
/// descended `path`. Empty path → ROOT. If `path` ends at a leaf,
/// the leaf's parent level is returned.
pub fn current_level(path: &[char]) -> &'static [MenuItem] {
    if path.is_empty() {
        return ROOT;
    }
    match resolve(path) {
        Some(item) => children(item),
        None => &[],
    }
}

// --------------------------------------------------------------------------
// The tree. Bottom-up so each level's parent can reference its children.
// --------------------------------------------------------------------------

const QUIT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not end 1-2-3 session",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "End 1-2-3 session",
        body: MenuBody::Action(Action::Quit),
    },
];

const WS_INSERT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Insert one or more columns at the cell pointer",
        body: MenuBody::Action(Action::WorksheetInsertColumn),
    },
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Insert one or more rows at the cell pointer",
        body: MenuBody::Action(Action::WorksheetInsertRow),
    },
    MenuItem {
        letter: 'S',
        name: "Sheet",
        help: "Insert one or more worksheets",
        body: MenuBody::NotImplemented("ws-insert-sheet"),
    },
];

const WS_DELETE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Delete one or more columns",
        body: MenuBody::Action(Action::WorksheetDeleteColumn),
    },
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Delete one or more rows",
        body: MenuBody::Action(Action::WorksheetDeleteRow),
    },
    MenuItem {
        letter: 'S',
        name: "Sheet",
        help: "Delete one or more worksheets",
        body: MenuBody::NotImplemented("ws-delete-sheet"),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Remove the current file from memory",
        body: MenuBody::NotImplemented("ws-delete-file"),
    },
];

const WS_COLUMN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set-Width",
        help: "Set the width of the current column",
        body: MenuBody::Action(Action::WorksheetColumnSetWidth),
    },
    MenuItem {
        letter: 'R',
        name: "Reset-Width",
        help: "Reset column width to the global default",
        body: MenuBody::NotImplemented("ws-col-reset"),
    },
    MenuItem {
        letter: 'H',
        name: "Hide",
        help: "Hide columns",
        body: MenuBody::NotImplemented("ws-col-hide"),
    },
    MenuItem {
        letter: 'D',
        name: "Display",
        help: "Redisplay hidden columns",
        body: MenuBody::NotImplemented("ws-col-display"),
    },
    MenuItem {
        letter: 'C',
        name: "Column-Range",
        help: "Set width for a range of columns",
        body: MenuBody::NotImplemented("ws-col-range"),
    },
];

const WS_ERASE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not erase the worksheet",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Erase ALL active files and start fresh",
        body: MenuBody::Action(Action::WorksheetEraseConfirm),
    },
];

const WS_GLOBAL_RECALC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Natural",
        help: "Natural-order recalculation",
        body: MenuBody::NotImplemented("wg-recalc-natural"),
    },
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "Columnwise recalculation order",
        body: MenuBody::NotImplemented("wg-recalc-col"),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "Rowwise recalculation order",
        body: MenuBody::NotImplemented("wg-recalc-row"),
    },
    MenuItem {
        letter: 'A',
        name: "Automatic",
        help: "Automatic recalculation after each entry",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcAutomatic),
    },
    MenuItem {
        letter: 'M',
        name: "Manual",
        help: "Manual recalculation — press F9 to recalc",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcManual),
    },
    MenuItem {
        letter: 'I',
        name: "Iteration",
        help: "Set iteration count (1-50)",
        body: MenuBody::NotImplemented("wg-recalc-iter"),
    },
];

const WS_GLOBAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Set global cell display format",
        body: MenuBody::NotImplemented("wg-format"),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Set global default label prefix",
        body: MenuBody::NotImplemented("wg-label"),
    },
    MenuItem {
        letter: 'C',
        name: "Col-Width",
        help: "Set global default column width",
        body: MenuBody::NotImplemented("wg-col-width"),
    },
    MenuItem {
        letter: 'P',
        name: "Prot",
        help: "Enable/disable worksheet protection",
        body: MenuBody::NotImplemented("wg-prot"),
    },
    MenuItem {
        letter: 'Z',
        name: "Zero",
        help: "Zero-value display: No/Yes/Label",
        body: MenuBody::NotImplemented("wg-zero"),
    },
    MenuItem {
        letter: 'R',
        name: "Recalc",
        help: "Recalculation mode",
        body: MenuBody::Submenu(WS_GLOBAL_RECALC_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Default",
        help: "Default settings (printer, dir, other, ...)",
        body: MenuBody::NotImplemented("wg-default"),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Enable/disable GROUP mode across sheets",
        body: MenuBody::NotImplemented("wg-group"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::NotImplemented("wg-quit"),
    },
];

const WORKSHEET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Global",
        help: "Set worksheet-wide options",
        body: MenuBody::Submenu(WS_GLOBAL_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Insert",
        help: "Insert a row, column, or sheet",
        body: MenuBody::Submenu(WS_INSERT_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Delete a row, column, sheet, or file",
        body: MenuBody::Submenu(WS_DELETE_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Column width and visibility",
        body: MenuBody::Submenu(WS_COLUMN_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Erase all active files",
        body: MenuBody::Submenu(WS_ERASE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Titles",
        help: "Freeze rows and/or columns as titles",
        body: MenuBody::NotImplemented("ws-titles"),
    },
    MenuItem {
        letter: 'W',
        name: "Window",
        help: "Split window into panes",
        body: MenuBody::NotImplemented("ws-window"),
    },
    MenuItem {
        letter: 'S',
        name: "Status",
        help: "Show worksheet status panel",
        body: MenuBody::NotImplemented("ws-status"),
    },
    MenuItem {
        letter: 'P',
        name: "Page",
        help: "Insert a page break",
        body: MenuBody::NotImplemented("ws-page"),
    },
    MenuItem {
        letter: 'H',
        name: "Hide",
        help: "Hide/show entire sheets",
        body: MenuBody::NotImplemented("ws-hide"),
    },
    MenuItem {
        letter: 'L',
        name: "Learn",
        help: "Define / cancel / erase the Learn range",
        body: MenuBody::NotImplemented("ws-learn"),
    },
];

const RANGE_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fixed",
        help: "Fixed number of decimal places",
        body: MenuBody::Action(Action::RangeFormatFixed),
    },
    MenuItem {
        letter: 'S',
        name: "Sci",
        help: "Scientific notation",
        body: MenuBody::Action(Action::RangeFormatScientific),
    },
    MenuItem {
        letter: 'C',
        name: "Currency",
        help: "Currency format with symbol",
        body: MenuBody::Action(Action::RangeFormatCurrency),
    },
    MenuItem {
        letter: ',',
        name: ",",
        help: "Comma-separated (no currency symbol)",
        body: MenuBody::Action(Action::RangeFormatComma),
    },
    MenuItem {
        letter: 'G',
        name: "General",
        help: "General format (default)",
        body: MenuBody::Action(Action::RangeFormatGeneral),
    },
    MenuItem {
        letter: 'P',
        name: "Percent",
        help: "Percent (value × 100) with % sign",
        body: MenuBody::Action(Action::RangeFormatPercent),
    },
    MenuItem {
        letter: 'D',
        name: "Date",
        help: "Date format (select D1..D5)",
        body: MenuBody::Action(Action::RangeFormatDate),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Show formulas instead of values",
        body: MenuBody::Action(Action::RangeFormatText),
    },
    MenuItem {
        letter: 'H',
        name: "Hidden",
        help: "Hide the cell display",
        body: MenuBody::NotImplemented("rf-hidden"),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Revert to global format",
        body: MenuBody::Action(Action::RangeFormatReset),
    },
];

const RANGE_LABEL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Left-align label prefix for range",
        body: MenuBody::Action(Action::RangeLabelLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Right-align label prefix for range",
        body: MenuBody::Action(Action::RangeLabelRight),
    },
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Center label prefix for range",
        body: MenuBody::Action(Action::RangeLabelCenter),
    },
];

const RANGE_NAME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Create a new named range",
        body: MenuBody::Action(Action::RangeNameCreate),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Delete a named range",
        body: MenuBody::Action(Action::RangeNameDelete),
    },
    MenuItem {
        letter: 'L',
        name: "Labels",
        help: "Create range names from adjacent labels",
        body: MenuBody::NotImplemented("rn-labels"),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Delete all range names",
        body: MenuBody::NotImplemented("rn-reset"),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Write a table of range names to the sheet",
        body: MenuBody::NotImplemented("rn-table"),
    },
    MenuItem {
        letter: 'U',
        name: "Undefine",
        help: "Remove a range name but preserve formulas",
        body: MenuBody::NotImplemented("rn-undefine"),
    },
    MenuItem {
        letter: 'N',
        name: "Note",
        help: "Cell notes attached to ranges",
        body: MenuBody::NotImplemented("rn-note"),
    },
];

const RANGE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Set display format for a range",
        body: MenuBody::Submenu(RANGE_FORMAT_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Change label prefix for a range",
        body: MenuBody::Submenu(RANGE_LABEL_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Erase the contents of a range",
        body: MenuBody::Action(Action::RangeErase),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Named ranges",
        body: MenuBody::Submenu(RANGE_NAME_MENU),
    },
    MenuItem {
        letter: 'J',
        name: "Justify",
        help: "Word-wrap long labels into a block",
        body: MenuBody::NotImplemented("r-justify"),
    },
    MenuItem {
        letter: 'P',
        name: "Prot",
        help: "Re-protect a range on a protected sheet",
        body: MenuBody::NotImplemented("r-prot"),
    },
    MenuItem {
        letter: 'U',
        name: "Unprot",
        help: "Mark a range as writable",
        body: MenuBody::NotImplemented("r-unprot"),
    },
    MenuItem {
        letter: 'I',
        name: "Input",
        help: "Form input limited to unprotected cells",
        body: MenuBody::NotImplemented("r-input"),
    },
    MenuItem {
        letter: 'V',
        name: "Value",
        help: "Copy a range converting formulas to values",
        body: MenuBody::NotImplemented("r-value"),
    },
    MenuItem {
        letter: 'T',
        name: "Trans",
        help: "Transpose rows and columns",
        body: MenuBody::NotImplemented("r-trans"),
    },
    MenuItem {
        letter: 'S',
        name: "Search",
        help: "Find / Replace across formulas and labels",
        body: MenuBody::NotImplemented("r-search"),
    },
];

const FILE_LIST_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "List worksheet files in the session directory",
        body: MenuBody::Action(Action::FileListWorksheet),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "List print settings files",
        body: MenuBody::NotImplemented("f-list-print"),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "List graph files",
        body: MenuBody::NotImplemented("f-list-graph"),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "List any file",
        body: MenuBody::NotImplemented("f-list-other"),
    },
    MenuItem {
        letter: 'A',
        name: "Active",
        help: "List currently active files",
        body: MenuBody::Action(Action::FileListActive),
    },
    MenuItem {
        letter: 'L',
        name: "Linked",
        help: "List files linked via formula references",
        body: MenuBody::NotImplemented("f-list-linked"),
    },
];

const FILE_NEW_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Create a new active file before the current one",
        body: MenuBody::Action(Action::FileNew),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Create a new active file after the current one",
        body: MenuBody::Action(Action::FileNew),
    },
];

const FILE_IMPORT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Import each line as a label in one column",
        body: MenuBody::NotImplemented("f-import-text"),
    },
    MenuItem {
        letter: 'N',
        name: "Numbers",
        help: "Parse CSV: numeric tokens as values, quoted strings as labels",
        body: MenuBody::Action(Action::FileImportNumbers),
    },
];

const FILE_XTRACT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Formulas",
        help: "Save the range with formulas intact",
        body: MenuBody::Action(Action::FileXtractFormulas),
    },
    MenuItem {
        letter: 'V',
        name: "Values",
        help: "Save the range with formulas replaced by their values",
        body: MenuBody::Action(Action::FileXtractValues),
    },
];

const FILE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Retrieve",
        help: "Replace all active files with one from disk",
        body: MenuBody::Action(Action::FileRetrieve),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Save all active files",
        body: MenuBody::Action(Action::FileSave),
    },
    MenuItem {
        letter: 'C',
        name: "Combine",
        help: "Merge a file into the current one",
        body: MenuBody::NotImplemented("f-combine"),
    },
    MenuItem {
        letter: 'X',
        name: "Xtract",
        help: "Save a range as a new file",
        body: MenuBody::Submenu(FILE_XTRACT_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Delete a file on disk",
        body: MenuBody::NotImplemented("f-erase"),
    },
    MenuItem {
        letter: 'L',
        name: "List",
        help: "Overlay list of files on disk",
        body: MenuBody::Submenu(FILE_LIST_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Import",
        help: "Import text or delimited numbers",
        body: MenuBody::Submenu(FILE_IMPORT_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Dir",
        help: "Change the session directory",
        body: MenuBody::Action(Action::FileDir),
    },
    MenuItem {
        letter: 'N',
        name: "New",
        help: "Create a new active file",
        body: MenuBody::Submenu(FILE_NEW_MENU),
    },
    MenuItem {
        letter: 'O',
        name: "Open",
        help: "Open another file alongside the current one",
        body: MenuBody::NotImplemented("f-open"),
    },
    MenuItem {
        letter: 'A',
        name: "Admin",
        help: "Reservation, seal, link-refresh",
        body: MenuBody::NotImplemented("f-admin"),
    },
];

const PRINT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Printer",
        help: "Print to printer",
        body: MenuBody::NotImplemented("p-printer"),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Print to .PRN text file",
        body: MenuBody::NotImplemented("p-file"),
    },
    MenuItem {
        letter: 'E',
        name: "Encoded",
        help: "Print to encoded file with printer codes",
        body: MenuBody::NotImplemented("p-encoded"),
    },
    MenuItem {
        letter: 'C',
        name: "Cancel",
        help: "Cancel current print job",
        body: MenuBody::NotImplemented("p-cancel"),
    },
];

const GRAPH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'T',
        name: "Type",
        help: "Select graph type",
        body: MenuBody::NotImplemented("g-type"),
    },
    MenuItem {
        letter: 'X',
        name: "X",
        help: "Set X-axis range",
        body: MenuBody::NotImplemented("g-x"),
    },
    MenuItem {
        letter: 'V',
        name: "View",
        help: "Display the current graph",
        body: MenuBody::NotImplemented("g-view"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::NotImplemented("g-quit"),
    },
];

const DATA_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fill",
        help: "Fill a range with a sequence",
        body: MenuBody::NotImplemented("d-fill"),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "What-if tables (1, 2, 3, Labeled)",
        body: MenuBody::NotImplemented("d-table"),
    },
    MenuItem {
        letter: 'S',
        name: "Sort",
        help: "Sort a range by keys",
        body: MenuBody::NotImplemented("d-sort"),
    },
    MenuItem {
        letter: 'Q',
        name: "Query",
        help: "Database query (find, extract, ...)",
        body: MenuBody::NotImplemented("d-query"),
    },
];

const ADDIN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Load",
        help: "Load an add-in into memory",
        body: MenuBody::NotImplemented("ai-load"),
    },
    MenuItem {
        letter: 'R',
        name: "Remove",
        help: "Unload an add-in",
        body: MenuBody::NotImplemented("ai-remove"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::NotImplemented("ai-quit"),
    },
];

/// Top-level slash menu.
pub const ROOT: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "Global settings, insert/delete, columns, titles...",
        body: MenuBody::Submenu(WORKSHEET_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Range",
        help: "Format, label, erase, name, justify, protect...",
        body: MenuBody::Submenu(RANGE_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Copy",
        help: "Copy a range to another location",
        body: MenuBody::Action(Action::Copy),
    },
    MenuItem {
        letter: 'M',
        name: "Move",
        help: "Move a range to another location",
        body: MenuBody::Action(Action::Move),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Retrieve, save, combine, import, export, ...",
        body: MenuBody::Submenu(FILE_MENU),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "Print to printer, file, or encoded file",
        body: MenuBody::Submenu(PRINT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Create and configure graphs",
        body: MenuBody::Submenu(GRAPH_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Data",
        help: "Fill, sort, query, table, regression, ...",
        body: MenuBody::Submenu(DATA_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "System",
        help: "Suspend to OS shell",
        body: MenuBody::NotImplemented("system"),
    },
    MenuItem {
        letter: 'A',
        name: "Add-In",
        help: "Load and invoke add-in programs",
        body: MenuBody::Submenu(ADDIN_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "End the 1-2-3 session",
        body: MenuBody::Submenu(QUIT_MENU),
    },
];

#[cfg(test)]
mod tests {
    use super::*;

    /// Walk every submenu and ensure letter accelerators are unique at
    /// each level (case-insensitive).
    fn assert_unique_letters(items: &[MenuItem], path: &str) {
        let mut seen: Vec<char> = Vec::new();
        for m in items {
            let up = m.letter.to_ascii_uppercase();
            assert!(
                !seen.contains(&up),
                "duplicate letter {up:?} at path {path:?}"
            );
            seen.push(up);
            if let MenuBody::Submenu(sub) = m.body {
                let sub_path = format!("{path}/{}", m.name);
                assert_unique_letters(sub, &sub_path);
            }
        }
    }

    #[test]
    fn letters_are_unique_at_every_level() {
        assert_unique_letters(ROOT, "");
    }

    #[test]
    fn all_eleven_top_level_items_present() {
        let names: Vec<&str> = ROOT.iter().map(|m| m.name).collect();
        assert_eq!(
            names,
            vec![
                "Worksheet", "Range", "Copy", "Move", "File", "Print", "Graph", "Data",
                "System", "Add-In", "Quit"
            ]
        );
    }

    #[test]
    fn resolve_quit_yes_is_action() {
        let node = resolve(&['Q', 'Y']).unwrap();
        assert!(matches!(node.body, MenuBody::Action(Action::Quit)));
    }

    #[test]
    fn resolve_is_case_insensitive() {
        let a = resolve(&['q', 'y']).unwrap();
        let b = resolve(&['Q', 'Y']).unwrap();
        assert_eq!(a.letter, b.letter);
    }

    #[test]
    fn resolve_nonexistent_returns_none() {
        assert!(resolve(&['Z']).is_none());
        assert!(resolve(&['Q', 'Z']).is_none());
    }

    #[test]
    fn resolve_ws_insert_row() {
        let node = resolve(&['W', 'I', 'R']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::WorksheetInsertRow)
        ));
    }

    #[test]
    fn resolve_ws_delete_column() {
        let node = resolve(&['W', 'D', 'C']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::WorksheetDeleteColumn)
        ));
    }

    #[test]
    fn resolve_range_name_create() {
        let node = resolve(&['R', 'N', 'C']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeNameCreate)
        ));
    }

    #[test]
    fn resolve_range_format_currency() {
        let node = resolve(&['R', 'F', 'C']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatCurrency)
        ));
    }

    #[test]
    fn root_names_all_start_with_capital() {
        for m in ROOT {
            let c = m.name.chars().next().unwrap();
            assert!(c.is_ascii_uppercase(), "{}", m.name);
        }
    }
}
