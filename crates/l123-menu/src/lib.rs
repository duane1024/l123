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
    WorksheetInsertSheetBefore,
    WorksheetInsertSheetAfter,
    WorksheetDeleteRow,
    WorksheetDeleteColumn,
    WorksheetColumnSetWidth,
    WorksheetColumnResetWidth,
    WorksheetColumnRangeSetWidth,
    WorksheetColumnRangeResetWidth,
    WorksheetColumnHide,
    WorksheetColumnDisplay,
    /// `/Worksheet Status` — full-screen overlay of recalculation
    /// mode, cell-display defaults, and environment probes.
    WorksheetStatus,
    WorksheetGlobalColWidth,
    WorksheetGlobalLabelLeft,
    WorksheetGlobalLabelRight,
    WorksheetGlobalLabelCenter,
    WorksheetGlobalRecalcAutomatic,
    WorksheetGlobalRecalcManual,
    /// `/Worksheet Global Recalc Natural` — dependency-order recalc.
    WorksheetGlobalRecalcNatural,
    /// `/Worksheet Global Recalc Columnwise`.
    WorksheetGlobalRecalcColumnwise,
    /// `/Worksheet Global Recalc Rowwise`.
    WorksheetGlobalRecalcRowwise,
    /// `/Worksheet Global Recalc Iteration` — prompts for 1..=50.
    WorksheetGlobalRecalcIteration,
    /// `/Worksheet Global Zero No` — show numeric zeros (default).
    WorksheetGlobalZeroNo,
    /// `/Worksheet Global Zero Yes` — blank numeric-zero cells.
    WorksheetGlobalZeroYes,
    /// `/Worksheet Global Protection Enable`.
    WorksheetGlobalProtectionEnable,
    /// `/Worksheet Global Protection Disable`.
    WorksheetGlobalProtectionDisable,
    WorksheetGlobalGroupEnable,
    WorksheetGlobalGroupDisable,
    WorksheetGlobalDefaultOtherUndoEnable,
    WorksheetGlobalDefaultOtherUndoDisable,
    /// `/Worksheet Global Default Other Beep Enable` — turn on the
    /// soft terminal bell that fires when the pointer hits an edge
    /// of the sheet or an equivalent invalid operation is attempted.
    WorksheetGlobalDefaultOtherBeepEnable,
    /// `/Worksheet Global Default Other Beep Disable`.
    WorksheetGlobalDefaultOtherBeepDisable,
    /// `/Worksheet Global Default Status` — full-screen overlay
    /// showing every persisted default (printer, dirs, autoexec, ext,
    /// graph). Read-only; any key dismisses.
    WorksheetGlobalDefaultStatus,
    /// `/Worksheet Global Default Update` — write the current defaults
    /// back to the L123.CNF config file so the next session starts with
    /// the same settings.
    WorksheetGlobalDefaultUpdate,
    /// `/Worksheet Global Default Dir` — prompt for the default
    /// session directory (used at next launch).
    WorksheetGlobalDefaultDir,
    /// `/Worksheet Global Default Temp` — prompt for the temporary-
    /// file directory.
    WorksheetGlobalDefaultTemp,
    /// `/Worksheet Global Default Autoexec Yes|No` — toggle the
    /// auto-run-`\0`-on-retrieve behavior.
    WorksheetGlobalDefaultAutoexecYes,
    WorksheetGlobalDefaultAutoexecNo,
    /// `/Worksheet Global Default Ext Save` — prompt for the default
    /// file extension used when saving.
    WorksheetGlobalDefaultExtSave,
    /// `/Worksheet Global Default Ext List` — prompt for the default
    /// file-extension filter used by /File List.
    WorksheetGlobalDefaultExtList,
    /// `/Worksheet Global Default Printer Interface` — numeric (1..=9)
    /// printer-interface index.
    WorksheetGlobalDefaultPrinterInterface,
    WorksheetGlobalDefaultPrinterAutoLfYes,
    WorksheetGlobalDefaultPrinterAutoLfNo,
    WorksheetGlobalDefaultPrinterMarginLeft,
    WorksheetGlobalDefaultPrinterMarginRight,
    WorksheetGlobalDefaultPrinterMarginTop,
    WorksheetGlobalDefaultPrinterMarginBottom,
    WorksheetGlobalDefaultPrinterPgLength,
    WorksheetGlobalDefaultPrinterWaitYes,
    WorksheetGlobalDefaultPrinterWaitNo,
    WorksheetGlobalDefaultPrinterSetup,
    WorksheetGlobalDefaultPrinterName,
    WorksheetGlobalDefaultPrinterQuit,
    /// `/Worksheet Global Default Graph Group Columnwise|Rowwise` —
    /// default orientation used by /Graph Group auto-graph.
    WorksheetGlobalDefaultGraphGroupColumnwise,
    WorksheetGlobalDefaultGraphGroupRowwise,
    /// `/Worksheet Global Default Graph Save Cgm|Pic` — default file
    /// type written by /Graph Save when no extension is supplied.
    WorksheetGlobalDefaultGraphSaveCgm,
    WorksheetGlobalDefaultGraphSavePic,
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
    RangeFormatDateDmy,
    RangeFormatDateDm,
    RangeFormatDateMy,
    RangeFormatDateLongIntl,
    RangeFormatDateShortIntl,
    RangeFormatText,
    RangeFormatReset,
    /// `/Worksheet Global Format` leaves — set the workbook-wide default
    /// cell format that cells without a per-cell `/RF` override inherit.
    /// Mirrors the [`RangeFormat*`](Action::RangeFormatFixed) family;
    /// each takes the same kind of argument (decimals where applicable)
    /// and applies immediately — no POINT step, since the global is a
    /// single-target setting.
    WorksheetGlobalFormatFixed,
    WorksheetGlobalFormatScientific,
    WorksheetGlobalFormatCurrency,
    WorksheetGlobalFormatComma,
    WorksheetGlobalFormatGeneral,
    WorksheetGlobalFormatPercent,
    WorksheetGlobalFormatDateDmy,
    WorksheetGlobalFormatDateDm,
    WorksheetGlobalFormatDateMy,
    WorksheetGlobalFormatDateLongIntl,
    WorksheetGlobalFormatDateShortIntl,
    WorksheetGlobalFormatText,
    WorksheetGlobalFormatReset,
    Copy,
    Move,
    FileSave,
    FileRetrieve,
    FileXtractFormulas,
    FileXtractValues,
    FileImportNumbers,
    FileNew,
    FileOpenBefore,
    FileOpenAfter,
    PrintFile,
    /// `/Print Printer` — start a printer session. Shares the session
    /// submenu (`PRINT_FILE_MENU`) with `/Print File`; Go branches on
    /// the session destination.
    PrintPrinter,
    /// `/Print Encoded` — write setup-string + ASCII page bytes to a
    /// `.ENC` file. Shares the session submenu with `/Print File` /
    /// `/Print Printer`; Go writes raw printer-ready bytes.
    PrintEncoded,
    /// `/Print Cancel` — drop any active [`PrintSession`] and return
    /// to READY. The "I changed my mind" exit at the top-level Print
    /// menu (before a destination is chosen).
    PrintCancel,
    PrintSessionRange,
    PrintSessionGo,
    PrintSessionQuit,
    PrintSessionAlign,
    PrintSessionClear,
    PrintSessionOptionsHeader,
    PrintSessionOptionsFooter,
    PrintSessionOptionsSetup,
    PrintSessionOptionsQuit,
    PrintSessionOptionsOtherAsDisplayed,
    PrintSessionOptionsOtherCellFormulas,
    PrintSessionOptionsOtherFormatted,
    PrintSessionOptionsOtherUnformatted,
    PrintSessionOptionsMarginLeft,
    PrintSessionOptionsMarginRight,
    PrintSessionOptionsMarginTop,
    PrintSessionOptionsMarginBottom,
    PrintSessionOptionsMarginsQuit,
    PrintSessionOptionsPgLength,
    /// `/Print Options Advanced Device` — set the CUPS queue name
    /// passed to `lp -d`. Stored on the active [`PrintSession`] and
    /// applied at Go-time when the destination is `Printer`.
    PrintSessionOptionsAdvancedDevice,
    /// `/Print Options Advanced Quit` — return to the Options
    /// submenu without changing settings.
    PrintSessionOptionsAdvancedQuit,
    RangeSearchFormulas,
    RangeSearchLabels,
    RangeSearchBoth,
    RangeSearchFind,
    RangeSearchReplace,
    FileDir,
    FileListWorksheet,
    FileListActive,
    GraphTypeLine,
    GraphTypeBar,
    GraphTypeXY,
    GraphTypeStack,
    GraphTypePie,
    GraphTypeHLCO,
    GraphTypeMixed,
    GraphX,
    GraphA,
    GraphB,
    GraphC,
    GraphD,
    GraphE,
    GraphF,
    /// `/Graph Reset Graph` — clear every range and restore the
    /// default type.
    GraphResetGraph,
    /// `/Graph View` — full-screen graph display (same as F10).
    GraphView,
    /// `/Graph Save` — prompt for a filename and write the graph to
    /// disk (SVG; a `.cgm` extension is preserved but the file body
    /// is still SVG, per project convention).
    GraphSave,
    /// `/Graph Quit` — close the `/Graph` menu back to READY.
    GraphQuit,

    // ---- WYSIWYG (`:`) colon-menu commands -----------------------------
    /// `:Format Bold Set` — apply bold to a range.
    FormatBoldSet,
    /// `:Format Bold Clear` — remove bold from a range.
    FormatBoldClear,
    /// `:Format Italic Set`.
    FormatItalicSet,
    /// `:Format Italic Clear`.
    FormatItalicClear,
    /// `:Format Underline Set`.
    FormatUnderlineSet,
    /// `:Format Underline Clear`.
    FormatUnderlineClear,
    /// `:Format Reset` — clear bold + italic + underline on a range.
    FormatReset,
}

/// Resolve a path of letter accelerators from the root menu.  Returns
/// `None` if any letter fails to match. Letters are case-insensitive.
pub fn resolve(path: &[char]) -> Option<&'static MenuItem> {
    resolve_within(ROOT, path)
}

/// Like [`resolve`] but starts from an arbitrary menu slice — used by
/// nested menus (e.g. the `/Print File` submenu rooted at
/// [`PRINT_FILE_MENU`]).
pub fn resolve_within(root: &'static [MenuItem], path: &[char]) -> Option<&'static MenuItem> {
    let mut items: &[MenuItem] = root;
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
    current_level_within(ROOT, path)
}

/// Nested variant of [`current_level`]: starts from `root` rather
/// than [`ROOT`].
pub fn current_level_within(root: &'static [MenuItem], path: &[char]) -> &'static [MenuItem] {
    if path.is_empty() {
        return root;
    }
    match resolve_within(root, path) {
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
        help: "Do not end l123 session",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "End l123 session",
        body: MenuBody::Action(Action::Quit),
    },
];

const WS_INSERT_SHEET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Insert a new worksheet before the current one",
        body: MenuBody::Action(Action::WorksheetInsertSheetBefore),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Insert a new worksheet after the current one",
        body: MenuBody::Action(Action::WorksheetInsertSheetAfter),
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
        help: "Insert a new worksheet before or after the current one",
        body: MenuBody::Submenu(WS_INSERT_SHEET_MENU),
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
        body: MenuBody::Action(Action::WorksheetColumnResetWidth),
    },
    MenuItem {
        letter: 'H',
        name: "Hide",
        help: "Hide columns",
        body: MenuBody::Action(Action::WorksheetColumnHide),
    },
    MenuItem {
        letter: 'D',
        name: "Display",
        help: "Redisplay hidden columns",
        body: MenuBody::Action(Action::WorksheetColumnDisplay),
    },
    MenuItem {
        letter: 'C',
        name: "Column-Range",
        help: "Set/reset width for a range of columns",
        body: MenuBody::Submenu(WS_COLUMN_RANGE_MENU),
    },
];

const WS_COLUMN_RANGE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set-Width",
        help: "Set a new width for each column in a range",
        body: MenuBody::Action(Action::WorksheetColumnRangeSetWidth),
    },
    MenuItem {
        letter: 'R',
        name: "Reset-Width",
        help: "Reset each column in a range to the global default",
        body: MenuBody::Action(Action::WorksheetColumnRangeResetWidth),
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
        body: MenuBody::Action(Action::WorksheetGlobalRecalcNatural),
    },
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "Columnwise recalculation order",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcColumnwise),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "Rowwise recalculation order",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcRowwise),
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
        body: MenuBody::Action(Action::WorksheetGlobalRecalcIteration),
    },
];

const WG_PROT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable global worksheet protection",
        body: MenuBody::Action(Action::WorksheetGlobalProtectionEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable global worksheet protection",
        body: MenuBody::Action(Action::WorksheetGlobalProtectionDisable),
    },
];

const WG_ZERO_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Show numeric zero values (default)",
        body: MenuBody::Action(Action::WorksheetGlobalZeroNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Hide numeric zero values",
        body: MenuBody::Action(Action::WorksheetGlobalZeroYes),
    },
];

const WG_DEFAULT_OTHER_UNDO_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable the undo journal (Alt-F4 reverts mutating commands)",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable the undo journal",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoDisable),
    },
];

const WG_DEFAULT_OTHER_BEEP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable the soft terminal bell on edge collisions",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable the soft terminal bell on edge collisions",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepDisable),
    },
];

const WG_DEFAULT_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "International",
        help: "Locale-specific punctuation, dates, and currency",
        body: MenuBody::NotImplemented("wgdo-intl"),
    },
    MenuItem {
        letter: 'H',
        name: "Help",
        help: "Help mode: Instant / Removable",
        body: MenuBody::NotImplemented("wgdo-help"),
    },
    MenuItem {
        letter: 'C',
        name: "Clock",
        help: "Clock display: Standard / International / None / Filename",
        body: MenuBody::NotImplemented("wgdo-clock"),
    },
    MenuItem {
        letter: 'U',
        name: "Undo",
        help: "Enable/disable the undo journal",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_UNDO_MENU),
    },
    MenuItem {
        letter: 'B',
        name: "Beep",
        help: "Enable/disable the soft error beep on edge collisions",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_BEEP_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Add-In",
        help: "Add-in auto-load settings",
        body: MenuBody::NotImplemented("wgdo-addin"),
    },
    MenuItem {
        letter: 'E',
        name: "Expanded-Memory",
        help: "Expanded-memory options (legacy DOS)",
        body: MenuBody::NotImplemented("wgdo-ems"),
    },
];

const WG_DEFAULT_PRINTER_AUTOLF_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Send a line-feed after every carriage return",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterAutoLfYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not auto-feed after a carriage return",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterAutoLfNo),
    },
];

const WG_DEFAULT_PRINTER_WAIT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Pause between pages so single-sheet feeders can be loaded",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterWaitYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Print without pausing between pages",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterWaitNo),
    },
];

const WG_DEFAULT_PRINTER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "Interface",
        help: "Printer interface number (1..9)",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterInterface),
    },
    MenuItem {
        letter: 'A',
        name: "AutoLf",
        help: "Auto line-feed after carriage return",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_AUTOLF_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Default left margin",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Default right margin",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginRight),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Default top margin",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginTop),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Default bottom margin",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginBottom),
    },
    MenuItem {
        letter: 'P',
        name: "Pg-Length",
        help: "Default page length",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterPgLength),
    },
    MenuItem {
        letter: 'W',
        name: "Wait",
        help: "Pause between pages",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_WAIT_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Setup",
        help: "Default printer setup string",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterSetup),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Default printer queue name",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterName),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Default menu",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterQuit),
    },
];

const WG_DEFAULT_AUTOEXEC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Run \\0 macro automatically when retrieving a file",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultAutoexecYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not run \\0 macro on retrieve",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultAutoexecNo),
    },
];

const WG_DEFAULT_EXT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Default extension applied when saving",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultExtSave),
    },
    MenuItem {
        letter: 'L',
        name: "List",
        help: "Default extension filter for /File List & /File Retrieve",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultExtList),
    },
];

const WG_DEFAULT_GRAPH_GROUP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "Auto-graph reads ranges columnwise",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphGroupColumnwise),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "Auto-graph reads ranges rowwise",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphGroupRowwise),
    },
];

const WG_DEFAULT_GRAPH_SAVE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Cgm",
        help: "Default /Graph Save format: CGM",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphSaveCgm),
    },
    MenuItem {
        letter: 'P',
        name: "Pic",
        help: "Default /Graph Save format: PIC",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphSavePic),
    },
];

const WG_DEFAULT_GRAPH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Default auto-graph orientation (Columnwise/Rowwise)",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_GROUP_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Default /Graph Save format (Cgm/Pic)",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_SAVE_MENU),
    },
];

const WG_DEFAULT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Printer",
        help: "Default printer settings",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Dir",
        help: "Default directory",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultDir),
    },
    MenuItem {
        letter: 'S',
        name: "Status",
        help: "Display global default settings",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultStatus),
    },
    MenuItem {
        letter: 'U',
        name: "Update",
        help: "Persist defaults to the config file",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultUpdate),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Miscellaneous defaults (Undo, International, Clock, …)",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Autoexec",
        help: "Run autoexec macro (\\0) on file retrieve",
        body: MenuBody::Submenu(WG_DEFAULT_AUTOEXEC_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Ext",
        help: "Default file extensions",
        body: MenuBody::Submenu(WG_DEFAULT_EXT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Default graph settings",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Temp",
        help: "Temporary file directory",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultTemp),
    },
];

const WG_LABEL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Default to left-aligned label prefix (')",
        body: MenuBody::Action(Action::WorksheetGlobalLabelLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Default to right-aligned label prefix (\")",
        body: MenuBody::Action(Action::WorksheetGlobalLabelRight),
    },
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Default to centered label prefix (^)",
        body: MenuBody::Action(Action::WorksheetGlobalLabelCenter),
    },
];

const WG_GROUP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable GROUP: format/column/row ops propagate across sheets",
        body: MenuBody::Action(Action::WorksheetGlobalGroupEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable GROUP: commands affect only the current sheet",
        body: MenuBody::Action(Action::WorksheetGlobalGroupDisable),
    },
];

const WS_GLOBAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Set global cell display format",
        body: MenuBody::Submenu(WG_FORMAT_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Set global default label prefix",
        body: MenuBody::Submenu(WG_LABEL_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Col-Width",
        help: "Set global default column width",
        body: MenuBody::Action(Action::WorksheetGlobalColWidth),
    },
    MenuItem {
        letter: 'P',
        name: "Prot",
        help: "Enable/disable worksheet protection",
        body: MenuBody::Submenu(WG_PROT_MENU),
    },
    MenuItem {
        letter: 'Z',
        name: "Zero",
        help: "Zero-value display: No/Yes",
        body: MenuBody::Submenu(WG_ZERO_MENU),
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
        body: MenuBody::Submenu(WG_DEFAULT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Enable/disable GROUP mode across sheets",
        body: MenuBody::Submenu(WG_GROUP_MENU),
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
        body: MenuBody::Action(Action::WorksheetStatus),
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
        help: "Date format (select D1..D5, Time for D6..D9)",
        body: MenuBody::Submenu(RANGE_FORMAT_DATE_MENU),
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

const RANGE_FORMAT_DATE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "DD-MMM-YY",
        body: MenuBody::Action(Action::RangeFormatDateDmy),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "DD-MMM",
        body: MenuBody::Action(Action::RangeFormatDateDm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "MMM-YY",
        body: MenuBody::Action(Action::RangeFormatDateMy),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Long international (MM/DD/YY)",
        body: MenuBody::Action(Action::RangeFormatDateLongIntl),
    },
    MenuItem {
        letter: '5',
        name: "5",
        help: "Short international (MM/DD)",
        body: MenuBody::Action(Action::RangeFormatDateShortIntl),
    },
    MenuItem {
        letter: 'T',
        name: "Time",
        help: "Time format (D6..D9)",
        body: MenuBody::NotImplemented("rf-date-time"),
    },
];

const WG_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fixed",
        help: "Default to fixed number of decimal places",
        body: MenuBody::Action(Action::WorksheetGlobalFormatFixed),
    },
    MenuItem {
        letter: 'S',
        name: "Sci",
        help: "Default to scientific notation",
        body: MenuBody::Action(Action::WorksheetGlobalFormatScientific),
    },
    MenuItem {
        letter: 'C',
        name: "Currency",
        help: "Default to currency format with symbol",
        body: MenuBody::Action(Action::WorksheetGlobalFormatCurrency),
    },
    MenuItem {
        letter: ',',
        name: ",",
        help: "Default to comma-separated (no currency symbol)",
        body: MenuBody::Action(Action::WorksheetGlobalFormatComma),
    },
    MenuItem {
        letter: 'G',
        name: "General",
        help: "Default to General format",
        body: MenuBody::Action(Action::WorksheetGlobalFormatGeneral),
    },
    MenuItem {
        letter: 'P',
        name: "Percent",
        help: "Default to Percent (value × 100) with % sign",
        body: MenuBody::Action(Action::WorksheetGlobalFormatPercent),
    },
    MenuItem {
        letter: 'D',
        name: "Date",
        help: "Default to a Date format (D1..D5, Time for D6..D9)",
        body: MenuBody::Submenu(WG_FORMAT_DATE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Default to showing formulas instead of values",
        body: MenuBody::Action(Action::WorksheetGlobalFormatText),
    },
    MenuItem {
        letter: 'H',
        name: "Hidden",
        help: "Default to hiding cell display",
        body: MenuBody::NotImplemented("wg-format-hidden"),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Reset global format to General",
        body: MenuBody::Action(Action::WorksheetGlobalFormatReset),
    },
];

const WG_FORMAT_DATE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "DD-MMM-YY",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateDmy),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "DD-MMM",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateDm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "MMM-YY",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateMy),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Long international (MM/DD/YY)",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateLongIntl),
    },
    MenuItem {
        letter: '5',
        name: "5",
        help: "Short international (MM/DD)",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateShortIntl),
    },
    MenuItem {
        letter: 'T',
        name: "Time",
        help: "Time format (D6..D9)",
        body: MenuBody::NotImplemented("wg-format-date-time"),
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

const RANGE_SEARCH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Formulas",
        help: "Search only formula cells",
        body: MenuBody::Action(Action::RangeSearchFormulas),
    },
    MenuItem {
        letter: 'L',
        name: "Labels",
        help: "Search only label cells",
        body: MenuBody::Action(Action::RangeSearchLabels),
    },
    MenuItem {
        letter: 'B',
        name: "Both",
        help: "Search both formulas and labels",
        body: MenuBody::Action(Action::RangeSearchBoth),
    },
];

/// Find|Replace sub-sub-menu shown after the search string commits.
/// Public so the UI layer can root a nested menu at it.
pub const RANGE_SEARCH_FIND_REPLACE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Find",
        help: "Move the pointer to the next match",
        body: MenuBody::Action(Action::RangeSearchFind),
    },
    MenuItem {
        letter: 'R',
        name: "Replace",
        help: "Replace all matches with another string",
        body: MenuBody::Action(Action::RangeSearchReplace),
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
        body: MenuBody::Submenu(RANGE_SEARCH_MENU),
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

const FILE_OPEN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Open a file as a new active file before the current one",
        body: MenuBody::Action(Action::FileOpenBefore),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Open a file as a new active file after the current one",
        body: MenuBody::Action(Action::FileOpenAfter),
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
        body: MenuBody::Submenu(FILE_OPEN_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Admin",
        help: "Reservation, seal, link-refresh",
        body: MenuBody::NotImplemented("f-admin"),
    },
];

const PRINT_FILE_OPTIONS_MARGINS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Set the left margin (spaces prefixed to each printed line)",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Set the right margin",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginRight),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Set the top margin (blank lines above the first row)",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginTop),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Set the bottom margin",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginBottom),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Options menu",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginsQuit),
    },
];

const PRINT_FILE_OPTIONS_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "As-Displayed",
        help: "Print cells as they appear on screen (default)",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherAsDisplayed),
    },
    MenuItem {
        letter: 'C',
        name: "Cell-Formulas",
        help: "Print the formula source in place of the computed value",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherCellFormulas),
    },
    MenuItem {
        letter: 'F',
        name: "Formatted",
        help: "Honor headers, footers, margins, and page breaks",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherFormatted),
    },
    MenuItem {
        letter: 'U',
        name: "Unformatted",
        help: "Dump the range with no page decorations",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherUnformatted),
    },
];

const PRINT_FILE_OPTIONS_ADVANCED_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "AutoLf",
        help: "Send LF after each CR (Auto Line Feed)",
        body: MenuBody::NotImplemented("pfo-advanced-autolf"),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Color print options",
        body: MenuBody::NotImplemented("pfo-advanced-color"),
    },
    MenuItem {
        letter: 'D',
        name: "Device",
        help: "CUPS printer queue name (lp -d <name>)",
        body: MenuBody::Action(Action::PrintSessionOptionsAdvancedDevice),
    },
    MenuItem {
        letter: 'F',
        name: "Fonts",
        help: "Font selection",
        body: MenuBody::NotImplemented("pfo-advanced-fonts"),
    },
    MenuItem {
        letter: 'I',
        name: "Images",
        help: "Image rendering options",
        body: MenuBody::NotImplemented("pfo-advanced-images"),
    },
    MenuItem {
        letter: 'L',
        name: "Layout",
        help: "Layout options",
        body: MenuBody::NotImplemented("pfo-advanced-layout"),
    },
    MenuItem {
        letter: 'P',
        name: "Priority",
        help: "Print job priority",
        body: MenuBody::NotImplemented("pfo-advanced-priority"),
    },
    MenuItem {
        letter: 'W',
        name: "Wait",
        help: "Wait between pages",
        body: MenuBody::NotImplemented("pfo-advanced-wait"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Options submenu",
        body: MenuBody::Action(Action::PrintSessionOptionsAdvancedQuit),
    },
];

const PRINT_FILE_OPTIONS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'H',
        name: "Header",
        help: "Set a three-part header (L|C|R), printed above each page",
        body: MenuBody::Action(Action::PrintSessionOptionsHeader),
    },
    MenuItem {
        letter: 'F',
        name: "Footer",
        help: "Set a three-part footer (L|C|R), printed below each page",
        body: MenuBody::Action(Action::PrintSessionOptionsFooter),
    },
    MenuItem {
        letter: 'M',
        name: "Margins",
        help: "Set Left / Right / Top / Bottom margins",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_MARGINS_MENU),
    },
    MenuItem {
        letter: 'P',
        name: "Pg-Length",
        help: "Set page length (lines per page)",
        body: MenuBody::Action(Action::PrintSessionOptionsPgLength),
    },
    MenuItem {
        letter: 'B',
        name: "Borders",
        help: "Repeat columns/rows at the top/left of each page",
        body: MenuBody::NotImplemented("pfo-borders"),
    },
    MenuItem {
        letter: 'S',
        name: "Setup",
        help: "Setup escape sequence for the printer",
        body: MenuBody::Action(Action::PrintSessionOptionsSetup),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "As-Displayed / Cell-Formulas / Formatted / Unformatted",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_OTHER_MENU),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Named print-settings sets",
        body: MenuBody::NotImplemented("pfo-name"),
    },
    MenuItem {
        letter: 'A',
        name: "Advanced",
        help: "Advanced options (AutoLf, Color, Device, Fonts, …)",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_ADVANCED_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the print menu",
        body: MenuBody::Action(Action::PrintSessionOptionsQuit),
    },
];

/// Sub-menu shown inside `/Print File` after the filename has been
/// committed. `pub` so the UI layer can root a nested menu at it.
pub const PRINT_FILE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Range",
        help: "Select the range of cells to print",
        body: MenuBody::Action(Action::PrintSessionRange),
    },
    MenuItem {
        letter: 'L',
        name: "Line",
        help: "Advance one line in the output file",
        body: MenuBody::NotImplemented("pf-line"),
    },
    MenuItem {
        letter: 'P',
        name: "Page",
        help: "Advance to the next page in the output file",
        body: MenuBody::NotImplemented("pf-page"),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Header, footer, margins, page length, and output format",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Clear print settings (header, footer, margins, …)",
        body: MenuBody::Action(Action::PrintSessionClear),
    },
    MenuItem {
        letter: 'A',
        name: "Align",
        help: "Reset the page counter to 1",
        body: MenuBody::Action(Action::PrintSessionAlign),
    },
    MenuItem {
        letter: 'G',
        name: "Go",
        help: "Write the selected range to the print file",
        body: MenuBody::Action(Action::PrintSessionGo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::Action(Action::PrintSessionQuit),
    },
];

const PRINT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Printer",
        help: "Print to printer",
        body: MenuBody::Action(Action::PrintPrinter),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Print to .PRN text file",
        body: MenuBody::Action(Action::PrintFile),
    },
    MenuItem {
        letter: 'E',
        name: "Encoded",
        help: "Print to encoded file with printer codes",
        body: MenuBody::Action(Action::PrintEncoded),
    },
    MenuItem {
        letter: 'C',
        name: "Cancel",
        help: "Cancel current print job",
        body: MenuBody::Action(Action::PrintCancel),
    },
];

const GRAPH_TYPE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Line",
        help: "Line graph",
        body: MenuBody::Action(Action::GraphTypeLine),
    },
    MenuItem {
        letter: 'B',
        name: "Bar",
        help: "Bar graph",
        body: MenuBody::Action(Action::GraphTypeBar),
    },
    MenuItem {
        letter: 'X',
        name: "XY",
        help: "XY (scatter) graph",
        body: MenuBody::Action(Action::GraphTypeXY),
    },
    MenuItem {
        letter: 'S',
        name: "Stack-Bar",
        help: "Stacked-bar graph",
        body: MenuBody::Action(Action::GraphTypeStack),
    },
    MenuItem {
        letter: 'P',
        name: "Pie",
        help: "Pie chart",
        body: MenuBody::Action(Action::GraphTypePie),
    },
    MenuItem {
        letter: 'H',
        name: "HLCO",
        help: "High/Low/Close/Open chart",
        body: MenuBody::Action(Action::GraphTypeHLCO),
    },
    MenuItem {
        letter: 'M',
        name: "Mixed",
        help: "Mixed (bar + line) chart",
        body: MenuBody::Action(Action::GraphTypeMixed),
    },
    MenuItem {
        letter: 'F',
        name: "Features",
        help: "Type features (stacked, 100%, 2Y, Y-ranges)",
        body: MenuBody::NotImplemented("gt-features"),
    },
];

const GRAPH_RESET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Clear every range and restore default type",
        body: MenuBody::Action(Action::GraphResetGraph),
    },
    MenuItem {
        letter: 'X',
        name: "X",
        help: "Clear X-axis range",
        body: MenuBody::NotImplemented("gr-x"),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Clear A range",
        body: MenuBody::NotImplemented("gr-a"),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Clear B range",
        body: MenuBody::NotImplemented("gr-b"),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Clear C range",
        body: MenuBody::NotImplemented("gr-c"),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Clear D range",
        body: MenuBody::NotImplemented("gr-d"),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Clear E range",
        body: MenuBody::NotImplemented("gr-e"),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Clear F range",
        body: MenuBody::NotImplemented("gr-f"),
    },
    MenuItem {
        letter: 'R',
        name: "Ranges",
        help: "Clear X and A..F together (keep options)",
        body: MenuBody::NotImplemented("gr-ranges"),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Reset graph options (keep ranges)",
        body: MenuBody::NotImplemented("gr-options"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph menu",
        body: MenuBody::NotImplemented("gr-quit"),
    },
];

const GRAPH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'T',
        name: "Type",
        help: "Select graph type",
        body: MenuBody::Submenu(GRAPH_TYPE_MENU),
    },
    MenuItem {
        letter: 'X',
        name: "X",
        help: "Set X-axis range",
        body: MenuBody::Action(Action::GraphX),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Set A data range",
        body: MenuBody::Action(Action::GraphA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Set B data range",
        body: MenuBody::Action(Action::GraphB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Set C data range",
        body: MenuBody::Action(Action::GraphC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Set D data range",
        body: MenuBody::Action(Action::GraphD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Set E data range",
        body: MenuBody::Action(Action::GraphE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Set F data range",
        body: MenuBody::Action(Action::GraphF),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Reset graph / ranges / options",
        body: MenuBody::Submenu(GRAPH_RESET_MENU),
    },
    MenuItem {
        letter: 'V',
        name: "View",
        help: "Display the current graph",
        body: MenuBody::Action(Action::GraphView),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Save graph to an SVG file",
        body: MenuBody::Action(Action::GraphSave),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Legend, Titles, Grid, Scale, Color, …",
        body: MenuBody::NotImplemented("g-options"),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Create, use, delete, reset named graphs",
        body: MenuBody::NotImplemented("g-name"),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Columnwise / Rowwise auto-assign",
        body: MenuBody::NotImplemented("g-group"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::Action(Action::GraphQuit),
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

// --------------------------------------------------------------------------
// WYSIWYG `:` colon-menu tree. R3.4a promoted WYSIWYG to an always-on
// feature; its commands live under the colon prefix to coexist with the
// classic `/` menu. Only `:Format` has live leaves today — the rest of
// the top level renders the muscle-memory path but surfaces
// "Not implemented yet" on commit.
// --------------------------------------------------------------------------

const WYSIWYG_FORMAT_BOLD_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Apply bold to a range",
        body: MenuBody::Action(Action::FormatBoldSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove bold from a range",
        body: MenuBody::Action(Action::FormatBoldClear),
    },
];

const WYSIWYG_FORMAT_ITALIC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Apply italic to a range",
        body: MenuBody::Action(Action::FormatItalicSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove italic from a range",
        body: MenuBody::Action(Action::FormatItalicClear),
    },
];

const WYSIWYG_FORMAT_UNDERLINE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Apply underline to a range",
        body: MenuBody::Action(Action::FormatUnderlineSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove underline from a range",
        body: MenuBody::Action(Action::FormatUnderlineClear),
    },
];

const WYSIWYG_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Bold",
        help: "Set or clear the bold attribute on a range",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_BOLD_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Italic",
        help: "Set or clear the italic attribute on a range",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_ITALIC_MENU),
    },
    MenuItem {
        letter: 'U',
        name: "Underline",
        help: "Set or clear the underline attribute on a range",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_UNDERLINE_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Font",
        help: "Change font on a range",
        body: MenuBody::NotImplemented("wysiwyg-format-font"),
    },
    MenuItem {
        letter: 'L',
        name: "Lines",
        help: "Draw lines around a range",
        body: MenuBody::NotImplemented("wysiwyg-format-lines"),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Change text or background color",
        body: MenuBody::NotImplemented("wysiwyg-format-color"),
    },
    MenuItem {
        letter: 'A',
        name: "Alignment",
        help: "Change alignment",
        body: MenuBody::NotImplemented("wysiwyg-format-alignment"),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear bold, italic and underline on a range",
        body: MenuBody::Action(Action::FormatReset),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::Action(Action::Cancel),
    },
];

/// Top-level WYSIWYG colon-menu.  Entered by pressing `:` in READY.
pub const WYSIWYG_ROOT: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "WYSIWYG worksheet display settings",
        body: MenuBody::NotImplemented("wysiwyg-worksheet"),
    },
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Bold, italic, underline, font, color, alignment...",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Embed a graph at the current cell",
        body: MenuBody::NotImplemented("wysiwyg-graph"),
    },
    MenuItem {
        letter: 'N',
        name: "Named-Style",
        help: "Named cell formatting style",
        body: MenuBody::NotImplemented("wysiwyg-named-style"),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "WYSIWYG print controls",
        body: MenuBody::NotImplemented("wysiwyg-print"),
    },
    MenuItem {
        letter: 'D',
        name: "Display",
        help: "Screen display settings",
        body: MenuBody::NotImplemented("wysiwyg-display"),
    },
    MenuItem {
        letter: 'S',
        name: "Special",
        help: "Copy / move / import formatting",
        body: MenuBody::NotImplemented("wysiwyg-special"),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Text editing and justification",
        body: MenuBody::NotImplemented("wysiwyg-text"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        body: MenuBody::Action(Action::Cancel),
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
        help: "End the l123 session",
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
                "Worksheet",
                "Range",
                "Copy",
                "Move",
                "File",
                "Print",
                "Graph",
                "Data",
                "System",
                "Add-In",
                "Quit"
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
    fn resolve_wgdo_undo_enable_and_disable() {
        let e = resolve(&['W', 'G', 'D', 'O', 'U', 'E']).unwrap();
        assert!(matches!(
            e.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoEnable)
        ));
        let d = resolve(&['W', 'G', 'D', 'O', 'U', 'D']).unwrap();
        assert!(matches!(
            d.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoDisable)
        ));
    }

    #[test]
    fn resolve_wgd_status_update_dir_temp() {
        for (path, expected) in [
            (
                &['W', 'G', 'D', 'S'][..],
                Action::WorksheetGlobalDefaultStatus,
            ),
            (&['W', 'G', 'D', 'U'], Action::WorksheetGlobalDefaultUpdate),
            (&['W', 'G', 'D', 'D'], Action::WorksheetGlobalDefaultDir),
            (&['W', 'G', 'D', 'T'], Action::WorksheetGlobalDefaultTemp),
        ] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            match n.body {
                MenuBody::Action(a) => assert_eq!(a, expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn resolve_wgd_autoexec_ext_graph_branches() {
        for (path, expected) in [
            (
                &['W', 'G', 'D', 'A', 'Y'][..],
                Action::WorksheetGlobalDefaultAutoexecYes,
            ),
            (
                &['W', 'G', 'D', 'A', 'N'],
                Action::WorksheetGlobalDefaultAutoexecNo,
            ),
            (
                &['W', 'G', 'D', 'E', 'S'],
                Action::WorksheetGlobalDefaultExtSave,
            ),
            (
                &['W', 'G', 'D', 'E', 'L'],
                Action::WorksheetGlobalDefaultExtList,
            ),
            (
                &['W', 'G', 'D', 'G', 'G', 'C'],
                Action::WorksheetGlobalDefaultGraphGroupColumnwise,
            ),
            (
                &['W', 'G', 'D', 'G', 'G', 'R'],
                Action::WorksheetGlobalDefaultGraphGroupRowwise,
            ),
            (
                &['W', 'G', 'D', 'G', 'S', 'C'],
                Action::WorksheetGlobalDefaultGraphSaveCgm,
            ),
            (
                &['W', 'G', 'D', 'G', 'S', 'P'],
                Action::WorksheetGlobalDefaultGraphSavePic,
            ),
        ] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            match n.body {
                MenuBody::Action(a) => assert_eq!(a, expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn resolve_wgd_printer_branches() {
        for (path, expected) in [
            (
                &['W', 'G', 'D', 'P', 'I'][..],
                Action::WorksheetGlobalDefaultPrinterInterface,
            ),
            (
                &['W', 'G', 'D', 'P', 'A', 'Y'],
                Action::WorksheetGlobalDefaultPrinterAutoLfYes,
            ),
            (
                &['W', 'G', 'D', 'P', 'A', 'N'],
                Action::WorksheetGlobalDefaultPrinterAutoLfNo,
            ),
            (
                &['W', 'G', 'D', 'P', 'L'],
                Action::WorksheetGlobalDefaultPrinterMarginLeft,
            ),
            (
                &['W', 'G', 'D', 'P', 'R'],
                Action::WorksheetGlobalDefaultPrinterMarginRight,
            ),
            (
                &['W', 'G', 'D', 'P', 'T'],
                Action::WorksheetGlobalDefaultPrinterMarginTop,
            ),
            (
                &['W', 'G', 'D', 'P', 'B'],
                Action::WorksheetGlobalDefaultPrinterMarginBottom,
            ),
            (
                &['W', 'G', 'D', 'P', 'P'],
                Action::WorksheetGlobalDefaultPrinterPgLength,
            ),
            (
                &['W', 'G', 'D', 'P', 'W', 'Y'],
                Action::WorksheetGlobalDefaultPrinterWaitYes,
            ),
            (
                &['W', 'G', 'D', 'P', 'W', 'N'],
                Action::WorksheetGlobalDefaultPrinterWaitNo,
            ),
            (
                &['W', 'G', 'D', 'P', 'S'],
                Action::WorksheetGlobalDefaultPrinterSetup,
            ),
            (
                &['W', 'G', 'D', 'P', 'N'],
                Action::WorksheetGlobalDefaultPrinterName,
            ),
            (
                &['W', 'G', 'D', 'P', 'Q'],
                Action::WorksheetGlobalDefaultPrinterQuit,
            ),
        ] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            match n.body {
                MenuBody::Action(a) => assert_eq!(a, expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn resolve_wgdo_beep_enable_and_disable() {
        let e = resolve(&['W', 'G', 'D', 'O', 'B', 'E']).unwrap();
        assert!(matches!(
            e.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepEnable)
        ));
        let d = resolve(&['W', 'G', 'D', 'O', 'B', 'D']).unwrap();
        assert!(matches!(
            d.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepDisable)
        ));
    }

    #[test]
    fn resolve_file_open_before_and_after() {
        let b = resolve(&['F', 'O', 'B']).unwrap();
        assert!(matches!(b.body, MenuBody::Action(Action::FileOpenBefore)));
        let a = resolve(&['F', 'O', 'A']).unwrap();
        assert!(matches!(a.body, MenuBody::Action(Action::FileOpenAfter)));
    }

    #[test]
    fn resolve_wg_group_enable_and_disable() {
        let e = resolve(&['W', 'G', 'G', 'E']).unwrap();
        assert!(matches!(
            e.body,
            MenuBody::Action(Action::WorksheetGlobalGroupEnable)
        ));
        let d = resolve(&['W', 'G', 'G', 'D']).unwrap();
        assert!(matches!(
            d.body,
            MenuBody::Action(Action::WorksheetGlobalGroupDisable)
        ));
    }

    #[test]
    fn resolve_ws_insert_sheet_before_and_after() {
        let b = resolve(&['W', 'I', 'S', 'B']).unwrap();
        assert!(matches!(
            b.body,
            MenuBody::Action(Action::WorksheetInsertSheetBefore)
        ));
        let a = resolve(&['W', 'I', 'S', 'A']).unwrap();
        assert!(matches!(
            a.body,
            MenuBody::Action(Action::WorksheetInsertSheetAfter)
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
    fn resolve_worksheet_global_format_leaves() {
        let cases: &[(&[char], Action)] = &[
            (&['W', 'G', 'F', 'F'], Action::WorksheetGlobalFormatFixed),
            (
                &['W', 'G', 'F', 'S'],
                Action::WorksheetGlobalFormatScientific,
            ),
            (&['W', 'G', 'F', 'C'], Action::WorksheetGlobalFormatCurrency),
            (&['W', 'G', 'F', ','], Action::WorksheetGlobalFormatComma),
            (&['W', 'G', 'F', 'G'], Action::WorksheetGlobalFormatGeneral),
            (&['W', 'G', 'F', 'P'], Action::WorksheetGlobalFormatPercent),
            (&['W', 'G', 'F', 'T'], Action::WorksheetGlobalFormatText),
            (&['W', 'G', 'F', 'R'], Action::WorksheetGlobalFormatReset),
            (
                &['W', 'G', 'F', 'D', '1'],
                Action::WorksheetGlobalFormatDateDmy,
            ),
            (
                &['W', 'G', 'F', 'D', '5'],
                Action::WorksheetGlobalFormatDateShortIntl,
            ),
        ];
        for (path, expected) in cases {
            let node = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            match node.body {
                MenuBody::Action(actual) => assert_eq!(actual, *expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn root_names_all_start_with_capital() {
        for m in ROOT {
            let c = m.name.chars().next().unwrap();
            assert!(c.is_ascii_uppercase(), "{}", m.name);
        }
    }

    #[test]
    fn wysiwyg_letters_are_unique_at_every_level() {
        assert_unique_letters(WYSIWYG_ROOT, ":");
    }

    #[test]
    fn wysiwyg_top_level_names() {
        let names: Vec<&str> = WYSIWYG_ROOT.iter().map(|m| m.name).collect();
        assert_eq!(
            names,
            vec![
                "Worksheet",
                "Format",
                "Graph",
                "Named-Style",
                "Print",
                "Display",
                "Special",
                "Text",
                "Quit",
            ]
        );
    }

    #[test]
    fn resolve_wysiwyg_format_bold_set_and_clear() {
        let s = resolve_within(WYSIWYG_ROOT, &['F', 'B', 'S']).unwrap();
        assert!(matches!(s.body, MenuBody::Action(Action::FormatBoldSet)));
        let c = resolve_within(WYSIWYG_ROOT, &['F', 'B', 'C']).unwrap();
        assert!(matches!(c.body, MenuBody::Action(Action::FormatBoldClear)));
    }

    #[test]
    fn resolve_wysiwyg_format_italic_and_underline() {
        let i = resolve_within(WYSIWYG_ROOT, &['F', 'I', 'S']).unwrap();
        assert!(matches!(i.body, MenuBody::Action(Action::FormatItalicSet)));
        let u = resolve_within(WYSIWYG_ROOT, &['F', 'U', 'C']).unwrap();
        assert!(matches!(
            u.body,
            MenuBody::Action(Action::FormatUnderlineClear)
        ));
    }

    #[test]
    fn resolve_wysiwyg_format_reset() {
        let r = resolve_within(WYSIWYG_ROOT, &['F', 'R']).unwrap();
        assert!(matches!(r.body, MenuBody::Action(Action::FormatReset)));
    }

    #[test]
    fn wysiwyg_format_font_is_not_implemented() {
        let f = resolve_within(WYSIWYG_ROOT, &['F', 'F']).unwrap();
        assert!(matches!(f.body, MenuBody::NotImplemented(_)));
    }

    #[test]
    fn wysiwyg_quit_maps_to_cancel() {
        let q = resolve_within(WYSIWYG_ROOT, &['Q']).unwrap();
        assert!(matches!(q.body, MenuBody::Action(Action::Cancel)));
    }
}
