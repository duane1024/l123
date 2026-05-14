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
    /// Filename of the L123 help page F1 jumps to when this item is the
    /// deepest one along the current menu path. Empty string means
    /// "inherit from the nearest ancestor that has one set".
    pub help_page: &'static str,
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
    /// `/QY` — request to end the session. Exits immediately when the
    /// workbook is clean; opens [`QUIT_DIRTY_MENU`] (a second No/Yes
    /// confirm) when there are unsaved changes. The second confirm's
    /// Yes leaf fires [`Action::Quit`].
    QuitConfirm,
    Quit,
    WorksheetEraseConfirm,
    WorksheetInsertRow,
    WorksheetInsertColumn,
    WorksheetInsertSheetBefore,
    WorksheetInsertSheetAfter,
    WorksheetDeleteRow,
    WorksheetDeleteColumn,
    /// `/Worksheet Delete Sheet` — drop the worksheet at the pointer.
    /// Sheets after it shift back one slot; refused on the last sheet.
    WorksheetDeleteSheet,
    /// `/Worksheet Delete File` — remove the foreground active file
    /// from memory. When the last active file is deleted, the
    /// workspace resets to a single blank workbook (mirrors
    /// `/Worksheet Erase Yes`).
    WorksheetDeleteFile,
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
    /// `/Worksheet Titles Both` — freeze rows above and columns left of
    /// the cell pointer.
    WorksheetTitlesBoth,
    /// `/Worksheet Titles Horizontal` — freeze the rows above the cell
    /// pointer.
    WorksheetTitlesHorizontal,
    /// `/Worksheet Titles Vertical` — freeze the columns left of the
    /// cell pointer.
    WorksheetTitlesVertical,
    /// `/Worksheet Titles Clear` — remove any frozen-pane setting on
    /// the current sheet.
    WorksheetTitlesClear,
    /// `/Worksheet Page Row` — insert a row at the pointer with `|::`
    /// in column A, marking a manual row page break for the print
    /// engine.
    WorksheetPageRow,
    /// `/Worksheet Page Column` — insert a column at the pointer with
    /// `|::` in row 1, marking a manual column page break for the
    /// print engine.
    WorksheetPageColumn,
    /// `/Worksheet Hide Enable` — hide the current worksheet so it
    /// disappears from `Ctrl-PgUp/PgDn` navigation. Refuses if the
    /// current sheet is the only visible one.
    WorksheetHideEnable,
    /// `/Worksheet Hide Disable` — unhide every hidden sheet on the
    /// current workbook.
    WorksheetHideDisable,
    /// `/Worksheet Learn Range` — POINT for the destination range
    /// where Alt-F5 LEARN recordings are written.
    WorksheetLearnRange,
    /// `/Worksheet Learn Cancel` — drop the learn-range definition
    /// and stop recording if active.
    WorksheetLearnCancel,
    /// `/Worksheet Learn Erase` — blank every cell of the current
    /// learn range without dropping the range definition itself.
    WorksheetLearnErase,
    WorksheetGlobalDefaultOtherUndoEnable,
    WorksheetGlobalDefaultOtherUndoDisable,
    /// `/Worksheet Global Default Other Beep Enable` — turn on the
    /// soft terminal bell that fires when the pointer hits an edge
    /// of the sheet or an equivalent invalid operation is attempted.
    WorksheetGlobalDefaultOtherBeepEnable,
    /// `/Worksheet Global Default Other Beep Disable`.
    WorksheetGlobalDefaultOtherBeepDisable,
    /// `/Worksheet Global Default Other International Punctuation A..H`
    /// — locale punctuation triple (decimal, argument, thousands).
    WorksheetGlobalDefaultOtherIntlPunctuationA,
    WorksheetGlobalDefaultOtherIntlPunctuationB,
    WorksheetGlobalDefaultOtherIntlPunctuationC,
    WorksheetGlobalDefaultOtherIntlPunctuationD,
    WorksheetGlobalDefaultOtherIntlPunctuationE,
    WorksheetGlobalDefaultOtherIntlPunctuationF,
    WorksheetGlobalDefaultOtherIntlPunctuationG,
    WorksheetGlobalDefaultOtherIntlPunctuationH,
    /// `/Worksheet Global Default Other International Currency Prefix|
    /// Suffix` — chooses where the currency symbol sits, then prompts
    /// for the symbol string.
    WorksheetGlobalDefaultOtherIntlCurrencyPrefix,
    WorksheetGlobalDefaultOtherIntlCurrencySuffix,
    /// `/Worksheet Global Default Other International Date A..D` —
    /// selects the international date style used by D4 (long) and
    /// D5 (short).
    WorksheetGlobalDefaultOtherIntlDateA,
    WorksheetGlobalDefaultOtherIntlDateB,
    WorksheetGlobalDefaultOtherIntlDateC,
    WorksheetGlobalDefaultOtherIntlDateD,
    /// `/Worksheet Global Default Other International Time A..D` —
    /// selects the international time style used by D8 (long) and
    /// D9 (short).
    WorksheetGlobalDefaultOtherIntlTimeA,
    WorksheetGlobalDefaultOtherIntlTimeB,
    WorksheetGlobalDefaultOtherIntlTimeC,
    WorksheetGlobalDefaultOtherIntlTimeD,
    /// `/Worksheet Global Default Other International Negative
    /// Parens|Sign` — controls how negative Currency/Comma values
    /// display.
    WorksheetGlobalDefaultOtherIntlNegativeParens,
    WorksheetGlobalDefaultOtherIntlNegativeSign,
    /// `/Worksheet Global Default Other Clock Standard` — show the
    /// date and time in the status line using the 12-hour Standard
    /// clock format (`DD-MMM-YY HH:MM AM/PM`).
    WorksheetGlobalDefaultOtherClockStandard,
    /// `/Worksheet Global Default Other Clock International` — show
    /// the date and time in the status line using the 24-hour
    /// International format (`DD-MMM-YYYY HH:MM`).
    WorksheetGlobalDefaultOtherClockInternational,
    /// `/Worksheet Global Default Other Clock None` — suppress the
    /// status-line clock entirely.
    WorksheetGlobalDefaultOtherClockNone,
    /// `/Worksheet Global Default Other Clock Filename` — show the
    /// active workbook's filename in the status-line slot instead of
    /// the clock.
    WorksheetGlobalDefaultOtherClockFilename,
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
    /// `/Range Name Reset` — drop every named range in the active
    /// file. Undo-aware: the previous map is captured as a single
    /// journal entry.
    RangeNameReset,
    /// `/Range Name Labels Right|Down|Left|Up` — for each label in
    /// the picked range, define a 1-cell range name (the label's
    /// text) pointing at the adjacent cell in the chosen direction.
    /// Labels that violate the name rules (>15 chars, non-letter
    /// first char, embedded whitespace) are skipped silently.
    RangeNameLabelsRight,
    RangeNameLabelsDown,
    RangeNameLabelsLeft,
    RangeNameLabelsUp,
    /// `/Range Name Table` — dump the active file's named-range
    /// table to a 2-column block starting at the picked cell.
    /// Column 1 is the name; column 2 is the range as a string.
    RangeNameTable,
    /// `/Range Name Note Create|Delete|Reset|Table` — manage notes
    /// on named-range definitions. Names with notes show them in
    /// the F3 NAMES picker.
    RangeNameNoteCreate,
    RangeNameNoteDelete,
    RangeNameNoteReset,
    RangeNameNoteTable,
    /// `/Range Name Undefine` — drop a name AND rewrite formulas
    /// that referenced it to use the literal range, preserving
    /// computed values across the deletion.
    RangeNameUndefine,
    /// `/Range Prot` — re-protect a range (default state). Effective
    /// when /WGP Enable is on.
    RangeProtect,
    /// `/Range Unprot` — mark a range as writable even when
    /// /WGP Enable is on.
    RangeUnprotect,
    /// `/Range Input` — restrict the pointer to unprotected cells
    /// within a range until the user presses Esc.
    RangeInput,
    /// `/Range Justify` — word-wrap a long label into a column
    /// block at the cell's column width.
    RangeJustify,
    /// `/Range Value` — copy a range, replacing formulas with their
    /// cached values at the destination.
    RangeValue,
    /// `/Range Trans` — transpose a rectangular range. Per R3.4a,
    /// formulas with relative references collapse to their cached
    /// values at the new orientation; absolute references travel
    /// unchanged.
    RangeTrans,
    /// `/Range Compare` — three-POINT diff between two ranges (v0.4).
    /// Prompts LEFT range, RIGHT range, OUTPUT anchor; writes one row
    /// per differing cell as `(addr, left_value, right_value,
    /// diff_kind)`. Equal cells produce no row; identical ranges
    /// report "No differences" on line 3 and write nothing. Shape
    /// mismatch raises ERROR mode.
    RangeCompare,
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
    RangeFormatHidden,
    RangeFormatTimeHmsAmPm,
    RangeFormatTimeHmAmPm,
    RangeFormatTimeLongIntl,
    RangeFormatTimeShortIntl,
    RangeFormatPlusMinus,
    RangeFormatAutomatic,
    RangeFormatLabelOnly,
    RangeFormatParensYes,
    RangeFormatParensNo,
    RangeFormatNegColorBlack,
    RangeFormatNegColorWhite,
    RangeFormatNegColorRed,
    RangeFormatNegColorGreen,
    RangeFormatNegColorBlue,
    RangeFormatNegColorYellow,
    RangeFormatNegColorCyan,
    RangeFormatNegColorMagenta,
    RangeFormatNegColorReset,
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
    WorksheetGlobalFormatTimeHmsAmPm,
    WorksheetGlobalFormatTimeHmAmPm,
    WorksheetGlobalFormatTimeLongIntl,
    WorksheetGlobalFormatTimeShortIntl,
    WorksheetGlobalFormatText,
    WorksheetGlobalFormatHidden,
    WorksheetGlobalFormatPlusMinus,
    WorksheetGlobalFormatAutomatic,
    WorksheetGlobalFormatLabelOnly,
    WorksheetGlobalFormatParensYes,
    WorksheetGlobalFormatParensNo,
    WorksheetGlobalFormatNegColorBlack,
    WorksheetGlobalFormatNegColorWhite,
    WorksheetGlobalFormatNegColorRed,
    WorksheetGlobalFormatNegColorGreen,
    WorksheetGlobalFormatNegColorBlue,
    WorksheetGlobalFormatNegColorYellow,
    WorksheetGlobalFormatNegColorCyan,
    WorksheetGlobalFormatNegColorMagenta,
    WorksheetGlobalFormatNegColorReset,
    WorksheetGlobalFormatReset,
    Copy,
    Move,
    FileSave,
    FileRetrieve,
    FileXtractFormulas,
    FileXtractValues,
    FileImportNumbers,
    /// `/File Import Text` — read a plain-text file and store each line
    /// as a label down a single column starting at the pointer. The
    /// counterpart to `FileImportNumbers`; no CSV parsing.
    FileImportText,
    /// `/File Combine Copy Entire-File` — overwrite cells starting at
    /// the pointer with the contents of every non-empty cell in the
    /// source file.
    FileCombineCopyEntire,
    /// `/File Combine Copy Named-Or-Specified-Range` — same as
    /// `FileCombineCopyEntire` but anchored to a user-typed source
    /// range like `A1..C5`.
    FileCombineCopyNamed,
    /// `/File Combine Add Entire-File` — numerically add each source
    /// cell to the corresponding target cell. Source labels and target
    /// labels are skipped (no overwrite).
    FileCombineAddEntire,
    /// `/File Combine Add Named-Or-Specified-Range`.
    FileCombineAddNamed,
    /// `/File Combine Subtract Entire-File` — numerically subtract each
    /// source cell from the corresponding target cell. Same label rules
    /// as `FileCombineAddEntire`.
    FileCombineSubtractEntire,
    /// `/File Combine Subtract Named-Or-Specified-Range`.
    FileCombineSubtractNamed,
    /// `/File Erase Worksheet` — delete a worksheet file from disk after
    /// a No/Yes confirm.
    FileEraseWorksheet,
    /// `/File Erase Print` — delete a print-settings file from disk.
    FileErasePrint,
    /// `/File Erase Graph` — delete a graph file from disk.
    FileEraseGraph,
    /// `/File Erase Other` — delete any file from disk.
    FileEraseOther,
    /// `/File Admin Reservation Get` — acquire the active file's edit
    /// reservation. Wired as a typed action; behavior is currently a
    /// menu-close (no in-memory reservation model yet).
    FileAdminReservationGet,
    /// `/File Admin Reservation Release` — release the active file's
    /// edit reservation. See `FileAdminReservationGet` for current
    /// behavior.
    FileAdminReservationRelease,
    /// `/File Admin Seal File` — password-seal the active file. Stub
    /// today (IronCalc 0.7 doesn't model encrypted xlsx).
    FileAdminSealFile,
    /// `/File Admin Seal Reservation-Setting` — password-seal the
    /// reservation behavior. Stub today.
    FileAdminSealReservationSetting,
    /// `/File Admin Seal Disable` — disable an existing seal (requires
    /// the seal password). Stub today.
    FileAdminSealDisable,
    /// `/File Admin Table {Worksheet|Print|Graph|Other|Active|Linked}`
    /// — build a table of files of the given kind in the session
    /// directory. Stubs today; the listing is provided by the simpler
    /// `/File List` family.
    FileAdminTableWorksheet,
    FileAdminTablePrint,
    FileAdminTableGraph,
    FileAdminTableOther,
    FileAdminTableActive,
    FileAdminTableLinked,
    /// `/File Admin Link-Refresh` — refresh formulas that reference
    /// linked files. Stub today.
    FileAdminLinkRefresh,
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
    FileListOther,
    /// `/System` — suspend the TUI and shell out to `$SHELL` (or
    /// `cmd.exe` on Windows). The shell's exit returns control to
    /// l123 with the workbook untouched.
    System,
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
    /// `/Graph Reset X` — clear the X range only.
    GraphResetX,
    /// `/Graph Reset A`..`/Graph Reset F` — clear one A-F range.
    GraphResetA,
    GraphResetB,
    GraphResetC,
    GraphResetD,
    GraphResetE,
    GraphResetF,
    /// `/Graph Reset Ranges` — clear X and every A-F range; keep
    /// type, features, and options.
    GraphResetRanges,
    /// `/Graph Reset Options` — restore options (titles, legends,
    /// grid, scale, …) to default; keep ranges, type, and features.
    GraphResetOptions,
    /// `/Graph View` — full-screen graph display (same as F10).
    GraphView,
    /// `/Graph Save` — prompt for a filename and write the graph to
    /// disk (SVG; a `.cgm` extension is preserved but the file body
    /// is still SVG, per project convention).
    GraphSave,
    /// `/Graph Quit` — close the `/Graph` menu back to READY.
    GraphQuit,

    // ---- /Graph Type Features (slice C1: orientation + boolean flags) --
    GraphFeaturesVertical,
    GraphFeaturesHorizontal,
    GraphFeaturesStackedYes,
    GraphFeaturesStackedNo,
    GraphFeaturesPercentYes,
    GraphFeaturesPercentNo,
    GraphFeaturesDropShadowYes,
    GraphFeaturesDropShadowNo,
    GraphFeaturesThreeDYes,
    GraphFeaturesThreeDNo,
    GraphFeaturesTableYes,
    GraphFeaturesTableNo,
    /// `/Graph Type Features Quit` — return to `/Graph Type` menu.
    /// Reused by the 2Y-Ranges, Y-Ranges, and Frame submenus' Quit
    /// leaves: each just pops one menu level.
    GraphFeaturesQuit,

    // ---- /Graph Type Features 2Y-Ranges and Y-Ranges (slice C2) -------
    GraphFeatures2YGraph,
    GraphFeatures2YA,
    GraphFeatures2YB,
    GraphFeatures2YC,
    GraphFeatures2YD,
    GraphFeatures2YE,
    GraphFeatures2YF,
    GraphFeaturesYGraph,
    GraphFeaturesYA,
    GraphFeaturesYB,
    GraphFeaturesYC,
    GraphFeaturesYD,
    GraphFeaturesYE,
    GraphFeaturesYF,

    // ---- /Graph Type Features Frame (slice C2) ------------------------
    GraphFeaturesFrameLeftYes,
    GraphFeaturesFrameLeftNo,
    GraphFeaturesFrameRightYes,
    GraphFeaturesFrameRightNo,
    GraphFeaturesFrameTopYes,
    GraphFeaturesFrameTopNo,
    GraphFeaturesFrameBottomYes,
    GraphFeaturesFrameBottomNo,
    /// `/Graph Type Features Frame Y-Axis {Yes|No}` — toggle the
    /// inner y-axis line drawn just inside the Left edge.
    GraphFeaturesFrameYAxisYes,
    GraphFeaturesFrameYAxisNo,
    /// `/Graph Type Features Frame All` — turn every frame side on.
    GraphFeaturesFrameAll,
    /// `/Graph Type Features Frame Clear` — turn every frame side off.
    GraphFeaturesFrameClear,

    // ---- /Graph Options (slice D of GRAPH_PLAN.md) -------------------
    /// `/Graph Options Color` — render the graph with colored series.
    GraphOptionsColor,
    /// `/Graph Options B&W` — render the graph in black and white.
    GraphOptionsBW,
    /// `/Graph Options Quit` — return to `/Graph` menu.
    GraphOptionsQuit,
    /// `/Graph Options Grid Horizontal` — turn on horizontal grid lines.
    /// `/Graph Options Grid Y-Axis {Y|2Y|Both}` — choose which
    /// y-axis the horizontal grid lines originate from.
    GraphOptionsGridYAxisFirst,
    GraphOptionsGridYAxisSecond,
    GraphOptionsGridYAxisBoth,
    GraphOptionsGridHorizontal,
    /// `/Graph Options Grid Vertical` — turn on vertical grid lines.
    GraphOptionsGridVertical,
    /// `/Graph Options Grid Both` — turn on horizontal and vertical
    /// grid lines together.
    GraphOptionsGridBoth,
    /// `/Graph Options Grid Clear` — remove every grid line.
    GraphOptionsGridClear,

    // ---- /Graph Options Format (slice D Format portion) ---------------
    // Per-series + Graph-wide × five LineFormat values. The `Graph`
    // slot applies the chosen format to all six A..F series at once.
    GraphFormatGraphLines,
    GraphFormatGraphSymbols,
    GraphFormatGraphBoth,
    GraphFormatGraphNeither,
    GraphFormatGraphArea,
    GraphFormatALines,
    GraphFormatASymbols,
    GraphFormatABoth,
    GraphFormatANeither,
    GraphFormatAArea,
    GraphFormatBLines,
    GraphFormatBSymbols,
    GraphFormatBBoth,
    GraphFormatBNeither,
    GraphFormatBArea,
    GraphFormatCLines,
    GraphFormatCSymbols,
    GraphFormatCBoth,
    GraphFormatCNeither,
    GraphFormatCArea,
    GraphFormatDLines,
    GraphFormatDSymbols,
    GraphFormatDBoth,
    GraphFormatDNeither,
    GraphFormatDArea,
    GraphFormatELines,
    GraphFormatESymbols,
    GraphFormatEBoth,
    GraphFormatENeither,
    GraphFormatEArea,
    GraphFormatFLines,
    GraphFormatFSymbols,
    GraphFormatFBoth,
    GraphFormatFNeither,
    GraphFormatFArea,

    // ---- /Graph Options Titles (slice D Titles portion) --------------
    /// `/Graph Options Titles First` — first (top) graph title line.
    GraphOptionsTitleFirst,
    /// `/Graph Options Titles Second` — second (subtitle) line.
    GraphOptionsTitleSecond,
    /// `/Graph Options Titles X-Axis` — x-axis caption.
    GraphOptionsTitleXAxis,
    /// `/Graph Options Titles Y-Axis` — first y-axis caption.
    GraphOptionsTitleYAxis,
    /// `/Graph Options Titles 2Y-Axis` — second y-axis caption.
    GraphOptionsTitle2YAxis,
    /// `/Graph Options Titles Note` — bottom-left footnote.
    GraphOptionsTitleNote,
    /// `/Graph Options Titles Other-Note` — bottom-right footnote.
    GraphOptionsTitleOtherNote,

    // ---- /Graph Options Legend (slice D Legend portion) --------------
    /// `/Graph Options Legend A` — text prompt for the A series legend.
    GraphOptionsLegendA,
    GraphOptionsLegendB,
    GraphOptionsLegendC,
    GraphOptionsLegendD,
    GraphOptionsLegendE,
    GraphOptionsLegendF,
    /// `/Graph Options Legend Range` — POINT prompt for a worksheet
    /// range; cell text fills the six legend slots in order.
    GraphOptionsLegendRange,

    // ---- /Graph Options Data-Labels (slice D Data-Labels portion) ---
    /// `/Graph Options Data-Labels A` — POINT prompt; commit stores
    /// the range in `current_graph.options.data_labels[0]`.
    GraphOptionsDataLabelsA,
    GraphOptionsDataLabelsB,
    GraphOptionsDataLabelsC,
    GraphOptionsDataLabelsD,
    GraphOptionsDataLabelsE,
    GraphOptionsDataLabelsF,
    /// Leaves of the data-label placement follow-up menu rooted
    /// after `/GOD {A-F}` commits the POINT range. Each one writes
    /// `current_graph.options.data_labels_placement[slot]` using
    /// the slot stashed on `App::pending_data_labels_slot`.
    GraphOptionsDataLabelsCenter,
    GraphOptionsDataLabelsLeft,
    GraphOptionsDataLabelsAbove,
    GraphOptionsDataLabelsRight,
    GraphOptionsDataLabelsBelow,

    // ---- /Graph Options Scale (slice D Scale portion) ----------------
    // Per-axis Auto/Manual mode, plus the global Skip prompt.
    // Lower/Upper/Format/Indicator/Type/Exponent/Width need extra
    // ScaleAxis fields and stay NotImplemented for now.
    GraphOptionsScaleYAuto,
    GraphOptionsScaleYManual,
    GraphOptionsScaleXAuto,
    GraphOptionsScaleXManual,
    GraphOptionsScale2YAuto,
    GraphOptionsScale2YManual,
    /// `/Graph Options Scale {axis} {Lower|Upper}` — open a numeric
    /// prompt for the manual axis bound. Pre-fills with the current
    /// value; an empty buffer + Enter clears the bound.
    GraphOptionsScaleYLower,
    GraphOptionsScaleYUpper,
    GraphOptionsScaleXLower,
    GraphOptionsScaleXUpper,
    GraphOptionsScale2YLower,
    GraphOptionsScale2YUpper,
    /// `/Graph Options Scale {axis} Type {Linear|Logarithmic}` —
    /// switch the per-axis mapping. Default Linear.
    GraphOptionsScaleYTypeLinear,
    GraphOptionsScaleYTypeLog,
    GraphOptionsScaleXTypeLinear,
    GraphOptionsScaleXTypeLog,
    GraphOptionsScale2YTypeLinear,
    GraphOptionsScale2YTypeLog,
    /// `/Graph Options Scale {axis} Width` — numeric prompt
    /// clamped 0..=40. 0 means "auto".
    GraphOptionsScaleYWidth,
    GraphOptionsScaleXWidth,
    GraphOptionsScale2YWidth,
    /// `/Graph Options Scale {axis} Exponent` — signed numeric
    /// prompt clamped -19..=19. 0 means "auto".
    GraphOptionsScaleYExponent,
    GraphOptionsScaleXExponent,
    GraphOptionsScale2YExponent,
    /// `/Graph Options Scale {axis} Indicator {Yes|No|Manual}` —
    /// magnitude indicator setting. Default Yes.
    GraphOptionsScaleYIndicatorYes,
    GraphOptionsScaleYIndicatorNo,
    GraphOptionsScaleYIndicatorManual,
    GraphOptionsScaleXIndicatorYes,
    GraphOptionsScaleXIndicatorNo,
    GraphOptionsScaleXIndicatorManual,
    GraphOptionsScale2YIndicatorYes,
    GraphOptionsScale2YIndicatorNo,
    GraphOptionsScale2YIndicatorManual,
    /// `/Graph Options Scale Skip` — numeric prompt; commit sets
    /// `current_graph.options.skip` to the new value (clamped 1..=8192).
    GraphOptionsScaleSkip,

    // ---- /Graph Name (slice E1 of GRAPH_PLAN.md) --------------------
    /// `/Graph Name Use` — text prompt; on commit replaces
    /// `current_graph` with the named graph's settings.
    GraphNameUse,
    /// `/Graph Name Create` — text prompt; on commit stores
    /// `current_graph.clone()` under the supplied name.
    GraphNameCreate,
    /// `/Graph Name Delete` — text prompt; on commit removes the
    /// named graph from `Workbook::graphs`.
    GraphNameDelete,
    /// `/Graph Name Reset` — immediately deletes every named graph
    /// in the current workbook. Reference is explicit: no confirmation.
    GraphNameReset,
    /// `/Graph Name Table` — POINT prompt for an anchor; writes one
    /// row per named graph, two columns: name + graph-type tag.
    GraphNameTable,

    // ---- /Graph Group (slice E2 of GRAPH_PLAN.md) -------------------
    /// `/Graph Group` — enters POINT for the group range. The
    /// per-axis assignment happens inside the
    /// `GraphGroupColumnwise|Rowwise` leaves of the rooted submenu.
    GraphGroup,
    /// `/Graph Group Columnwise` — first column of the stashed range
    /// becomes X; succeeding columns A..F.
    GraphGroupColumnwise,
    /// `/Graph Group Rowwise` — first row of the stashed range
    /// becomes X; succeeding rows A..F.
    GraphGroupRowwise,

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
    /// `:Format Lines` — paint thin borders on a POINT-selected range.
    /// Each `*Set` variant adds the corresponding edges; each `*Clear`
    /// removes them. `Outline` touches the perimeter of the range
    /// (top of top row, bottom of bottom row, left of leftmost col,
    /// right of rightmost col); `Left/Right/Top/Bottom` touch that one
    /// edge of every cell; `All` is every edge of every cell. Top and
    /// bottom edges are stored and round-trip via xlsx but the TTY
    /// renders only the vertical seams.
    FormatLinesOutlineSet,
    FormatLinesLeftSet,
    FormatLinesRightSet,
    FormatLinesTopSet,
    FormatLinesBottomSet,
    FormatLinesAllSet,
    FormatLinesOutlineClear,
    FormatLinesLeftClear,
    FormatLinesRightClear,
    FormatLinesTopClear,
    FormatLinesBottomClear,
    FormatLinesAllClear,
    /// `:Format Alignment Left` — left-align text in a range.
    FormatAlignmentLeft,
    /// `:Format Alignment Right` — right-align text in a range.
    FormatAlignmentRight,
    /// `:Format Alignment Center` — center text in a range.
    FormatAlignmentCenter,
    /// `:Format Alignment General` — clear the alignment override so
    /// the cell falls back to 1-2-3's default (label-prefix for
    /// labels, right for numbers).
    FormatAlignmentGeneral,
    /// `:Format Color Background <color>` — paint the cell background.
    FormatColorBgBlack,
    FormatColorBgWhite,
    FormatColorBgRed,
    FormatColorBgGreen,
    FormatColorBgBlue,
    FormatColorBgYellow,
    FormatColorBgCyan,
    FormatColorBgMagenta,
    /// `:Format Color Text <color>` — tint the cell foreground.
    FormatColorTextBlack,
    FormatColorTextWhite,
    FormatColorTextRed,
    FormatColorTextGreen,
    FormatColorTextBlue,
    FormatColorTextYellow,
    FormatColorTextCyan,
    FormatColorTextMagenta,
    /// `:Format Color Reset` — strip both background fill and font
    /// color from a range.
    FormatColorReset,
    /// `:Display Mode Color` — paper-look default cell background.
    DisplayModeColor,
    /// `:Display Mode B&W` — strip default cell color (terminal default).
    DisplayModeBW,
    /// `:Display Mode Reverse` — invert default cell colors.
    DisplayModeReverse,
    /// `:Display Options Grid Yes` — show row/column gutter.
    DisplayOptionsGridYes,
    /// `:Display Options Grid No` — hide row/column gutter.
    DisplayOptionsGridNo,
    /// `:Special Copy` — POINT for source range, then for destination.
    /// Replicates *every* formatting attribute (number format, text
    /// style, alignment, fill, font color/size/strike, borders) from
    /// each source cell onto the matching destination cell. Cell
    /// contents are untouched. Destination attributes the source
    /// doesn't carry are cleared (the dest ends up matching the source,
    /// not merged).
    SpecialCopy,
    /// `:Special Move` — same shape as `SpecialCopy`, but after the
    /// destination is written, every source cell *outside* the
    /// destination block has its formatting cleared.
    SpecialMove,

    /// `/Data Fill` — POINT for the fill range, then prompt for
    /// Start, Step, and Stop. Writes the arithmetic sequence into
    /// the range column-major (down each column, left to right),
    /// stopping when the next computed value would exceed Stop.
    DataFill,

    /// `/Data Sort Data-Range` — POINT for the rectangular range
    /// of records to sort. Stored on `App.data_sort` and surviving
    /// across visits to the Sort menu.
    DataSortDataRange,
    /// `/Data Sort Primary-Key` — POINT for any cell in the desired
    /// key column, then descend into the Asc/Desc submenu.
    DataSortPrimaryKey,
    /// `/Data Sort Secondary-Key` — same shape as Primary-Key, used
    /// as the tiebreaker when primary keys compare equal.
    DataSortSecondaryKey,
    /// `/Data Sort Reset` — clear data range and all key
    /// definitions; return to the Sort menu.
    DataSortReset,
    /// `/Data Sort Go` — execute the sort using the currently
    /// configured data range and keys. Refuses (no-op) when the
    /// data range or primary key is unset.
    DataSortGo,
    /// `/Data Sort Quit` — return to READY without executing.
    DataSortQuit,
    /// `/Data Sort … Ascending` — bind the in-flight key direction
    /// to ascending and re-enter the Sort menu.
    DataSortAscending,
    /// `/Data Sort … Descending` — bind the in-flight key direction
    /// to descending and re-enter the Sort menu.
    DataSortDescending,
    /// `/Data Sort Extra-Key` — POINT for any cell in a third
    /// tiebreaker column. Asc/Desc submenu writes the slot.
    DataSortExtraKey,

    /// `/Data Distribution` — POINT for the values range, then POINT
    /// for the bin range (single column). Writes a frequency
    /// distribution into the column immediately right of the bins,
    /// with one extra row at the bottom for the over-the-largest-bin
    /// overflow count.
    DataDistribution,

    /// `/Data Regression X-Range` — POINT for the independent-variable
    /// column. Univariate today; multivariate (multi-column X) is
    /// deferred.
    DataRegressionXRange,
    /// `/Data Regression Y-Range` — POINT for the dependent-variable
    /// column. Must be the same length as X-Range to participate in Go.
    DataRegressionYRange,
    /// `/Data Regression Output-Range` — POINT for the top-left cell
    /// of the labeled output table.
    DataRegressionOutputRange,
    /// `/Data Regression Intercept Compute` — fit the intercept term
    /// (default).
    DataRegressionInterceptCompute,
    /// `/Data Regression Intercept Zero` — fix the intercept at zero
    /// (force-through-origin regression).
    DataRegressionInterceptZero,
    /// `/Data Regression Reset` — clear all ranges and revert
    /// intercept to Compute.
    DataRegressionReset,
    /// `/Data Regression Go` — compute the regression and write the
    /// output table at the configured anchor.
    DataRegressionGo,
    /// `/Data Regression Quit` — return to READY without executing.
    DataRegressionQuit,

    /// `/Data Matrix Invert` — POINT a square matrix and an output
    /// anchor, then write the Gauss-Jordan inverse to the output.
    DataMatrixInvert,
    /// `/Data Matrix Multiply` — POINT matrix A, matrix B, and an
    /// output anchor; write A*B (cols(A) must equal rows(B)).
    DataMatrixMultiply,

    /// `/Data Parse Input-Column` — POINT for the column of long
    /// labels to split, including the format-line row at the top.
    DataParseInputColumn,
    /// `/Data Parse Output-Range` — POINT for the top-left cell of
    /// the parsed output rectangle.
    DataParseOutputRange,
    /// `/Data Parse Reset` — clear input column and output anchor.
    DataParseReset,
    /// `/Data Parse Go` — execute the parse using the configured
    /// input/output ranges and the in-cell format line.
    DataParseGo,
    /// `/Data Parse Quit` — return to READY without executing.
    DataParseQuit,

    /// `/Data Table 1` — POINT for the table range, then for
    /// Input cell 1; substitute each variable value in the left
    /// column into Input cell 1 and fill the body with the
    /// recalculated top-row formulas.
    DataTable1,
    /// `/Data Table 2` — POINT for the table range, then for
    /// Input cells 1 and 2. The corner cell holds the formula;
    /// the left column supplies values for Input cell 1 and the
    /// top row for Input cell 2. Body[i,j] is the formula
    /// evaluated with both substitutions applied.
    DataTable2,
    /// `/Data Table Reset` — return to READY without executing.
    /// Today the Table executors are stateless one-shots, so this
    /// is just a clean menu close.
    DataTableReset,
    /// `/Data Parse Format-Line Create` — auto-generate a format
    /// line above the first data row by splitting it on whitespace
    /// runs (each run becomes an `L` field). Refuses with a
    /// status-line error when no input column is set.
    DataParseFormatLineCreate,
    /// `/Data Parse Format-Line Edit` — open the format-line cell
    /// (top of the input column) in EDIT mode.
    DataParseFormatLineEdit,

    /// `/Data Query Input` — POINT for the input range. Top row
    /// is the field-name header; subsequent rows are records.
    DataQueryInput,
    /// `/Data Query Criteria` — POINT for the criteria range.
    /// Top row holds field-name headers (subset of the input
    /// header); subsequent rows are criterion records (rows OR'd,
    /// cells in a row AND'd).
    DataQueryCriteria,
    /// `/Data Query Output` — POINT for the output range. If the
    /// top row contains field names, Extract writes only those
    /// fields; otherwise all input fields are copied in input
    /// order.
    DataQueryOutput,
    /// `/Data Query Find` — move the pointer to the first record
    /// in the input range that matches the criteria.
    DataQueryFind,
    /// `/Data Query Extract` — copy all matching records into the
    /// output range below its header row.
    DataQueryExtract,
    /// `/Data Query Unique` — like Extract, but skips records whose
    /// extracted-field tuple duplicates an earlier-extracted one.
    DataQueryUnique,
    /// `/Data Query Del` — delete every matching record from the
    /// input range and shift subsequent records up. The header
    /// row stays put.
    DataQueryDel,
    /// `/Data Query Reset` — clear input, criteria, and output
    /// settings.
    DataQueryReset,
    /// `/Data Query Quit` — return to READY without executing.
    DataQueryQuit,

    /// `/Data Table 3` — placeholder action that surfaces a
    /// status-line "not yet implemented" message. The 3D table
    /// flavor needs multi-sheet POINT plumbing we haven't wired.
    DataTable3Stub,
    /// `/Data Table Labeled` — placeholder action; needs the
    /// lookup-by-named-range table semantics we haven't wired.
    DataTableLabeledStub,
    /// `/Data Query Modify` — placeholder action; the
    /// Extract/Replace/Insert sub-flow isn't wired.
    DataQueryModifyStub,
    /// Shared by every `/Data External` leaf — no external
    /// database driver is configured in L123.
    DataExternalStub,
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

/// Help page shown when the user presses F1 at the empty-path root of
/// the slash menu (`/`). Distinct from any per-`MenuItem` mapping
/// because the root has no item to attach to.
pub const ROOT_HELP_PAGE: &str = "0005-1-2-3-commands.html";

/// Walk the chain of items along `path` from the deepest one back
/// toward the root, returning the first non-empty
/// [`MenuItem::help_page`]. `None` means no item along the chain has
/// been wired yet — including the empty-path case, which the caller
/// can fall back to [`ROOT_HELP_PAGE`].
pub fn help_page_for_path(path: &[char]) -> Option<&'static str> {
    help_page_within(ROOT, path)
}

/// Nested variant of [`help_page_for_path`].
pub fn help_page_within(root: &'static [MenuItem], path: &[char]) -> Option<&'static str> {
    for end in (1..=path.len()).rev() {
        if let Some(item) = resolve_within(root, &path[..end]) {
            if !item.help_page.is_empty() {
                return Some(item.help_page);
            }
        }
    }
    None
}

// --------------------------------------------------------------------------
// The tree. Bottom-up so each level's parent can reference its children.
// --------------------------------------------------------------------------

const QUIT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not end l123 session; return to READY mode",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "End l123 session",
        help_page: "",
        body: MenuBody::Action(Action::QuitConfirm),
    },
];

/// Second confirmation shown when the user requests `/QY` against a
/// workbook with unsaved changes. Mirrors 1-2-3 R3.4a's "WORKSHEET
/// CHANGES NOT SAVED" guard. Surfaced via the highlighted item's
/// help text on control-panel line 3.
pub const QUIT_DIRTY_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "WORKSHEET CHANGES NOT SAVED — End l123 anyway?",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "WORKSHEET CHANGES NOT SAVED — End l123 anyway?",
        help_page: "",
        body: MenuBody::Action(Action::Quit),
    },
];

const WS_INSERT_SHEET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Insert a new worksheet before the current one",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetInsertSheetBefore),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Insert a new worksheet after the current one",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetInsertSheetAfter),
    },
];

const WS_INSERT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Insert one or more columns at the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetInsertColumn),
    },
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Insert one or more rows at the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetInsertRow),
    },
    MenuItem {
        letter: 'S',
        name: "Sheet",
        help: "Insert a new worksheet before or after the current one",
        help_page: "",
        body: MenuBody::Submenu(WS_INSERT_SHEET_MENU),
    },
];

const WS_DELETE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Delete one or more columns",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetDeleteColumn),
    },
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Delete one or more rows",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetDeleteRow),
    },
    MenuItem {
        letter: 'S',
        name: "Sheet",
        help: "Delete one or more worksheets",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetDeleteSheet),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Remove the current file from memory",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetDeleteFile),
    },
];

const WS_COLUMN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set-Width",
        help: "Set the width of the current column",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnSetWidth),
    },
    MenuItem {
        letter: 'R',
        name: "Reset-Width",
        help: "Reset column width to the global default",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnResetWidth),
    },
    MenuItem {
        letter: 'H',
        name: "Hide",
        help: "Hide columns",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnHide),
    },
    MenuItem {
        letter: 'D',
        name: "Display",
        help: "Redisplay hidden columns",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnDisplay),
    },
    MenuItem {
        letter: 'C',
        name: "Column-Range",
        help: "Set/reset width for a range of columns",
        help_page: "",
        body: MenuBody::Submenu(WS_COLUMN_RANGE_MENU),
    },
];

const WS_COLUMN_RANGE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set-Width",
        help: "Set a new width for each column in a range",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnRangeSetWidth),
    },
    MenuItem {
        letter: 'R',
        name: "Reset-Width",
        help: "Reset each column in a range to the global default",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnRangeResetWidth),
    },
];

const WS_ERASE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not erase the worksheet",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Erase ALL active files and start fresh",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetEraseConfirm),
    },
];

const WS_GLOBAL_RECALC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Natural",
        help: "Natural-order recalculation",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcNatural),
    },
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "Columnwise recalculation order",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcColumnwise),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "Rowwise recalculation order",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcRowwise),
    },
    MenuItem {
        letter: 'A',
        name: "Automatic",
        help: "Automatic recalculation after each entry",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcAutomatic),
    },
    MenuItem {
        letter: 'M',
        name: "Manual",
        help: "Manual recalculation — press F9 to recalc",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcManual),
    },
    MenuItem {
        letter: 'I',
        name: "Iteration",
        help: "Set iteration count (1-50)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalRecalcIteration),
    },
];

const WG_PROT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable global worksheet protection",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalProtectionEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable global worksheet protection",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalProtectionDisable),
    },
];

const WG_ZERO_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Show numeric zero values (default)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalZeroNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Hide numeric zero values",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalZeroYes),
    },
];

const WG_DEFAULT_OTHER_UNDO_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable the undo journal (Alt-F4 reverts mutating commands)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable the undo journal",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherUndoDisable),
    },
];

const WG_DEFAULT_OTHER_BEEP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable the soft terminal bell on edge collisions",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable the soft terminal bell on edge collisions",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherBeepDisable),
    },
];

const WG_DEFAULT_OTHER_INTL_PUNCTUATION_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Decimal . | argument , | thousands , (US default)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Decimal , | argument . | thousands .",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Decimal . | argument , | thousands (space)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Decimal , | argument . | thousands (space)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Decimal . | argument ; | thousands ,",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Decimal , | argument ; | thousands .",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationF),
    },
    MenuItem {
        letter: 'G',
        name: "G",
        help: "Decimal . | argument ; | thousands (space)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationG),
    },
    MenuItem {
        letter: 'H',
        name: "H",
        help: "Decimal , | argument ; | thousands (space)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlPunctuationH),
    },
];

const WG_DEFAULT_OTHER_INTL_CURRENCY_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Prefix",
        help: "Currency symbol leads the number ($1234)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlCurrencyPrefix),
    },
    MenuItem {
        letter: 'S',
        name: "Suffix",
        help: "Currency symbol trails the number (1234€)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlCurrencySuffix),
    },
];

const WG_DEFAULT_OTHER_INTL_DATE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "A",
        help: "MM/DD/YY (long), MM/DD (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlDateA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "DD/MM/YY (long), DD/MM (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlDateB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "DD.MM.YY (long), DD.MM (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlDateC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "YY-MM-DD (long), MM-DD (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlDateD),
    },
];

const WG_DEFAULT_OTHER_INTL_TIME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "A",
        help: "HH:MM:SS (long), HH:MM (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlTimeA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "HH.MM.SS (long), HH.MM (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlTimeB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "HH,MM,SS (long), HH,MM (short)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlTimeC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "HH:MM:SS (long), HH:MM (short) — colon fallback",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlTimeD),
    },
];

const WG_DEFAULT_OTHER_INTL_NEGATIVE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Parens",
        help: "Show negatives in parentheses: (1234.50)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlNegativeParens),
    },
    MenuItem {
        letter: 'S',
        name: "Sign",
        help: "Show negatives with a leading minus: -1234.50",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherIntlNegativeSign),
    },
];

const WG_DEFAULT_OTHER_INTL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Punctuation",
        help: "Decimal / argument / thousands character triple (A..H)",
        help_page: "0435-worksheet-global-default-other-international.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_PUNCTUATION_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Currency",
        help: "Currency symbol position (Prefix / Suffix) and string",
        help_page: "0436-worksheet-global-default-other-international-continued.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_CURRENCY_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Date",
        help: "International date style (D4 long, D5 short)",
        help_page: "0436-worksheet-global-default-other-international-continued.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_DATE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Time",
        help: "International time style (D8 long, D9 short)",
        help_page: "0436-worksheet-global-default-other-international-continued.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_TIME_MENU),
    },
    MenuItem {
        letter: 'N',
        name: "Negative",
        help: "Negative-number display: Parens or Sign",
        help_page: "0436-worksheet-global-default-other-international-continued.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_NEGATIVE_MENU),
    },
];

const WG_DEFAULT_OTHER_CLOCK_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Standard",
        help: "12-hour clock (DD-MMM-YY HH:MM AM/PM)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockStandard),
    },
    MenuItem {
        letter: 'I',
        name: "International",
        help: "24-hour clock (DD-MMM-YYYY HH:MM)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockInternational),
    },
    MenuItem {
        letter: 'N',
        name: "None",
        help: "Suppress the status-line clock",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockNone),
    },
    MenuItem {
        letter: 'F',
        name: "Filename",
        help: "Show the active workbook's filename instead of a clock",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockFilename),
    },
];

const WG_DEFAULT_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "International",
        help: "Locale-specific punctuation, dates, and currency",
        help_page: "0435-worksheet-global-default-other-international.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_INTL_MENU),
    },
    MenuItem {
        letter: 'H',
        name: "Help",
        help: "Help mode: Instant / Removable",
        help_page: "0434-worksheet-global-default-other-help.html",
        body: MenuBody::NotImplemented("wgdo-help"),
    },
    MenuItem {
        letter: 'C',
        name: "Clock",
        help: "Clock display: Standard / International / None / Filename",
        help_page: "0433-worksheet-global-default-other-clock.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_CLOCK_MENU),
    },
    MenuItem {
        letter: 'U',
        name: "Undo",
        help: "Enable/disable the undo journal",
        help_page: "0437-worksheet-global-default-other-undo.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_UNDO_MENU),
    },
    MenuItem {
        letter: 'B',
        name: "Beep",
        help: "Enable/disable the soft error beep on edge collisions",
        help_page: "0432-worksheet-global-default-other-beep.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_BEEP_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Expanded-Memory",
        help: "Expanded-memory options (legacy DOS)",
        help_page: "",
        body: MenuBody::NotImplemented("wgdo-ems"),
    },
];

const WG_DEFAULT_PRINTER_AUTOLF_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Send a line-feed after every carriage return",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterAutoLfYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not auto-feed after a carriage return",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterAutoLfNo),
    },
];

const WG_DEFAULT_PRINTER_WAIT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Pause between pages so single-sheet feeders can be loaded",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterWaitYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Print without pausing between pages",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterWaitNo),
    },
];

const WG_DEFAULT_PRINTER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "Interface",
        help: "Printer interface number (1..9)",
        help_page: "0442-worksheet-global-default-printer-interface.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterInterface),
    },
    MenuItem {
        letter: 'A',
        name: "AutoLf",
        help: "Auto line-feed after carriage return",
        help_page: "0443-worksheet-global-default-printer-autolf.html",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_AUTOLF_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Default left margin",
        help_page: "0444-worksheet-global-default-printer-left-right-top-bottom.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Default right margin",
        help_page: "0444-worksheet-global-default-printer-left-right-top-bottom.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginRight),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Default top margin",
        help_page: "0444-worksheet-global-default-printer-left-right-top-bottom.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginTop),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Default bottom margin",
        help_page: "0444-worksheet-global-default-printer-left-right-top-bottom.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterMarginBottom),
    },
    MenuItem {
        letter: 'P',
        name: "Pg-Length",
        help: "Default page length",
        help_page: "0445-worksheet-global-default-printer-pg-length.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterPgLength),
    },
    MenuItem {
        letter: 'W',
        name: "Wait",
        help: "Pause between pages",
        help_page: "0446-worksheet-global-default-printer-wait.html",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_WAIT_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Setup",
        help: "Default printer setup string",
        help_page: "0447-worksheet-global-default-printer-setup.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterSetup),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Default printer queue name",
        help_page: "0448-worksheet-global-default-printer-name.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterName),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Default menu",
        help_page: "0449-worksheet-global-default-quit.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultPrinterQuit),
    },
];

const WG_DEFAULT_AUTOEXEC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Run \\0 macro automatically when retrieving a file",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultAutoexecYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not run \\0 macro on retrieve",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultAutoexecNo),
    },
];

const WG_DEFAULT_EXT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Default extension applied when saving",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultExtSave),
    },
    MenuItem {
        letter: 'L',
        name: "List",
        help: "Default extension filter for /File List & /File Retrieve",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultExtList),
    },
];

const WG_DEFAULT_GRAPH_GROUP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "Auto-graph reads ranges columnwise",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphGroupColumnwise),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "Auto-graph reads ranges rowwise",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphGroupRowwise),
    },
];

const WG_DEFAULT_GRAPH_SAVE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Cgm",
        help: "Default /Graph Save format: CGM",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphSaveCgm),
    },
    MenuItem {
        letter: 'P',
        name: "Pic",
        help: "Default /Graph Save format: PIC",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultGraphSavePic),
    },
];

const WG_DEFAULT_GRAPH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Default auto-graph orientation (Columnwise/Rowwise)",
        help_page: "",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_GROUP_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Default /Graph Save format (Cgm/Pic)",
        help_page: "",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_SAVE_MENU),
    },
];

const WG_DEFAULT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Printer",
        help: "Default printer settings",
        help_page: "0441-worksheet-global-default-printer.html",
        body: MenuBody::Submenu(WG_DEFAULT_PRINTER_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Dir",
        help: "Default directory",
        help_page: "0428-worksheet-global-default-dir.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultDir),
    },
    MenuItem {
        letter: 'S',
        name: "Status",
        help: "Display global default settings",
        help_page: "0450-worksheet-global-default-status.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultStatus),
    },
    MenuItem {
        letter: 'U',
        name: "Update",
        help: "Persist defaults to the config file",
        help_page: "0454-worksheet-global-default-update.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultUpdate),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Miscellaneous defaults (Undo, International, Clock, …)",
        help_page: "0431-worksheet-global-default-other.html",
        body: MenuBody::Submenu(WG_DEFAULT_OTHER_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Autoexec",
        help: "Run autoexec macro (\\0) on file retrieve",
        help_page: "0427-worksheet-global-default-autoexec.html",
        body: MenuBody::Submenu(WG_DEFAULT_AUTOEXEC_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Ext",
        help: "Default file extensions",
        help_page: "0429-worksheet-global-default-ext.html",
        body: MenuBody::Submenu(WG_DEFAULT_EXT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Default graph settings",
        help_page: "0430-worksheet-global-default-graph.html",
        body: MenuBody::Submenu(WG_DEFAULT_GRAPH_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Temp",
        help: "Temporary file directory",
        help_page: "0453-worksheet-global-default-temp.html",
        body: MenuBody::Action(Action::WorksheetGlobalDefaultTemp),
    },
];

const WG_LABEL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Default to left-aligned label prefix (')",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalLabelLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Default to right-aligned label prefix (\")",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalLabelRight),
    },
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Default to centered label prefix (^)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalLabelCenter),
    },
];

const WG_GROUP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Enable GROUP: format/column/row ops propagate across sheets",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalGroupEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable GROUP: commands affect only the current sheet",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalGroupDisable),
    },
];

const WS_GLOBAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Set global cell display format",
        help_page: "0426-worksheet-global-format.html",
        body: MenuBody::Submenu(WG_FORMAT_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Set global default label prefix",
        help_page: "0420-worksheet-global-label.html",
        body: MenuBody::Submenu(WG_LABEL_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Col-Width",
        help: "Set global default column width",
        help_page: "0418-worksheet-global-col-width.html",
        body: MenuBody::Action(Action::WorksheetGlobalColWidth),
    },
    MenuItem {
        letter: 'P',
        name: "Prot",
        help: "Enable/disable worksheet protection",
        help_page: "0421-worksheet-global-prot.html",
        body: MenuBody::Submenu(WG_PROT_MENU),
    },
    MenuItem {
        letter: 'Z',
        name: "Zero",
        help: "Zero-value display: No/Yes",
        help_page: "0424-worksheet-global-zero.html",
        body: MenuBody::Submenu(WG_ZERO_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Recalc",
        help: "Recalculation mode",
        help_page: "0422-worksheet-global-recalc.html",
        body: MenuBody::Submenu(WS_GLOBAL_RECALC_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Default",
        help: "Default settings (printer, dir, other, ...)",
        help_page: "0425-worksheet-global-default.html",
        body: MenuBody::Submenu(WG_DEFAULT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Enable/disable GROUP mode across sheets",
        help_page: "0419-worksheet-global-group.html",
        body: MenuBody::Submenu(WG_GROUP_MENU),
    },
];

const WS_TITLES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Both",
        help: "Freeze rows above and columns left of the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetTitlesBoth),
    },
    MenuItem {
        letter: 'H',
        name: "Horizontal",
        help: "Freeze the rows above the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetTitlesHorizontal),
    },
    MenuItem {
        letter: 'V',
        name: "Vertical",
        help: "Freeze the columns left of the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetTitlesVertical),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Clear frozen titles on the current sheet",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetTitlesClear),
    },
];

const WS_HIDE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Enable",
        help: "Hide the current worksheet",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetHideEnable),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Redisplay every hidden worksheet",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetHideDisable),
    },
];

const WORKSHEET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Global",
        help: "Set worksheet-wide options",
        help_page: "0417-worksheet-global.html",
        body: MenuBody::Submenu(WS_GLOBAL_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Insert",
        help: "Insert a row, column, or sheet",
        help_page: "0462-worksheet-insert.html",
        body: MenuBody::Submenu(WS_INSERT_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Delete a row, column, sheet, or file",
        help_page: "0458-worksheet-delete.html",
        body: MenuBody::Submenu(WS_DELETE_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Column width and visibility",
        help_page: "0455-worksheet-column.html",
        body: MenuBody::Submenu(WS_COLUMN_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Erase all active files",
        help_page: "0460-worksheet-erase.html",
        body: MenuBody::Submenu(WS_ERASE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Titles",
        help: "Freeze rows and/or columns as titles",
        help_page: "0465-worksheet-titles.html",
        body: MenuBody::Submenu(WS_TITLES_MENU),
    },
    MenuItem {
        letter: 'W',
        name: "Window",
        help: "Split window into panes",
        help_page: "0467-worksheet-window.html",
        body: MenuBody::NotImplemented("ws-window"),
    },
    MenuItem {
        letter: 'S',
        name: "Status",
        help: "Show worksheet status panel",
        help_page: "0464-worksheet-status.html",
        body: MenuBody::Action(Action::WorksheetStatus),
    },
    MenuItem {
        letter: 'P',
        name: "Page",
        help: "Insert a manual row or column page break at the pointer",
        help_page: "0463-worksheet-page.html",
        body: MenuBody::Submenu(WS_PAGE_MENU),
    },
    MenuItem {
        letter: 'H',
        name: "Hide",
        help: "Hide/show entire sheets",
        help_page: "0461-worksheet-hide.html",
        body: MenuBody::Submenu(WS_HIDE_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Learn",
        help: "Define / cancel / erase the Learn range",
        help_page: "",
        body: MenuBody::Submenu(WS_LEARN_MENU),
    },
];

const WS_PAGE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Insert a manual row page break at the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetPageRow),
    },
    MenuItem {
        letter: 'C',
        name: "Column",
        help: "Insert a manual column page break at the cell pointer",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetPageColumn),
    },
];

const WS_LEARN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Range",
        help: "Set the destination range for Alt-F5 LEARN recordings",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetLearnRange),
    },
    MenuItem {
        letter: 'C',
        name: "Cancel",
        help: "Forget the learn range and stop any in-progress recording",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetLearnCancel),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Blank the cells of the learn range, keeping the range itself",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetLearnErase),
    },
];

const RANGE_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fixed",
        help: "Fixed number of decimal places",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatFixed),
    },
    MenuItem {
        letter: 'S',
        name: "Sci",
        help: "Scientific notation",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatScientific),
    },
    MenuItem {
        letter: 'C',
        name: "Currency",
        help: "Currency format with symbol",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatCurrency),
    },
    MenuItem {
        letter: ',',
        name: ",",
        help: "Comma-separated (no currency symbol)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatComma),
    },
    MenuItem {
        letter: 'G',
        name: "General",
        help: "General format (default)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatGeneral),
    },
    MenuItem {
        letter: 'P',
        name: "Percent",
        help: "Percent (value × 100) with % sign",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatPercent),
    },
    MenuItem {
        letter: '+',
        name: "+/-",
        help: "Bar chart: each unit prints as + (positive) or - (negative)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatPlusMinus),
    },
    MenuItem {
        letter: 'D',
        name: "Date",
        help: "Date format (select D1..D5, Time for D6..D9)",
        help_page: "",
        body: MenuBody::Submenu(RANGE_FORMAT_DATE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Show formulas instead of values",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatText),
    },
    MenuItem {
        letter: 'H',
        name: "Hidden",
        help: "Hide the cell display",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatHidden),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Automatic, negative-color, label-only, parentheses",
        help_page: "0398-range-format-other-and-worksheet-global-format-other.html",
        body: MenuBody::Submenu(RANGE_FORMAT_OTHER_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Revert to global format",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatReset),
    },
];

const RANGE_FORMAT_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "Automatic",
        help: "Pick a format based on the cell's current value",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatAutomatic),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Color negative values; or reset to default",
        help_page: "",
        body: MenuBody::Submenu(RANGE_FORMAT_OTHER_COLOR_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Treat the cell as label-only (numeric input becomes a label)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatLabelOnly),
    },
    MenuItem {
        letter: 'P',
        name: "Parentheses",
        help: "Wrap numeric output in parentheses",
        help_page: "",
        body: MenuBody::Submenu(RANGE_FORMAT_OTHER_PAREN_MENU),
    },
];

const RANGE_FORMAT_OTHER_COLOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Negative",
        help: "Render negative values in a distinct color",
        help_page: "",
        body: MenuBody::Submenu(RANGE_FORMAT_OTHER_COLOR_NEG_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Drop the negative-value color override",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorReset),
    },
];

const RANGE_FORMAT_OTHER_COLOR_NEG_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Black",
        help: "Tint negative values black",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorBlack),
    },
    MenuItem {
        letter: 'W',
        name: "White",
        help: "Tint negative values white",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorWhite),
    },
    MenuItem {
        letter: 'R',
        name: "Red",
        help: "Tint negative values red",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorRed),
    },
    MenuItem {
        letter: 'G',
        name: "Green",
        help: "Tint negative values green",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorGreen),
    },
    MenuItem {
        letter: 'L',
        name: "Blue",
        help: "Tint negative values blue",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorBlue),
    },
    MenuItem {
        letter: 'Y',
        name: "Yellow",
        help: "Tint negative values yellow",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorYellow),
    },
    MenuItem {
        letter: 'C',
        name: "Cyan",
        help: "Tint negative values cyan",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorCyan),
    },
    MenuItem {
        letter: 'M',
        name: "Magenta",
        help: "Tint negative values magenta",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatNegColorMagenta),
    },
];

const RANGE_FORMAT_OTHER_PAREN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Wrap numeric output in parentheses",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatParensYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Stop wrapping numeric output in parentheses",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatParensNo),
    },
];

const RANGE_FORMAT_DATE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "DD-MMM-YY",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatDateDmy),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "DD-MMM",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatDateDm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "MMM-YY",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatDateMy),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Long international (MM/DD/YY)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatDateLongIntl),
    },
    MenuItem {
        letter: '5',
        name: "5",
        help: "Short international (MM/DD)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatDateShortIntl),
    },
    MenuItem {
        letter: 'T',
        name: "Time",
        help: "Time format (D6..D9)",
        help_page: "",
        body: MenuBody::Submenu(RANGE_FORMAT_TIME_MENU),
    },
];

const RANGE_FORMAT_TIME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "HH:MM:SS AM/PM (D6)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatTimeHmsAmPm),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "HH:MM AM/PM (D7)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatTimeHmAmPm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "Long international time (D8)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatTimeLongIntl),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Short international time (D9)",
        help_page: "",
        body: MenuBody::Action(Action::RangeFormatTimeShortIntl),
    },
];

const WG_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fixed",
        help: "Default to fixed number of decimal places",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatFixed),
    },
    MenuItem {
        letter: 'S',
        name: "Sci",
        help: "Default to scientific notation",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatScientific),
    },
    MenuItem {
        letter: 'C',
        name: "Currency",
        help: "Default to currency format with symbol",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatCurrency),
    },
    MenuItem {
        letter: ',',
        name: ",",
        help: "Default to comma-separated (no currency symbol)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatComma),
    },
    MenuItem {
        letter: 'G',
        name: "General",
        help: "Default to General format",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatGeneral),
    },
    MenuItem {
        letter: 'P',
        name: "Percent",
        help: "Default to Percent (value × 100) with % sign",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatPercent),
    },
    MenuItem {
        letter: '+',
        name: "+/-",
        help: "Default to bar chart (each unit prints as + or -)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatPlusMinus),
    },
    MenuItem {
        letter: 'D',
        name: "Date",
        help: "Default to a Date format (D1..D5, Time for D6..D9)",
        help_page: "",
        body: MenuBody::Submenu(WG_FORMAT_DATE_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Default to showing formulas instead of values",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatText),
    },
    MenuItem {
        letter: 'H',
        name: "Hidden",
        help: "Default to hiding cell display",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatHidden),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Automatic, negative-color, label-only, parentheses",
        help_page: "0398-range-format-other-and-worksheet-global-format-other.html",
        body: MenuBody::Submenu(WG_FORMAT_OTHER_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Reset global format to General",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatReset),
    },
];

const WG_FORMAT_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "Automatic",
        help: "Default to picking a format from the cell's current value",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatAutomatic),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Color negative values; or reset to default",
        help_page: "",
        body: MenuBody::Submenu(WG_FORMAT_OTHER_COLOR_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Default to treating new cells as label-only",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatLabelOnly),
    },
    MenuItem {
        letter: 'P',
        name: "Parentheses",
        help: "Default to wrapping numeric output in parentheses",
        help_page: "",
        body: MenuBody::Submenu(WG_FORMAT_OTHER_PAREN_MENU),
    },
];

const WG_FORMAT_OTHER_COLOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Negative",
        help: "Render negative values in a distinct color",
        help_page: "",
        body: MenuBody::Submenu(WG_FORMAT_OTHER_COLOR_NEG_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Drop the negative-value color override",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorReset),
    },
];

const WG_FORMAT_OTHER_COLOR_NEG_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Black",
        help: "Default to tinting negative values black",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorBlack),
    },
    MenuItem {
        letter: 'W',
        name: "White",
        help: "Default to tinting negative values white",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorWhite),
    },
    MenuItem {
        letter: 'R',
        name: "Red",
        help: "Default to tinting negative values red",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorRed),
    },
    MenuItem {
        letter: 'G',
        name: "Green",
        help: "Default to tinting negative values green",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorGreen),
    },
    MenuItem {
        letter: 'L',
        name: "Blue",
        help: "Default to tinting negative values blue",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorBlue),
    },
    MenuItem {
        letter: 'Y',
        name: "Yellow",
        help: "Default to tinting negative values yellow",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorYellow),
    },
    MenuItem {
        letter: 'C',
        name: "Cyan",
        help: "Default to tinting negative values cyan",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorCyan),
    },
    MenuItem {
        letter: 'M',
        name: "Magenta",
        help: "Default to tinting negative values magenta",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatNegColorMagenta),
    },
];

const WG_FORMAT_OTHER_PAREN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Wrap numeric output in parentheses",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatParensYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Stop wrapping numeric output in parentheses",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatParensNo),
    },
];

const WG_FORMAT_DATE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "DD-MMM-YY",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateDmy),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "DD-MMM",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateDm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "MMM-YY",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateMy),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Long international (MM/DD/YY)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateLongIntl),
    },
    MenuItem {
        letter: '5',
        name: "5",
        help: "Short international (MM/DD)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatDateShortIntl),
    },
    MenuItem {
        letter: 'T',
        name: "Time",
        help: "Time format (D6..D9)",
        help_page: "",
        body: MenuBody::Submenu(WG_FORMAT_DATE_TIME_MENU),
    },
];

const WG_FORMAT_DATE_TIME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "HH:MM:SS AM/PM (D6)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatTimeHmsAmPm),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "HH:MM AM/PM (D7)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatTimeHmAmPm),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "Long international time (D8)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatTimeLongIntl),
    },
    MenuItem {
        letter: '4',
        name: "4",
        help: "Short international time (D9)",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetGlobalFormatTimeShortIntl),
    },
];

const RANGE_LABEL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Left-align label prefix for range",
        help_page: "",
        body: MenuBody::Action(Action::RangeLabelLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Right-align label prefix for range",
        help_page: "",
        body: MenuBody::Action(Action::RangeLabelRight),
    },
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Center label prefix for range",
        help_page: "",
        body: MenuBody::Action(Action::RangeLabelCenter),
    },
];

const RANGE_SEARCH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Formulas",
        help: "Search only formula cells",
        help_page: "",
        body: MenuBody::Action(Action::RangeSearchFormulas),
    },
    MenuItem {
        letter: 'L',
        name: "Labels",
        help: "Search only label cells",
        help_page: "",
        body: MenuBody::Action(Action::RangeSearchLabels),
    },
    MenuItem {
        letter: 'B',
        name: "Both",
        help: "Search both formulas and labels",
        help_page: "",
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
        help_page: "",
        body: MenuBody::Action(Action::RangeSearchFind),
    },
    MenuItem {
        letter: 'R',
        name: "Replace",
        help: "Replace all matches with another string",
        help_page: "",
        body: MenuBody::Action(Action::RangeSearchReplace),
    },
];

const RANGE_NAME_LABELS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Names point one cell to the right of each label",
        help_page: "",
        body: MenuBody::Action(Action::RangeNameLabelsRight),
    },
    MenuItem {
        letter: 'D',
        name: "Down",
        help: "Names point one cell below each label",
        help_page: "",
        body: MenuBody::Action(Action::RangeNameLabelsDown),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Names point one cell to the left of each label",
        help_page: "",
        body: MenuBody::Action(Action::RangeNameLabelsLeft),
    },
    MenuItem {
        letter: 'U',
        name: "Up",
        help: "Names point one cell above each label",
        help_page: "",
        body: MenuBody::Action(Action::RangeNameLabelsUp),
    },
];

const RANGE_NAME_NOTE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Attach a note to a named range",
        help_page: "0380-range-name-note-create.html",
        body: MenuBody::Action(Action::RangeNameNoteCreate),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Remove a named range's note",
        help_page: "0381-range-name-note-delete.html",
        body: MenuBody::Action(Action::RangeNameNoteDelete),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Remove every named-range note",
        help_page: "0382-range-name-note-reset.html",
        body: MenuBody::Action(Action::RangeNameNoteReset),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Dump names + notes to a 3-column block",
        help_page: "0383-range-name-note-table.html",
        body: MenuBody::Action(Action::RangeNameNoteTable),
    },
];

const RANGE_NAME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Create a new named range",
        help_page: "0375-range-name-create.html",
        body: MenuBody::Action(Action::RangeNameCreate),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Delete a named range",
        help_page: "0377-range-name-delete.html",
        body: MenuBody::Action(Action::RangeNameDelete),
    },
    MenuItem {
        letter: 'L',
        name: "Labels",
        help: "Create range names from adjacent labels",
        help_page: "0378-range-name-labels.html",
        body: MenuBody::Submenu(RANGE_NAME_LABELS_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Delete all range names",
        help_page: "0384-range-name-reset.html",
        body: MenuBody::Action(Action::RangeNameReset),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Write a table of range names to the sheet",
        help_page: "0385-range-name-table.html",
        body: MenuBody::Action(Action::RangeNameTable),
    },
    MenuItem {
        letter: 'U',
        name: "Undefine",
        help: "Remove a range name but preserve formula values",
        help_page: "0386-range-name-undefine.html",
        body: MenuBody::Action(Action::RangeNameUndefine),
    },
    MenuItem {
        letter: 'N',
        name: "Note",
        help: "Notes attached to named ranges",
        help_page: "0379-range-name-note.html",
        body: MenuBody::Submenu(RANGE_NAME_NOTE_MENU),
    },
];

const RANGE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Set display format for a range",
        help_page: "0394-range-format.html",
        body: MenuBody::Submenu(RANGE_FORMAT_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Label",
        help: "Change label prefix for a range",
        help_page: "0371-range-label.html",
        body: MenuBody::Submenu(RANGE_LABEL_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Erase the contents of a range",
        help_page: "0368-range-erase.html",
        body: MenuBody::Action(Action::RangeErase),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Named ranges",
        help_page: "0372-range-name.html",
        body: MenuBody::Submenu(RANGE_NAME_MENU),
    },
    MenuItem {
        letter: 'J',
        name: "Justify",
        help: "Word-wrap long labels into a block",
        help_page: "0370-range-justify.html",
        body: MenuBody::Action(Action::RangeJustify),
    },
    MenuItem {
        letter: 'P',
        name: "Prot",
        help: "Re-protect a range on a protected sheet",
        help_page: "0387-range-prot.html",
        body: MenuBody::Action(Action::RangeProtect),
    },
    MenuItem {
        letter: 'U',
        name: "Unprot",
        help: "Mark a range as writable",
        help_page: "0392-range-unprot.html",
        body: MenuBody::Action(Action::RangeUnprotect),
    },
    MenuItem {
        letter: 'I',
        name: "Input",
        help: "Form input limited to unprotected cells",
        help_page: "0369-range-input.html",
        body: MenuBody::Action(Action::RangeInput),
    },
    MenuItem {
        letter: 'V',
        name: "Value",
        help: "Copy a range converting formulas to values",
        help_page: "0393-range-value.html",
        body: MenuBody::Action(Action::RangeValue),
    },
    MenuItem {
        letter: 'T',
        name: "Trans",
        help: "Transpose rows and columns",
        help_page: "0390-range-trans.html",
        body: MenuBody::Action(Action::RangeTrans),
    },
    MenuItem {
        letter: 'S',
        name: "Search",
        help: "Find / Replace across formulas and labels",
        help_page: "0388-range-search.html",
        body: MenuBody::Submenu(RANGE_SEARCH_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Compare",
        help: "Compare two ranges into a third (POINT both, then anchor)",
        help_page: "",
        body: MenuBody::Action(Action::RangeCompare),
    },
];

const FILE_LIST_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "List worksheet files in the session directory",
        help_page: "",
        body: MenuBody::Action(Action::FileListWorksheet),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "List print settings files",
        help_page: "",
        body: MenuBody::NotImplemented("f-list-print"),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "List graph files",
        help_page: "",
        body: MenuBody::NotImplemented("f-list-graph"),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "List any file",
        help_page: "",
        body: MenuBody::Action(Action::FileListOther),
    },
    MenuItem {
        letter: 'A',
        name: "Active",
        help: "List currently active files",
        help_page: "",
        body: MenuBody::Action(Action::FileListActive),
    },
    MenuItem {
        letter: 'L',
        name: "Linked",
        help: "List files linked via formula references",
        help_page: "",
        body: MenuBody::NotImplemented("f-list-linked"),
    },
];

const FILE_NEW_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Create a new active file before the current one",
        help_page: "",
        body: MenuBody::Action(Action::FileNew),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Create a new active file after the current one",
        help_page: "",
        body: MenuBody::Action(Action::FileNew),
    },
];

const FILE_OPEN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Before",
        help: "Open a file as a new active file before the current one",
        help_page: "",
        body: MenuBody::Action(Action::FileOpenBefore),
    },
    MenuItem {
        letter: 'A',
        name: "After",
        help: "Open a file as a new active file after the current one",
        help_page: "",
        body: MenuBody::Action(Action::FileOpenAfter),
    },
];

const FILE_IMPORT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Import each line as a label in one column",
        help_page: "",
        body: MenuBody::Action(Action::FileImportText),
    },
    MenuItem {
        letter: 'N',
        name: "Numbers",
        help: "Parse CSV: numeric tokens as values, quoted strings as labels",
        help_page: "",
        body: MenuBody::Action(Action::FileImportNumbers),
    },
];

const FILE_COMBINE_RANGE_KIND_HELP_COPY: &str = "Combine source by overwriting target cells";
const FILE_COMBINE_RANGE_KIND_HELP_ADD: &str = "Combine source by adding to target cells";
const FILE_COMBINE_RANGE_KIND_HELP_SUBTRACT: &str =
    "Combine source by subtracting from target cells";

const FILE_COMBINE_COPY_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Entire-File",
        help: FILE_COMBINE_RANGE_KIND_HELP_COPY,
        help_page: "",
        body: MenuBody::Action(Action::FileCombineCopyEntire),
    },
    MenuItem {
        letter: 'N',
        name: "Named/Specified-Range",
        help: "Combine a source range like A1..C5 by overwriting target cells",
        help_page: "",
        body: MenuBody::Action(Action::FileCombineCopyNamed),
    },
];

const FILE_COMBINE_ADD_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Entire-File",
        help: FILE_COMBINE_RANGE_KIND_HELP_ADD,
        help_page: "",
        body: MenuBody::Action(Action::FileCombineAddEntire),
    },
    MenuItem {
        letter: 'N',
        name: "Named/Specified-Range",
        help: "Combine a source range like A1..C5 by adding to target cells",
        help_page: "",
        body: MenuBody::Action(Action::FileCombineAddNamed),
    },
];

const FILE_COMBINE_SUBTRACT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'E',
        name: "Entire-File",
        help: FILE_COMBINE_RANGE_KIND_HELP_SUBTRACT,
        help_page: "",
        body: MenuBody::Action(Action::FileCombineSubtractEntire),
    },
    MenuItem {
        letter: 'N',
        name: "Named/Specified-Range",
        help: "Combine a source range like A1..C5 by subtracting from target cells",
        help_page: "",
        body: MenuBody::Action(Action::FileCombineSubtractNamed),
    },
];

const FILE_COMBINE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Copy",
        help: "Overwrite the target cells with the source",
        help_page: "",
        body: MenuBody::Submenu(FILE_COMBINE_COPY_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Add",
        help: "Add the source values to the target cells",
        help_page: "",
        body: MenuBody::Submenu(FILE_COMBINE_ADD_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Subtract",
        help: "Subtract the source values from the target cells",
        help_page: "",
        body: MenuBody::Submenu(FILE_COMBINE_SUBTRACT_MENU),
    },
];

const FILE_ERASE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "Delete a worksheet file (.xlsx, .wk*) from disk",
        help_page: "",
        body: MenuBody::Action(Action::FileEraseWorksheet),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "Delete a print-settings file from disk",
        help_page: "",
        body: MenuBody::Action(Action::FileErasePrint),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Delete a graph file from disk",
        help_page: "",
        body: MenuBody::Action(Action::FileEraseGraph),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Delete any file from disk",
        help_page: "",
        body: MenuBody::Action(Action::FileEraseOther),
    },
];

const FILE_ADMIN_RESERVATION_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Get",
        help: "Acquire the file's reservation for editing",
        help_page: "0020-file-admin-reservation-get-and-release.html",
        body: MenuBody::Action(Action::FileAdminReservationGet),
    },
    MenuItem {
        letter: 'R',
        name: "Release",
        help: "Release the file's reservation",
        help_page: "0020-file-admin-reservation-get-and-release.html",
        body: MenuBody::Action(Action::FileAdminReservationRelease),
    },
];

const FILE_ADMIN_SEAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Seal the file with a password",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminSealFile),
    },
    MenuItem {
        letter: 'R',
        name: "Reservation-Setting",
        help: "Seal the reservation behavior of this file",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminSealReservationSetting),
    },
    MenuItem {
        letter: 'D',
        name: "Disable",
        help: "Disable an existing seal (requires the seal password)",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminSealDisable),
    },
];

const FILE_ADMIN_TABLE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "Build a table of worksheet files in the session directory",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTableWorksheet),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "Build a table of print-settings files",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTablePrint),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Build a table of graph files",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTableGraph),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Build a table of any file type",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTableOther),
    },
    MenuItem {
        letter: 'A',
        name: "Active",
        help: "Build a table of currently active files",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTableActive),
    },
    MenuItem {
        letter: 'L',
        name: "Linked",
        help: "Build a table of files linked via formula references",
        help_page: "",
        body: MenuBody::Action(Action::FileAdminTableLinked),
    },
];

const FILE_ADMIN_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Reservation",
        help: "Get or release the file's edit reservation",
        help_page: "0019-file-admin-reservation.html",
        body: MenuBody::Submenu(FILE_ADMIN_RESERVATION_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Seal",
        help: "Seal the file or its reservation setting with a password",
        help_page: "0024-file-admin-seal.html",
        body: MenuBody::Submenu(FILE_ADMIN_SEAL_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Build a table of files of a given type",
        help_page: "0026-file-admin-table.html",
        body: MenuBody::Submenu(FILE_ADMIN_TABLE_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "Link-Refresh",
        help: "Refresh formulas that reference linked files",
        help_page: "0018-file-admin-link-refresh.html",
        body: MenuBody::Action(Action::FileAdminLinkRefresh),
    },
];

const FILE_XTRACT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Formulas",
        help: "Save the range with formulas intact",
        help_page: "",
        body: MenuBody::Action(Action::FileXtractFormulas),
    },
    MenuItem {
        letter: 'V',
        name: "Values",
        help: "Save the range with formulas replaced by their values",
        help_page: "",
        body: MenuBody::Action(Action::FileXtractValues),
    },
];

const FILE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'R',
        name: "Retrieve",
        help: "Replace all active files with one from disk",
        help_page: "0039-file-retrieve.html",
        body: MenuBody::Action(Action::FileRetrieve),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Save all active files",
        help_page: "0041-file-save.html",
        body: MenuBody::Action(Action::FileSave),
    },
    MenuItem {
        letter: 'C',
        name: "Combine",
        help: "Merge a file into the current one",
        help_page: "0028-file-combine.html",
        body: MenuBody::Submenu(FILE_COMBINE_MENU),
    },
    MenuItem {
        letter: 'X',
        name: "Xtract",
        help: "Save a range as a new file",
        help_page: "0045-file-xtract.html",
        body: MenuBody::Submenu(FILE_XTRACT_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "Erase",
        help: "Delete a file on disk",
        help_page: "0031-file-erase.html",
        body: MenuBody::Submenu(FILE_ERASE_MENU),
    },
    MenuItem {
        letter: 'L',
        name: "List",
        help: "Overlay list of files on disk",
        help_page: "0035-file-list.html",
        body: MenuBody::Submenu(FILE_LIST_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Import",
        help: "Import text or delimited numbers",
        help_page: "0032-file-import.html",
        body: MenuBody::Submenu(FILE_IMPORT_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Dir",
        help: "Change the session directory",
        help_page: "0030-file-dir.html",
        body: MenuBody::Action(Action::FileDir),
    },
    MenuItem {
        letter: 'N',
        name: "New",
        help: "Create a new active file",
        help_page: "0036-file-new.html",
        body: MenuBody::Submenu(FILE_NEW_MENU),
    },
    MenuItem {
        letter: 'O',
        name: "Open",
        help: "Open another file alongside the current one",
        help_page: "0037-file-open.html",
        body: MenuBody::Submenu(FILE_OPEN_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Admin",
        help: "Reservation, seal, link-refresh",
        help_page: "0017-file-admin.html",
        body: MenuBody::Submenu(FILE_ADMIN_MENU),
    },
];

const PRINT_FILE_OPTIONS_MARGINS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Set the left margin (spaces prefixed to each printed line)",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Set the right margin",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginRight),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Set the top margin (blank lines above the first row)",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginTop),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Set the bottom margin",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginBottom),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Options menu",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsMarginsQuit),
    },
];

const PRINT_FILE_OPTIONS_OTHER_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "As-Displayed",
        help: "Print cells as they appear on screen (default)",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherAsDisplayed),
    },
    MenuItem {
        letter: 'C',
        name: "Cell-Formulas",
        help: "Print the formula source in place of the computed value",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherCellFormulas),
    },
    MenuItem {
        letter: 'F',
        name: "Formatted",
        help: "Honor headers, footers, margins, and page breaks",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherFormatted),
    },
    MenuItem {
        letter: 'U',
        name: "Unformatted",
        help: "Dump the range with no page decorations",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsOtherUnformatted),
    },
];

const PRINT_FILE_OPTIONS_ADVANCED_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "AutoLf",
        help: "Send LF after each CR (Auto Line Feed)",
        help_page: "0302-print-background-encoded-printer-options-advanced-autolf.html",
        body: MenuBody::NotImplemented("pfo-advanced-autolf"),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Color print options",
        help_page: "0303-print-background-encoded-printer-options-advanced-color.html",
        body: MenuBody::NotImplemented("pfo-advanced-color"),
    },
    MenuItem {
        letter: 'D',
        name: "Device",
        help: "CUPS printer queue name (lp -d <name>)",
        help_page: "0304-print-background-encoded-printer-options-advanced-device.html",
        body: MenuBody::Action(Action::PrintSessionOptionsAdvancedDevice),
    },
    MenuItem {
        letter: 'F',
        name: "Fonts",
        help: "Font selection",
        help_page: "0305-print-background-encoded-printer-options-advanced-fonts.html",
        body: MenuBody::NotImplemented("pfo-advanced-fonts"),
    },
    MenuItem {
        letter: 'I',
        name: "Images",
        help: "Image rendering options",
        help_page: "0307-print-background-encoded-printer-options-advanced-image.html",
        body: MenuBody::NotImplemented("pfo-advanced-images"),
    },
    MenuItem {
        letter: 'L',
        name: "Layout",
        help: "Layout options",
        help_page: "0311-print-background-encoded-printer-options-advanced-layout.html",
        body: MenuBody::NotImplemented("pfo-advanced-layout"),
    },
    MenuItem {
        letter: 'P',
        name: "Priority",
        help: "Print job priority",
        help_page: "0315-print-background-printer-options-advanced-priority.html",
        body: MenuBody::NotImplemented("pfo-advanced-priority"),
    },
    MenuItem {
        letter: 'W',
        name: "Wait",
        help: "Wait between pages",
        help_page: "0316-print-printer-options-advanced-wait.html",
        body: MenuBody::NotImplemented("pfo-advanced-wait"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the Options submenu",
        help_page: "",
        body: MenuBody::Action(Action::PrintSessionOptionsAdvancedQuit),
    },
];

const PRINT_FILE_OPTIONS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'H',
        name: "Header",
        help: "Set a three-part header (L|C|R), printed above each page",
        help_page: "0283-print-background-encoded-file-printer-options-header-and-footer.html",
        body: MenuBody::Action(Action::PrintSessionOptionsHeader),
    },
    MenuItem {
        letter: 'F',
        name: "Footer",
        help: "Set a three-part footer (L|C|R), printed below each page",
        help_page: "0283-print-background-encoded-file-printer-options-header-and-footer.html",
        body: MenuBody::Action(Action::PrintSessionOptionsFooter),
    },
    MenuItem {
        letter: 'M',
        name: "Margins",
        help: "Set Left / Right / Top / Bottom margins",
        help_page: "0285-print-background-encoded-file-printer-options-margins.html",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_MARGINS_MENU),
    },
    MenuItem {
        letter: 'P',
        name: "Pg-Length",
        help: "Set page length (lines per page)",
        help_page: "0297-print-background-encoded-file-printer-options-pg-length.html",
        body: MenuBody::Action(Action::PrintSessionOptionsPgLength),
    },
    MenuItem {
        letter: 'B',
        name: "Borders",
        help: "Repeat columns/rows at the top/left of each page",
        help_page: "0282-print-background-encoded-file-printer-options-borders.html",
        body: MenuBody::NotImplemented("pfo-borders"),
    },
    MenuItem {
        letter: 'S',
        name: "Setup",
        help: "Setup escape sequence for the printer",
        help_page: "0299-print-background-encoded-printer-options-setup.html",
        body: MenuBody::Action(Action::PrintSessionOptionsSetup),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "As-Displayed / Cell-Formulas / Formatted / Unformatted",
        help_page: "0295-print-background-encoded-file-printer-options-other.html",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_OTHER_MENU),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Named print-settings sets",
        help_page: "0288-print-background-encoded-file-printer-options-name.html",
        body: MenuBody::NotImplemented("pfo-name"),
    },
    MenuItem {
        letter: 'A',
        name: "Advanced",
        help: "Advanced options (AutoLf, Color, Device, Fonts, …)",
        help_page: "0301-print-background-encoded-file-printer-options-advanced.html",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_ADVANCED_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to the print menu",
        help_page: "",
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
        help_page: "0276-print-background-encoded-file-printer-range.html",
        body: MenuBody::Action(Action::PrintSessionRange),
    },
    MenuItem {
        letter: 'L',
        name: "Line",
        help: "Advance one line in the output file",
        help_page: "0273-print-background-encoded-file-printer-line.html",
        body: MenuBody::NotImplemented("pf-line"),
    },
    MenuItem {
        letter: 'P',
        name: "Page",
        help: "Advance to the next page in the output file",
        help_page: "0274-print-background-encoded-file-printer-page.html",
        body: MenuBody::NotImplemented("pf-page"),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Header, footer, margins, page length, and output format",
        help_page: "0281-print-background-encoded-file-printer-options.html",
        body: MenuBody::Submenu(PRINT_FILE_OPTIONS_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Clear print settings (header, footer, margins, …)",
        help_page: "0270-print-background-encoded-file-printer-clear.html",
        body: MenuBody::Action(Action::PrintSessionClear),
    },
    MenuItem {
        letter: 'A',
        name: "Align",
        help: "Reset the page counter to 1",
        help_page: "0269-print-background-encoded-file-printer-align.html",
        body: MenuBody::Action(Action::PrintSessionAlign),
    },
    MenuItem {
        letter: 'G',
        name: "Go",
        help: "Write the selected range to the print file",
        help_page: "0271-print-background-encoded-file-printer-go.html",
        body: MenuBody::Action(Action::PrintSessionGo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "0275-print-background-encoded-file-printer-quit.html",
        body: MenuBody::Action(Action::PrintSessionQuit),
    },
];

const PRINT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'P',
        name: "Printer",
        help: "Print to printer",
        help_page: "0261-print-printer.html",
        body: MenuBody::Action(Action::PrintPrinter),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Print to .PRN text file",
        help_page: "0262-print-file.html",
        body: MenuBody::Action(Action::PrintFile),
    },
    MenuItem {
        letter: 'E',
        name: "Encoded",
        help: "Print to encoded file with printer codes",
        help_page: "0264-print-encoded.html",
        body: MenuBody::Action(Action::PrintEncoded),
    },
    MenuItem {
        letter: 'C',
        name: "Cancel",
        help: "Cancel current print job",
        help_page: "0267-print-cancel.html",
        body: MenuBody::Action(Action::PrintCancel),
    },
];

// /Graph Type Features Yes|No submenus. Letter accelerator order
// matches Lotus muscle memory: No before Yes.

const GTF_STACKED_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Do not stack data ranges",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesStackedNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Stack data ranges",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesStackedYes),
    },
];

const GTF_PERCENT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Plot absolute values",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesPercentNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Plot values as percentages of their column total",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesPercentYes),
    },
];

const GTF_DROP_SHADOW_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Remove drop-shadow",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesDropShadowNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Add drop-shadow",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesDropShadowYes),
    },
];

const GTF_THREE_D_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Plot in 2-D",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesThreeDNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Plot in 3-D",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesThreeDYes),
    },
];

const GTF_TABLE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide value table",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesTableNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show value table below the graph",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesTableYes),
    },
];

// 2Y-Ranges — assign series to the second y-axis.
const GTF_2Y_RANGES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Assign every data range to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YGraph),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Assign A to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Assign B to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Assign C to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Assign D to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Assign E to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Assign F to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeatures2YF),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Type Features",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesQuit),
    },
];

// Y-Ranges — reassign series back to the first y-axis.
const GTF_Y_RANGES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Reassign every data range to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYGraph),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Reassign A to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Reassign B to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Reassign C to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Reassign D to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Reassign E to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Reassign F to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesYF),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Type Features",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesQuit),
    },
];

// Frame Y-Axis Yes/No submenu.
const GTF_FRAME_Y_AXIS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the inner y-axis line",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameYAxisNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the inner y-axis line",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameYAxisYes),
    },
];

// Frame side Yes/No submenus.
const GTF_FRAME_LEFT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the left frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameLeftNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the left frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameLeftYes),
    },
];

const GTF_FRAME_RIGHT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the right frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameRightNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the right frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameRightYes),
    },
];

const GTF_FRAME_TOP_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the top frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameTopNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the top frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameTopYes),
    },
];

const GTF_FRAME_BOTTOM_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the bottom frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameBottomNo),
    },
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the bottom frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameBottomYes),
    },
];

const GTF_FRAME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Toggle the left frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_LEFT_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Toggle the right frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_RIGHT_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Toggle the top frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_TOP_MENU),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Toggle the bottom frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_BOTTOM_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "All",
        help: "Show every frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameAll),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Hide every frame edge",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Action(Action::GraphFeaturesFrameClear),
    },
    // The Y-Axis frame line is conceptually distinct from the four
    // outer edges — it sits just inside the Left edge between the
    // outer frame and the data plot.
    MenuItem {
        letter: 'Y',
        name: "Y-Axis",
        help: "Toggle the inner y-axis line",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_Y_AXIS_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Type Features",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesQuit),
    },
];

const GRAPH_TYPE_FEATURES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'V',
        name: "Vertical",
        help: "X-axis at bottom, y-axes on left and right (default)",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesVertical),
    },
    MenuItem {
        letter: 'H',
        name: "Horizontal",
        help: "X-axis at left, y-axes top and bottom",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Action(Action::GraphFeaturesHorizontal),
    },
    MenuItem {
        letter: 'S',
        name: "Stacked",
        help: "Stack data ranges (line/bar/mixed/XY)",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Submenu(GTF_STACKED_MENU),
    },
    MenuItem {
        letter: '1',
        name: "100%",
        help: "Plot data ranges as percentages of their total",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Submenu(GTF_PERCENT_MENU),
    },
    MenuItem {
        letter: '2',
        name: "2Y-Ranges",
        help: "Assign data ranges to the second y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_2Y_RANGES_MENU),
    },
    MenuItem {
        letter: 'Y',
        name: "Y-Ranges",
        help: "Reassign data ranges to the first y-axis",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_Y_RANGES_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Frame",
        help: "Toggle the graph frame edges",
        help_page: "0065-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_FRAME_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Drop-Shadow",
        help: "Add or remove drop-shadow",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_DROP_SHADOW_MENU),
    },
    MenuItem {
        letter: '3',
        name: "3-D",
        help: "Display the graph in 3-D",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_THREE_D_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Show a value table below the graph",
        help_page: "0066-graph-type-features-continued.html",
        body: MenuBody::Submenu(GTF_TABLE_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Type",
        help_page: "0060-graph-type.html",
        body: MenuBody::Action(Action::GraphFeaturesQuit),
    },
];

/// /Graph Group orientation submenu. Public because the UI roots
/// into it after the POINT step commits the group range. Reference
/// p. 2-172 (0070-graph-group.html).
pub const GRAPH_GROUP_ORIENT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Columnwise",
        help: "First column → X; succeeding columns → A, B, … F",
        help_page: "0070-graph-group.html",
        body: MenuBody::Action(Action::GraphGroupColumnwise),
    },
    MenuItem {
        letter: 'R',
        name: "Rowwise",
        help: "First row → X; succeeding rows → A, B, … F",
        help_page: "0070-graph-group.html",
        body: MenuBody::Action(Action::GraphGroupRowwise),
    },
];

// /Graph Name — Reference p. 2-218 → 2-224 (0071-graph-name.html).
// Use, Create, Delete are string prompts; Reset is immediate (no
// confirmation per the Reference's explicit CAUTION); Table is
// deferred to slice F (needs cell-write orchestration).
const GRAPH_NAME_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'U',
        name: "Use",
        help: "Load a named graph as the current graph",
        help_page: "0076-graph-name-use.html",
        body: MenuBody::Action(Action::GraphNameUse),
    },
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Save the current graph under a name",
        help_page: "0072-graph-name-create.html",
        body: MenuBody::Action(Action::GraphNameCreate),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Drop one named graph",
        help_page: "0073-graph-name-delete.html",
        body: MenuBody::Action(Action::GraphNameDelete),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Delete every named graph (no confirmation)",
        help_page: "0074-graph-name-reset.html",
        body: MenuBody::Action(Action::GraphNameReset),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "Write a table of named graphs to a worksheet range",
        help_page: "0075-graph-name-table.html",
        body: MenuBody::Action(Action::GraphNameTable),
    },
];

const GRAPH_TYPE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Line",
        help: "Line graph",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeLine),
    },
    MenuItem {
        letter: 'B',
        name: "Bar",
        help: "Bar graph",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeBar),
    },
    MenuItem {
        letter: 'X',
        name: "XY",
        help: "XY (scatter) graph",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeXY),
    },
    MenuItem {
        letter: 'S',
        name: "Stack-Bar",
        help: "Stacked-bar graph",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeStack),
    },
    MenuItem {
        letter: 'P',
        name: "Pie",
        help: "Pie chart",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypePie),
    },
    MenuItem {
        letter: 'H',
        name: "HLCO",
        help: "High/Low/Close/Open chart",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeHLCO),
    },
    MenuItem {
        letter: 'M',
        name: "Mixed",
        help: "Mixed (bar + line) chart",
        help_page: "",
        body: MenuBody::Action(Action::GraphTypeMixed),
    },
    MenuItem {
        letter: 'F',
        name: "Features",
        help: "Type features (orientation, stacked, 100%, frame, 3-D, …)",
        help_page: "0064-graph-type-features.html",
        body: MenuBody::Submenu(GRAPH_TYPE_FEATURES_MENU),
    },
];

// /Graph Options Advanced — Reference p. 2-200,
// 0078-graph-options-advanced.html. Affects raster output only and
// needs data-model fields not present today; this menu is wired as a
// navigable shell so the tree is traversable, with each leaf
// surfacing a NotImplemented tag until slice F lands the model.
const GO_ADVANCED_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Colors",
        help: "Per-series colors / Hide / Range",
        help_page: "0079-graph-options-advanced-colors-sets-colors-for-or-hides-the-a-f-data.html",
        body: MenuBody::NotImplemented("goa-colors"),
    },
    MenuItem {
        letter: 'H',
        name: "Hatches",
        help: "Per-series hatch patterns",
        help_page: "0081-graph-options-advanced-hatches.html",
        body: MenuBody::NotImplemented("goa-hatches"),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Color, font, and size for graph text",
        help_page: "0083-graph-options-advanced-text.html",
        body: MenuBody::NotImplemented("goa-text"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Options",
        help_page: "0077-graph-options.html",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

// /Graph Options Scale {Y|X|2Y}-Scale — per-axis settings. Reference
// p. 2-206 (0093-graph-options-scale-continued.html). Auto and Manual
// drive `ScaleAxis::mode`; the rest are stubs until the model carries
// `lower`, `upper`, `format`, `indicator`, `type_`, `exponent`, and
// `width` fields.
macro_rules! gos_axis_menu {
    ($auto:expr, $manual:expr, $lower:expr, $upper:expr, $format:literal,
     $indicator_menu:expr, $type_menu:expr, $exponent:expr, $width:expr) => {
        &[
            MenuItem {
                letter: 'A',
                name: "Automatic",
                help: "Compute scale limits from the data (default)",
                help_page: "0094-graph-options-scale-y-x-2y-automatic-manual-lower-or-upper.html",
                body: MenuBody::Action($auto),
            },
            MenuItem {
                letter: 'M',
                name: "Manual",
                help: "Use the manually-set lower and upper limits",
                help_page: "0094-graph-options-scale-y-x-2y-automatic-manual-lower-or-upper.html",
                body: MenuBody::Action($manual),
            },
            MenuItem {
                letter: 'L',
                name: "Lower",
                help: "Manual lower limit",
                help_page: "0094-graph-options-scale-y-x-2y-automatic-manual-lower-or-upper.html",
                body: MenuBody::Action($lower),
            },
            MenuItem {
                letter: 'U',
                name: "Upper",
                help: "Manual upper limit",
                help_page: "0094-graph-options-scale-y-x-2y-automatic-manual-lower-or-upper.html",
                body: MenuBody::Action($upper),
            },
            MenuItem {
                letter: 'F',
                name: "Format",
                help: "Number format for scale labels",
                help_page: "0096-graph-options-scale-y-scale-x-scale-2y-scale-format.html",
                body: MenuBody::NotImplemented($format),
            },
            MenuItem {
                letter: 'I',
                name: "Indicator",
                help: "Scale indicator (Yes/No/Manual)",
                help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
                body: MenuBody::Submenu($indicator_menu),
            },
            MenuItem {
                letter: 'T',
                name: "Type",
                help: "Linear or logarithmic scale",
                help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
                body: MenuBody::Submenu($type_menu),
            },
            MenuItem {
                letter: 'E',
                name: "Exponent",
                help: "Order-of-magnitude shift",
                help_page: "0095-graph-options-scale-y-scale-x-scale-2y-scale-exponent.html",
                body: MenuBody::Action($exponent),
            },
            MenuItem {
                letter: 'W',
                name: "Width",
                help: "Maximum width of scale labels",
                help_page: "0099-graph-options-scale-y-scale-x-scale-2y-scale-width.html",
                body: MenuBody::Action($width),
            },
            MenuItem {
                letter: 'Q',
                name: "Quit",
                help: "Return to /Graph Options Scale",
                help_page: "0092-graph-options-scale.html",
                body: MenuBody::Action(Action::GraphOptionsQuit),
            },
        ]
    };
}

// Yes/No/Manual submenus rooted under `/Graph Options Scale
// {axis} Indicator`.
const GO_SCALE_Y_INDICATOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleYIndicatorYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleYIndicatorNo),
    },
    MenuItem {
        letter: 'M',
        name: "Manual",
        help: "Use a manual y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleYIndicatorManual),
    },
];
const GO_SCALE_X_INDICATOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the x-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleXIndicatorYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the x-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleXIndicatorNo),
    },
    MenuItem {
        letter: 'M',
        name: "Manual",
        help: "Use a manual x-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScaleXIndicatorManual),
    },
];
const GO_SCALE_2Y_INDICATOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the 2y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScale2YIndicatorYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the 2y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScale2YIndicatorNo),
    },
    MenuItem {
        letter: 'M',
        name: "Manual",
        help: "Use a manual 2y-axis magnitude indicator",
        help_page: "0097-graph-options-scale-y-scale-x-scale-2y-scale-indicator.html",
        body: MenuBody::Action(Action::GraphOptionsScale2YIndicatorManual),
    },
];

// Linear/Logarithmic submenus rooted under `/Graph Options Scale
// {axis} Type`. Letter mapping: N (liNear), L (Logarithmic) — N
// chosen to avoid collision with L=Logarithmic.
const GO_SCALE_Y_TYPE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Linear",
        help: "Linear y-axis (default)",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScaleYTypeLinear),
    },
    MenuItem {
        letter: 'L',
        name: "Logarithmic",
        help: "Logarithmic y-axis",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScaleYTypeLog),
    },
];
const GO_SCALE_X_TYPE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Linear",
        help: "Linear x-axis (default)",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScaleXTypeLinear),
    },
    MenuItem {
        letter: 'L',
        name: "Logarithmic",
        help: "Logarithmic x-axis",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScaleXTypeLog),
    },
];
const GO_SCALE_2Y_TYPE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'N',
        name: "Linear",
        help: "Linear 2y-axis (default)",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScale2YTypeLinear),
    },
    MenuItem {
        letter: 'L',
        name: "Logarithmic",
        help: "Logarithmic 2y-axis",
        help_page: "0098-graph-options-scale-y-scale-x-scale-2y-scale-type.html",
        body: MenuBody::Action(Action::GraphOptionsScale2YTypeLog),
    },
];

const GO_SCALE_Y_MENU: &[MenuItem] = gos_axis_menu!(
    Action::GraphOptionsScaleYAuto,
    Action::GraphOptionsScaleYManual,
    Action::GraphOptionsScaleYLower,
    Action::GraphOptionsScaleYUpper,
    "gosy-format",
    GO_SCALE_Y_INDICATOR_MENU,
    GO_SCALE_Y_TYPE_MENU,
    Action::GraphOptionsScaleYExponent,
    Action::GraphOptionsScaleYWidth
);
const GO_SCALE_X_MENU: &[MenuItem] = gos_axis_menu!(
    Action::GraphOptionsScaleXAuto,
    Action::GraphOptionsScaleXManual,
    Action::GraphOptionsScaleXLower,
    Action::GraphOptionsScaleXUpper,
    "gosx-format",
    GO_SCALE_X_INDICATOR_MENU,
    GO_SCALE_X_TYPE_MENU,
    Action::GraphOptionsScaleXExponent,
    Action::GraphOptionsScaleXWidth
);
const GO_SCALE_2Y_MENU: &[MenuItem] = gos_axis_menu!(
    Action::GraphOptionsScale2YAuto,
    Action::GraphOptionsScale2YManual,
    Action::GraphOptionsScale2YLower,
    Action::GraphOptionsScale2YUpper,
    "gos2-format",
    GO_SCALE_2Y_INDICATOR_MENU,
    GO_SCALE_2Y_TYPE_MENU,
    Action::GraphOptionsScale2YExponent,
    Action::GraphOptionsScale2YWidth
);

const GO_SCALE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Y-Scale",
        help: "First y-axis scaling",
        help_page: "0092-graph-options-scale.html",
        body: MenuBody::Submenu(GO_SCALE_Y_MENU),
    },
    MenuItem {
        letter: 'X',
        name: "X-Scale",
        help: "X-axis scaling (XY graphs)",
        help_page: "0092-graph-options-scale.html",
        body: MenuBody::Submenu(GO_SCALE_X_MENU),
    },
    MenuItem {
        letter: '2',
        name: "2Y-Scale",
        help: "Second y-axis scaling",
        help_page: "0092-graph-options-scale.html",
        body: MenuBody::Submenu(GO_SCALE_2Y_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Skip",
        help: "Every Nth x-axis label",
        help_page: "0092-graph-options-scale.html",
        body: MenuBody::Action(Action::GraphOptionsScaleSkip),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Options",
        help_page: "0077-graph-options.html",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

// /Graph Options Data-Labels — Reference p. 2-204,
// 0088-graph-options-data-labels.html. A-F leaves enter POINT mode for
// the data-label range. Group (Columnwise/Rowwise) and the placement
// follow-up (Center/Left/Above/Right/Below) are deferred.
const GO_DATA_LABELS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Data labels for the A range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Data labels for the B range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Data labels for the C range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Data labels for the D range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Data labels for the E range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Data labels for the F range",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsF),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Distribute one range across A-F",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::NotImplemented("god-group"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Options",
        help_page: "0077-graph-options.html",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

/// Placement follow-up rooted after `/GOD {A-F}` commits a range.
/// Letters mirror the original 1-2-3 R3.4a "Centered Left Above
/// Right Below" prompt; each leaf writes the placement and pops
/// back to READY.
pub const GO_DATA_LABELS_PLACEMENT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Place data labels at the data point",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsCenter),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Place data labels to the left of the point",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsLeft),
    },
    MenuItem {
        letter: 'A',
        name: "Above",
        help: "Place data labels above the point",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsAbove),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Place data labels to the right of the point",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsRight),
    },
    MenuItem {
        letter: 'B',
        name: "Below",
        help: "Place data labels below the point",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Action(Action::GraphOptionsDataLabelsBelow),
    },
];

// /Graph Options Legend — Reference p. 2-200, 0091-graph-options-legend.html.
// A-F leaves open per-series text prompts; the Range leaf points at a
// worksheet range whose cell values are read into the six legend
// slots in order. Range is deferred (slice F).
const GO_LEGEND_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Legend for the A data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Legend for the B data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Legend for the C data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Legend for the D data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Legend for the E data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Legend for the F data range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendF),
    },
    MenuItem {
        letter: 'R',
        name: "Range",
        help: "Read legends from a worksheet range",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Action(Action::GraphOptionsLegendRange),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Options",
        help_page: "0077-graph-options.html",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

// /Graph Options Titles — Reference p. 2-216, 0100-graph-options-titles.html.
// Each leaf opens a single-line text prompt; the committed buffer
// replaces the matching slot in `current_graph.options.titles` (or
// clears it if the buffer is empty).
const GO_TITLES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "First",
        help: "First (top) line of graph title",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleFirst),
    },
    MenuItem {
        letter: 'S',
        name: "Second",
        help: "Second (subtitle) line",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleSecond),
    },
    MenuItem {
        letter: 'X',
        name: "X-Axis",
        help: "Caption beneath the x-axis",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleXAxis),
    },
    MenuItem {
        letter: 'Y',
        name: "Y-Axis",
        help: "Caption alongside the first y-axis",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleYAxis),
    },
    MenuItem {
        letter: '2',
        name: "2Y-Axis",
        help: "Caption alongside the second y-axis",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitle2YAxis),
    },
    MenuItem {
        letter: 'N',
        name: "Note",
        help: "Bottom-left footnote",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleNote),
    },
    MenuItem {
        letter: 'O',
        name: "Other-Note",
        help: "Bottom-right footnote",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Action(Action::GraphOptionsTitleOtherNote),
    },
];

// /Graph Options Format value submenus — one per series (and one for
// "Graph" = all six). Each carries the same five values (Lines /
// Symbols / Both / Neither / Area) and a Quit that pops back to the
// /GOF entry point. Reference p. 2-200, 0089-graph-options-format.html.

macro_rules! gof_value_menu {
    ($lines:expr, $symbols:expr, $both:expr, $neither:expr, $area:expr) => {
        &[
            MenuItem {
                letter: 'L',
                name: "Lines",
                help: "Connecting lines only",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action($lines),
            },
            MenuItem {
                letter: 'S',
                name: "Symbols",
                help: "Symbols at each data point",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action($symbols),
            },
            MenuItem {
                letter: 'B',
                name: "Both",
                help: "Lines and symbols (default)",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action($both),
            },
            MenuItem {
                letter: 'N',
                name: "Neither",
                help: "Hide both lines and symbols",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action($neither),
            },
            MenuItem {
                letter: 'A',
                name: "Area",
                help: "Fill area below the line",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action($area),
            },
            MenuItem {
                letter: 'Q',
                name: "Quit",
                help: "Return to /Graph Options Format",
                help_page: "0089-graph-options-format.html",
                body: MenuBody::Action(Action::GraphOptionsQuit),
            },
        ]
    };
}

const GO_FORMAT_GRAPH_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatGraphLines,
    Action::GraphFormatGraphSymbols,
    Action::GraphFormatGraphBoth,
    Action::GraphFormatGraphNeither,
    Action::GraphFormatGraphArea
);
const GO_FORMAT_A_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatALines,
    Action::GraphFormatASymbols,
    Action::GraphFormatABoth,
    Action::GraphFormatANeither,
    Action::GraphFormatAArea
);
const GO_FORMAT_B_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatBLines,
    Action::GraphFormatBSymbols,
    Action::GraphFormatBBoth,
    Action::GraphFormatBNeither,
    Action::GraphFormatBArea
);
const GO_FORMAT_C_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatCLines,
    Action::GraphFormatCSymbols,
    Action::GraphFormatCBoth,
    Action::GraphFormatCNeither,
    Action::GraphFormatCArea
);
const GO_FORMAT_D_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatDLines,
    Action::GraphFormatDSymbols,
    Action::GraphFormatDBoth,
    Action::GraphFormatDNeither,
    Action::GraphFormatDArea
);
const GO_FORMAT_E_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatELines,
    Action::GraphFormatESymbols,
    Action::GraphFormatEBoth,
    Action::GraphFormatENeither,
    Action::GraphFormatEArea
);
const GO_FORMAT_F_VALUES_MENU: &[MenuItem] = gof_value_menu!(
    Action::GraphFormatFLines,
    Action::GraphFormatFSymbols,
    Action::GraphFormatFBoth,
    Action::GraphFormatFNeither,
    Action::GraphFormatFArea
);

const GO_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Format every data range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_GRAPH_VALUES_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Format the A range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_A_VALUES_MENU),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Format the B range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_B_VALUES_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Format the C range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_C_VALUES_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Format the D range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_D_VALUES_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Format the E range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_E_VALUES_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Format the F range",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_F_VALUES_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph Options",
        help_page: "0077-graph-options.html",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

// /Graph Options Grid — direct add-only actions per Reference p. 2-200
// (0090-graph-options-grid.html). Single sides cannot be removed
// individually; use Clear and re-add.
const GO_GRID_Y_AXIS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Y",
        help: "Anchor horizontal grid lines to the first y-axis",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridYAxisFirst),
    },
    MenuItem {
        letter: '2',
        name: "2Y",
        help: "Anchor horizontal grid lines to the second y-axis",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridYAxisSecond),
    },
    MenuItem {
        letter: 'B',
        name: "Both",
        help: "Anchor horizontal grid lines to both y-axes",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridYAxisBoth),
    },
];

const GO_GRID_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'H',
        name: "Horizontal",
        help: "Add horizontal grid lines",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridHorizontal),
    },
    MenuItem {
        letter: 'V',
        name: "Vertical",
        help: "Add vertical grid lines",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridVertical),
    },
    MenuItem {
        letter: 'B',
        name: "Both",
        help: "Add horizontal and vertical grid lines",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridBoth),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove every grid line",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Action(Action::GraphOptionsGridClear),
    },
    MenuItem {
        letter: 'Y',
        name: "Y-Axis",
        help: "Choose which y-axis horizontal grids originate from",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Submenu(GO_GRID_Y_AXIS_MENU),
    },
];

// /Graph Options — Reference p. 2-200 (0077-graph-options.html). Slice
// D of GRAPH_PLAN.md fills these in over several sub-slices; D1 wires
// Color and B&W as the smallest meaningful pair. The remaining leaves
// (Legend, Format, Titles, Grid, Scale, Data-Labels, Advanced) sit as
// NotImplemented stubs so the menu shape is complete and muscle-memory
// works today.
const GRAPH_OPTIONS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Legend",
        help: "Set per-series legend text",
        help_page: "0091-graph-options-legend.html",
        body: MenuBody::Submenu(GO_LEGEND_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Lines / Symbols / Both / Neither / Area",
        help_page: "0089-graph-options-format.html",
        body: MenuBody::Submenu(GO_FORMAT_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Titles",
        help: "Graph, axis, and footnote titles",
        help_page: "0100-graph-options-titles.html",
        body: MenuBody::Submenu(GO_TITLES_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Grid",
        help: "Horizontal / Vertical / Y-Axis grid lines",
        help_page: "0090-graph-options-grid.html",
        body: MenuBody::Submenu(GO_GRID_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Scale",
        help: "Y / X / 2Y axis scaling",
        help_page: "0092-graph-options-scale.html",
        body: MenuBody::Submenu(GO_SCALE_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Render graph in color",
        help_page: "0087-graph-options-color-and-graph-options-b-w.html",
        body: MenuBody::Action(Action::GraphOptionsColor),
    },
    MenuItem {
        letter: 'B',
        name: "B&W",
        help: "Render graph in black and white",
        help_page: "0087-graph-options-color-and-graph-options-b-w.html",
        body: MenuBody::Action(Action::GraphOptionsBW),
    },
    MenuItem {
        letter: 'D',
        name: "Data-Labels",
        help: "Per-series data labels",
        help_page: "0088-graph-options-data-labels.html",
        body: MenuBody::Submenu(GO_DATA_LABELS_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Advanced",
        help: "Colors, hatches, fonts, and text sizes",
        help_page: "0078-graph-options-advanced.html",
        body: MenuBody::Submenu(GO_ADVANCED_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to /Graph",
        help_page: "",
        body: MenuBody::Action(Action::GraphOptionsQuit),
    },
];

const GRAPH_RESET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Clear every range and restore default type",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetGraph),
    },
    MenuItem {
        letter: 'X',
        name: "X",
        help: "Clear X-axis range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetX),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Clear A range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Clear B range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Clear C range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Clear D range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Clear E range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Clear F range",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetF),
    },
    MenuItem {
        letter: 'R',
        name: "Ranges",
        help: "Clear X and A..F together (keep options)",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetRanges),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Reset graph options (keep ranges)",
        help_page: "",
        body: MenuBody::Action(Action::GraphResetOptions),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const GRAPH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'T',
        name: "Type",
        help: "Select graph type",
        help_page: "0060-graph-type.html",
        body: MenuBody::Submenu(GRAPH_TYPE_MENU),
    },
    MenuItem {
        letter: 'X',
        name: "X",
        help: "Set X-axis range",
        help_page: "0067-graph-x.html",
        body: MenuBody::Action(Action::GraphX),
    },
    MenuItem {
        letter: 'A',
        name: "A",
        help: "Set A data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphA),
    },
    MenuItem {
        letter: 'B',
        name: "B",
        help: "Set B data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphB),
    },
    MenuItem {
        letter: 'C',
        name: "C",
        help: "Set C data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphC),
    },
    MenuItem {
        letter: 'D',
        name: "D",
        help: "Set D data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphD),
    },
    MenuItem {
        letter: 'E',
        name: "E",
        help: "Set E data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphE),
    },
    MenuItem {
        letter: 'F',
        name: "F",
        help: "Set F data range",
        help_page: "0068-graph-a-b-c-d-e-f.html",
        body: MenuBody::Action(Action::GraphF),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Reset graph / ranges / options",
        help_page: "0102-graph-reset.html",
        body: MenuBody::Submenu(GRAPH_RESET_MENU),
    },
    MenuItem {
        letter: 'V',
        name: "View",
        help: "Display the current graph",
        help_page: "0104-graph-view.html",
        body: MenuBody::Action(Action::GraphView),
    },
    MenuItem {
        letter: 'S',
        name: "Save",
        help: "Save graph to an SVG file",
        help_page: "0103-graph-save.html",
        body: MenuBody::Action(Action::GraphSave),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Legend, Titles, Grid, Scale, Color, …",
        help_page: "0077-graph-options.html",
        body: MenuBody::Submenu(GRAPH_OPTIONS_MENU),
    },
    MenuItem {
        letter: 'N',
        name: "Name",
        help: "Create, use, delete, reset named graphs",
        help_page: "0071-graph-name.html",
        body: MenuBody::Submenu(GRAPH_NAME_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Group",
        help: "Columnwise / Rowwise auto-assign",
        help_page: "0070-graph-group.html",
        body: MenuBody::Action(Action::GraphGroup),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::GraphQuit),
    },
];

const DATA_TABLE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: '1',
        name: "1",
        help: "One-variable what-if table",
        help_page: "0172-data-table-1.html",
        body: MenuBody::Action(Action::DataTable1),
    },
    MenuItem {
        letter: '2',
        name: "2",
        help: "Two-variable what-if table",
        help_page: "0174-data-table-2.html",
        body: MenuBody::Action(Action::DataTable2),
    },
    MenuItem {
        letter: '3',
        name: "3",
        help: "Three-variable what-if table (R3+) — not yet implemented",
        help_page: "0177-data-table-3.html",
        body: MenuBody::Action(Action::DataTable3Stub),
    },
    MenuItem {
        letter: 'L',
        name: "Labeled",
        help: "Labeled what-if table (R3+) — not yet implemented",
        help_page: "0180-data-table-labeled.html",
        body: MenuBody::Action(Action::DataTableLabeledStub),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear table ranges and input cells",
        help_page: "",
        body: MenuBody::Action(Action::DataTableReset),
    },
];

/// Direction submenu shared by the Primary-Key and Secondary-Key
/// flows. Which slot is being set is tracked on the App, not in the
/// menu — both leaves resolve to the same Action regardless of which
/// key invoked them.
pub const DATA_SORT_DIR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'A',
        name: "Ascending",
        help: "Sort low-to-high (numbers ascending, labels A→Z)",
        help_page: "",
        body: MenuBody::Action(Action::DataSortAscending),
    },
    MenuItem {
        letter: 'D',
        name: "Descending",
        help: "Sort high-to-low (numbers descending, labels Z→A)",
        help_page: "",
        body: MenuBody::Action(Action::DataSortDescending),
    },
];

pub const DATA_PARSE_FORMAT_LINE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Auto-create a format line above the input column",
        help_page: "0115-data-parse-format-line-create.html",
        body: MenuBody::Action(Action::DataParseFormatLineCreate),
    },
    MenuItem {
        letter: 'E',
        name: "Edit",
        help: "Edit the format line in the EDIT-mode buffer",
        help_page: "0116-data-parse-format-line-edit.html",
        body: MenuBody::Action(Action::DataParseFormatLineEdit),
    },
];

pub const DATA_PARSE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Format-Line",
        help: "Auto-create or edit the format-line label",
        help_page: "0113-data-parse-format-line.html",
        body: MenuBody::Submenu(DATA_PARSE_FORMAT_LINE_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Input-Column",
        help: "Column of long labels to parse (includes the format-line row)",
        help_page: "0118-data-parse-input-column.html",
        body: MenuBody::Action(Action::DataParseInputColumn),
    },
    MenuItem {
        letter: 'O',
        name: "Output-Range",
        help: "Top-left cell for the parsed-out fields",
        help_page: "0119-data-parse-output-range.html",
        body: MenuBody::Action(Action::DataParseOutputRange),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear input and output settings",
        help_page: "",
        body: MenuBody::Action(Action::DataParseReset),
    },
    MenuItem {
        letter: 'G',
        name: "Go",
        help: "Parse each label according to the format line",
        help_page: "",
        body: MenuBody::Action(Action::DataParseGo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::DataParseQuit),
    },
];

pub const DATA_REGRESSION_INTERCEPT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Compute",
        help: "Fit the intercept term (default)",
        help_page: "",
        body: MenuBody::Action(Action::DataRegressionInterceptCompute),
    },
    MenuItem {
        letter: 'Z',
        name: "Zero",
        help: "Force the intercept to zero (regression through origin)",
        help_page: "",
        body: MenuBody::Action(Action::DataRegressionInterceptZero),
    },
];

pub const DATA_REGRESSION_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'X',
        name: "X-Range",
        help: "Independent variable column",
        help_page: "0154-data-regression-x-range.html",
        body: MenuBody::Action(Action::DataRegressionXRange),
    },
    MenuItem {
        letter: 'Y',
        name: "Y-Range",
        help: "Dependent variable column",
        help_page: "0153-data-regression-y-range.html",
        body: MenuBody::Action(Action::DataRegressionYRange),
    },
    MenuItem {
        letter: 'O',
        name: "Output-Range",
        help: "Top-left cell for the labeled output table",
        help_page: "0155-data-regression-output-range.html",
        body: MenuBody::Action(Action::DataRegressionOutputRange),
    },
    MenuItem {
        letter: 'I',
        name: "Intercept",
        help: "Compute or zero the intercept term",
        help_page: "0156-data-regression-intercept.html",
        body: MenuBody::Submenu(DATA_REGRESSION_INTERCEPT_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear all ranges and revert Intercept to Compute",
        help_page: "",
        body: MenuBody::Action(Action::DataRegressionReset),
    },
    MenuItem {
        letter: 'G',
        name: "Go",
        help: "Run the regression and write the output table",
        help_page: "",
        body: MenuBody::Action(Action::DataRegressionGo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::DataRegressionQuit),
    },
];

pub const DATA_SORT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'D',
        name: "Data-Range",
        help: "Range of records to sort",
        help_page: "0159-data-sort-data-range.html",
        body: MenuBody::Action(Action::DataSortDataRange),
    },
    MenuItem {
        letter: 'P',
        name: "Primary-Key",
        help: "Primary sort key cell + Ascending/Descending",
        help_page: "0160-data-sort-primary-key.html",
        body: MenuBody::Action(Action::DataSortPrimaryKey),
    },
    MenuItem {
        letter: 'S',
        name: "Secondary-Key",
        help: "Secondary sort key cell + Ascending/Descending",
        help_page: "0165-data-sort-secondary-key.html",
        body: MenuBody::Action(Action::DataSortSecondaryKey),
    },
    MenuItem {
        letter: 'E',
        name: "Extra-Key",
        help: "Additional sort key, used as third tiebreaker (R3+)",
        help_page: "0166-data-sort-extra-key.html",
        body: MenuBody::Action(Action::DataSortExtraKey),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear all sort settings",
        help_page: "",
        body: MenuBody::Action(Action::DataSortReset),
    },
    MenuItem {
        letter: 'G',
        name: "Go",
        help: "Perform the sort",
        help_page: "",
        body: MenuBody::Action(Action::DataSortGo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::DataSortQuit),
    },
];

pub const DATA_QUERY_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "Input",
        help: "Range of records (with field-name header row)",
        help_page: "0123-data-query-input.html",
        body: MenuBody::Action(Action::DataQueryInput),
    },
    MenuItem {
        letter: 'C',
        name: "Criteria",
        help: "Range of search criteria",
        help_page: "0126-data-query-criteria.html",
        body: MenuBody::Action(Action::DataQueryCriteria),
    },
    MenuItem {
        letter: 'O',
        name: "Output",
        help: "Range to receive matching records",
        help_page: "0136-data-query-output.html",
        body: MenuBody::Action(Action::DataQueryOutput),
    },
    MenuItem {
        letter: 'F',
        name: "Find",
        help: "Highlight records that match the criteria",
        help_page: "0142-data-query-find.html",
        body: MenuBody::Action(Action::DataQueryFind),
    },
    MenuItem {
        letter: 'E',
        name: "Extract",
        help: "Copy matching records to the output range",
        help_page: "0141-data-query-extract.html",
        body: MenuBody::Action(Action::DataQueryExtract),
    },
    MenuItem {
        letter: 'U',
        name: "Unique",
        help: "Extract unique matching records",
        help_page: "0143-data-query-unique.html",
        body: MenuBody::Action(Action::DataQueryUnique),
    },
    MenuItem {
        letter: 'D',
        name: "Del",
        help: "Delete matching records from the input range",
        help_page: "0140-data-query-del.html",
        body: MenuBody::Action(Action::DataQueryDel),
    },
    MenuItem {
        letter: 'M',
        name: "Modify",
        help: "Edit matching records in place — not yet implemented",
        help_page: "0144-data-query-modify.html",
        body: MenuBody::Action(Action::DataQueryModifyStub),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear input, criteria, and output ranges",
        help_page: "0149-data-query-reset.html",
        body: MenuBody::Action(Action::DataQueryReset),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::DataQueryQuit),
    },
];

const DATA_MATRIX_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'I',
        name: "Invert",
        help: "Invert a square matrix (up to 90×90)",
        help_page: "0109-data-matrix-invert.html",
        body: MenuBody::Action(Action::DataMatrixInvert),
    },
    MenuItem {
        letter: 'M',
        name: "Multiply",
        help: "Multiply two matrices",
        help_page: "0110-data-matrix-multiply.html",
        body: MenuBody::Action(Action::DataMatrixMultiply),
    },
];

const DATA_EXTERNAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'U',
        name: "Use",
        help: "Connect to an external database driver",
        help_page: "0212-data-external-use.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'L',
        name: "List",
        help: "List available external tables / fields",
        help_page: "0191-data-external-list.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'C',
        name: "Create",
        help: "Create an external table",
        help_page: "0194-data-external-create.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'D',
        name: "Delete",
        help: "Delete an external table",
        help_page: "0206-data-external-delete.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'O',
        name: "Other",
        help: "Driver-specific options (Send / Translation)",
        help_page: "0207-data-external-other.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Disconnect all external tables",
        help_page: "0211-data-external-reset.html",
        body: MenuBody::Action(Action::DataExternalStub),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::DataExternalStub),
    },
];

const DATA_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'F',
        name: "Fill",
        help: "Fill a range with a sequence",
        help_page: "0106-data-fill.html",
        body: MenuBody::Action(Action::DataFill),
    },
    MenuItem {
        letter: 'T',
        name: "Table",
        help: "What-if tables (1, 2, 3, Labeled)",
        help_page: "0169-data-table.html",
        body: MenuBody::Submenu(DATA_TABLE_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Sort",
        help: "Sort a range by keys",
        help_page: "0157-data-sort.html",
        body: MenuBody::Submenu(DATA_SORT_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Query",
        help: "Database query (find, extract, ...)",
        help_page: "0120-data-query.html",
        body: MenuBody::Submenu(DATA_QUERY_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Distribution",
        help: "Frequency distribution into bins",
        help_page: "0105-data-distribution.html",
        body: MenuBody::Action(Action::DataDistribution),
    },
    MenuItem {
        letter: 'M',
        name: "Matrix",
        help: "Invert / Multiply matrices",
        help_page: "0108-data-matrix.html",
        body: MenuBody::Submenu(DATA_MATRIX_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Regression",
        help: "Linear regression on X and Y ranges",
        help_page: "0150-data-regression.html",
        body: MenuBody::Submenu(DATA_REGRESSION_MENU),
    },
    MenuItem {
        letter: 'P',
        name: "Parse",
        help: "Parse a column of long labels into fields",
        help_page: "0111-data-parse.html",
        body: MenuBody::Submenu(DATA_PARSE_MENU),
    },
    MenuItem {
        letter: 'E',
        name: "External",
        help: "External database tables (Use, List, ...)",
        help_page: "0190-data-external.html",
        body: MenuBody::Submenu(DATA_EXTERNAL_MENU),
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
        help_page: "",
        body: MenuBody::Action(Action::FormatBoldSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove bold from a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatBoldClear),
    },
];

const WYSIWYG_FORMAT_ITALIC_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Apply italic to a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatItalicSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove italic from a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatItalicClear),
    },
];

const WYSIWYG_FORMAT_UNDERLINE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Apply underline to a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatUnderlineSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Remove underline from a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatUnderlineClear),
    },
];

const WYSIWYG_FORMAT_COLOR_BG_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Black",
        help: "Paint the background black",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgBlack),
    },
    MenuItem {
        letter: 'W',
        name: "White",
        help: "Paint the background white",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgWhite),
    },
    MenuItem {
        letter: 'R',
        name: "Red",
        help: "Paint the background red",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgRed),
    },
    MenuItem {
        letter: 'G',
        name: "Green",
        help: "Paint the background green",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgGreen),
    },
    MenuItem {
        letter: 'L',
        name: "Blue",
        help: "Paint the background blue",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgBlue),
    },
    MenuItem {
        letter: 'Y',
        name: "Yellow",
        help: "Paint the background yellow",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgYellow),
    },
    MenuItem {
        letter: 'C',
        name: "Cyan",
        help: "Paint the background cyan",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgCyan),
    },
    MenuItem {
        letter: 'M',
        name: "Magenta",
        help: "Paint the background magenta",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorBgMagenta),
    },
];

const WYSIWYG_FORMAT_COLOR_TEXT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Black",
        help: "Tint the text black",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextBlack),
    },
    MenuItem {
        letter: 'W',
        name: "White",
        help: "Tint the text white",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextWhite),
    },
    MenuItem {
        letter: 'R',
        name: "Red",
        help: "Tint the text red",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextRed),
    },
    MenuItem {
        letter: 'G',
        name: "Green",
        help: "Tint the text green",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextGreen),
    },
    MenuItem {
        letter: 'L',
        name: "Blue",
        help: "Tint the text blue",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextBlue),
    },
    MenuItem {
        letter: 'Y',
        name: "Yellow",
        help: "Tint the text yellow",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextYellow),
    },
    MenuItem {
        letter: 'C',
        name: "Cyan",
        help: "Tint the text cyan",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextCyan),
    },
    MenuItem {
        letter: 'M',
        name: "Magenta",
        help: "Tint the text magenta",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorTextMagenta),
    },
];

const WYSIWYG_FORMAT_COLOR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Background",
        help: "Pick a background color for a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_COLOR_BG_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Pick a text color for a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_COLOR_TEXT_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear background and text colors on a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatColorReset),
    },
];

const WYSIWYG_FORMAT_ALIGNMENT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Left-align cell contents in a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatAlignmentLeft),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Right-align cell contents in a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatAlignmentRight),
    },
    MenuItem {
        letter: 'C',
        name: "Center",
        help: "Center cell contents in a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatAlignmentCenter),
    },
    MenuItem {
        letter: 'G',
        name: "General",
        help: "Clear the alignment override (labels left, numbers right)",
        help_page: "",
        body: MenuBody::Action(Action::FormatAlignmentGeneral),
    },
];

const WYSIWYG_FORMAT_LINES_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'O',
        name: "Outline",
        help: "Draw a thin border around the perimeter of the range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesOutlineSet),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Draw a thin border on the left edge of every cell",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesLeftSet),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Draw a thin border on the right edge of every cell",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesRightSet),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Draw a thin border on the top edge of every cell",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesTopSet),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Draw a thin border on the bottom edge of every cell",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesBottomSet),
    },
    MenuItem {
        letter: 'A',
        name: "All",
        help: "Draw thin borders on every edge of every cell",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesAllSet),
    },
    MenuItem {
        letter: 'C',
        name: "Clear",
        help: "Clear lines from a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_LINES_CLEAR_MENU),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_FORMAT_LINES_CLEAR_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'O',
        name: "Outline",
        help: "Clear the perimeter outline on a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesOutlineClear),
    },
    MenuItem {
        letter: 'L',
        name: "Left",
        help: "Clear left edges on every cell of a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesLeftClear),
    },
    MenuItem {
        letter: 'R',
        name: "Right",
        help: "Clear right edges on every cell of a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesRightClear),
    },
    MenuItem {
        letter: 'T',
        name: "Top",
        help: "Clear top edges on every cell of a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesTopClear),
    },
    MenuItem {
        letter: 'B',
        name: "Bottom",
        help: "Clear bottom edges on every cell of a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesBottomClear),
    },
    MenuItem {
        letter: 'A',
        name: "All",
        help: "Clear every edge on every cell of a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatLinesAllClear),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_FORMAT_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'B',
        name: "Bold",
        help: "Set or clear the bold attribute on a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_BOLD_MENU),
    },
    MenuItem {
        letter: 'I',
        name: "Italic",
        help: "Set or clear the italic attribute on a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_ITALIC_MENU),
    },
    MenuItem {
        letter: 'U',
        name: "Underline",
        help: "Set or clear the underline attribute on a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_UNDERLINE_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Font",
        help: "Change font on a range",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-format-font"),
    },
    MenuItem {
        letter: 'L',
        name: "Lines",
        help: "Draw lines around a range",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_LINES_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Change text or background color",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_COLOR_MENU),
    },
    MenuItem {
        letter: 'A',
        name: "Alignment",
        help: "Change alignment",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_ALIGNMENT_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Clear bold, italic and underline on a range",
        help_page: "",
        body: MenuBody::Action(Action::FormatReset),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_WORKSHEET_COLUMN_WIDTH_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'S',
        name: "Set",
        help: "Set the width of the current column",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnSetWidth),
    },
    MenuItem {
        letter: 'R',
        name: "Reset",
        help: "Reset column width to the global default",
        help_page: "",
        body: MenuBody::Action(Action::WorksheetColumnResetWidth),
    },
];

const WYSIWYG_WORKSHEET_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Column-Width",
        help: "Set or reset the current column's width",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_WORKSHEET_COLUMN_WIDTH_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Row",
        help: "Set or auto-fit row height",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-worksheet-row"),
    },
    MenuItem {
        letter: 'P',
        name: "Page",
        help: "Insert or remove manual page breaks",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-worksheet-page"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_DISPLAY_MODE_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Color",
        help: "Default to white worksheet background, black text",
        help_page: "",
        body: MenuBody::Action(Action::DisplayModeColor),
    },
    MenuItem {
        letter: 'B',
        name: "B&W",
        help: "Strip default cell color (terminal default)",
        help_page: "",
        body: MenuBody::Action(Action::DisplayModeBW),
    },
    MenuItem {
        letter: 'R',
        name: "Reverse",
        help: "Invert default cell colors",
        help_page: "",
        body: MenuBody::Action(Action::DisplayModeReverse),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_DISPLAY_OPTIONS_GRID_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'Y',
        name: "Yes",
        help: "Show the row/column gutter",
        help_page: "",
        body: MenuBody::Action(Action::DisplayOptionsGridYes),
    },
    MenuItem {
        letter: 'N',
        name: "No",
        help: "Hide the row/column gutter",
        help_page: "",
        body: MenuBody::Action(Action::DisplayOptionsGridNo),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_DISPLAY_OPTIONS_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'G',
        name: "Grid",
        help: "Show or hide the row/column gutter",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_DISPLAY_OPTIONS_GRID_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Frame",
        help: "Frame style",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-options-frame"),
    },
    MenuItem {
        letter: 'P',
        name: "Page-Breaks",
        help: "Show or hide page break markers",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-options-page-breaks"),
    },
    MenuItem {
        letter: 'C',
        name: "Cell-Pointer",
        help: "Cell pointer style",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-options-cell-pointer"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_DISPLAY_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'M',
        name: "Mode",
        help: "Color, B&W, or Reverse default cell colors",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_DISPLAY_MODE_MENU),
    },
    MenuItem {
        letter: 'O',
        name: "Options",
        help: "Grid, frame, page break, cell pointer options",
        help_page: "",
        body: MenuBody::Submenu(WYSIWYG_DISPLAY_OPTIONS_MENU),
    },
    MenuItem {
        letter: 'Z',
        name: "Zoom",
        help: "Zoom level",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-zoom"),
    },
    MenuItem {
        letter: 'C',
        name: "Colors",
        help: "Display palette tweaks",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-colors"),
    },
    MenuItem {
        letter: 'R',
        name: "Rows",
        help: "Visible row count",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-rows"),
    },
    MenuItem {
        letter: 'F',
        name: "Font-Directory",
        help: "Set or reset font directory",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-font-directory"),
    },
    MenuItem {
        letter: 'D',
        name: "Default",
        help: "Update or restore default settings",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-display-default"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

const WYSIWYG_SPECIAL_MENU: &[MenuItem] = &[
    MenuItem {
        letter: 'C',
        name: "Copy",
        help: "Copy formatting from a source range to a destination",
        help_page: "",
        body: MenuBody::Action(Action::SpecialCopy),
    },
    MenuItem {
        letter: 'M',
        name: "Move",
        help: "Move formatting from a source range to a destination",
        help_page: "",
        body: MenuBody::Action(Action::SpecialMove),
    },
    MenuItem {
        letter: 'I',
        name: "Import",
        help: "Import formatting from a saved .FMT file",
        help_page: "",
        body: MenuBody::NotImplemented("wysiwyg-special-import-fmt"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

/// Top-level WYSIWYG colon-menu.  Entered by pressing `:` in READY.
pub const WYSIWYG_ROOT: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "WYSIWYG worksheet display settings",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::Submenu(WYSIWYG_WORKSHEET_MENU),
    },
    MenuItem {
        letter: 'F',
        name: "Format",
        help: "Bold, italic, underline, font, color, alignment...",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::Submenu(WYSIWYG_FORMAT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Embed a graph at the current cell",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::NotImplemented("wysiwyg-graph"),
    },
    MenuItem {
        letter: 'N',
        name: "Named-Style",
        help: "Named cell formatting style",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::NotImplemented("wysiwyg-named-style"),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "WYSIWYG print controls",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::NotImplemented("wysiwyg-print"),
    },
    MenuItem {
        letter: 'D',
        name: "Display",
        help: "Screen display settings",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::Submenu(WYSIWYG_DISPLAY_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "Special",
        help: "Copy / move / import formatting",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::Submenu(WYSIWYG_SPECIAL_MENU),
    },
    MenuItem {
        letter: 'T',
        name: "Text",
        help: "Text editing and justification",
        help_page: "0816-using-wysiwyg.html",
        body: MenuBody::NotImplemented("wysiwyg-text"),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "Return to READY",
        help_page: "",
        body: MenuBody::Action(Action::Cancel),
    },
];

/// Top-level slash menu.
pub const ROOT: &[MenuItem] = &[
    MenuItem {
        letter: 'W',
        name: "Worksheet",
        help: "Global settings, insert/delete, columns, titles...",
        help_page: "0016-the-worksheet-commands.html",
        body: MenuBody::Submenu(WORKSHEET_MENU),
    },
    MenuItem {
        letter: 'R',
        name: "Range",
        help: "Format, label, erase, name, justify, protect...",
        help_page: "0014-the-range-commands.html",
        body: MenuBody::Submenu(RANGE_MENU),
    },
    MenuItem {
        letter: 'C',
        name: "Copy",
        help: "Copy a range to another location",
        help_page: "0006-copy.html",
        body: MenuBody::Action(Action::Copy),
    },
    MenuItem {
        letter: 'M',
        name: "Move",
        help: "Move a range to another location",
        help_page: "0011-move.html",
        body: MenuBody::Action(Action::Move),
    },
    MenuItem {
        letter: 'F',
        name: "File",
        help: "Retrieve, save, combine, import, export, ...",
        help_page: "0009-the-file-commands.html",
        body: MenuBody::Submenu(FILE_MENU),
    },
    MenuItem {
        letter: 'P',
        name: "Print",
        help: "Print to printer, file, or encoded file",
        help_page: "0012-the-print-commands.html",
        body: MenuBody::Submenu(PRINT_MENU),
    },
    MenuItem {
        letter: 'G',
        name: "Graph",
        help: "Create and configure graphs",
        help_page: "0010-the-graph-commands.html",
        body: MenuBody::Submenu(GRAPH_MENU),
    },
    MenuItem {
        letter: 'D',
        name: "Data",
        help: "Fill, sort, query, table, regression, ...",
        help_page: "0008-the-data-commands.html",
        body: MenuBody::Submenu(DATA_MENU),
    },
    MenuItem {
        letter: 'S',
        name: "System",
        help: "Suspend to OS shell",
        help_page: "0015-system.html",
        body: MenuBody::Action(Action::System),
    },
    MenuItem {
        letter: 'Q',
        name: "Quit",
        help: "End the l123 session",
        help_page: "0013-quit.html",
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
    fn all_ten_top_level_items_present() {
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
                "Quit"
            ]
        );
    }

    #[test]
    fn resolve_quit_yes_is_action() {
        let node = resolve(&['Q', 'Y']).unwrap();
        assert!(matches!(node.body, MenuBody::Action(Action::QuitConfirm)));
    }

    #[test]
    fn resolve_graph_reset_quit_is_cancel() {
        let node = resolve(&['G', 'R', 'Q']).unwrap();
        assert!(matches!(node.body, MenuBody::Action(Action::Cancel)));
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
    fn resolve_wgdo_clock_leaves() {
        let s = resolve(&['W', 'G', 'D', 'O', 'C', 'S']).unwrap();
        assert!(matches!(
            s.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockStandard)
        ));
        let i = resolve(&['W', 'G', 'D', 'O', 'C', 'I']).unwrap();
        assert!(matches!(
            i.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockInternational)
        ));
        let n = resolve(&['W', 'G', 'D', 'O', 'C', 'N']).unwrap();
        assert!(matches!(
            n.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockNone)
        ));
        let f = resolve(&['W', 'G', 'D', 'O', 'C', 'F']).unwrap();
        assert!(matches!(
            f.body,
            MenuBody::Action(Action::WorksheetGlobalDefaultOtherClockFilename)
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
    fn resolve_range_format_plus_minus() {
        let node = resolve(&['R', 'F', '+']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatPlusMinus)
        ));
    }

    /// /Range Format Other exposes the four extras from SPEC §12 / MENU §43:
    /// Automatic, Color, Label, Parentheses (in that order).
    #[test]
    fn range_format_other_has_four_leaves() {
        let other = resolve(&['R', 'F', 'O']).unwrap();
        let kids = children(other);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(names, vec!["Automatic", "Color", "Label", "Parentheses"]);
    }

    #[test]
    fn resolve_range_format_other_automatic() {
        let node = resolve(&['R', 'F', 'O', 'A']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatAutomatic)
        ));
    }

    #[test]
    fn resolve_wg_format_other_automatic() {
        let node = resolve(&['W', 'G', 'F', 'O', 'A']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::WorksheetGlobalFormatAutomatic)
        ));
    }

    #[test]
    fn resolve_range_format_other_label() {
        let node = resolve(&['R', 'F', 'O', 'L']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatLabelOnly)
        ));
    }

    #[test]
    fn resolve_wg_format_other_label() {
        let node = resolve(&['W', 'G', 'F', 'O', 'L']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::WorksheetGlobalFormatLabelOnly)
        ));
    }

    #[test]
    fn resolve_range_format_other_parens_yes() {
        let node = resolve(&['R', 'F', 'O', 'P', 'Y']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatParensYes)
        ));
    }

    #[test]
    fn resolve_range_format_other_parens_no() {
        let node = resolve(&['R', 'F', 'O', 'P', 'N']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatParensNo)
        ));
    }

    #[test]
    fn resolve_range_format_other_color_neg_red() {
        let node = resolve(&['R', 'F', 'O', 'C', 'N', 'R']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatNegColorRed)
        ));
    }

    #[test]
    fn resolve_range_format_other_color_reset() {
        let node = resolve(&['R', 'F', 'O', 'C', 'R']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::RangeFormatNegColorReset)
        ));
    }

    #[test]
    fn range_format_other_color_negative_has_eight_palette_entries() {
        let neg = resolve(&['R', 'F', 'O', 'C', 'N']).unwrap();
        let kids = children(neg);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(
            names,
            vec!["Black", "White", "Red", "Green", "Blue", "Yellow", "Cyan", "Magenta",]
        );
    }

    #[test]
    fn resolve_wg_format_other_parens_yes() {
        let node = resolve(&['W', 'G', 'F', 'O', 'P', 'Y']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::WorksheetGlobalFormatParensYes)
        ));
    }

    #[test]
    fn range_format_other_color_has_negative_and_reset() {
        let color = resolve(&['R', 'F', 'O', 'C']).unwrap();
        let kids = children(color);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(names, vec!["Negative", "Reset"]);
    }

    #[test]
    fn range_format_other_parens_has_yes_and_no() {
        let parens = resolve(&['R', 'F', 'O', 'P']).unwrap();
        let kids = children(parens);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(names, vec!["Yes", "No"]);
    }

    #[test]
    fn wg_format_other_has_four_leaves() {
        let other = resolve(&['W', 'G', 'F', 'O']).unwrap();
        let kids = children(other);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(names, vec!["Automatic", "Color", "Label", "Parentheses"]);
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
            (
                &['W', 'G', 'F', '+'],
                Action::WorksheetGlobalFormatPlusMinus,
            ),
            (&['W', 'G', 'F', 'T'], Action::WorksheetGlobalFormatText),
            (&['W', 'G', 'F', 'H'], Action::WorksheetGlobalFormatHidden),
            (&['W', 'G', 'F', 'R'], Action::WorksheetGlobalFormatReset),
            (
                &['W', 'G', 'F', 'D', '1'],
                Action::WorksheetGlobalFormatDateDmy,
            ),
            (
                &['W', 'G', 'F', 'D', '5'],
                Action::WorksheetGlobalFormatDateShortIntl,
            ),
            (
                &['W', 'G', 'F', 'D', 'T', '1'],
                Action::WorksheetGlobalFormatTimeHmsAmPm,
            ),
            (
                &['W', 'G', 'F', 'D', 'T', '2'],
                Action::WorksheetGlobalFormatTimeHmAmPm,
            ),
            (
                &['W', 'G', 'F', 'D', 'T', '3'],
                Action::WorksheetGlobalFormatTimeLongIntl,
            ),
            (
                &['W', 'G', 'F', 'D', 'T', '4'],
                Action::WorksheetGlobalFormatTimeShortIntl,
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
    fn resolve_wgdo_intl_punctuation_leaves() {
        let cases: &[(&[char], Action)] = &[
            (
                &['W', 'G', 'D', 'O', 'I', 'P', 'A'],
                Action::WorksheetGlobalDefaultOtherIntlPunctuationA,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'P', 'B'],
                Action::WorksheetGlobalDefaultOtherIntlPunctuationB,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'P', 'H'],
                Action::WorksheetGlobalDefaultOtherIntlPunctuationH,
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
    fn resolve_wgdo_intl_date_time_negative_currency_leaves() {
        let cases: &[(&[char], Action)] = &[
            (
                &['W', 'G', 'D', 'O', 'I', 'D', 'A'],
                Action::WorksheetGlobalDefaultOtherIntlDateA,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'D', 'D'],
                Action::WorksheetGlobalDefaultOtherIntlDateD,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'T', 'A'],
                Action::WorksheetGlobalDefaultOtherIntlTimeA,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'T', 'D'],
                Action::WorksheetGlobalDefaultOtherIntlTimeD,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'N', 'P'],
                Action::WorksheetGlobalDefaultOtherIntlNegativeParens,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'N', 'S'],
                Action::WorksheetGlobalDefaultOtherIntlNegativeSign,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'C', 'P'],
                Action::WorksheetGlobalDefaultOtherIntlCurrencyPrefix,
            ),
            (
                &['W', 'G', 'D', 'O', 'I', 'C', 'S'],
                Action::WorksheetGlobalDefaultOtherIntlCurrencySuffix,
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
    fn resolve_worksheet_titles_leaves() {
        let cases: &[(&[char], Action)] = &[
            (&['W', 'T', 'B'], Action::WorksheetTitlesBoth),
            (&['W', 'T', 'H'], Action::WorksheetTitlesHorizontal),
            (&['W', 'T', 'V'], Action::WorksheetTitlesVertical),
            (&['W', 'T', 'C'], Action::WorksheetTitlesClear),
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
    fn resolve_worksheet_page() {
        let row = resolve(&['W', 'P', 'R']).unwrap();
        assert!(matches!(
            row.body,
            MenuBody::Action(Action::WorksheetPageRow)
        ));
        let col = resolve(&['W', 'P', 'C']).unwrap();
        assert!(matches!(
            col.body,
            MenuBody::Action(Action::WorksheetPageColumn)
        ));
    }

    #[test]
    fn resolve_worksheet_hide_enable_and_disable() {
        let e = resolve(&['W', 'H', 'E']).unwrap();
        assert!(matches!(
            e.body,
            MenuBody::Action(Action::WorksheetHideEnable)
        ));
        let d = resolve(&['W', 'H', 'D']).unwrap();
        assert!(matches!(
            d.body,
            MenuBody::Action(Action::WorksheetHideDisable)
        ));
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
    fn resolve_wysiwyg_format_lines_leaves() {
        let cases: &[(&[char], Action)] = &[
            (&['F', 'L', 'O'], Action::FormatLinesOutlineSet),
            (&['F', 'L', 'L'], Action::FormatLinesLeftSet),
            (&['F', 'L', 'R'], Action::FormatLinesRightSet),
            (&['F', 'L', 'T'], Action::FormatLinesTopSet),
            (&['F', 'L', 'B'], Action::FormatLinesBottomSet),
            (&['F', 'L', 'A'], Action::FormatLinesAllSet),
            (&['F', 'L', 'C', 'O'], Action::FormatLinesOutlineClear),
            (&['F', 'L', 'C', 'L'], Action::FormatLinesLeftClear),
            (&['F', 'L', 'C', 'R'], Action::FormatLinesRightClear),
            (&['F', 'L', 'C', 'T'], Action::FormatLinesTopClear),
            (&['F', 'L', 'C', 'B'], Action::FormatLinesBottomClear),
            (&['F', 'L', 'C', 'A'], Action::FormatLinesAllClear),
        ];
        for (path, expected) in cases {
            let node = resolve_within(WYSIWYG_ROOT, path)
                .unwrap_or_else(|| panic!("resolve {path:?}"));
            match node.body {
                MenuBody::Action(actual) => assert_eq!(actual, *expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    #[test]
    fn wysiwyg_quit_maps_to_cancel() {
        let q = resolve_within(WYSIWYG_ROOT, &['Q']).unwrap();
        assert!(matches!(q.body, MenuBody::Action(Action::Cancel)));
    }

    #[test]
    fn resolve_wysiwyg_worksheet_column_set_and_reset() {
        let s = resolve_within(WYSIWYG_ROOT, &['W', 'C', 'S']).unwrap();
        assert!(matches!(
            s.body,
            MenuBody::Action(Action::WorksheetColumnSetWidth)
        ));
        let r = resolve_within(WYSIWYG_ROOT, &['W', 'C', 'R']).unwrap();
        assert!(matches!(
            r.body,
            MenuBody::Action(Action::WorksheetColumnResetWidth)
        ));
    }

    #[test]
    fn wysiwyg_worksheet_row_and_page_are_not_implemented() {
        let row = resolve_within(WYSIWYG_ROOT, &['W', 'R']).unwrap();
        assert!(matches!(row.body, MenuBody::NotImplemented(_)));
        let page = resolve_within(WYSIWYG_ROOT, &['W', 'P']).unwrap();
        assert!(matches!(page.body, MenuBody::NotImplemented(_)));
    }

    #[test]
    fn wysiwyg_worksheet_quit_maps_to_cancel() {
        let q = resolve_within(WYSIWYG_ROOT, &['W', 'Q']).unwrap();
        assert!(matches!(q.body, MenuBody::Action(Action::Cancel)));
    }

    #[test]
    fn resolve_wysiwyg_display_mode_actions() {
        let c = resolve_within(WYSIWYG_ROOT, &['D', 'M', 'C']).unwrap();
        assert!(matches!(c.body, MenuBody::Action(Action::DisplayModeColor)));
        let b = resolve_within(WYSIWYG_ROOT, &['D', 'M', 'B']).unwrap();
        assert!(matches!(b.body, MenuBody::Action(Action::DisplayModeBW)));
        let r = resolve_within(WYSIWYG_ROOT, &['D', 'M', 'R']).unwrap();
        assert!(matches!(
            r.body,
            MenuBody::Action(Action::DisplayModeReverse)
        ));
    }

    #[test]
    fn resolve_wysiwyg_display_options_grid_actions() {
        let y = resolve_within(WYSIWYG_ROOT, &['D', 'O', 'G', 'Y']).unwrap();
        assert!(matches!(
            y.body,
            MenuBody::Action(Action::DisplayOptionsGridYes)
        ));
        let n = resolve_within(WYSIWYG_ROOT, &['D', 'O', 'G', 'N']).unwrap();
        assert!(matches!(
            n.body,
            MenuBody::Action(Action::DisplayOptionsGridNo)
        ));
    }

    #[test]
    fn wysiwyg_display_options_siblings_are_not_implemented() {
        for (path, _label) in [
            (&['D', 'O', 'F'][..], "frame"),
            (&['D', 'O', 'P'][..], "page-breaks"),
            (&['D', 'O', 'C'][..], "cell-pointer"),
        ] {
            let n = resolve_within(WYSIWYG_ROOT, path).unwrap();
            assert!(matches!(n.body, MenuBody::NotImplemented(_)));
        }
    }

    #[test]
    fn wysiwyg_display_top_level_siblings_are_not_implemented() {
        for path in [
            &['D', 'Z'][..],
            &['D', 'C'][..],
            &['D', 'R'][..],
            &['D', 'F'][..],
            &['D', 'D'][..],
        ] {
            let n = resolve_within(WYSIWYG_ROOT, path).unwrap();
            assert!(matches!(n.body, MenuBody::NotImplemented(_)));
        }
    }

    #[test]
    fn resolve_file_import_text() {
        let node = resolve(&['F', 'I', 'T']).unwrap();
        assert!(matches!(
            node.body,
            MenuBody::Action(Action::FileImportText)
        ));
    }

    #[test]
    fn resolve_file_combine_leaves() {
        let cases: &[(&[char], Action)] = &[
            (&['F', 'C', 'C', 'E'], Action::FileCombineCopyEntire),
            (&['F', 'C', 'C', 'N'], Action::FileCombineCopyNamed),
            (&['F', 'C', 'A', 'E'], Action::FileCombineAddEntire),
            (&['F', 'C', 'A', 'N'], Action::FileCombineAddNamed),
            (&['F', 'C', 'S', 'E'], Action::FileCombineSubtractEntire),
            (&['F', 'C', 'S', 'N'], Action::FileCombineSubtractNamed),
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
    fn resolve_file_erase_leaves() {
        let cases: &[(&[char], Action)] = &[
            (&['F', 'E', 'W'], Action::FileEraseWorksheet),
            (&['F', 'E', 'P'], Action::FileErasePrint),
            (&['F', 'E', 'G'], Action::FileEraseGraph),
            (&['F', 'E', 'O'], Action::FileEraseOther),
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
    fn file_admin_leaves_resolve_to_typed_actions() {
        let admin = resolve(&['F', 'A']).unwrap();
        assert_eq!(admin.name, "Admin");
        assert!(matches!(admin.body, MenuBody::Submenu(_)));

        let cases: &[(&[char], Action)] = &[
            (&['F', 'A', 'R', 'G'], Action::FileAdminReservationGet),
            (&['F', 'A', 'R', 'R'], Action::FileAdminReservationRelease),
            (&['F', 'A', 'S', 'F'], Action::FileAdminSealFile),
            (
                &['F', 'A', 'S', 'R'],
                Action::FileAdminSealReservationSetting,
            ),
            (&['F', 'A', 'S', 'D'], Action::FileAdminSealDisable),
            (&['F', 'A', 'T', 'W'], Action::FileAdminTableWorksheet),
            (&['F', 'A', 'T', 'P'], Action::FileAdminTablePrint),
            (&['F', 'A', 'T', 'G'], Action::FileAdminTableGraph),
            (&['F', 'A', 'T', 'O'], Action::FileAdminTableOther),
            (&['F', 'A', 'T', 'A'], Action::FileAdminTableActive),
            (&['F', 'A', 'T', 'L'], Action::FileAdminTableLinked),
            (&['F', 'A', 'L'], Action::FileAdminLinkRefresh),
        ];
        for (path, expected) in cases {
            let node = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            match node.body {
                MenuBody::Action(actual) => assert_eq!(actual, *expected, "{path:?}"),
                other => panic!("expected Action for {path:?}, got {other:?}"),
            }
        }
    }

    /// /Data top level should expose all nine entries from MENU.md, in
    /// canonical order, with the first letter of each as accelerator.
    #[test]
    fn data_top_level_has_all_nine_entries() {
        let data = resolve(&['D']).unwrap();
        let kids = children(data);
        let names: Vec<&str> = kids.iter().map(|m| m.name).collect();
        assert_eq!(
            names,
            vec![
                "Fill",
                "Table",
                "Sort",
                "Query",
                "Distribution",
                "Matrix",
                "Regression",
                "Parse",
                "External",
            ]
        );
    }

    /// Every leaf reachable through /Data must be either an Action or a
    /// NotImplemented placeholder; no Submenu nodes that don't lead to
    /// a real terminal — the muscle-memory promise (SPEC §10) is that
    /// every path resolves to *something*, even if it just says "Not
    /// implemented yet" in line 3.
    #[test]
    fn data_subtrees_are_complete_per_menu_md() {
        // Table → 1, 2, 3, Labeled, Reset
        for c in ['1', '2', '3', 'L', 'R'] {
            let n = resolve(&['D', 'T', c]).unwrap_or_else(|| panic!("D T {c}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "D T {c} should be a leaf"
            );
        }
        // Sort → Data-Range, Primary-Key, Secondary-Key, Extra-Key,
        // Reset, Go, Quit
        for c in ['D', 'P', 'S', 'E', 'R', 'G', 'Q'] {
            let n = resolve(&['D', 'S', c]).unwrap_or_else(|| panic!("D S {c}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "D S {c} should be a leaf"
            );
        }
        // Query → Input, Criteria, Output, Find, Extract, Unique,
        // Del, Modify, Reset, Quit
        for c in ['I', 'C', 'O', 'F', 'E', 'U', 'D', 'M', 'R', 'Q'] {
            let n = resolve(&['D', 'Q', c]).unwrap_or_else(|| panic!("D Q {c}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "D Q {c} should be a leaf"
            );
        }
        // Matrix → Invert, Multiply
        for c in ['I', 'M'] {
            let n = resolve(&['D', 'M', c]).unwrap_or_else(|| panic!("D M {c}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "D M {c} should be a leaf"
            );
        }
        // External → Use, List, Create, Delete, Other, Reset, Quit
        for c in ['U', 'L', 'C', 'D', 'O', 'R', 'Q'] {
            let n = resolve(&['D', 'E', c]).unwrap_or_else(|| panic!("D E {c}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "D E {c} should be a leaf"
            );
        }
        // Fill, Distribution — single-step leaves at this level
        // (Regression and Parse have graduated to their own
        // submenus since M8).
        for path in [&['D', 'F'][..], &['D', 'D']] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            assert!(
                matches!(n.body, MenuBody::NotImplemented(_) | MenuBody::Action(_)),
                "{path:?} should be a leaf"
            );
        }
        // /Data Regression — sticky submenu (X-Range, Y-Range,
        // Output-Range, Intercept, Reset, Go, Quit). Spot-check
        // X-Range and Go to confirm the subtree is wired.
        for path in [&['D', 'R', 'X'][..], &['D', 'R', 'G']] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            assert!(
                matches!(n.body, MenuBody::Action(_)),
                "{path:?} should be an Action leaf"
            );
        }
        // /Data Parse — sticky submenu (Format-Line, Input-Column,
        // Output-Range, Reset, Go, Quit). Spot-check Input-Column
        // and Go.
        for path in [&['D', 'P', 'I'][..], &['D', 'P', 'G']] {
            let n = resolve(path).unwrap_or_else(|| panic!("resolve {path:?}"));
            assert!(
                matches!(n.body, MenuBody::Action(_)),
                "{path:?} should be an Action leaf"
            );
        }
    }

    #[test]
    fn help_page_for_empty_path_is_none() {
        assert_eq!(help_page_for_path(&[]), None);
    }

    #[test]
    fn help_page_for_top_level_items() {
        let cases: &[(&[char], &str)] = &[
            (&['W'], "0016-the-worksheet-commands.html"),
            (&['R'], "0014-the-range-commands.html"),
            (&['C'], "0006-copy.html"),
            (&['M'], "0011-move.html"),
            (&['F'], "0009-the-file-commands.html"),
            (&['P'], "0012-the-print-commands.html"),
            (&['G'], "0010-the-graph-commands.html"),
            (&['D'], "0008-the-data-commands.html"),
            (&['S'], "0015-system.html"),
            (&['Q'], "0013-quit.html"),
        ];
        for (path, expected) in cases {
            assert_eq!(help_page_for_path(path), Some(*expected), "{path:?}");
        }
    }

    #[test]
    fn help_page_walks_up_to_nearest_ancestor() {
        // /QY ("Yes" leaf) hasn't been wired, so the lookup should
        // walk up to the /Q parent's page.
        assert_eq!(
            help_page_for_path(&['Q', 'Y']),
            Some("0013-quit.html"),
        );
    }

    #[test]
    fn help_page_for_worksheet_submenu_heads() {
        let cases: &[(&[char], &str)] = &[
            (&['W', 'G'], "0417-worksheet-global.html"),
            (&['W', 'G', 'F'], "0426-worksheet-global-format.html"),
            (&['W', 'G', 'R'], "0422-worksheet-global-recalc.html"),
            (&['W', 'G', 'D'], "0425-worksheet-global-default.html"),
            (&['W', 'I'], "0462-worksheet-insert.html"),
            (&['W', 'D'], "0458-worksheet-delete.html"),
            (&['W', 'C'], "0455-worksheet-column.html"),
            (&['W', 'E'], "0460-worksheet-erase.html"),
            (&['W', 'T'], "0465-worksheet-titles.html"),
            (&['W', 'W'], "0467-worksheet-window.html"),
            (&['W', 'S'], "0464-worksheet-status.html"),
            (&['W', 'P'], "0463-worksheet-page.html"),
            (&['W', 'H'], "0461-worksheet-hide.html"),
        ];
        for (path, expected) in cases {
            assert_eq!(help_page_for_path(path), Some(*expected), "{path:?}");
        }
    }

    #[test]
    fn help_page_unknown_top_level_letter_is_none() {
        assert_eq!(help_page_for_path(&['Z']), None);
    }

    #[test]
    fn help_page_partial_invalid_falls_back_to_valid_prefix() {
        // The walk-up tries each prefix; if the deepest letter doesn't
        // resolve, the next-deepest prefix that does is used.
        assert_eq!(help_page_for_path(&['Q', 'Z']), Some("0013-quit.html"));
    }

    #[test]
    fn help_page_within_nested_root() {
        // QUIT_DIRTY_MENU has no items wired with help_page, so the
        // within-variant should return None at every depth.
        assert_eq!(help_page_within(QUIT_DIRTY_MENU, &[]), None);
        assert_eq!(help_page_within(QUIT_DIRTY_MENU, &['N']), None);
    }

    #[test]
    fn help_page_for_range_file_print_graph_data_heads() {
        let cases: &[(&[char], &str)] = &[
            (&['R', 'F'], "0394-range-format.html"),
            (&['R', 'N'], "0372-range-name.html"),
            (&['R', 'N', 'C'], "0375-range-name-create.html"),
            (&['F', 'C'], "0028-file-combine.html"),
            (&['F', 'A'], "0017-file-admin.html"),
            (&['F', 'A', 'R'], "0019-file-admin-reservation.html"),
            (&['F', 'A', 'S'], "0024-file-admin-seal.html"),
            (&['F', 'A', 'T'], "0026-file-admin-table.html"),
            (&['G', 'T'], "0060-graph-type.html"),
            (&['D', 'S'], "0157-data-sort.html"),
            (&['D', 'S', 'D'], "0159-data-sort-data-range.html"),
            (&['D', 'Q'], "0120-data-query.html"),
            (&['D', 'Q', 'I'], "0123-data-query-input.html"),
            (&['D', 'R'], "0150-data-regression.html"),
            (&['D', 'R', 'X'], "0154-data-regression-x-range.html"),
            (&['D', 'P'], "0111-data-parse.html"),
            (&['D', 'P', 'I'], "0118-data-parse-input-column.html"),
            (&['D', 'M'], "0108-data-matrix.html"),
            (&['D', 'M', 'I'], "0109-data-matrix-invert.html"),
        ];
        for (path, expected) in cases {
            assert_eq!(help_page_for_path(path), Some(*expected), "{path:?}");
        }
    }

    #[test]
    fn help_page_for_wysiwyg_root_uses_using_wysiwyg() {
        // Every WYSIWYG_ROOT submenu head except Quit points at the
        // overview page so the colon-menu has consistent F1.
        let cases: &[&[char]] = &[
            &['W'], &['F'], &['G'], &['N'], &['P'], &['D'], &['S'], &['T'],
        ];
        for path in cases {
            assert_eq!(
                help_page_within(WYSIWYG_ROOT, path),
                Some("0816-using-wysiwyg.html"),
                "{path:?}",
            );
        }
    }
}
