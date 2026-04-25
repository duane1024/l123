//! Excel "tables" (a.k.a. ListObjects): named rectangular ranges with
//! header rows, optional totals row, autofilter metadata, and a
//! style.  See `docs/XLSX_IMPORT_PLAN.md` §3.2.
//!
//! 1-2-3 R3.4a had no first-class tables surface.  v1 of this type is
//! round-trip-only: data flows xlsx → engine → UI → engine → xlsx
//! (subject to IronCalc 0.7's exporter gap, see the engine adapter
//! tests).  No filter UI yet; the eventual home is `/Data Query Define`
//! per `docs/MENU.md`.
//!
//! ## Fidelity
//!
//! IronCalc's `Table` carries `dxf_id` fields for header / data /
//! totals row formatting.  L123 drops these — they reference
//! workbook-wide differential-format records that L123 doesn't model.
//! Documented round-trip loss; styling on tables degrades to the
//! workbook's regular cell-style overrides.
//!
//! Threaded comments / pivot tables / charts referencing tables are
//! all out of scope here.

use crate::{Address, Range, SheetId};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    /// Internal identifier — unique within the workbook.  Used by xlsx
    /// formulas (e.g. `Table1[Year]`) so it must match what the cells
    /// reference.
    pub name: String,
    /// User-facing display name.  Often equals `name` in practice;
    /// stored separately because the xlsx schema distinguishes them.
    pub display_name: String,
    /// The data range covered by the table — top-left to bottom-right
    /// inclusive of the header row (when present).
    pub range: Range,
    /// True when the table has an autofilter row (Excel's "filter
    /// arrows" UI).  L123 doesn't yet render the filter widgets, but
    /// the flag round-trips.
    pub has_filters: bool,
    /// True when the table has a header row at the top of `range`.
    /// Almost always true for real tables.
    pub has_header_row: bool,
    /// True when the table reserves the bottom row for totals.  When
    /// set, [`TableColumn::totals_row_function`] / `totals_row_label`
    /// describe what each column shows.
    pub has_totals_row: bool,
    pub columns: Vec<TableColumn>,
    pub style: TableStyle,
}

impl Default for Table {
    fn default() -> Self {
        // `Range` deliberately has no Default impl — a "zero range"
        // isn't meaningful at the type level — so we synthesize a
        // single-cell A1:A1 placeholder.  Real callers always
        // overwrite the range immediately.
        let zero = Address::new(SheetId::A, 0, 0);
        Self {
            name: String::new(),
            display_name: String::new(),
            range: Range {
                start: zero,
                end: zero,
            },
            has_filters: false,
            has_header_row: false,
            has_totals_row: false,
            columns: Vec::new(),
            style: TableStyle::default(),
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableColumn {
    /// 1-based column id within the table.  Excel uses these to
    /// disambiguate columns across rename operations.
    pub id: u32,
    /// Column header label (the text shown in the header row).
    pub name: String,
    /// Optional totals-row function (e.g. `"sum"`, `"average"`).
    /// `None` when the totals row uses a literal label or is absent.
    pub totals_row_function: Option<String>,
    /// Optional totals-row literal label (mutually exclusive with
    /// `totals_row_function` in well-formed tables).
    pub totals_row_label: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TableStyle {
    /// Built-in or custom style name (e.g. `"TableStyleMedium2"`).
    /// `None` means use the workbook default.
    pub name: Option<String>,
    pub show_first_column: bool,
    pub show_last_column: bool,
    pub show_row_stripes: bool,
    pub show_column_stripes: bool,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_is_empty() {
        let t = Table::default();
        assert!(t.name.is_empty());
        assert!(t.columns.is_empty());
        assert!(!t.has_filters);
        assert!(!t.has_header_row);
        assert!(!t.has_totals_row);
        assert_eq!(t.style, TableStyle::default());
    }

    #[test]
    fn columns_round_trip_through_struct() {
        let t = Table {
            name: "Sales".into(),
            display_name: "Sales".into(),
            range: Range {
                start: Address::new(SheetId::A, 0, 0),
                end: Address::new(SheetId::A, 3, 5),
            },
            has_filters: true,
            has_header_row: true,
            has_totals_row: false,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Year".into(),
                    ..Default::default()
                },
                TableColumn {
                    id: 2,
                    name: "Total".into(),
                    totals_row_function: Some("sum".into()),
                    ..Default::default()
                },
            ],
            style: TableStyle {
                name: Some("TableStyleMedium2".into()),
                show_row_stripes: true,
                ..Default::default()
            },
        };
        assert_eq!(t.columns.len(), 2);
        assert_eq!(t.columns[1].totals_row_function.as_deref(), Some("sum"));
    }
}
