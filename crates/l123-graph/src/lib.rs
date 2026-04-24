//! Pure-data graph model for L123.
//!
//! This crate defines the shape of a 1-2-3 graph: its type, its X-axis
//! source range, and its six data-series ranges (A through F). It does
//! no rendering — see later slices for the unicode and raster backends.
//!
//! A [`GraphDef`] is the value stored in [`Workbook::graphs`] by name,
//! and also as the workbook's unnamed "current graph" that `/Graph`
//! menu commands mutate in place.

use std::collections::BTreeMap;

use l123_core::Range;

pub mod icon_data;
pub mod icons;
pub mod render;
pub mod raster;
pub use icons::{icon_action, render_panel_png, IconAction, Panel, SysAction};
pub use render::{render as render_unicode, GraphValues};
pub use raster::{render_png, render_png_sized, render_svg, render_svg_sized};

/// The seven graph types 1-2-3 R3.4a supports.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum GraphType {
    #[default]
    Line,
    Bar,
    XY,
    Stack,
    Pie,
    /// High/Low/Close/Open (financial candlestick-ish).
    HLCO,
    Mixed,
}

/// Which slot of a [`GraphDef`] a range fills.
///
/// `X` is the independent-axis range; `A`..`F` are the six data series.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub enum Series {
    X,
    A,
    B,
    C,
    D,
    E,
    F,
}

/// All state that makes up a single graph definition.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct GraphDef {
    pub graph_type: GraphType,
    pub x: Option<Range>,
    /// Data series A..F in positional order.
    pub data: [Option<Range>; 6],
}

impl GraphDef {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn get(&self, s: Series) -> Option<Range> {
        match s {
            Series::X => self.x,
            Series::A => self.data[0],
            Series::B => self.data[1],
            Series::C => self.data[2],
            Series::D => self.data[3],
            Series::E => self.data[4],
            Series::F => self.data[5],
        }
    }

    pub fn set(&mut self, s: Series, r: Range) {
        let slot = self.slot_mut(s);
        *slot = Some(r);
    }

    pub fn clear(&mut self, s: Series) {
        *self.slot_mut(s) = None;
    }

    fn slot_mut(&mut self, s: Series) -> &mut Option<Range> {
        match s {
            Series::X => &mut self.x,
            Series::A => &mut self.data[0],
            Series::B => &mut self.data[1],
            Series::C => &mut self.data[2],
            Series::D => &mut self.data[3],
            Series::E => &mut self.data[4],
            Series::F => &mut self.data[5],
        }
    }

    /// Equivalent to `/Graph Reset Graph`: clear every range and return
    /// the type to the default.
    pub fn reset(&mut self) {
        *self = Self::default();
    }

    pub fn is_empty(&self) -> bool {
        self.x.is_none() && self.data.iter().all(Option::is_none)
    }
}

/// Collection of named graph definitions persisted with a workbook.
pub type NamedGraphs = BTreeMap<String, GraphDef>;

#[cfg(test)]
mod tests {
    use super::*;
    use l123_core::{Address, SheetId};

    fn r(c0: u16, r0: u32, c1: u16, r1: u32) -> Range {
        Range {
            start: Address { sheet: SheetId(0), col: c0, row: r0 },
            end: Address { sheet: SheetId(0), col: c1, row: r1 },
        }
    }

    #[test]
    fn default_is_line_with_no_ranges() {
        let g = GraphDef::default();
        assert_eq!(g.graph_type, GraphType::Line);
        assert!(g.is_empty());
        assert!(g.x.is_none());
        for slot in &g.data {
            assert!(slot.is_none());
        }
    }

    #[test]
    fn set_x_stores_in_x_slot() {
        let mut g = GraphDef::default();
        let range = r(0, 0, 0, 9);
        g.set(Series::X, range);
        assert_eq!(g.get(Series::X), Some(range));
        assert_eq!(g.get(Series::A), None);
        assert!(!g.is_empty());
    }

    #[test]
    fn set_a_through_f_independent_slots() {
        let mut g = GraphDef::default();
        for (i, s) in [Series::A, Series::B, Series::C, Series::D, Series::E, Series::F]
            .iter()
            .enumerate()
        {
            g.set(*s, r(i as u16, 0, i as u16, 9));
        }
        assert_eq!(g.get(Series::A).unwrap().start.col, 0);
        assert_eq!(g.get(Series::B).unwrap().start.col, 1);
        assert_eq!(g.get(Series::C).unwrap().start.col, 2);
        assert_eq!(g.get(Series::D).unwrap().start.col, 3);
        assert_eq!(g.get(Series::E).unwrap().start.col, 4);
        assert_eq!(g.get(Series::F).unwrap().start.col, 5);
        assert_eq!(g.get(Series::X), None);
    }

    #[test]
    fn clear_removes_only_that_slot() {
        let mut g = GraphDef::default();
        let range = r(0, 0, 0, 9);
        g.set(Series::A, range);
        g.set(Series::B, range);
        g.clear(Series::A);
        assert_eq!(g.get(Series::A), None);
        assert_eq!(g.get(Series::B), Some(range));
    }

    #[test]
    fn reset_restores_default_type_and_clears_ranges() {
        let mut g = GraphDef { graph_type: GraphType::Bar, ..Default::default() };
        g.set(Series::X, r(0, 0, 0, 5));
        g.set(Series::A, r(1, 0, 1, 5));
        g.set(Series::F, r(5, 0, 5, 5));
        g.reset();
        assert_eq!(g, GraphDef::default());
    }

    #[test]
    fn graph_type_defaults_to_line() {
        assert_eq!(GraphType::default(), GraphType::Line);
    }

    #[test]
    fn named_graphs_is_btreemap_sorted() {
        let mut m: NamedGraphs = NamedGraphs::new();
        m.insert("sales_q4".into(), GraphDef::default());
        m.insert("sales_q1".into(), GraphDef::default());
        m.insert("sales_q2".into(), GraphDef::default());
        let keys: Vec<&String> = m.keys().collect();
        assert_eq!(keys, vec!["sales_q1", "sales_q2", "sales_q4"]);
    }
}
