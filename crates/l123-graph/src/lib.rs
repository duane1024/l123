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
pub mod icon_rle;
pub mod icons;
pub mod raster;
pub mod render;
pub use icon_rle::{decode_color_bitmap, ColorBitmap, RleError, ICON_DIM, ICON_PIXELS};
pub use icons::{icon_action, render_panel_png, slot_description, IconAction, Panel, SysAction};
pub use raster::{
    render_dynamic_image_sized, render_png, render_png_sized, render_svg, render_svg_sized,
};
pub use render::{render as render_unicode, GraphValues};

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

/// `/Graph Type Features Vertical|Horizontal` — orientation of the
/// independent axis. R3.1 Reference p. 2-194: Vertical is the default.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum Orientation {
    #[default]
    Vertical,
    Horizontal,
}

/// `/Graph Type Features {Y|2Y}-Ranges` — which y-axis a data series
/// is plotted against. Default is the first y-axis.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum YAxis {
    #[default]
    First,
    Second,
}

/// `/Graph Options Format` — how a series renders in line / mixed / XY
/// / HLCO graphs. Reference p. 2-200, 0089-graph-options-format.html:
/// "Both (Default)".
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum LineFormat {
    Lines,
    Symbols,
    #[default]
    Both,
    Neither,
    Area,
}

/// `/Graph Options Scale {Y|X|2Y} Automatic|Manual`. Slice D4 of the
/// graph plan fleshes out the rest of the per-axis settings (lower,
/// upper, format, indicator, type, exponent, width).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum ScaleMode {
    #[default]
    Automatic,
    Manual,
}

/// Which sides of a `/Graph Type Features Frame` are drawn.
///
/// 1-2-3 ships graphs with all four outer sides on (the user can
/// clear any of them with `/Graph Type Features Frame {Side} No`);
/// the manual `Default` impl below sets every outer flag to `true`
/// to match. `y_axis` is the **inner** y-axis line that 1-2-3 draws
/// just inside the Left edge — distinct from the outer Left edge.
/// It is off by default; `/GTF Frame Y-Axis Yes` turns it on.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash)]
pub struct FrameMask {
    pub left: bool,
    pub right: bool,
    pub top: bool,
    pub bottom: bool,
    pub y_axis: bool,
}

impl Default for FrameMask {
    fn default() -> Self {
        Self {
            left: true,
            right: true,
            top: true,
            bottom: true,
            y_axis: false,
        }
    }
}

/// `/Graph Options Grid Y-Axis` — which y-axis the horizontal grid
/// lines originate from when the user enables a Y-Axis-anchored
/// grid. `Both` means horizontal grids are drawn at every tick of
/// each axis. Reference p. 2-203.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum GridYAxisOrigin {
    #[default]
    First,
    Second,
    Both,
}

impl GridYAxisOrigin {
    pub fn tag(self) -> &'static str {
        match self {
            GridYAxisOrigin::First => "Y",
            GridYAxisOrigin::Second => "2Y",
            GridYAxisOrigin::Both => "Both",
        }
    }
}

/// `/Graph Options Grid` — independent toggles for horizontal /
/// vertical grid lines plus a Y-Axis-anchored grid origin. Default
/// is no grid in any direction (`y_axis = None`).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub struct GridMask {
    pub horizontal: bool,
    pub vertical: bool,
    pub y_axis: Option<GridYAxisOrigin>,
}

/// `/Graph Options Data-Labels {slot}` placement — where the label
/// sits relative to its data point. Default `Above` matches 1-2-3
/// R3.4a's "Center" rendering for unspecified placement (the
/// dialog's first leaf in the original menu). Reference p. 2-217.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum DataLabelPlacement {
    Center,
    Left,
    #[default]
    Above,
    Right,
    Below,
}

impl DataLabelPlacement {
    /// Stable ASCII-cased tag for the placement, used by the
    /// settings panel and `ASSERT_GRAPH_DATA_LABELS_PLACEMENT`
    /// transcript directive.
    pub fn tag(self) -> &'static str {
        match self {
            DataLabelPlacement::Center => "Center",
            DataLabelPlacement::Left => "Left",
            DataLabelPlacement::Above => "Above",
            DataLabelPlacement::Right => "Right",
            DataLabelPlacement::Below => "Below",
        }
    }
}

/// `/Graph Options Titles` — the seven independently editable text
/// strings that decorate a graph. Reference p. 2-216.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Titles {
    pub first: Option<String>,
    pub second: Option<String>,
    pub x_axis: Option<String>,
    pub y_axis: Option<String>,
    pub two_y_axis: Option<String>,
    pub note: Option<String>,
    pub other_note: Option<String>,
}

/// `/Graph Options Scale {axis} Type` — Linear (default) or
/// Logarithmic. Reference p. 2-204.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum ScaleType {
    #[default]
    Linear,
    Logarithmic,
}

impl ScaleType {
    /// Stable ASCII tag, used by the `ASSERT_GRAPH_SCALE_TYPE`
    /// transcript directive and the settings panel.
    pub fn tag(self) -> &'static str {
        match self {
            ScaleType::Linear => "Linear",
            ScaleType::Logarithmic => "Logarithmic",
        }
    }
}

/// `/Graph Options Scale {axis} Indicator` — controls whether the
/// per-axis magnitude indicator (e.g. "× 1000") appears next to
/// the tick labels.
///
/// `Yes` (default) lets the renderer auto-pick when the data
/// range warrants one. `No` suppresses it. `Manual` is
/// authentic-Lotus terminology for "I'll supply the indicator
/// text" — the per-axis text prompt that drives Manual is a
/// follow-up; for now the variant is stored but the indicator
/// behaves like `Yes`. Reference p. 2-204.
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Default)]
pub enum ScaleIndicator {
    #[default]
    Yes,
    No,
    Manual,
}

impl ScaleIndicator {
    /// Stable ASCII tag, used by the
    /// `ASSERT_GRAPH_SCALE_INDICATOR` transcript directive.
    pub fn tag(self) -> &'static str {
        match self {
            ScaleIndicator::Yes => "Yes",
            ScaleIndicator::No => "No",
            ScaleIndicator::Manual => "Manual",
        }
    }
}

/// `/Graph Options Scale {axis}` settings, per axis.
///
/// `lower` / `upper` are the manual bounds set via
/// `/Graph Options Scale {Y|X|2Y} {Lower|Upper}`. They're stored
/// independently of `mode` — 1-2-3 retains them when you toggle
/// back to Automatic so re-entering Manual restores the prior
/// limits. `type_` selects the linear / logarithmic mapping.
/// `width` is the maximum character width for the per-axis tick
/// labels; 0 means "auto" (renderer picks). `exponent` is the
/// order-of-magnitude shift applied to the labels (e.g. `3` shows
/// "× 1000" indicator and divides labels by 1000); 0 lets the
/// renderer pick automatically.
#[derive(Copy, Clone, Debug, PartialEq, Default)]
pub struct ScaleAxis {
    pub mode: ScaleMode,
    pub lower: Option<f64>,
    pub upper: Option<f64>,
    pub type_: ScaleType,
    pub width: u8,
    pub exponent: i8,
    pub indicator: ScaleIndicator,
}

impl ScaleAxis {
    /// Apply this axis' scale settings to a data-derived
    /// `(lo, hi)` pair. With `mode == Automatic` the data range
    /// passes through unchanged. With `Manual`, `lower` and
    /// `upper` (when set) override the corresponding side; the
    /// other side falls back to the data extent. After overriding,
    /// callers should still ensure `lo <= hi` and treat a zero
    /// span as a unit interval to avoid div-by-zero.
    pub fn apply(&self, data_lo: f64, data_hi: f64) -> (f64, f64) {
        if self.mode != ScaleMode::Manual {
            return (data_lo, data_hi);
        }
        (self.lower.unwrap_or(data_lo), self.upper.unwrap_or(data_hi))
    }
}

/// `/Graph Options Advanced` — colors, hatch patterns, and text
/// fonts/sizes. Empty placeholder for slice F of the graph plan; all
/// the leaves under Advanced affect raster output only, not the
/// settings panel.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct Advanced {}

/// `/Graph Type Features` — the ten orientation / stacking / framing
/// flags exposed under the Features submenu. Reference pp. 2-194 →
/// 2-198.
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct GraphFeatures {
    pub orientation: Orientation,
    pub stacked: bool,
    pub percent: bool,
    /// Which y-axis each of A..F plots against.
    pub y_axis: [YAxis; 6],
    pub frame: FrameMask,
    pub drop_shadow: bool,
    pub three_d: bool,
    pub table: bool,
}

/// `/Graph Options` — everything reachable from the Options submenu.
/// Reference pp. 2-200 → 2-218.
///
/// Manual `Default` impl below: `color = true` (color on, B&W off) and
/// `skip = 1` (every x-axis label drawn). Every other field uses its
/// type's derived default.
#[derive(Clone, Debug, PartialEq)]
pub struct GraphOptions {
    pub legend: [Option<String>; 6],
    pub format: [LineFormat; 6],
    pub titles: Titles,
    pub grid: GridMask,
    pub scale_y: ScaleAxis,
    pub scale_x: ScaleAxis,
    pub scale_2y: ScaleAxis,
    /// `/Graph Options Scale Skip` — every Nth x-axis label rendered.
    pub skip: u32,
    /// `true` = `/Graph Options Color`; `false` = `/Graph Options B&W`.
    pub color: bool,
    pub data_labels: [Option<Range>; 6],
    /// Per-series placement of the data-label glyph relative to its
    /// data point. Set independently from `data_labels` so changing
    /// the placement after binding the range doesn't have to re-walk
    /// the cells.
    pub data_labels_placement: [DataLabelPlacement; 6],
    pub advanced: Advanced,
}

impl Default for GraphOptions {
    fn default() -> Self {
        Self {
            legend: Default::default(),
            format: Default::default(),
            titles: Titles::default(),
            grid: GridMask::default(),
            scale_y: ScaleAxis::default(),
            scale_x: ScaleAxis::default(),
            scale_2y: ScaleAxis::default(),
            skip: 1,
            color: true,
            data_labels: Default::default(),
            data_labels_placement: [DataLabelPlacement::default(); 6],
            advanced: Advanced::default(),
        }
    }
}

/// All state that makes up a single graph definition.
#[derive(Clone, Debug, PartialEq, Default)]
pub struct GraphDef {
    pub graph_type: GraphType,
    pub x: Option<Range>,
    /// Data series A..F in positional order.
    pub data: [Option<Range>; 6],
    pub features: GraphFeatures,
    pub options: GraphOptions,
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

    /// Equivalent to `/Graph Reset Graph`: clear every range, every
    /// feature, and every option, and return the type to the default.
    pub fn reset(&mut self) {
        *self = Self::default();
    }

    /// Equivalent to `/Graph Reset Ranges`: clear X and A..F, leave
    /// `graph_type`, `features`, and `options` alone.
    pub fn reset_ranges(&mut self) {
        self.x = None;
        self.data = Default::default();
    }

    /// Equivalent to `/Graph Reset Options`: restore `options` to its
    /// default. Ranges, type, and Type Features are preserved.
    pub fn reset_options(&mut self) {
        self.options = GraphOptions::default();
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
            start: Address {
                sheet: SheetId(0),
                col: c0,
                row: r0,
            },
            end: Address {
                sheet: SheetId(0),
                col: c1,
                row: r1,
            },
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
        for (i, s) in [
            Series::A,
            Series::B,
            Series::C,
            Series::D,
            Series::E,
            Series::F,
        ]
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
        let mut g = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
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

    #[test]
    fn features_default_matches_reference_p_2_194() {
        let f = GraphFeatures::default();
        assert_eq!(f.orientation, Orientation::Vertical);
        assert!(!f.stacked);
        assert!(!f.percent);
        assert_eq!(f.y_axis, [YAxis::First; 6]);
        assert!(!f.drop_shadow);
        assert!(!f.three_d);
        assert!(!f.table);
    }

    #[test]
    fn frame_default_has_all_four_sides_on() {
        let frame = FrameMask::default();
        assert!(frame.left);
        assert!(frame.right);
        assert!(frame.top);
        assert!(frame.bottom);
    }

    #[test]
    fn grid_default_has_no_lines() {
        let g = GridMask::default();
        assert!(!g.horizontal);
        assert!(!g.vertical);
        assert!(g.y_axis.is_none());
    }

    #[test]
    fn options_default_color_on_skip_one_no_titles_no_grid() {
        let o = GraphOptions::default();
        assert!(o.color, "Reference default is /Graph Options Color");
        assert_eq!(o.skip, 1, "every x-axis label drawn by default");
        assert_eq!(o.scale_y.mode, ScaleMode::Automatic);
        assert_eq!(o.scale_x.mode, ScaleMode::Automatic);
        assert_eq!(o.scale_2y.mode, ScaleMode::Automatic);
        assert_eq!(o.titles, Titles::default());
        assert_eq!(o.grid, GridMask::default());
        assert!(o.legend.iter().all(Option::is_none));
        assert!(o.data_labels.iter().all(Option::is_none));
        assert_eq!(
            o.format,
            [LineFormat::Both; 6],
            "Reference default is Both (lines + symbols)"
        );
    }

    #[test]
    fn graph_def_carries_features_and_options() {
        let g = GraphDef::default();
        assert_eq!(g.features, GraphFeatures::default());
        assert_eq!(g.options, GraphOptions::default());
    }

    #[test]
    fn reset_ranges_clears_x_and_a_through_f_only() {
        let mut g = GraphDef {
            graph_type: GraphType::Bar,
            features: GraphFeatures {
                orientation: Orientation::Horizontal,
                ..Default::default()
            },
            options: GraphOptions {
                color: false,
                ..Default::default()
            },
            ..Default::default()
        };
        g.set(Series::X, r(0, 0, 0, 5));
        g.set(Series::A, r(1, 0, 1, 5));
        g.set(Series::F, r(5, 0, 5, 5));

        g.reset_ranges();

        assert!(g.is_empty(), "ranges cleared");
        assert_eq!(g.graph_type, GraphType::Bar, "graph_type preserved");
        assert_eq!(
            g.features.orientation,
            Orientation::Horizontal,
            "features preserved"
        );
        assert!(!g.options.color, "options preserved");
    }

    #[test]
    fn reset_options_clears_options_only() {
        let mut g = GraphDef {
            graph_type: GraphType::Bar,
            features: GraphFeatures {
                orientation: Orientation::Horizontal,
                ..Default::default()
            },
            options: GraphOptions {
                color: false,
                skip: 7,
                ..Default::default()
            },
            ..Default::default()
        };
        g.set(Series::A, r(1, 0, 1, 5));

        g.reset_options();

        assert_eq!(g.options, GraphOptions::default(), "options reset");
        assert_eq!(g.graph_type, GraphType::Bar, "graph_type preserved");
        assert_eq!(
            g.features.orientation,
            Orientation::Horizontal,
            "features preserved"
        );
        assert_eq!(g.get(Series::A), Some(r(1, 0, 1, 5)), "ranges preserved");
    }

    #[test]
    fn reset_clears_features_and_options_too() {
        let mut g = GraphDef {
            graph_type: GraphType::Bar,
            features: GraphFeatures {
                orientation: Orientation::Horizontal,
                three_d: true,
                ..Default::default()
            },
            options: GraphOptions {
                color: false,
                ..Default::default()
            },
            ..Default::default()
        };
        g.reset();
        assert_eq!(g, GraphDef::default());
    }
}
