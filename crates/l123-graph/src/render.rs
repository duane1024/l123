//! Unicode rendering backend for the full-screen graph view (F10).
//!
//! Produces plain ratatui `Buffer` output using block characters for
//! bar charts and single-character sparkline dots for line charts.
//! Other graph types show a placeholder until the raster path (slice 5)
//! lands.
//!
//! Values arrive pre-resolved as [`GraphValues`]; this crate intentionally
//! knows nothing about [`Engine`](l123_engine) — the UI layer walks the
//! ranges and hands numbers in.

use ratatui::{
    buffer::Buffer,
    layout::Rect,
    style::{Color, Style},
};

use crate::{GraphDef, GraphType};

/// Numeric values resolved per series slot.
///
/// `None` means the corresponding series has no range set. An empty
/// `Vec` means the range is set but yielded no numeric cells. Cells
/// that are blank or non-numeric show up as `f64::NAN` so the positional
/// alignment between series is preserved.
#[derive(Clone, Debug, Default)]
pub struct GraphValues {
    pub x: Option<Vec<f64>>,
    pub data: [Option<Vec<f64>>; 6],
}

impl GraphValues {
    /// True when every series slot (including X) is unset. A graph
    /// with no data is not worth rendering; the caller should show
    /// a "Define ranges first" message instead.
    pub fn is_empty(&self) -> bool {
        self.x.is_none() && self.data.iter().all(Option::is_none)
    }

    /// The first data series that has any numeric values, in A..F order.
    pub fn first_series(&self) -> Option<&[f64]> {
        self.data.iter().find_map(|opt| opt.as_deref())
    }
}

/// Render a graph into `buf` over `area`. Falls back to a one-line
/// placeholder for types not yet implemented (slice 5 fills in).
pub fn render(def: &GraphDef, vals: &GraphValues, area: Rect, buf: &mut Buffer) {
    clear(area, buf);
    if area.width < 8 || area.height < 4 {
        return;
    }
    if vals.is_empty() {
        write_centered(
            area,
            buf,
            "No graph ranges set — /Graph X and /Graph A..F define them.",
        );
        return;
    }
    match def.graph_type {
        GraphType::Bar => render_bar(vals, area, buf),
        GraphType::Line => render_line(vals, area, buf),
        other => write_centered(
            area,
            buf,
            &format!("{other:?} graphs render in a later slice; press Esc to return."),
        ),
    }
}

fn clear(area: Rect, buf: &mut Buffer) {
    let blank = Style::default().bg(Color::Reset).fg(Color::Reset);
    for y in area.top()..area.bottom() {
        for x in area.left()..area.right() {
            let cell = &mut buf[(x, y)];
            cell.set_symbol(" ");
            cell.set_style(blank);
        }
    }
}

fn write_centered(area: Rect, buf: &mut Buffer, msg: &str) {
    let y = area.top() + area.height / 2;
    let x0 = area.left() + area.width.saturating_sub(msg.chars().count() as u16) / 2;
    buf.set_string(x0, y, msg, Style::default());
}

/// Half-block bar chart. One column per A-series value. Height in
/// half-rows is `2 * (row_height) * val / max_abs`, rounded; odd half
/// steps use `▄` (lower half block) / `▀` (upper half block) for the
/// fractional cell at the top of the bar.
fn render_bar(vals: &GraphValues, area: Rect, buf: &mut Buffer) {
    let Some(series) = vals.first_series() else {
        write_centered(area, buf, "No numeric A-series values to plot.");
        return;
    };
    // Leave the bottom row for the baseline axis.
    let plot_top = area.top();
    let plot_bottom = area.bottom().saturating_sub(1);
    let plot_height = plot_bottom.saturating_sub(plot_top);
    if plot_height < 2 || series.is_empty() {
        return;
    }
    let max = series
        .iter()
        .copied()
        .filter(|v| v.is_finite())
        .fold(f64::NEG_INFINITY, f64::max);
    let max = if max <= 0.0 || !max.is_finite() { 1.0 } else { max };

    let bar_count = series.len() as u16;
    let plot_width = area.width;
    // One column per bar, one column of padding between them when we
    // can afford it.
    let step = (plot_width / bar_count).max(1);
    let bar_width = step.saturating_sub(1).max(1);

    let full = "█";
    let half = "▄";
    let style = Style::default().fg(Color::Cyan);

    for (i, v) in series.iter().copied().enumerate() {
        if !v.is_finite() || v <= 0.0 {
            continue;
        }
        let x0 = area.left() + (i as u16) * step;
        let height_halves = ((v / max) * plot_height as f64 * 2.0).round() as u16;
        let full_rows = height_halves / 2;
        let has_half = height_halves % 2 == 1;
        for row in 0..full_rows {
            let y = plot_bottom.saturating_sub(1 + row);
            for bx in 0..bar_width {
                let x = x0 + bx;
                if x < area.right() && y >= area.top() {
                    let cell = &mut buf[(x, y)];
                    cell.set_symbol(full);
                    cell.set_style(style);
                }
            }
        }
        if has_half {
            let y = plot_bottom.saturating_sub(1 + full_rows);
            for bx in 0..bar_width {
                let x = x0 + bx;
                if x < area.right() && y >= area.top() {
                    let cell = &mut buf[(x, y)];
                    cell.set_symbol(half);
                    cell.set_style(style);
                }
            }
        }
    }
    // Baseline row of `─`.
    for x in area.left()..area.right() {
        let cell = &mut buf[(x, plot_bottom)];
        cell.set_symbol("─");
        cell.set_style(Style::default().fg(Color::Gray));
    }
}

/// Dot-per-sample line chart. Each A-series point is a `•` placed
/// at its y-position. Spans the full plot width evenly regardless
/// of how many samples there are.
fn render_line(vals: &GraphValues, area: Rect, buf: &mut Buffer) {
    let Some(series) = vals.first_series() else {
        write_centered(area, buf, "No numeric A-series values to plot.");
        return;
    };
    let plot_top = area.top();
    let plot_bottom = area.bottom().saturating_sub(1);
    let plot_height = plot_bottom.saturating_sub(plot_top);
    if plot_height < 2 || series.is_empty() {
        return;
    }
    let (min, max) = series.iter().copied().filter(|v| v.is_finite()).fold(
        (f64::INFINITY, f64::NEG_INFINITY),
        |(lo, hi), v| (lo.min(v), hi.max(v)),
    );
    let (min, max) = if !min.is_finite() || !max.is_finite() || min == max {
        (0.0, 1.0)
    } else {
        (min, max)
    };
    let style = Style::default().fg(Color::Cyan);
    let span = max - min;
    let n = series.len().max(1);
    let denom = (n - 1).max(1) as f64;
    for (i, v) in series.iter().copied().enumerate() {
        if !v.is_finite() {
            continue;
        }
        let frac_x = if n == 1 { 0.5 } else { i as f64 / denom };
        let x = area.left() + (frac_x * (area.width as f64 - 1.0)).round() as u16;
        let frac_y = (v - min) / span;
        let y_from_bottom = (frac_y * (plot_height as f64 - 1.0)).round() as u16;
        let y = plot_bottom.saturating_sub(1 + y_from_bottom);
        if x < area.right() && y >= area.top() && y < area.bottom() {
            let cell = &mut buf[(x, y)];
            cell.set_symbol("•");
            cell.set_style(style);
        }
    }
    for x in area.left()..area.right() {
        let cell = &mut buf[(x, plot_bottom)];
        cell.set_symbol("─");
        cell.set_style(Style::default().fg(Color::Gray));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn render_to(def: GraphType, a: Vec<f64>, w: u16, h: u16) -> Buffer {
        let area = Rect::new(0, 0, w, h);
        let mut buf = Buffer::empty(area);
        let mut vals = GraphValues::default();
        vals.data[0] = Some(a);
        let gdef = GraphDef { graph_type: def, ..Default::default() };
        render(&gdef, &vals, area, &mut buf);
        buf
    }

    fn contains(buf: &Buffer, needle: &str) -> bool {
        for y in 0..buf.area.height {
            let mut row = String::new();
            for x in 0..buf.area.width {
                row.push_str(buf[(x, y)].symbol());
            }
            if row.contains(needle) {
                return true;
            }
        }
        false
    }

    #[test]
    fn bar_renders_full_block() {
        let buf = render_to(GraphType::Bar, vec![1.0, 2.0, 3.0], 40, 10);
        assert!(contains(&buf, "█"), "no █ in bar chart");
        assert!(contains(&buf, "─"), "no baseline");
    }

    #[test]
    fn line_renders_dot_and_baseline() {
        let buf = render_to(GraphType::Line, vec![1.0, 5.0, 3.0, 7.0], 40, 10);
        assert!(contains(&buf, "•"), "no • in line chart");
        assert!(contains(&buf, "─"), "no baseline");
    }

    #[test]
    fn empty_values_shows_placeholder() {
        let area = Rect::new(0, 0, 60, 10);
        let mut buf = Buffer::empty(area);
        let vals = GraphValues::default();
        render(&GraphDef::default(), &vals, area, &mut buf);
        assert!(contains(&buf, "No graph ranges set"));
    }

    #[test]
    fn unimplemented_type_shows_placeholder() {
        let buf = render_to(GraphType::Pie, vec![1.0, 2.0, 3.0], 80, 10);
        assert!(contains(&buf, "Pie"));
        assert!(contains(&buf, "later slice"));
    }

    #[test]
    fn tiny_area_is_safe() {
        let area = Rect::new(0, 0, 4, 2);
        let mut buf = Buffer::empty(area);
        let mut vals = GraphValues::default();
        vals.data[0] = Some(vec![1.0, 2.0, 3.0]);
        render(
            &GraphDef { graph_type: GraphType::Bar, ..Default::default() },
            &vals,
            area,
            &mut buf,
        );
        // No panic = pass. Buffer may be empty.
    }

    #[test]
    fn bar_bar_height_monotonic_with_value() {
        let buf = render_to(GraphType::Bar, vec![1.0, 10.0], 20, 12);
        // Count `█` cells by column. The 10.0 bar should be taller
        // than the 1.0 bar.
        let mut heights = [0u16; 20];
        for y in 0..buf.area.height {
            for x in 0..buf.area.width {
                if buf[(x, y)].symbol() == "█" {
                    heights[x as usize] += 1;
                }
            }
        }
        let small = heights.iter().filter(|&&h| h > 0).copied().min().unwrap_or(0);
        let big = heights.iter().copied().max().unwrap_or(0);
        assert!(big > small, "bigger value should produce taller bar ({big} vs {small})");
    }
}
