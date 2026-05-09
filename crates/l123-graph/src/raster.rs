//! Raster rendering: SVG (for `/Graph Save`) and PNG (for ratatui-image
//! in slice 6). Uses plotters for the heavy lifting.
//!
//! All seven 1-2-3 graph types are supported. Semantics roughly follow
//! Lotus conventions:
//!
//! - **Line** — X is the category axis (falls back to 1-based indices),
//!   each populated A..F slot is a line series.
//! - **Bar** — series A is bars at each X sample.
//! - **XY** — X is the independent axis, A is the dependent.
//! - **Stack** — A..F stacked bars.
//! - **Pie** — A slices summed.
//! - **HLCO** — A=High, B=Low, C=Close, D=Open (by 1-2-3 convention).
//! - **Mixed** — series A as bars with B as a line overlay.
//!
//! The actual prettiness is basic on purpose; this slice's deliverable
//! is "all seven types produce a valid SVG / PNG without panicking and
//! with recognizably-the-right visual structure."

use std::io::Cursor;

use plotters::prelude::*;

use crate::{GraphDef, GraphType, GraphValues};

pub const DEFAULT_WIDTH: u32 = 800;
pub const DEFAULT_HEIGHT: u32 = 600;

type DrawResult = std::result::Result<(), Box<dyn std::error::Error + Send + Sync>>;

/// Render the graph to an SVG document. The `/Graph Save` command
/// uses this output directly.
pub fn render_svg(def: &GraphDef, vals: &GraphValues) -> String {
    render_svg_sized(def, vals, DEFAULT_WIDTH, DEFAULT_HEIGHT)
}

pub fn render_svg_sized(def: &GraphDef, vals: &GraphValues, w: u32, h: u32) -> String {
    let mut out = String::new();
    {
        let backend = SVGBackend::with_string(&mut out, (w, h));
        let root = backend.into_drawing_area();
        // Errors here can only come from malformed plot requests. We
        // catch them and paint an explanatory message rather than bail.
        if let Err(e) = draw_graph(def, vals, &root) {
            let _ = paint_error(&root, &format!("{e}"));
        }
        let _ = root.present();
    }
    out
}

/// Render the graph to a PNG byte buffer. Slice 6's ratatui-image
/// path feeds these bytes straight to the Picker.
pub fn render_png(def: &GraphDef, vals: &GraphValues) -> Vec<u8> {
    render_png_sized(def, vals, DEFAULT_WIDTH, DEFAULT_HEIGHT)
}

pub fn render_png_sized(def: &GraphDef, vals: &GraphValues, w: u32, h: u32) -> Vec<u8> {
    let img = match render_dynamic_image_sized(def, vals, w, h) {
        Some(i) => i,
        None => return Vec::new(),
    };
    let mut out = Cursor::new(Vec::new());
    match img.write_to(&mut out, image::ImageFormat::Png) {
        Ok(_) => out.into_inner(),
        Err(_) => Vec::new(),
    }
}

/// Render the graph straight to a [`image::DynamicImage`], skipping
/// the PNG encode/decode round-trip. The graph view feeds this
/// directly to ratatui-image's Picker so the image is sized to the
/// terminal at render time. Returns `None` only when `w`/`h` is zero
/// or `RgbImage::from_raw` rejects the buffer (the buffer is sized
/// to match `(w, h)`, so this is a degenerate-input guard).
pub fn render_dynamic_image_sized(
    def: &GraphDef,
    vals: &GraphValues,
    w: u32,
    h: u32,
) -> Option<image::DynamicImage> {
    if w == 0 || h == 0 {
        return None;
    }
    let pixels = (w as usize) * (h as usize) * 3;
    let mut rgb = vec![0u8; pixels];
    {
        let backend = BitMapBackend::with_buffer(&mut rgb, (w, h));
        let root = backend.into_drawing_area();
        if let Err(e) = draw_graph(def, vals, &root) {
            let _ = paint_error(&root, &format!("{e}"));
        }
        let _ = root.present();
    }
    image::RgbImage::from_raw(w, h, rgb).map(image::DynamicImage::ImageRgb8)
}

fn draw_graph<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    root.fill(&WHITE)?;
    if vals.is_empty() {
        paint_error(root, "Define /Graph ranges first.")?;
        return Ok(());
    }
    match def.graph_type {
        GraphType::Line => draw_line(def, vals, root),
        GraphType::Bar => draw_bar(def, vals, root),
        GraphType::XY => draw_xy(def, vals, root),
        GraphType::Stack => draw_stack(def, vals, root),
        GraphType::Pie => draw_pie(def, vals, root),
        GraphType::HLCO => draw_hlco(def, vals, root),
        GraphType::Mixed => draw_mixed(def, vals, root),
    }
}

/// Top caption for the chart, derived from First / Second titles.
/// Pie charts don't carry x/y descriptions, but they do carry
/// captions; everything else uses both.
fn caption_string(def: &GraphDef) -> Option<String> {
    let t = &def.options.titles;
    match (t.first.as_deref(), t.second.as_deref()) {
        (Some(a), Some(b)) => Some(format!("{a} — {b}")),
        (Some(a), None) => Some(a.to_owned()),
        (None, Some(b)) => Some(b.to_owned()),
        (None, None) => None,
    }
}

fn paint_error<DB>(root: &DrawingArea<DB, plotters::coord::Shift>, msg: &str) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let dim = root.dim_in_pixel();
    root.draw_text(
        msg,
        &TextStyle::from(("sans-serif", 20).into_font()).color(&RED),
        (10, (dim.1 / 2) as i32),
    )?;
    Ok(())
}

/// Series colours A..F. Eyeball-matched to the 1-2-3 R3.4a graph
/// view in DOS VGA mode 12h: A=red, B=green, C=blue, D=yellow,
/// E=magenta, F=cyan. Cycles if more series somehow appear.
const SERIES_PALETTE: &[RGBColor] = &[RED, GREEN, BLUE, YELLOW, MAGENTA, CYAN];

fn axis_bounds(series: &[&[f64]]) -> (f64, f64) {
    let mut lo = f64::INFINITY;
    let mut hi = f64::NEG_INFINITY;
    for s in series {
        for &v in *s {
            if !v.is_finite() {
                continue;
            }
            if v < lo {
                lo = v;
            }
            if v > hi {
                hi = v;
            }
        }
    }
    if !lo.is_finite() || !hi.is_finite() {
        return (0.0, 1.0);
    }
    if lo == hi {
        if lo == 0.0 {
            (0.0, 1.0)
        } else {
            (lo.min(0.0), hi.max(lo.abs()))
        }
    } else {
        let pad = (hi - lo).abs() * 0.05;
        (lo - pad, hi + pad)
    }
}

fn series_len_max(series: &[&[f64]]) -> usize {
    series.iter().map(|s| s.len()).max().unwrap_or(0)
}

/// Collect every populated A..F into a Vec of slices, in slot order.
/// Used by the multi-series types (Line, Stack, Mixed).
fn collected_data(vals: &GraphValues) -> Vec<&[f64]> {
    vals.data.iter().filter_map(|o| o.as_deref()).collect()
}

fn draw_line<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let series = collected_data(vals);
    let n = series_len_max(&series);
    if n == 0 {
        return paint_error(root, "No A..F data.");
    }
    let (y_lo, y_hi) = axis_bounds(&series);
    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    let mut chart =
        builder.build_cartesian_2d(0f64..(n.saturating_sub(1).max(1) as f64), y_lo..y_hi)?;
    let mut mesh = chart.configure_mesh();
    // Grid is off-by-default in 1-2-3 (Reference p. 2-200). plotters
    // draws a full mesh by default, so we explicitly disable each
    // direction whose flag the user hasn't set.
    if !def.options.grid.vertical {
        mesh.disable_x_mesh();
    }
    if !def.options.grid.horizontal {
        mesh.disable_y_mesh();
    }
    if let Some(t) = def.options.titles.x_axis.as_deref() {
        mesh.x_desc(t);
    }
    if let Some(t) = def.options.titles.y_axis.as_deref() {
        mesh.y_desc(t);
    }
    mesh.draw()?;
    // Pair each populated series with its A..F slot so data labels
    // and the legend text both look up by slot, not by populated
    // position.
    let series_with_slot: Vec<(usize, &[f64])> = vals
        .data
        .iter()
        .enumerate()
        .filter_map(|(slot, o)| o.as_deref().map(|s| (slot, s)))
        .collect();
    for (palette_i, (slot, s)) in series_with_slot.iter().enumerate() {
        let color = SERIES_PALETTE[palette_i % SERIES_PALETTE.len()];
        let points: Vec<(usize, f64, f64)> = s
            .iter()
            .enumerate()
            .filter_map(|(idx, &v)| {
                if v.is_finite() {
                    Some((idx, idx as f64, v))
                } else {
                    None
                }
            })
            .collect();
        let xy_points: Vec<(f64, f64)> =
            points.iter().map(|(_, x, y)| (*x, *y)).collect();
        let label = def
            .options
            .legend
            .get(*slot)
            .and_then(|opt| opt.clone())
            .unwrap_or_else(|| format!("Series {}", (b'A' + *slot as u8) as char));
        chart
            .draw_series(LineSeries::new(xy_points.clone(), color.stroke_width(2)))?
            .label(label)
            .legend(move |(x, y)| PathElement::new(vec![(x, y), (x + 20, y)], color));
        // Visible markers at each point.
        chart.draw_series(
            xy_points
                .iter()
                .map(|(x, y)| Circle::new((*x, *y), 3, color.filled())),
        )?;
        // Per-point data labels, when bound for this slot. Anchor
        // each label at the data point and let `placement_pos`
        // translate the placement enum into a TextStyle alignment
        // so the label sits above / below / to the side of the dot.
        if let Some(labels) = vals.data_label_text.get(*slot).and_then(|o| o.as_deref()) {
            let placement = def.options.data_labels_placement[*slot];
            let pos = placement_pos(placement);
            chart.draw_series(points.iter().filter_map(|(orig_i, x, y)| {
                labels
                    .get(*orig_i)
                    .filter(|s| !s.is_empty())
                    .map(|text| {
                        Text::new(
                            text.clone(),
                            (*x, *y),
                            ("sans-serif", 14)
                                .into_font()
                                .color(&BLACK)
                                .pos(pos),
                        )
                    })
            }))?;
        }
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

/// Translate a [`DataLabelPlacement`] into a plotters `Pos` so the
/// `Text` element anchored at a data point lands at the right
/// offset relative to the point. plotters positions text by where
/// its bounding box's anchor point falls, so e.g. `(Center, Bottom)`
/// puts the bottom-centre of the label at the data coord — making
/// the label float **above** the dot.
fn placement_pos(p: crate::DataLabelPlacement) -> plotters::style::text_anchor::Pos {
    use plotters::style::text_anchor::{HPos, Pos, VPos};
    use crate::DataLabelPlacement::*;
    match p {
        Above => Pos::new(HPos::Center, VPos::Bottom),
        Below => Pos::new(HPos::Center, VPos::Top),
        Left => Pos::new(HPos::Right, VPos::Center),
        Right => Pos::new(HPos::Left, VPos::Center),
        Center => Pos::new(HPos::Center, VPos::Center),
    }
}

fn draw_bar<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let a = match vals.data[0].as_deref() {
        Some(a) if !a.is_empty() => a,
        _ => return paint_error(root, "Set /Graph A to plot bars."),
    };
    let (y_lo, y_hi) = axis_bounds(&[a]);
    let y_lo = y_lo.min(0.0);
    let bar_color = SERIES_PALETTE[0];
    let label = def
        .options
        .legend
        .first()
        .and_then(|o| o.clone())
        .unwrap_or_else(|| "Series A".into());
    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    match def.features.orientation {
        crate::Orientation::Vertical => {
            let mut chart = builder
                .build_cartesian_2d((0..a.len() as i32).into_segmented(), y_lo..y_hi)?;
            let mut mesh = chart.configure_mesh();
            if !def.options.grid.vertical {
                mesh.disable_x_mesh();
            }
            if !def.options.grid.horizontal {
                mesh.disable_y_mesh();
            }
            if let Some(t) = def.options.titles.x_axis.as_deref() {
                mesh.x_desc(t);
            }
            if let Some(t) = def.options.titles.y_axis.as_deref() {
                mesh.y_desc(t);
            }
            mesh.draw()?;
            chart
                .draw_series(
                    Histogram::vertical(&chart)
                        .style(bar_color.filled())
                        .margin(6)
                        .data(a.iter().enumerate().filter_map(|(i, &v)| {
                            if v.is_finite() {
                                Some((i as i32, v))
                            } else {
                                None
                            }
                        })),
                )?
                .label(label)
                .legend(move |(x, y)| {
                    Rectangle::new([(x, y - 5), (x + 12, y + 5)], bar_color.filled())
                });
            // Per-bar data labels, when bound for slot 0. Anchor at
            // the bar's top: (segment center, v).
            if let Some(labels) =
                vals.data_label_text[0].as_deref()
            {
                let placement = def.options.data_labels_placement[0];
                let pos = placement_pos(placement);
                chart.draw_series(a.iter().enumerate().filter_map(|(i, &v)| {
                    if !v.is_finite() {
                        return None;
                    }
                    let text = labels.get(i)?.clone();
                    if text.is_empty() {
                        return None;
                    }
                    Some(Text::new(
                        text,
                        (SegmentValue::CenterOf(i as i32), v),
                        ("sans-serif", 14).into_font().color(&BLACK).pos(pos),
                    ))
                }))?;
            }
            chart.configure_series_labels().border_style(BLACK).draw()?;
        }
        crate::Orientation::Horizontal => {
            // Horizontal bars: value axis becomes the X axis, the
            // segmented category axis becomes Y. Histogram::horizontal
            // emits one rectangle per (value, segment) pair, growing
            // rightward from the y-axis baseline. Grid axes swap
            // meaning along with the orientation.
            let mut chart = builder
                .build_cartesian_2d(y_lo..y_hi, (0..a.len() as i32).into_segmented())?;
            let mut mesh = chart.configure_mesh();
            // Vertical grid lines (in 1-2-3 terms) → x-axis grid in
            // plotters terms when the value axis is X.
            if !def.options.grid.vertical {
                mesh.disable_x_mesh();
            }
            if !def.options.grid.horizontal {
                mesh.disable_y_mesh();
            }
            if let Some(t) = def.options.titles.x_axis.as_deref() {
                mesh.x_desc(t);
            }
            if let Some(t) = def.options.titles.y_axis.as_deref() {
                mesh.y_desc(t);
            }
            mesh.draw()?;
            chart
                .draw_series(
                    Histogram::horizontal(&chart)
                        .style(bar_color.filled())
                        .margin(6)
                        .data(a.iter().enumerate().filter_map(|(i, &v)| {
                            if v.is_finite() {
                                Some((i as i32, v))
                            } else {
                                None
                            }
                        })),
                )?
                .label(label)
                .legend(move |(x, y)| {
                    Rectangle::new([(x, y - 5), (x + 12, y + 5)], bar_color.filled())
                });
            // Per-bar data labels, when bound for slot 0. Anchor at
            // the bar's right end: (v, segment center).
            if let Some(labels) =
                vals.data_label_text[0].as_deref()
            {
                let placement = def.options.data_labels_placement[0];
                let pos = placement_pos(placement);
                chart.draw_series(a.iter().enumerate().filter_map(|(i, &v)| {
                    if !v.is_finite() {
                        return None;
                    }
                    let text = labels.get(i)?.clone();
                    if text.is_empty() {
                        return None;
                    }
                    Some(Text::new(
                        text,
                        (v, SegmentValue::CenterOf(i as i32)),
                        ("sans-serif", 14).into_font().color(&BLACK).pos(pos),
                    ))
                }))?;
            }
            chart.configure_series_labels().border_style(BLACK).draw()?;
        }
    }
    Ok(())
}

fn draw_xy<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let x = match vals.x.as_deref() {
        Some(x) if !x.is_empty() => x,
        _ => return paint_error(root, "XY graphs need an X range."),
    };
    let y = match vals.data[0].as_deref() {
        Some(y) if !y.is_empty() => y,
        _ => return paint_error(root, "XY graphs need series A."),
    };
    let (x_lo, x_hi) = axis_bounds(&[x]);
    let (y_lo, y_hi) = axis_bounds(&[y]);
    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    let mut chart = builder.build_cartesian_2d(x_lo..x_hi, y_lo..y_hi)?;
    let mut mesh = chart.configure_mesh();
    // Grid is off-by-default in 1-2-3 (Reference p. 2-200). plotters
    // draws a full mesh by default, so we explicitly disable each
    // direction whose flag the user hasn't set.
    if !def.options.grid.vertical {
        mesh.disable_x_mesh();
    }
    if !def.options.grid.horizontal {
        mesh.disable_y_mesh();
    }
    if let Some(t) = def.options.titles.x_axis.as_deref() {
        mesh.x_desc(t);
    }
    if let Some(t) = def.options.titles.y_axis.as_deref() {
        mesh.y_desc(t);
    }
    mesh.draw()?;
    let n = x.len().min(y.len());
    // Keep the original index so `data_label_text[0][orig_i]` lines
    // up after the non-finite filter.
    let points: Vec<(usize, f64, f64)> = (0..n)
        .filter_map(|i| {
            if x[i].is_finite() && y[i].is_finite() {
                Some((i, x[i], y[i]))
            } else {
                None
            }
        })
        .collect();
    let color = SERIES_PALETTE[0];
    let label = def
        .options
        .legend
        .first()
        .and_then(|o| o.clone())
        .unwrap_or_else(|| "Series A".into());
    chart
        .draw_series(
            points
                .iter()
                .map(|(_, px, py)| Circle::new((*px, *py), 4, color.filled())),
        )?
        .label(label)
        .legend(move |(x, y)| Circle::new((x + 6, y), 4, color.filled()));
    // Per-point data labels, when bound for slot 0.
    if let Some(labels) = vals.data_label_text[0].as_deref() {
        let placement = def.options.data_labels_placement[0];
        let pos = placement_pos(placement);
        chart.draw_series(points.iter().filter_map(|(orig_i, px, py)| {
            let text = labels.get(*orig_i)?.clone();
            if text.is_empty() {
                return None;
            }
            Some(Text::new(
                text,
                (*px, *py),
                ("sans-serif", 14).into_font().color(&BLACK).pos(pos),
            ))
        }))?;
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

fn draw_stack<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let series = collected_data(vals);
    if series.is_empty() {
        return paint_error(root, "No A..F data.");
    }
    let n = series_len_max(&series);
    // For each x index, stack sum.
    let mut stacked: Vec<f64> = vec![0.0; n];
    let mut y_max = 0f64;
    for s in &series {
        for (i, &v) in s.iter().enumerate() {
            if v.is_finite() && v > 0.0 && i < stacked.len() {
                stacked[i] += v;
            }
        }
    }
    for &v in &stacked {
        if v > y_max {
            y_max = v;
        }
    }
    if y_max == 0.0 {
        y_max = 1.0;
    }
    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    let mut chart =
        builder.build_cartesian_2d((0..n as i32).into_segmented(), 0f64..(y_max * 1.05))?;
    let mut mesh = chart.configure_mesh();
    // Grid is off-by-default in 1-2-3 (Reference p. 2-200). plotters
    // draws a full mesh by default, so we explicitly disable each
    // direction whose flag the user hasn't set.
    if !def.options.grid.vertical {
        mesh.disable_x_mesh();
    }
    if !def.options.grid.horizontal {
        mesh.disable_y_mesh();
    }
    if let Some(t) = def.options.titles.x_axis.as_deref() {
        mesh.x_desc(t);
    }
    if let Some(t) = def.options.titles.y_axis.as_deref() {
        mesh.y_desc(t);
    }
    mesh.draw()?;
    // Draw each layer as its own Histogram on top of the cumulative
    // sums. Pair every populated series with its original A..F slot
    // so the legend text and the stacking colour stay aligned.
    let series_with_slot: Vec<(usize, &[f64])> = vals
        .data
        .iter()
        .enumerate()
        .filter_map(|(slot, o)| o.as_deref().map(|s| (slot, s)))
        .collect();
    let mut base: Vec<f64> = vec![0.0; n];
    for (layer_i, (slot, s)) in series_with_slot.iter().enumerate() {
        let color = SERIES_PALETTE[layer_i % SERIES_PALETTE.len()];
        let points: Vec<(i32, f64)> = (0..n as i32)
            .filter_map(|i| {
                let idx = i as usize;
                let v = s.get(idx).copied().unwrap_or(f64::NAN);
                if !v.is_finite() || v <= 0.0 {
                    return None;
                }
                let top = base[idx] + v;
                base[idx] = top;
                Some((i, top))
            })
            .collect();
        let label = def
            .options
            .legend
            .get(*slot)
            .and_then(|opt| opt.clone())
            .unwrap_or_else(|| format!("Series {}", (b'A' + *slot as u8) as char));
        chart
            .draw_series(
                Histogram::vertical(&chart)
                    .style(color.filled())
                    .margin(6)
                    .data(points.clone()),
            )?
            .label(label)
            .legend(move |(x, y)| {
                Rectangle::new([(x, y - 5), (x + 12, y + 5)], color.filled())
            });
        // Per-segment data labels, anchored at the segment's top
        // (the running cumulative total after this layer).
        if let Some(labels) = vals.data_label_text[*slot].as_deref() {
            let placement = def.options.data_labels_placement[*slot];
            let pos = placement_pos(placement);
            chart.draw_series(points.iter().filter_map(|(i, top)| {
                let text = labels.get(*i as usize)?.clone();
                if text.is_empty() {
                    return None;
                }
                Some(Text::new(
                    text,
                    (SegmentValue::CenterOf(*i), *top),
                    ("sans-serif", 14).into_font().color(&BLACK).pos(pos),
                ))
            }))?;
        }
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

/// Wedge label for a pie slice. Prefers the cell-text string at
/// `orig_i` in `x_labels` when set and non-empty; otherwise falls
/// back to a 1-based positional index using the surviving wedge's
/// position so the labels stay sequential ("1", "2", "3").
fn wedge_label(x_labels: Option<&[String]>, orig_i: usize, wedge_i: usize) -> String {
    if let Some(labels) = x_labels {
        if let Some(s) = labels.get(orig_i) {
            if !s.trim().is_empty() {
                return s.clone();
            }
        }
    }
    format!("{}", wedge_i + 1)
}

fn draw_pie<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let a = match vals.data[0].as_deref() {
        Some(a) if !a.is_empty() => a,
        _ => return paint_error(root, "Set /Graph A to plot a pie."),
    };
    // Walk the A series alongside its index so each surviving wedge
    // can pull its label from the parallel x_labels slot. Filter must
    // happen after pairing so x_labels stays aligned even when some
    // A values are non-positive and get dropped.
    let pairs: Vec<(usize, f64)> = a
        .iter()
        .copied()
        .enumerate()
        .filter(|(_, v)| v.is_finite() && *v > 0.0)
        .collect();
    if pairs.is_empty() {
        return paint_error(root, "Pie needs positive values.");
    }
    // Pie charts have no Cartesian axes, so x_desc/y_desc don't apply;
    // the caption is drawn directly via root.draw_text instead.
    let (w, h) = root.dim_in_pixel();
    if let Some(caption) = caption_string(def) {
        root.draw_text(
            &caption,
            &TextStyle::from(("sans-serif", 24).into_font()).color(&BLACK),
            (10, 10),
        )?;
    }
    let cx = (w / 2) as i32;
    let cy = (h / 2) as i32;
    let radius = (w.min(h) as f64 * 0.4).max(20.0);
    let positive: Vec<f64> = pairs.iter().map(|(_, v)| *v).collect();
    let labels: Vec<String> = pairs
        .iter()
        .enumerate()
        .map(|(wedge_i, (orig_i, _))| wedge_label(vals.x_labels.as_deref(), *orig_i, wedge_i))
        .collect();
    let colors: Vec<RGBColor> = positive
        .iter()
        .enumerate()
        .map(|(i, _)| SERIES_PALETTE[i % SERIES_PALETTE.len()])
        .collect();
    let center = (cx, cy);
    let mut pie = Pie::new(&center, &radius, &positive, &colors, &labels);
    pie.label_style(("sans-serif", 18).into_font().color(&BLACK));
    pie.percentages(("sans-serif", 14).into_font().color(&WHITE));
    root.draw(&pie)?;
    Ok(())
}

fn draw_hlco<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let high = match vals.data[0].as_deref() {
        Some(a) if !a.is_empty() => a,
        _ => return paint_error(root, "HLCO needs series A (high)."),
    };
    let low = vals.data[1].as_deref().unwrap_or(&[]);
    let close = vals.data[2].as_deref().unwrap_or(&[]);
    let open = vals.data[3].as_deref().unwrap_or(&[]);
    let n = high.len();

    let mut y_lo = f64::INFINITY;
    let mut y_hi = f64::NEG_INFINITY;
    for s in [high, low, close, open] {
        for &v in s {
            if !v.is_finite() {
                continue;
            }
            if v < y_lo {
                y_lo = v;
            }
            if v > y_hi {
                y_hi = v;
            }
        }
    }
    if !y_lo.is_finite() || !y_hi.is_finite() {
        return paint_error(root, "HLCO has no numeric values.");
    }
    if y_lo == y_hi {
        y_hi = y_lo + 1.0;
    }

    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    let mut chart = builder.build_cartesian_2d(0f64..(n.max(1) as f64), y_lo..y_hi)?;
    let mut mesh = chart.configure_mesh();
    // Grid is off-by-default in 1-2-3 (Reference p. 2-200). plotters
    // draws a full mesh by default, so we explicitly disable each
    // direction whose flag the user hasn't set.
    if !def.options.grid.vertical {
        mesh.disable_x_mesh();
    }
    if !def.options.grid.horizontal {
        mesh.disable_y_mesh();
    }
    if let Some(t) = def.options.titles.x_axis.as_deref() {
        mesh.x_desc(t);
    }
    if let Some(t) = def.options.titles.y_axis.as_deref() {
        mesh.y_desc(t);
    }
    mesh.draw()?;
    // Aggregate per-x line segments into three single-series
    // pipelines so each contributes one legend entry. Legend names
    // come from the matching A/B/C/D slot with sensible fallbacks.
    let mut bar_segs: Vec<PathElement<(f64, f64)>> = Vec::new();
    let mut close_segs: Vec<PathElement<(f64, f64)>> = Vec::new();
    let mut open_segs: Vec<PathElement<(f64, f64)>> = Vec::new();
    for (i, &h) in high.iter().enumerate().take(n) {
        let l = low.get(i).copied().unwrap_or(f64::NAN);
        let c = close.get(i).copied().unwrap_or(f64::NAN);
        let o = open.get(i).copied().unwrap_or(f64::NAN);
        if !h.is_finite() || !l.is_finite() {
            continue;
        }
        let x = i as f64;
        bar_segs.push(PathElement::new(
            vec![(x, l), (x, h)],
            BLACK.stroke_width(1),
        ));
        if c.is_finite() {
            close_segs.push(PathElement::new(
                vec![(x, c), (x + 0.3, c)],
                GREEN.stroke_width(2),
            ));
        }
        if o.is_finite() {
            open_segs.push(PathElement::new(
                vec![(x - 0.3, o), (x, o)],
                RED.stroke_width(2),
            ));
        }
    }
    let high_label = def
        .options
        .legend
        .first()
        .and_then(|o| o.clone())
        .unwrap_or_else(|| "High-Low".into());
    chart
        .draw_series(bar_segs)?
        .label(high_label)
        .legend(|(x, y)| PathElement::new(vec![(x + 6, y - 5), (x + 6, y + 5)], BLACK.stroke_width(2)));
    if !close_segs.is_empty() {
        let close_label = def
            .options
            .legend
            .get(2)
            .and_then(|o| o.clone())
            .unwrap_or_else(|| "Close".into());
        chart
            .draw_series(close_segs)?
            .label(close_label)
            .legend(|(x, y)| PathElement::new(vec![(x, y), (x + 12, y)], GREEN.stroke_width(2)));
    }
    if !open_segs.is_empty() {
        let open_label = def
            .options
            .legend
            .get(3)
            .and_then(|o| o.clone())
            .unwrap_or_else(|| "Open".into());
        chart
            .draw_series(open_segs)?
            .label(open_label)
            .legend(|(x, y)| PathElement::new(vec![(x, y), (x + 12, y)], RED.stroke_width(2)));
    }
    // Per-bar data labels, anchored at the bar's high point. Slot 0
    // (the High series) is HLCO's primary anchor for labels — same
    // convention as the unicode F10 view.
    if let Some(labels) = vals.data_label_text[0].as_deref() {
        let placement = def.options.data_labels_placement[0];
        let pos = placement_pos(placement);
        chart.draw_series(high.iter().enumerate().take(n).filter_map(|(i, &h)| {
            let l = low.get(i).copied().unwrap_or(f64::NAN);
            if !h.is_finite() || !l.is_finite() {
                return None;
            }
            let text = labels.get(i)?.clone();
            if text.is_empty() {
                return None;
            }
            Some(Text::new(
                text,
                (i as f64, h.max(l)),
                ("sans-serif", 14).into_font().color(&BLACK).pos(pos),
            ))
        }))?;
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

fn draw_mixed<DB>(
    def: &GraphDef,
    vals: &GraphValues,
    root: &DrawingArea<DB, plotters::coord::Shift>,
) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let a = vals.data[0].as_deref().unwrap_or(&[]);
    let b = vals.data[1].as_deref().unwrap_or(&[]);
    if a.is_empty() && b.is_empty() {
        return paint_error(root, "Mixed needs at least one of A or B.");
    }
    let n = a.len().max(b.len());
    let (y_lo, y_hi) = axis_bounds(&[a, b]);
    let y_lo = y_lo.min(0.0);
    let mut builder = ChartBuilder::on(root);
    builder
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40);
    if let Some(c) = caption_string(def) {
        builder.caption(c, ("sans-serif", 24));
    }
    let mut chart =
        builder.build_cartesian_2d((0..n as i32).into_segmented(), y_lo..y_hi)?;
    let mut mesh = chart.configure_mesh();
    // Grid is off-by-default in 1-2-3 (Reference p. 2-200). plotters
    // draws a full mesh by default, so we explicitly disable each
    // direction whose flag the user hasn't set.
    if !def.options.grid.vertical {
        mesh.disable_x_mesh();
    }
    if !def.options.grid.horizontal {
        mesh.disable_y_mesh();
    }
    if let Some(t) = def.options.titles.x_axis.as_deref() {
        mesh.x_desc(t);
    }
    if let Some(t) = def.options.titles.y_axis.as_deref() {
        mesh.y_desc(t);
    }
    mesh.draw()?;
    let bar_color = SERIES_PALETTE[0];
    let line_color = SERIES_PALETTE[1];
    if !a.is_empty() {
        let bar_label = def
            .options
            .legend
            .first()
            .and_then(|o| o.clone())
            .unwrap_or_else(|| "Series A".into());
        chart
            .draw_series(
                Histogram::vertical(&chart)
                    .style(bar_color.filled())
                    .margin(6)
                    .data(a.iter().enumerate().filter_map(|(i, &v)| {
                        if v.is_finite() {
                            Some((i as i32, v))
                        } else {
                            None
                        }
                    })),
            )?
            .label(bar_label)
            .legend(move |(x, y)| {
                Rectangle::new([(x, y - 5), (x + 12, y + 5)], bar_color.filled())
            });
    }
    // Line overlay on series B.
    if !b.is_empty() {
        // Register the line legend on the bar chart via a phantom
        // empty histogram so both rows render in the same legend
        // box. The visual line itself draws on a second chart with
        // Cartesian f64 coords, since Histogram uses SegmentValue.
        let line_label = def
            .options
            .legend
            .get(1)
            .and_then(|o| o.clone())
            .unwrap_or_else(|| "Series B".into());
        chart
            .draw_series(
                Histogram::vertical(&chart)
                    .style(line_color.filled())
                    .data(std::iter::empty::<(i32, f64)>()),
            )?
            .label(line_label)
            .legend(move |(x, y)| {
                PathElement::new(vec![(x, y), (x + 12, y)], line_color.stroke_width(2))
            });
        let mut over = ChartBuilder::on(root)
            .margin(20)
            .x_label_area_size(30)
            .y_label_area_size(40)
            .build_cartesian_2d(-0.5f64..(n as f64 - 0.5), y_lo..y_hi)?;
        over.plotting_area().strip_coord_spec();
        let pts: Vec<(f64, f64)> = b
            .iter()
            .enumerate()
            .filter_map(|(i, &v)| {
                if v.is_finite() {
                    Some((i as f64, v))
                } else {
                    None
                }
            })
            .collect();
        over.draw_series(LineSeries::new(pts.clone(), line_color.stroke_width(2)))?;
        over.draw_series(
            pts.into_iter()
                .map(|(x, y)| Circle::new((x, y), 3, line_color.filled())),
        )?;
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_vals(types: &[(usize, Vec<f64>)]) -> GraphValues {
        // `types` uses slot indices: 0..=5 for A..F, 6 for X.
        let mut v = GraphValues::default();
        for (slot, data) in types {
            if *slot == 6 {
                v.x = Some(data.clone());
            } else {
                v.data[*slot] = Some(data.clone());
            }
        }
        v
    }

    fn a(values: Vec<f64>) -> GraphValues {
        make_vals(&[(0, values)])
    }

    #[test]
    fn svg_line_has_svg_envelope_and_polyline() {
        let def = GraphDef {
            graph_type: GraphType::Line,
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 5.0, 3.0, 4.0]));
        assert!(svg.contains("<svg"), "no <svg root: {svg:.120}");
        assert!(svg.contains("</svg>"), "no </svg> close");
        // plotters' LineSeries emits a <polyline> or repeated <path> in SVG.
        assert!(
            svg.contains("polyline") || svg.contains("<path"),
            "no line shape in SVG"
        );
    }

    #[test]
    fn svg_bar_has_rect_elements() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 3.0]));
        assert!(svg.contains("<rect"), "no <rect in bar SVG");
    }

    #[test]
    fn svg_xy_draws_when_x_and_a_are_set() {
        let vals = make_vals(&[(6, vec![1.0, 2.0, 3.0]), (0, vec![10.0, 20.0, 15.0])]);
        let def = GraphDef {
            graph_type: GraphType::XY,
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("<svg"));
        assert!(
            svg.contains("<circle") || svg.contains("fill"),
            "no data marks in XY"
        );
    }

    #[test]
    fn svg_stack_draws_multiple_series() {
        let vals = make_vals(&[
            (0, vec![1.0, 2.0, 3.0]),
            (1, vec![2.0, 1.0, 4.0]),
            (2, vec![0.5, 0.5, 0.5]),
        ]);
        let def = GraphDef {
            graph_type: GraphType::Stack,
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("<rect"), "stack should emit rects");
    }

    #[test]
    fn svg_pie_draws_slices() {
        let def = GraphDef {
            graph_type: GraphType::Pie,
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![30.0, 20.0, 50.0]));
        // plotters' Pie emits one <polygon> per slice.
        let slices = svg.matches("<polygon").count();
        assert!(slices >= 3, "pie should draw 3+ polygons, got {slices}");
    }

    #[test]
    fn svg_pie_slice_colors_match_lotus_palette() {
        // 1-2-3 R3.4a assigned slice colors from a fixed palette in
        // series-A order. Eyeball-matched from the reference
        // screenshot in docs/SPEC: A=red, B=green, C=blue, D=yellow,
        // E=magenta, F=cyan. Plotters emits each slice's fill via the
        // `<polygon ... fill="#RRGGBB"` SVG attribute.
        let def = GraphDef {
            graph_type: GraphType::Pie,
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![10.0, 20.0, 30.0, 40.0, 50.0, 60.0]));
        let expected = [
            "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        ];
        for color in expected {
            let needle = format!("fill=\"{color}\"");
            assert!(
                svg.contains(&needle),
                "pie SVG should contain a {color} slice; got: {}",
                &svg[..svg.len().min(400)]
            );
        }
    }

    #[test]
    fn svg_pie_uses_x_labels_for_wedges() {
        let def = GraphDef {
            graph_type: GraphType::Pie,
            ..Default::default()
        };
        let mut vals = a(vec![30.0, 20.0, 50.0]);
        vals.x_labels = Some(vec!["Apples".into(), "Pears".into(), "Plums".into()]);
        let svg = render_svg(&def, &vals);
        for label in ["Apples", "Pears", "Plums"] {
            assert!(
                svg.contains(label),
                "pie should label wedges with x_labels, missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_hlco_draws_vertical_bars() {
        let vals = make_vals(&[
            (0, vec![10.0, 11.0, 12.0]), // high
            (1, vec![8.0, 9.0, 10.0]),   // low
            (2, vec![9.0, 10.5, 11.0]),  // close
            (3, vec![8.5, 10.0, 10.5]),  // open
        ]);
        let def = GraphDef {
            graph_type: GraphType::HLCO,
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("<svg"));
        assert!(
            svg.contains("<path") || svg.contains("<line") || svg.contains("<polyline"),
            "no strokes in HLCO"
        );
    }

    #[test]
    fn svg_mixed_has_both_bars_and_line() {
        let vals = make_vals(&[(0, vec![1.0, 2.0, 3.0]), (1, vec![3.0, 2.0, 1.0])]);
        let def = GraphDef {
            graph_type: GraphType::Mixed,
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("<rect"), "mixed should include bars");
        assert!(
            svg.contains("polyline") || svg.contains("<path"),
            "mixed should include a line"
        );
    }

    #[test]
    fn svg_empty_values_shows_placeholder() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let svg = render_svg(&def, &GraphValues::default());
        assert!(svg.contains("Define"), "placeholder text missing");
    }

    #[test]
    fn png_has_png_magic() {
        let def = GraphDef {
            graph_type: GraphType::Line,
            ..Default::default()
        };
        let bytes = render_png(&def, &a(vec![1.0, 2.0, 3.0]));
        assert!(bytes.len() > 8, "png bytes too short");
        assert_eq!(&bytes[..8], b"\x89PNG\r\n\x1a\n", "missing PNG magic");
    }

    #[test]
    fn png_is_decodable() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let bytes = render_png(&def, &a(vec![5.0, 10.0, 3.0, 7.0, 2.0]));
        let img = image::load_from_memory(&bytes).expect("valid PNG");
        assert_eq!(img.width(), DEFAULT_WIDTH);
        assert_eq!(img.height(), DEFAULT_HEIGHT);
    }

    #[test]
    fn dynamic_image_sized_returns_image_at_requested_dims() {
        let def = GraphDef {
            graph_type: GraphType::Pie,
            ..Default::default()
        };
        let img = render_dynamic_image_sized(&def, &a(vec![10.0, 20.0, 30.0]), 1280, 720)
            .expect("valid image");
        assert_eq!(img.width(), 1280);
        assert_eq!(img.height(), 720);
    }

    #[test]
    fn dynamic_image_sized_rejects_zero_dims() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        assert!(render_dynamic_image_sized(&def, &a(vec![1.0]), 0, 100).is_none());
        assert!(render_dynamic_image_sized(&def, &a(vec![1.0]), 100, 0).is_none());
    }

    #[test]
    fn all_seven_types_render_without_panic() {
        let vals = make_vals(&[
            (6, vec![1.0, 2.0, 3.0, 4.0]),
            (0, vec![10.0, 20.0, 30.0, 40.0]),
            (1, vec![5.0, 15.0, 25.0, 35.0]),
            (2, vec![8.0, 12.0, 28.0, 38.0]),
            (3, vec![6.0, 16.0, 22.0, 36.0]),
        ]);
        for t in [
            GraphType::Line,
            GraphType::Bar,
            GraphType::XY,
            GraphType::Stack,
            GraphType::Pie,
            GraphType::HLCO,
            GraphType::Mixed,
        ] {
            let def = GraphDef {
                graph_type: t,
                ..Default::default()
            };
            let svg = render_svg(&def, &vals);
            assert!(svg.contains("<svg"), "{t:?} produced no SVG");
            let png = render_png(&def, &vals);
            assert_eq!(&png[..8], b"\x89PNG\r\n\x1a\n", "{t:?} bad PNG magic");
        }
    }

    #[test]
    fn svg_caption_combines_first_and_second_titles() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            options: crate::GraphOptions {
                titles: crate::Titles {
                    first: Some("Sales 1991".into()),
                    second: Some("by Quarter".into()),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 3.0]));
        assert!(svg.contains("Sales 1991"), "missing first title in SVG");
        assert!(svg.contains("by Quarter"), "missing second title in SVG");
    }

    #[test]
    fn svg_uses_x_axis_and_y_axis_descriptions() {
        let def = GraphDef {
            graph_type: GraphType::Line,
            options: crate::GraphOptions {
                titles: crate::Titles {
                    x_axis: Some("Quarter".into()),
                    y_axis: Some("Dollars".into()),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 3.0, 4.0]));
        assert!(svg.contains("Quarter"), "missing x-axis description in SVG");
        assert!(svg.contains("Dollars"), "missing y-axis description in SVG");
    }

    #[test]
    fn svg_default_has_no_mesh_grid_in_plot_area() {
        // plotters draws a mesh by default; we suppress it when the
        // user has not set options.grid. The SVG still contains axis
        // ticks (short tick marks), but full-length grid lines should
        // be absent. We approximate by counting <line> path elements:
        // a no-grid SVG has the axis frame, axis ticks, and a
        // vertical/horizontal axis line, but no row of horizontal
        // mesh lines spanning the plot.
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 3.0]));
        let no_grid_lines = svg.match_indices("<line").count();
        let with_grid = GraphDef {
            graph_type: GraphType::Bar,
            options: crate::GraphOptions {
                grid: crate::GridMask {
                    horizontal: true,
                    vertical: true,
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };
        let svg2 = render_svg(&with_grid, &a(vec![1.0, 2.0, 3.0]));
        let with_grid_lines = svg2.match_indices("<line").count();
        assert!(
            with_grid_lines > no_grid_lines,
            "enabling grid should add SVG <line> elements (no_grid={no_grid_lines}, with_grid={with_grid_lines})"
        );
    }

    #[test]
    fn svg_mixed_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::Mixed,
            options: crate::GraphOptions {
                legend: [
                    Some("ABars".into()),
                    Some("BLine".into()),
                    None,
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let vals = make_vals(&[
            (0, vec![1.0, 2.0, 3.0]),
            (1, vec![3.0, 2.0, 1.0]),
        ]);
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("ABars"), "Mixed SVG missing A-bars legend");
        assert!(svg.contains("BLine"), "Mixed SVG missing B-line legend");
    }

    #[test]
    fn svg_hlco_emits_data_labels_when_bound() {
        let def = GraphDef {
            graph_type: GraphType::HLCO,
            ..Default::default()
        };
        let vals = GraphValues {
            data: [
                Some(vec![10.0, 11.0, 12.0]),
                Some(vec![8.0, 9.0, 10.0]),
                Some(vec![9.0, 10.5, 11.0]),
                Some(vec![8.5, 10.0, 10.5]),
                None,
                None,
            ],
            data_label_text: [
                Some(vec!["D1".into(), "D2".into(), "D3".into()]),
                None,
                None,
                None,
                None,
                None,
            ],
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        for label in ["D1", "D2", "D3"] {
            assert!(
                svg.contains(label),
                "raster HLCO SVG should embed bound data labels; missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_hlco_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::HLCO,
            options: crate::GraphOptions {
                legend: [
                    Some("HighA".into()),
                    None,
                    Some("CloseC".into()),
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let vals = make_vals(&[
            (0, vec![10.0, 11.0, 12.0]),
            (1, vec![8.0, 9.0, 10.0]),
            (2, vec![9.0, 10.5, 11.0]),
            (3, vec![8.5, 10.0, 10.5]),
        ]);
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("HighA"), "HLCO SVG missing High legend");
        assert!(svg.contains("CloseC"), "HLCO SVG missing Close legend");
    }

    #[test]
    fn svg_xy_emits_data_labels_when_bound() {
        let def = GraphDef {
            graph_type: GraphType::XY,
            ..Default::default()
        };
        let vals = GraphValues {
            x: Some(vec![1.0, 2.0, 3.0]),
            data: [
                Some(vec![10.0, 20.0, 30.0]),
                None,
                None,
                None,
                None,
                None,
            ],
            data_label_text: [
                Some(vec!["P1".into(), "P2".into(), "P3".into()]),
                None,
                None,
                None,
                None,
                None,
            ],
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        for label in ["P1", "P2", "P3"] {
            assert!(
                svg.contains(label),
                "raster XY SVG should embed bound data labels; missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_xy_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::XY,
            options: crate::GraphOptions {
                legend: [
                    Some("Sample".into()),
                    None,
                    None,
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let vals = make_vals(&[
            (6, vec![1.0, 2.0, 3.0]), // X
            (0, vec![10.0, 20.0, 30.0]), // A
        ]);
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("Sample"), "XY SVG missing user legend text");
        assert!(
            !svg.contains("Series A"),
            "fallback legend leaked through"
        );
    }

    #[test]
    fn svg_stack_emits_data_labels_when_bound() {
        let def = GraphDef {
            graph_type: GraphType::Stack,
            ..Default::default()
        };
        let vals = GraphValues {
            data: [
                Some(vec![1.0, 2.0, 3.0]),
                Some(vec![1.0, 1.0, 1.0]),
                None,
                None,
                None,
                None,
            ],
            data_label_text: [
                Some(vec!["A1".into(), "A2".into(), "A3".into()]),
                Some(vec!["B1".into(), "B2".into(), "B3".into()]),
                None,
                None,
                None,
                None,
            ],
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        for label in ["A1", "A2", "A3", "B1", "B2", "B3"] {
            assert!(
                svg.contains(label),
                "raster stack SVG should embed bound data labels; missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_stack_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::Stack,
            options: crate::GraphOptions {
                legend: [
                    Some("A1".into()),
                    Some("B1".into()),
                    None,
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let vals = make_vals(&[
            (0, vec![1.0, 2.0, 3.0]),
            (1, vec![1.0, 1.0, 1.0]),
        ]);
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("A1"), "stack SVG missing A1 legend");
        assert!(svg.contains("B1"), "stack SVG missing B1 legend");
        assert!(
            !svg.contains("Series A") && !svg.contains("Series B"),
            "fallback legend leaked through"
        );
    }

    #[test]
    fn svg_bar_emits_data_labels_when_bound() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let vals = GraphValues {
            data: [Some(vec![10.0, 20.0, 30.0]), None, None, None, None, None],
            data_label_text: [
                Some(vec!["Q1".into(), "Q2".into(), "Q3".into()]),
                None,
                None,
                None,
                None,
                None,
            ],
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        for label in ["Q1", "Q2", "Q3"] {
            assert!(
                svg.contains(label),
                "raster bar SVG should embed bound data labels; missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_bar_horizontal_differs_from_vertical() {
        // Same data, two orientations. The SVGs should render
        // distinct geometry: vertical bars extend up from the
        // x-axis, horizontal bars extend right from the y-axis.
        let vals = a(vec![1.0, 2.0, 3.0]);
        let v_def = GraphDef {
            graph_type: GraphType::Bar,
            ..Default::default()
        };
        let h_def = GraphDef {
            graph_type: GraphType::Bar,
            features: crate::GraphFeatures {
                orientation: crate::Orientation::Horizontal,
                ..Default::default()
            },
            ..Default::default()
        };
        let v_svg = render_svg(&v_def, &vals);
        let h_svg = render_svg(&h_def, &vals);

        // Pull every bar-coloured rect (#FF0000 fill from
        // SERIES_PALETTE[0]) and find the LARGEST one in each SVG.
        // Plotters emits the rect attrs as `x=".." y=".." width="N"
        // height="N"`. The largest such rect is the tallest bar in
        // vertical orientation, or the widest bar in horizontal —
        // its aspect ratio reveals which axis is the value axis.
        fn largest_bar_dims(svg: &str) -> Option<(u32, u32)> {
            let mut best: Option<(u32, u32, u32)> = None;
            for piece in svg.split("<rect ").skip(1) {
                if !piece.contains("fill=\"#FF0000\"") {
                    continue;
                }
                let pull = |name: &str| -> Option<u32> {
                    let needle = format!(" {name}=\"");
                    let i = piece.find(&needle)? + needle.len();
                    let rest = &piece[i..];
                    let end = rest.find('"')?;
                    rest[..end].parse().ok()
                };
                let w = pull("width")?;
                let h = pull("height")?;
                let area = w * h;
                if best.is_none_or(|(_, _, a)| area > a) {
                    best = Some((w, h, area));
                }
            }
            best.map(|(w, h, _)| (w, h))
        }

        let (v_w, v_h) =
            largest_bar_dims(&v_svg).expect("vertical SVG missing #FF0000 bar rect");
        let (h_w, h_h) =
            largest_bar_dims(&h_svg).expect("horizontal SVG missing #FF0000 bar rect");
        assert!(
            v_h > v_w * 2,
            "vertical bar should be much taller than wide; got {v_w}x{v_h}"
        );
        assert!(
            h_w > h_h * 2,
            "horizontal bar should be much wider than tall; got {h_w}x{h_h}"
        );
    }

    #[test]
    fn svg_bar_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::Bar,
            options: crate::GraphOptions {
                legend: [
                    Some("Q1 Sales".into()),
                    None,
                    None,
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![10.0, 20.0, 30.0]));
        assert!(
            svg.contains("Q1 Sales"),
            "bar SVG should embed the user-set legend text"
        );
    }

    #[test]
    fn svg_line_emits_data_labels_when_bound() {
        let def = GraphDef {
            graph_type: GraphType::Line,
            ..Default::default()
        };
        let vals = GraphValues {
            data: [
                Some(vec![10.0, 20.0, 30.0]),
                None,
                None,
                None,
                None,
                None,
            ],
            data_label_text: [
                Some(vec!["Q1".into(), "Q2".into(), "Q3".into()]),
                None,
                None,
                None,
                None,
                None,
            ],
            ..Default::default()
        };
        let svg = render_svg(&def, &vals);
        for label in ["Q1", "Q2", "Q3"] {
            assert!(
                svg.contains(label),
                "raster line SVG should embed bound data labels; missing {label:?}"
            );
        }
    }

    #[test]
    fn svg_line_uses_user_legend_text_when_set() {
        let def = GraphDef {
            graph_type: GraphType::Line,
            options: crate::GraphOptions {
                legend: [
                    Some("Net Sales".into()),
                    Some("YTD".into()),
                    None,
                    None,
                    None,
                    None,
                ],
                ..Default::default()
            },
            ..Default::default()
        };
        let vals = make_vals(&[
            (0, vec![1.0, 2.0, 3.0]),
            (1, vec![4.0, 5.0, 6.0]),
        ]);
        let svg = render_svg(&def, &vals);
        assert!(svg.contains("Net Sales"), "missing A legend in SVG");
        assert!(svg.contains("YTD"), "missing B legend in SVG");
        assert!(
            !svg.contains("Series A") && !svg.contains("Series B"),
            "fallback legend leaked through"
        );
    }

    #[test]
    fn svg_pie_caption_drawn_directly_when_titles_set() {
        // Pie chart has no Cartesian mesh; the caption is painted by
        // root.draw_text rather than ChartBuilder.caption.
        let def = GraphDef {
            graph_type: GraphType::Pie,
            options: crate::GraphOptions {
                titles: crate::Titles {
                    first: Some("Cost Breakdown".into()),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        };
        let svg = render_svg(&def, &a(vec![1.0, 2.0, 3.0]));
        assert!(
            svg.contains("Cost Breakdown"),
            "pie caption missing in SVG"
        );
    }
}
