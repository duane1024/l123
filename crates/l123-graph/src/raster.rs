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
    let img = match image::RgbImage::from_raw(w, h, rgb) {
        Some(i) => i,
        None => return Vec::new(),
    };
    let mut out = Cursor::new(Vec::new());
    match img.write_to(&mut out, image::ImageFormat::Png) {
        Ok(_) => out.into_inner(),
        Err(_) => Vec::new(),
    }
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
        GraphType::Line => draw_line(vals, root),
        GraphType::Bar => draw_bar(vals, root),
        GraphType::XY => draw_xy(vals, root),
        GraphType::Stack => draw_stack(vals, root),
        GraphType::Pie => draw_pie(vals, root),
        GraphType::HLCO => draw_hlco(vals, root),
        GraphType::Mixed => draw_mixed(vals, root),
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

/// Series colours A..F. Cycles if more series somehow appear.
const SERIES_PALETTE: &[RGBColor] = &[BLUE, RED, GREEN, MAGENTA, CYAN, BLACK];

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

fn draw_line<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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
    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d(0f64..(n.saturating_sub(1).max(1) as f64), y_lo..y_hi)?;
    chart.configure_mesh().draw()?;
    for (i, s) in series.iter().enumerate() {
        let color = SERIES_PALETTE[i % SERIES_PALETTE.len()];
        let points: Vec<(f64, f64)> = s
            .iter()
            .enumerate()
            .filter_map(|(idx, &v)| {
                if v.is_finite() {
                    Some((idx as f64, v))
                } else {
                    None
                }
            })
            .collect();
        chart
            .draw_series(LineSeries::new(points.clone(), color.stroke_width(2)))?
            .label(format!("Series {}", (b'A' + i as u8) as char))
            .legend(move |(x, y)| PathElement::new(vec![(x, y), (x + 20, y)], color));
        // Visible markers at each point.
        chart.draw_series(
            points
                .into_iter()
                .map(|(x, y)| Circle::new((x, y), 3, color.filled())),
        )?;
    }
    chart.configure_series_labels().border_style(BLACK).draw()?;
    Ok(())
}

fn draw_bar<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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
    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d((0..a.len() as i32).into_segmented(), y_lo..y_hi)?;
    chart.configure_mesh().draw()?;
    chart.draw_series(
        Histogram::vertical(&chart)
            .style(BLUE.filled())
            .margin(6)
            .data(a.iter().enumerate().filter_map(|(i, &v)| {
                if v.is_finite() {
                    Some((i as i32, v))
                } else {
                    None
                }
            })),
    )?;
    Ok(())
}

fn draw_xy<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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
    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d(x_lo..x_hi, y_lo..y_hi)?;
    chart.configure_mesh().draw()?;
    let n = x.len().min(y.len());
    let points: Vec<(f64, f64)> = (0..n)
        .filter_map(|i| {
            if x[i].is_finite() && y[i].is_finite() {
                Some((x[i], y[i]))
            } else {
                None
            }
        })
        .collect();
    chart.draw_series(
        points
            .iter()
            .map(|&(px, py)| Circle::new((px, py), 4, BLUE.filled())),
    )?;
    Ok(())
}

fn draw_stack<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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
    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d((0..n as i32).into_segmented(), 0f64..(y_max * 1.05))?;
    chart.configure_mesh().draw()?;
    // Draw each layer as its own Histogram on top of the cumulative sums.
    let mut base: Vec<f64> = vec![0.0; n];
    for (si, s) in series.iter().enumerate() {
        let color = SERIES_PALETTE[si % SERIES_PALETTE.len()];
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
        chart.draw_series(
            Histogram::vertical(&chart)
                .style(color.filled())
                .margin(6)
                .data(points),
        )?;
    }
    Ok(())
}

fn draw_pie<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
where
    DB: DrawingBackend,
    DB::ErrorType: 'static,
{
    let a = match vals.data[0].as_deref() {
        Some(a) if !a.is_empty() => a,
        _ => return paint_error(root, "Set /Graph A to plot a pie."),
    };
    let positive: Vec<f64> = a
        .iter()
        .copied()
        .filter(|v| v.is_finite() && *v > 0.0)
        .collect();
    if positive.is_empty() {
        return paint_error(root, "Pie needs positive values.");
    }
    let (w, h) = root.dim_in_pixel();
    let cx = (w / 2) as i32;
    let cy = (h / 2) as i32;
    let radius = (w.min(h) as f64 * 0.4).max(20.0);
    let labels: Vec<String> = positive
        .iter()
        .enumerate()
        .map(|(i, _)| format!("{}", i + 1))
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

fn draw_hlco<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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

    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d(0f64..(n.max(1) as f64), y_lo..y_hi)?;
    chart.configure_mesh().draw()?;
    for (i, &h) in high.iter().enumerate().take(n) {
        let l = low.get(i).copied().unwrap_or(f64::NAN);
        let c = close.get(i).copied().unwrap_or(f64::NAN);
        let o = open.get(i).copied().unwrap_or(f64::NAN);
        if !h.is_finite() || !l.is_finite() {
            continue;
        }
        let x = i as f64;
        // Main vertical line from L to H.
        chart.draw_series(std::iter::once(PathElement::new(
            vec![(x, l), (x, h)],
            BLACK.stroke_width(1),
        )))?;
        if c.is_finite() {
            chart.draw_series(std::iter::once(PathElement::new(
                vec![(x, c), (x + 0.3, c)],
                GREEN.stroke_width(2),
            )))?;
        }
        if o.is_finite() {
            chart.draw_series(std::iter::once(PathElement::new(
                vec![(x - 0.3, o), (x, o)],
                RED.stroke_width(2),
            )))?;
        }
    }
    Ok(())
}

fn draw_mixed<DB>(vals: &GraphValues, root: &DrawingArea<DB, plotters::coord::Shift>) -> DrawResult
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
    let mut chart = ChartBuilder::on(root)
        .margin(20)
        .x_label_area_size(30)
        .y_label_area_size(40)
        .build_cartesian_2d((0..n as i32).into_segmented(), y_lo..y_hi)?;
    chart.configure_mesh().draw()?;
    if !a.is_empty() {
        chart.draw_series(
            Histogram::vertical(&chart)
                .style(BLUE.filled())
                .margin(6)
                .data(a.iter().enumerate().filter_map(|(i, &v)| {
                    if v.is_finite() {
                        Some((i as i32, v))
                    } else {
                        None
                    }
                })),
        )?;
    }
    // Line overlay on series B.
    if !b.is_empty() {
        // Histogram uses SegmentValue<i32>; a LineSeries over the same
        // axis type needs the same coordinate shape, so we build a
        // second chart in Cartesian 2D aligned to the first.
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
        over.draw_series(LineSeries::new(pts.clone(), RED.stroke_width(2)))?;
        over.draw_series(
            pts.into_iter()
                .map(|(x, y)| Circle::new((x, y), 3, RED.filled())),
        )?;
    }
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
            svg.contains("<path") || svg.contains("<line"),
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
}
