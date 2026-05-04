//! Build a [`PageGrid`] from a [`WorkbookView`] and [`PrintSettings`].
//!
//! All layout decisions (column widths, label prefixes, numeric
//! alignment, overflow-to-asterisks, right-margin truncation, left
//! margin, header/footer token substitution + three-part formatting,
//! page chunking) happen here so encoders stay dumb.

use l123_core::{render_label, render_value_in_cell, Address, CellContents, LabelPrefix, Range};

use crate::grid::{Page, PageGrid};
use crate::settings::{PrintContentMode, PrintFormatMode, PrintSettings};
use crate::view::WorkbookView;

/// Render `range` into a paginated grid using `settings`.
pub fn render<V: WorkbookView + ?Sized>(
    view: &V,
    range: Range,
    settings: &PrintSettings,
) -> PageGrid {
    let r = range.normalized();

    // Column-break segmentation. `|::` in row r.start.row of any
    // column marks a manual column page break (`/Worksheet Page
    // Column` inserts that label); each such column closes the current
    // segment and opens a new one. Other pipe-prefix columns still
    // suppress output without breaking pages.
    let mut col_segments: Vec<Vec<u16>> = vec![Vec::new()];
    for col in r.start.col..=r.end.col {
        let first = Address::new(r.start.sheet, col, r.start.row);
        if let Some(CellContents::Label {
            prefix: LabelPrefix::Pipe,
            text,
        }) = view.cell(first)
        {
            if text.starts_with("::") {
                col_segments.push(Vec::new());
            }
            continue;
        }
        col_segments.last_mut().unwrap().push(col);
    }
    while col_segments.len() > 1 && col_segments.last().is_some_and(Vec::is_empty) {
        col_segments.pop();
    }
    let segment_widths: Vec<usize> = col_segments
        .iter()
        .map(|seg| {
            seg.iter()
                .map(|c| view.col_width(r.start.sheet, *c) as usize)
                .sum()
        })
        .collect();
    // Grid-level page_width: first non-empty segment, falling back to
    // the full range width when every segment is empty.
    let total_page_width: usize = (r.start.col..=r.end.col)
        .map(|c| view.col_width(r.start.sheet, c) as usize)
        .sum();
    let grid_page_width = segment_widths
        .iter()
        .copied()
        .find(|w| *w > 0)
        .unwrap_or(total_page_width)
        .max(1);

    let (header, footer) = match settings.format_mode {
        PrintFormatMode::Formatted => (settings.header.as_str(), settings.footer.as_str()),
        PrintFormatMode::Unformatted => ("", ""),
    };
    let content_mode = settings.content_mode;
    let left_pad: String = " ".repeat(settings.margin_left as usize);

    // Build pages segment-by-segment (column-major: all rows of the
    // leftmost segment first, then the next segment, etc.).
    let mut pages_with_width: Vec<(Vec<String>, usize)> = Vec::new();
    for (seg_idx, cols) in col_segments.iter().enumerate() {
        let seg_width = segment_widths[seg_idx].max(1);
        let effective_width = if settings.margin_right == 0 {
            seg_width
        } else {
            seg_width.saturating_sub(settings.margin_right as usize)
        };

        // Collect content rows into row-sections — `|::` in column
        // r.start.col marks a manual row page break (`/Worksheet Page
        // Row` inserts that label), closing the current section and
        // opening a new one. Other pipe-prefix rows still suppress
        // output without breaking pages.  Each entry has `left_pad`
        // prepended and a trailing `\n`.
        let mut sections: Vec<Vec<String>> = vec![Vec::new()];
        for row in r.start.row..=r.end.row {
            let first = Address::new(r.start.sheet, r.start.col, row);
            if let Some(CellContents::Label {
                prefix: LabelPrefix::Pipe,
                text,
            }) = view.cell(first)
            {
                if text.starts_with("::") {
                    sections.push(Vec::new());
                }
                continue;
            }
            let mut line = String::new();
            for &col in cols {
                let addr = Address::new(r.start.sheet, col, row);
                let w = view.col_width(r.start.sheet, col) as usize;
                let piece = match view.cell(addr) {
                    Some(CellContents::Empty) | None => " ".repeat(w),
                    Some(CellContents::Label { prefix, text }) => render_label(*prefix, text, w),
                    Some(CellContents::Constant(v)) => {
                        let fmt = view.format_for_cell(addr);
                        render_value_in_cell(v, w, fmt, view.international(), view.date_formats())
                            .unwrap_or_else(|| " ".repeat(w))
                    }
                    Some(CellContents::Formula { expr, cached_value }) => match content_mode {
                        PrintContentMode::CellFormulas => {
                            let src = format!("@{expr}");
                            let pad = w.saturating_sub(src.chars().count());
                            let mut s = src;
                            s.extend(std::iter::repeat_n(' ', pad));
                            s
                        }
                        PrintContentMode::AsDisplayed => match cached_value {
                            Some(v) => {
                                let fmt = view.format_for_cell(addr);
                                render_value_in_cell(v, w, fmt, view.international(), view.date_formats())
                                    .unwrap_or_else(|| " ".repeat(w))
                            }
                            None => " ".repeat(w),
                        },
                    },
                };
                line.push_str(&piece);
            }
            let truncated: String = line.chars().take(effective_width).collect();
            let trimmed: String = truncated.trim_end().to_string();
            let mut entry = String::with_capacity(left_pad.len() + trimmed.len() + 1);
            entry.push_str(&left_pad);
            entry.push_str(&trimmed);
            entry.push('\n');
            sections.last_mut().unwrap().push(entry);
        }
        // Trailing `|::` shouldn't yield a blank page on its own.
        while sections.len() > 1 && sections.last().is_some_and(Vec::is_empty) {
            sections.pop();
        }

        // Chunk each row-section into pages. pg_length == 0 means no
        // pagination within a section.  Sections still split into
        // separate pages, since a manual break is itself a page
        // boundary.
        let chunked: Vec<Vec<String>> = if sections.iter().all(Vec::is_empty) {
            vec![Vec::new()]
        } else {
            let mut out: Vec<Vec<String>> = Vec::new();
            for section in sections {
                if section.is_empty() {
                    out.push(Vec::new());
                    continue;
                }
                let per_page = if settings.pg_length == 0 {
                    section.len()
                } else {
                    settings.pg_length as usize
                };
                out.extend(section.chunks(per_page).map(<[String]>::to_vec));
            }
            out
        };

        for page_rows in chunked {
            pages_with_width.push((page_rows, seg_width));
        }
    }

    let today = today_ddmmmyy();
    let mut pages: Vec<Page> = Vec::with_capacity(pages_with_width.len());
    for (i, (page_rows, pw)) in pages_with_width.into_iter().enumerate() {
        let page_no = settings.start_page as usize + i;
        let header_line = if header.is_empty() {
            None
        } else {
            let substituted = substitute_tokens(header, page_no, &today);
            let mut line = String::with_capacity(left_pad.len() + pw);
            line.push_str(&left_pad);
            line.push_str(&format_three_part(&substituted, pw));
            Some(line)
        };
        let footer_line = if footer.is_empty() {
            None
        } else {
            let substituted = substitute_tokens(footer, page_no, &today);
            let mut line = String::with_capacity(left_pad.len() + pw);
            line.push_str(&left_pad);
            line.push_str(&format_three_part(&substituted, pw));
            Some(line)
        };
        pages.push(Page {
            number: page_no as u32,
            header: header_line,
            footer: footer_line,
            rows: page_rows,
            top_blank: settings.margin_top,
            bottom_blank: settings.margin_bottom,
        });
    }

    PageGrid {
        pages,
        page_width: grid_page_width as u16,
    }
}

/// Substitute 1-2-3 header/footer tokens:
///   `#` → `page_no` (as a decimal number)
///   `@` → `today` (pre-formatted DD-MMM-YY date string)
/// Other characters pass through untouched. `\name` (named-range
/// substitution) is a later milestone.
fn substitute_tokens(s: &str, page_no: usize, today: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '#' => out.push_str(&page_no.to_string()),
            '@' => out.push_str(today),
            other => out.push(other),
        }
    }
    out
}

/// Today's date formatted as `DD-MMM-YY` (1-2-3's default D1 date
/// format). Falls back to an empty string if the system clock is
/// somehow earlier than the Unix epoch.
fn today_ddmmmyy() -> String {
    use std::time::{SystemTime, UNIX_EPOCH};
    let secs = match SystemTime::now().duration_since(UNIX_EPOCH) {
        Ok(d) => d.as_secs() as i64,
        Err(_) => return String::new(),
    };
    let days = secs.div_euclid(86_400);
    let (y, m, d) = days_to_ymd(days);
    let month_name = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ][(m - 1) as usize];
    format!("{d:02}-{month_name}-{:02}", (y % 100 + 100) % 100)
}

/// Days since Unix epoch (1970-01-01) → (year, month [1..12],
/// day-of-month [1..31]). Hinnant's civil-from-days algorithm.
fn days_to_ymd(z: i64) -> (i32, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = (z - era * 146_097) as u32;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe as i64 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = y + if m <= 2 { 1 } else { 0 };
    (y as i32, m, d)
}

/// Format a `|`-split three-part header/footer line to `width`
/// characters. Parts past the third are ignored; missing parts are
/// treated as empty. Left-aligned | centered | right-aligned,
/// truncated to width if over.
fn format_three_part(s: &str, width: usize) -> String {
    let mut parts = s.splitn(3, '|');
    let left = parts.next().unwrap_or("");
    let center = parts.next().unwrap_or("");
    let right = parts.next().unwrap_or("");
    let lcount = left.chars().count();
    let ccount = center.chars().count();
    let rcount = right.chars().count();
    if lcount + ccount + rcount >= width {
        let joined = format!("{left}{center}{right}");
        return joined.chars().take(width).collect();
    }
    let c_start = (width.saturating_sub(ccount)) / 2;
    let c_end = c_start + ccount;
    let r_start = width - rcount;
    let mut out = String::with_capacity(width);
    out.push_str(left);
    let pad1 = c_start.saturating_sub(lcount);
    out.extend(std::iter::repeat_n(' ', pad1));
    out.push_str(center);
    let pad2 = r_start.saturating_sub(c_end);
    out.extend(std::iter::repeat_n(' ', pad2));
    out.push_str(right);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn days_to_ymd_known_dates() {
        assert_eq!(days_to_ymd(0), (1970, 1, 1));
        assert_eq!(days_to_ymd(31), (1970, 2, 1));
        assert_eq!(days_to_ymd(365), (1971, 1, 1));
        assert_eq!(days_to_ymd(11016), (2000, 2, 29));
        assert_eq!(days_to_ymd(11017), (2000, 3, 1));
        assert_eq!(days_to_ymd(20566), (2026, 4, 23));
    }

    #[test]
    fn substitute_tokens_expands_hash_and_at() {
        assert_eq!(
            substitute_tokens("Page # of 5 (@)", 3, "23-Apr-26"),
            "Page 3 of 5 (23-Apr-26)"
        );
        assert_eq!(substitute_tokens("no tokens", 7, "X"), "no tokens");
    }
}
