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
    // Page width = sum of column widths in the selected range.
    let page_width_usize: usize = (r.start.col..=r.end.col)
        .map(|c| view.col_width(r.start.sheet, c) as usize)
        .sum();
    let page_width = page_width_usize.max(1);
    // Right margin trims content lines (after the left pad) down to
    // `page_width - margin_right`. `0` = no limit.
    let effective_width = if settings.margin_right == 0 {
        page_width
    } else {
        page_width.saturating_sub(settings.margin_right as usize)
    };
    let (header, footer) = match settings.format_mode {
        PrintFormatMode::Formatted => (settings.header.as_str(), settings.footer.as_str()),
        PrintFormatMode::Unformatted => ("", ""),
    };
    let content_mode = settings.content_mode;
    let left_pad: String = " ".repeat(settings.margin_left as usize);

    // Collect content rows (after pipe-row suppression, content-mode
    // rendering, and right-margin truncation). Each entry already has
    // `left_pad` prepended and a trailing `\n`.
    let mut rows: Vec<String> = Vec::new();
    for row in r.start.row..=r.end.row {
        let first = Address::new(r.start.sheet, r.start.col, row);
        if let Some(CellContents::Label {
            prefix: LabelPrefix::Pipe,
            ..
        }) = view.cell(first)
        {
            continue;
        }
        let mut line = String::new();
        for col in r.start.col..=r.end.col {
            let addr = Address::new(r.start.sheet, col, row);
            let w = view.col_width(r.start.sheet, col) as usize;
            let piece = match view.cell(addr) {
                Some(CellContents::Empty) | None => " ".repeat(w),
                Some(CellContents::Label { prefix, text }) => render_label(*prefix, text, w),
                Some(CellContents::Constant(v)) => {
                    let fmt = view.format_for_cell(addr);
                    render_value_in_cell(v, w, fmt).unwrap_or_else(|| " ".repeat(w))
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
                            render_value_in_cell(v, w, fmt).unwrap_or_else(|| " ".repeat(w))
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
        rows.push(entry);
    }

    // Chunk into pages. pg_length == 0 means no pagination — one
    // page with every row.
    let per_page = if settings.pg_length == 0 {
        rows.len().max(1)
    } else {
        settings.pg_length as usize
    };
    let chunked: Vec<Vec<String>> = if rows.is_empty() {
        vec![Vec::new()]
    } else {
        rows.chunks(per_page).map(<[String]>::to_vec).collect()
    };

    let today = today_ddmmmyy();
    let mut pages: Vec<Page> = Vec::with_capacity(chunked.len());
    for (i, page_rows) in chunked.into_iter().enumerate() {
        let page_no = settings.start_page as usize + i;
        let header_line = if header.is_empty() {
            None
        } else {
            let substituted = substitute_tokens(header, page_no, &today);
            let mut line = String::with_capacity(left_pad.len() + page_width);
            line.push_str(&left_pad);
            line.push_str(&format_three_part(&substituted, page_width));
            Some(line)
        };
        let footer_line = if footer.is_empty() {
            None
        } else {
            let substituted = substitute_tokens(footer, page_no, &today);
            let mut line = String::with_capacity(left_pad.len() + page_width);
            line.push_str(&left_pad);
            line.push_str(&format_three_part(&substituted, page_width));
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
        page_width: page_width as u16,
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
