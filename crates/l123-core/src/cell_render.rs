//! Pure functions that render a cell's contents into a fixed-width
//! monospace slot. Shared by the on-screen renderer in `l123-ui` and
//! the print renderer in `l123-print`: alignment rules don't depend
//! on the output medium.

use crate::{Format, FormatKind, LabelPrefix, Value};

/// One slot in a row for label-spill planning. Callers classify each
/// visible column into one of these variants — see [`plan_row_spill`].
#[derive(Debug, Clone)]
pub enum SpillSlot<'a> {
    /// No content. Can be claimed by an adjacent label that overflows.
    Empty,
    /// A label. May spill into adjacent [`SpillSlot::Empty`] cells.
    Label { prefix: LabelPrefix, text: &'a str },
    /// Already-rendered content, exactly `width` chars wide. Blocks
    /// adjacent labels from spilling through.
    Rendered(String),
}

/// Plan the text to paint into each column slot of one row, honoring
/// the 1-2-3 label-spill contract:
///
/// * apostrophe / pipe labels overflow rightward into consecutive
///   empty cells;
/// * quote labels overflow leftward;
/// * caret labels overflow both ways, centered in the combined span;
/// * backslash (fill) labels never spill;
/// * values, formula results, and non-empty neighbors hard-stop the
///   spill, and the label truncates at its own column boundary.
///
/// Returns a `Vec` the same length as `slots`; entry `i` is exactly
/// `widths[i]` chars wide. Callers paint each slot with its own
/// styling (e.g. per-column highlight) — spill does not fold slot
/// boundaries into a single style run.
pub fn plan_row_spill(slots: &[SpillSlot<'_>], widths: &[usize]) -> Vec<String> {
    assert_eq!(
        slots.len(),
        widths.len(),
        "plan_row_spill: slots/widths length mismatch"
    );
    let mut out: Vec<String> = widths.iter().map(|w| " ".repeat(*w)).collect();
    let mut claimed: Vec<bool> = vec![false; slots.len()];

    for i in 0..slots.len() {
        match &slots[i] {
            SpillSlot::Rendered(s) => {
                out[i] = s.clone();
                claimed[i] = true;
            }
            SpillSlot::Empty => {}
            SpillSlot::Label { prefix, text } => {
                let own_w = widths[i];
                let text_len = text.chars().count();
                if *prefix == LabelPrefix::Backslash || text_len <= own_w {
                    out[i] = render_label(*prefix, text, own_w);
                    claimed[i] = true;
                    continue;
                }
                let (start, end, span_w) =
                    spill_extent(*prefix, i, own_w, text_len, slots, widths, &claimed);
                let full = match prefix {
                    LabelPrefix::Apostrophe | LabelPrefix::Pipe => right_pad(text, span_w, false),
                    LabelPrefix::Quote => right_pad(text, span_w, true),
                    LabelPrefix::Caret => center_pad(text, span_w),
                    LabelPrefix::Backslash => unreachable!("handled above"),
                };
                let chars: Vec<char> = full.chars().collect();
                let mut idx = 0;
                for k in start..=end {
                    let w = widths[k];
                    let slice: String = chars.iter().skip(idx).take(w).collect();
                    idx += w;
                    out[k] = slice;
                    claimed[k] = true;
                }
            }
        }
    }
    out
}

fn is_available(slots: &[SpillSlot<'_>], claimed: &[bool], idx: usize) -> bool {
    matches!(slots[idx], SpillSlot::Empty) && !claimed[idx]
}

fn spill_extent(
    prefix: LabelPrefix,
    i: usize,
    own_w: usize,
    text_len: usize,
    slots: &[SpillSlot<'_>],
    widths: &[usize],
    claimed: &[bool],
) -> (usize, usize, usize) {
    match prefix {
        LabelPrefix::Apostrophe | LabelPrefix::Pipe => {
            let mut end = i;
            let mut span = own_w;
            while text_len > span && end + 1 < slots.len() && is_available(slots, claimed, end + 1)
            {
                end += 1;
                span += widths[end];
            }
            (i, end, span)
        }
        LabelPrefix::Quote => {
            let mut start = i;
            let mut span = own_w;
            while text_len > span && start > 0 && is_available(slots, claimed, start - 1) {
                start -= 1;
                span += widths[start];
            }
            (start, i, span)
        }
        LabelPrefix::Caret => {
            // Symmetric extension: while both sides are empty, claim
            // one cell on each side per iteration. When only one side
            // remains, keep extending that side alone. This matches
            // 1-2-3's centered-label spill — the text is centered
            // within the combined span rather than leaning toward
            // whichever side the algorithm visits first.
            let mut start = i;
            let mut end = i;
            let mut span = own_w;
            while text_len > span {
                let can_left = start > 0 && is_available(slots, claimed, start - 1);
                let can_right = end + 1 < slots.len() && is_available(slots, claimed, end + 1);
                if !can_left && !can_right {
                    break;
                }
                if can_left {
                    start -= 1;
                    span += widths[start];
                }
                if can_right {
                    end += 1;
                    span += widths[end];
                }
            }
            (start, end, span)
        }
        LabelPrefix::Backslash => (i, i, own_w),
    }
}

/// Render a value into a `width`-character cell using `format`.
///
/// Numbers are right-aligned and overflow to asterisks (Authenticity
/// Contract §20.9); `General` format short-circuits the overflow
/// check for now since it would normally switch to scientific first.
/// Text is left-aligned; booleans and errors right-aligned. `Empty`
/// yields `None` so callers can decide whether to blank or leave the
/// slot untouched.
pub fn render_value_in_cell(v: &Value, width: usize, format: Format) -> Option<String> {
    match v {
        Value::Number(n) => {
            let s = crate::format_number(*n, format);
            if s.chars().count() > width && !matches!(format.kind, FormatKind::General) {
                Some("*".repeat(width))
            } else {
                Some(right_pad(&s, width, true))
            }
        }
        Value::Text(s) => Some(right_pad(s, width, false)),
        Value::Bool(b) => Some(right_pad(if *b { "TRUE" } else { "FALSE" }, width, true)),
        Value::Error(e) => Some(right_pad(e.lotus_tag(), width, true)),
        Value::Empty => None,
    }
}

/// Render a label into a `width`-character cell. The prefix character
/// selects alignment per 1-2-3 convention: `'` left, `"` right, `^`
/// centered, `\` repeated to fill, `|` left (and suppressed from
/// print by the caller).
pub fn render_label(prefix: LabelPrefix, text: &str, width: usize) -> String {
    if text.is_empty() {
        return " ".repeat(width);
    }
    match prefix {
        LabelPrefix::Apostrophe | LabelPrefix::Pipe => right_pad(text, width, false),
        LabelPrefix::Quote => right_pad(text, width, true),
        LabelPrefix::Caret => center_pad(text, width),
        LabelPrefix::Backslash => repeat_to_width(text, width),
    }
}

/// Left-pad (`right_align = true`) or right-pad with spaces to
/// `width`, or truncate to `width`.
pub fn right_pad(text: &str, width: usize, right_align: bool) -> String {
    let chars: Vec<char> = text.chars().collect();
    if chars.len() >= width {
        return chars.into_iter().take(width).collect();
    }
    let pad = width - chars.len();
    if right_align {
        format!("{}{text}", " ".repeat(pad))
    } else {
        format!("{text}{}", " ".repeat(pad))
    }
}

pub fn center_pad(text: &str, width: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    if chars.len() >= width {
        return chars.into_iter().take(width).collect();
    }
    let pad = width - chars.len();
    let left = pad / 2;
    let right = pad - left;
    format!("{}{text}{}", " ".repeat(left), " ".repeat(right))
}

pub fn repeat_to_width(pattern: &str, width: usize) -> String {
    if pattern.is_empty() {
        return " ".repeat(width);
    }
    let mut out = String::with_capacity(width);
    let pattern_chars: Vec<char> = pattern.chars().collect();
    while out.chars().count() < width {
        for ch in &pattern_chars {
            if out.chars().count() == width {
                break;
            }
            out.push(*ch);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn render_label_alignments() {
        assert_eq!(render_label(LabelPrefix::Apostrophe, "hi", 9), "hi       ");
        assert_eq!(render_label(LabelPrefix::Quote, "hi", 9), "       hi");
        assert_eq!(render_label(LabelPrefix::Caret, "hi", 9), "   hi    ");
        assert_eq!(render_label(LabelPrefix::Backslash, "-", 9), "---------");
        assert_eq!(render_label(LabelPrefix::Backslash, "ab", 9), "ababababa");
        assert_eq!(render_label(LabelPrefix::Backslash, "abc", 9), "abcabcabc");
    }

    fn label(prefix: LabelPrefix, text: &str) -> SpillSlot<'_> {
        SpillSlot::Label { prefix, text }
    }

    fn rendered(s: &str) -> SpillSlot<'static> {
        SpillSlot::Rendered(s.to_string())
    }

    #[test]
    fn apostrophe_label_fits_in_own_width() {
        let slots = [
            label(LabelPrefix::Apostrophe, "hi"),
            SpillSlot::Empty,
            SpillSlot::Empty,
        ];
        let got = plan_row_spill(&slots, &[9, 9, 9]);
        assert_eq!(got[0], "hi       ");
        assert_eq!(got[1], "         ");
        assert_eq!(got[2], "         ");
    }

    #[test]
    fn apostrophe_label_spills_right_into_empty_neighbors() {
        let slots = [
            label(LabelPrefix::Apostrophe, "INCOME SUMMARY 1991"),
            SpillSlot::Empty,
            SpillSlot::Empty,
        ];
        let got = plan_row_spill(&slots, &[9, 9, 9]);
        assert_eq!(got[0], "INCOME SU");
        assert_eq!(got[1], "MMARY 199");
        assert_eq!(got[2], "1        ");
    }

    #[test]
    fn left_spill_hard_stops_at_non_empty_neighbor() {
        let slots = [
            label(LabelPrefix::Apostrophe, "blocked spill"),
            label(LabelPrefix::Apostrophe, "X"),
            SpillSlot::Empty,
        ];
        let got = plan_row_spill(&slots, &[9, 9, 9]);
        assert_eq!(got[0], "blocked s");
        assert_eq!(got[1], "X        ");
        assert_eq!(got[2], "         ");
    }

    #[test]
    fn left_spill_blocked_by_rendered_value() {
        let slots = [
            label(LabelPrefix::Apostrophe, "overflowing label"),
            rendered("     1234"),
        ];
        let got = plan_row_spill(&slots, &[9, 9]);
        assert_eq!(got[0], "overflowi");
        assert_eq!(got[1], "     1234");
    }

    #[test]
    fn quote_label_spills_left_into_empty_neighbor() {
        let slots = [
            SpillSlot::Empty,
            label(LabelPrefix::Quote, "right-align spill"),
        ];
        let got = plan_row_spill(&slots, &[9, 9]);
        assert_eq!(got[0], " right-al");
        assert_eq!(got[1], "ign spill");
    }

    #[test]
    fn right_spill_hard_stops_at_non_empty_neighbor() {
        let slots = [
            label(LabelPrefix::Apostrophe, "L"),
            label(LabelPrefix::Quote, "cant spill left"),
        ];
        let got = plan_row_spill(&slots, &[9, 9]);
        assert_eq!(got[0], "L        ");
        // No empty neighbor on the left → falls back to own-width
        // truncation (first-9-chars, matching `right_pad`).
        assert_eq!(got[1], "cant spil");
    }

    #[test]
    fn caret_label_spills_both_directions_centered() {
        let slots = [
            SpillSlot::Empty,
            label(LabelPrefix::Caret, "centered text"),
            SpillSlot::Empty,
        ];
        let got = plan_row_spill(&slots, &[9, 9, 9]);
        // "centered text" = 13 chars, span = 27 → 7 pad each side.
        assert_eq!(got[0], "       ce");
        assert_eq!(got[1], "ntered te");
        assert_eq!(got[2], "xt       ");
    }

    #[test]
    fn backslash_label_never_spills() {
        let slots = [label(LabelPrefix::Backslash, "-"), SpillSlot::Empty];
        let got = plan_row_spill(&slots, &[9, 9]);
        assert_eq!(got[0], "---------");
        assert_eq!(got[1], "         ");
    }

    #[test]
    fn first_spill_claims_middle_empty_blocking_later_spiller() {
        // Left-align label at 0 claims slot 1 for its spill; a
        // right-align label at 2 cannot then spill into slot 1.
        let slots = [
            label(LabelPrefix::Apostrophe, "first label spills"),
            SpillSlot::Empty,
            label(LabelPrefix::Quote, "second"),
        ];
        let got = plan_row_spill(&slots, &[9, 9, 9]);
        assert_eq!(got[0], "first lab");
        assert_eq!(got[1], "el spills");
        assert_eq!(got[2], "   second");
    }
}
