//! Runtime state for the F1 help overlay.
//!
//! The help corpus itself lives in `l123-help` as a sorted slice of
//! `(filename, html)` pairs embedded at build time. This module owns
//! only the navigation state — current page, focused link index, and
//! a back-stack of previously-visited filenames so BACKSPACE can pop.

use l123_core::Mode;
use l123_help::HelpPage;

/// F1 HELP overlay state.
#[derive(Debug)]
pub(crate) struct HelpState {
    /// Currently displayed help page.
    pub page: HelpPage,
    /// Index into [`HelpPage::links`] of the link the up/down cursor is
    /// resting on. `0` is valid only when the page has at least one
    /// link; pages with no links keep this at `0` and ignore navigation.
    pub focus: usize,
    /// First body row visible at the top of the rendered overlay. The
    /// renderer auto-scrolls to keep the focused link in view, so the
    /// user rarely sees this directly.
    pub scroll: u16,
    /// Stack of previously-visited page filenames. ENTER pushes the
    /// current page's filename before navigating; BACKSPACE pops.
    pub history: Vec<&'static str>,
    /// Mode that was active when F1 was first pressed; restored on Esc.
    pub return_mode: Mode,
}

impl HelpState {
    pub(crate) fn open(return_mode: Mode) -> Option<Self> {
        let page = l123_help::load_page(l123_help::INDEX_FILENAME)?;
        Some(Self {
            page,
            focus: 0,
            scroll: 0,
            history: Vec::new(),
            return_mode,
        })
    }

    /// Navigate to `target`, pushing the current page onto the history
    /// stack. Returns `true` if the target was found.
    pub(crate) fn follow(&mut self, target: &str) -> bool {
        let Some(page) = l123_help::load_page(target) else {
            return false;
        };
        self.history.push(self.page.filename);
        self.page = page;
        self.focus = 0;
        self.scroll = 0;
        true
    }

    /// Pop the history stack. Returns `true` if a page was popped.
    pub(crate) fn pop(&mut self) -> bool {
        let Some(prev) = self.history.pop() else {
            return false;
        };
        match l123_help::load_page(prev) {
            Some(page) => {
                self.page = page;
                self.focus = 0;
                self.scroll = 0;
                true
            }
            None => false,
        }
    }

    pub(crate) fn focus_down(&mut self) {
        self.focus_2d(|cur, cand| {
            if cand.0 <= cur.0 {
                None
            } else {
                Some((cand.0 - cur.0, cand.1.abs_diff(cur.1)))
            }
        });
    }

    pub(crate) fn focus_up(&mut self) {
        self.focus_2d(|cur, cand| {
            if cand.0 >= cur.0 {
                None
            } else {
                Some((cur.0 - cand.0, cand.1.abs_diff(cur.1)))
            }
        });
    }

    pub(crate) fn focus_right(&mut self) {
        self.focus_2d(|cur, cand| {
            if cand.0 != cur.0 || cand.1 <= cur.1 {
                None
            } else {
                Some((0, cand.1 - cur.1))
            }
        });
    }

    pub(crate) fn focus_left(&mut self) {
        self.focus_2d(|cur, cand| {
            if cand.0 != cur.0 || cand.1 >= cur.1 {
                None
            } else {
                Some((0, cur.1 - cand.1))
            }
        });
    }

    /// Pick the link whose `(row_distance, col_distance)` from the
    /// current focus is smallest, after `score` has filtered out
    /// candidates outside the desired direction. Candidates returning
    /// `None` are ignored. Tuple ordering means smaller row distance
    /// always wins over column distance.
    fn focus_2d<F>(&mut self, score: F)
    where
        F: Fn((usize, usize), (usize, usize)) -> Option<(usize, usize)>,
    {
        let positions = self.page.link_positions();
        let Some(&cur) = positions.get(self.focus) else {
            return;
        };
        let mut best: Option<(usize, usize, usize)> = None;
        for (i, &cand) in positions.iter().enumerate() {
            if i == self.focus {
                continue;
            }
            if let Some((rd, cd)) = score(cur, cand) {
                let key = (rd, cd, i);
                if best.is_none_or(|b| key < b) {
                    best = Some(key);
                }
            }
        }
        if let Some((_, _, i)) = best {
            self.focus = i;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_loads_index() {
        let s = HelpState::open(Mode::Ready).expect("index loads");
        assert_eq!(s.page.filename, l123_help::INDEX_FILENAME);
        assert!(!s.page.links.is_empty());
        assert!(s.history.is_empty());
    }

    #[test]
    fn follow_pushes_history_and_swaps_page() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        let target = s.page.links[0].target.clone();
        assert!(s.follow(&target));
        assert_eq!(s.page.filename, target);
        assert_eq!(s.history, vec![l123_help::INDEX_FILENAME]);
        assert_eq!(s.focus, 0);
    }

    #[test]
    fn follow_unknown_target_is_noop() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        let before = s.page.filename;
        assert!(!s.follow("nope.html"));
        assert_eq!(s.page.filename, before);
        assert!(s.history.is_empty());
    }

    #[test]
    fn pop_returns_to_previous_page() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        let target = s.page.links[0].target.clone();
        s.follow(&target);
        assert!(s.pop());
        assert_eq!(s.page.filename, l123_help::INDEX_FILENAME);
        assert!(s.history.is_empty());
    }

    #[test]
    fn pop_with_empty_history_is_noop() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        assert!(!s.pop());
    }

    /// Snapshot a (row, col, target) tuple for the focused link so
    /// tests can assert against the visual layout.
    fn focus_position(s: &HelpState) -> (usize, usize, &str) {
        let positions = s.page.link_positions();
        let (r, c) = positions[s.focus];
        (r, c, s.page.links[s.focus].target.as_str())
    }

    #[test]
    fn focus_down_stays_in_same_column_on_index() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        // The first link sits at row 0 inset by 6 columns
        // (`1-2-3 Help Index`). Step down once to drop into the
        // 3-column grid, then step down again — that second hop is
        // what should stay in column.
        s.focus_down();
        let (r0, c0, _) = focus_position(&s);
        s.focus_down();
        let (r1, c1, _) = focus_position(&s);
        assert!(r1 > r0, "row must advance");
        assert_eq!(c1, c0, "down should stay in column {c0}, landed at {c1}");
    }

    #[test]
    fn focus_right_walks_same_row() {
        // Pick a page whose body has multiple links on one row. The
        // /Copy footer has three on its last line.
        let mut s = HelpState::open(Mode::Ready).unwrap();
        // Manually load the /Copy page.
        s.page = l123_help::load_page("0006-copy.html").unwrap();
        s.focus = 0;
        let (r0, c0, _) = focus_position(&s);
        s.focus_right();
        let (r1, c1, _) = focus_position(&s);
        assert_eq!(r0, r1, "right must stay on the same row");
        assert!(c1 > c0, "right must advance the column");
    }

    #[test]
    fn focus_right_at_end_of_row_is_noop() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        s.page = l123_help::load_page("0006-copy.html").unwrap();
        // Walk all the way right.
        for _ in 0..s.page.links.len() {
            s.focus_right();
        }
        let before = s.focus;
        s.focus_right();
        assert_eq!(s.focus, before);
    }

    #[test]
    fn focus_left_inverts_focus_right() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        s.page = l123_help::load_page("0006-copy.html").unwrap();
        s.focus = 0;
        s.focus_right();
        let mid = s.focus;
        s.focus_left();
        assert_ne!(s.focus, mid);
    }

    #[test]
    fn focus_up_inverts_focus_down() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        let start = s.focus;
        s.focus_down();
        s.focus_up();
        assert_eq!(s.focus, start);
    }

    #[test]
    fn focus_left_at_start_of_row_is_noop() {
        let mut s = HelpState::open(Mode::Ready).unwrap();
        let before = s.focus;
        s.focus_left();
        assert_eq!(s.focus, before);
    }
}
