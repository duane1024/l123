//! F1 help system. Topics are static `(id, title, body)` triples
//! authored in this file for now. Phase 2 of the help work will
//! replace this with a parser that lifts content from
//! `123.HLP` (the original Lotus 1-2-3 R3.4a help file) — at that
//! point this module becomes the runtime view layer over a richer
//! topic graph.
//!
//! Rendering and key handling live in [`crate::app`]; this module
//! holds only the topic data.
//!
//! Body text uses `\n` for line breaks. Bodies are kept short enough
//! that the overlay can show them without paging on a default 80×25
//! screen, but the renderer handles vertical scroll for completeness.
//!
//! Topic ordering matters: `Tab` walks the slice in order, so put the
//! index/overview first.
#[derive(Debug, Clone, Copy)]
pub struct HelpTopic {
    pub id: &'static str,
    pub title: &'static str,
    pub body: &'static str,
}

/// Static topic table. Tab/Shift-Tab walks this slice in order.
pub const HELP_TOPICS: &[HelpTopic] = &[
    HelpTopic {
        id: "index",
        title: "Help — Index",
        body: "\
L123 — a Lotus 1-2-3 R3.4a TUI clone.

Press TAB to walk the topic list, SHIFT-TAB to go back, PGUP/PGDN to
scroll within a topic, and ESC to close Help.

Topics:
  - Modes (READY, LABEL, VALUE, EDIT, POINT, MENU)
  - Function keys (F1..F10)
  - Slash menu (/)
  - Range commands (/Range …)
  - File commands (/File …)
  - Print commands (/Print …)
  - Quit
",
    },
    HelpTopic {
        id: "modes",
        title: "Modes",
        body: "\
The mode indicator on the right side of the control panel tells you
what the keyboard does next.

READY  Cell pointer is idle. Type a value to enter VALUE mode, a
       label-starter character (anything else) to enter LABEL mode,
       or `/` to open the menu.
LABEL  You are typing a label into the current cell. ENTER commits.
VALUE  You are typing a number or formula. ENTER commits.
EDIT   F2 was pressed; the cursor moves freely within the buffer.
POINT  A command is asking you to highlight a range. Move the
       pointer with arrows, type `,` to add another range to the
       list, then ENTER.
MENU   The slash menu is active. Letters / arrows navigate; ENTER
       selects.
",
    },
    HelpTopic {
        id: "function-keys",
        title: "Function keys",
        body: "\
F1   Help (this overlay).
F2   EDIT — re-open the current cell for in-place editing.
F3   NAMES — list defined range names. In POINT or in the GOTO and
     /Range Name Delete prompts.
F4   ABS — toggle absolute / relative cell references (formula
     entry; not yet implemented).
F5   GOTO — jump the pointer to a typed address or named range.
F9   CALC — recompute the workbook (when /Worksheet Global Recalc
     is set to Manual).
F10  GRAPH — show the current graph view.
ALT+F4  UNDO — revert the most recent edit.
",
    },
    HelpTopic {
        id: "menu",
        title: "Slash menu (/)",
        body: "\
Press `/` to open the main menu. Letters typed (or arrow + ENTER)
descend; ESC backs out one level. The colon menu (`:`) opens the
WYSIWYG submenu.

Top-level: Worksheet  Range  Copy  Move  File  Print  Graph  Data
            Exit  System  Quit

Each branch is documented in its own topic.
",
    },
    HelpTopic {
        id: "range",
        title: "/Range commands",
        body: "\
Operate on a typed range or a highlighted block.

  /Range Format        Apply a display format to one or more ranges.
  /Range Label         Realign label text within a range.
  /Range Erase         Clear cells (keeps formats).
  /Range Name Create   Define a name pointing at a range.
  /Range Name Delete   Remove a name.
  /Range Search        Find / replace within a range.

In the POINT step, type comma-separated lists for multi-range
inputs: `A1..B2,D4..E5`.
",
    },
    HelpTopic {
        id: "file",
        title: "/File commands",
        body: "\
  /File Save        Save the workbook to a path.
  /File Retrieve    Replace the workbook with one loaded from disk.
                    Recognizes .xlsx and .WK3 (R3 format).
  /File Open        Add a file to the active set.
  /File List        Show the file picker overlay.
  /File Erase       Remove a file from disk.
  /File Xtract      Save a range to a new file.
  /File Import      Read a CSV into the current sheet.
  /File Dir         Change the session working directory.
",
    },
    HelpTopic {
        id: "print",
        title: "/Print commands",
        body: "\
  /Print File       Write to a .prn (ASCII) or .pdf file.
  /Print Printer    Pipe to the system printer (CUPS lp).
  /Print Encoded    Raw printer-ready output to a path.

Each shares the Range / Line / Page / Options / Clear / Align / Go
/ Quit submenu. Range accepts a comma-separated list to print
multiple disjoint blocks as one job.
",
    },
    HelpTopic {
        id: "quit",
        title: "Quit",
        body: "\
/Quit Yes        Exit L123. Unsaved changes are lost.
/Quit No         Cancel and stay.

Ctrl+C from the terminal also exits, but skipping /Quit is rude
to anyone who configured an autosave hook.
",
    },
];

/// Lookup a topic by id. Returns None if the id is unknown.
pub fn find_topic(id: &str) -> Option<&'static HelpTopic> {
    HELP_TOPICS.iter().find(|t| t.id == id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn topic_table_nonempty_and_unique_ids() {
        assert!(!HELP_TOPICS.is_empty());
        let mut ids: Vec<&str> = HELP_TOPICS.iter().map(|t| t.id).collect();
        ids.sort();
        let count = ids.len();
        ids.dedup();
        assert_eq!(ids.len(), count, "duplicate topic ids");
    }

    #[test]
    fn first_topic_is_index() {
        assert_eq!(HELP_TOPICS[0].id, "index");
    }

    #[test]
    fn find_topic_lookup() {
        assert!(find_topic("index").is_some());
        assert!(find_topic("function-keys").is_some());
        assert!(find_topic("nope").is_none());
    }

    #[test]
    fn every_topic_has_nonempty_title_and_body() {
        for t in HELP_TOPICS {
            assert!(!t.title.is_empty(), "title empty: {}", t.id);
            assert!(!t.body.is_empty(), "body empty: {}", t.id);
        }
    }
}
