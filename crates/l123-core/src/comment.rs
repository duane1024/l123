//! Per-cell comments (Excel "notes" / threaded-comment legacy form)
//! carried from xlsx imports.  See `docs/XLSX_IMPORT_PLAN.md` §2.7.
//!
//! 1-2-3 R3.4a had `/Range Note` for cell notes; this type is the
//! L123 representation.  v1 imports legacy `comments1.xml` only —
//! threaded comments (`threadedComments.xml`) are not yet handled,
//! matching IronCalc's own scope.
//!
//! Rich-text formatting inside the comment body is flattened to a
//! plain string at the engine boundary; round-trip is lossy.

use crate::Address;

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct Comment {
    pub addr: Address,
    pub author: String,
    pub text: String,
}

impl Comment {
    pub fn new(addr: Address, author: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            addr,
            author: author.into(),
            text: text.into(),
        }
    }

    /// Single-line summary used by the control panel when the pointer
    /// lands on a commented cell.  Format: `<author>: <text>`.
    /// Newlines in the body are flattened to a single space so the
    /// summary stays on one row; long bodies are truncated by the
    /// caller to fit the panel width.
    pub fn summary(&self) -> String {
        let body: String = self
            .text
            .replace(['\r', '\n'], " ")
            .split_whitespace()
            .collect::<Vec<_>>()
            .join(" ");
        if self.author.is_empty() {
            body
        } else {
            format!("{}: {body}", self.author)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::SheetId;

    #[test]
    fn summary_includes_author_when_present() {
        let c = Comment::new(Address::new(SheetId::A, 0, 0), "Alice", "looks high");
        assert_eq!(c.summary(), "Alice: looks high");
    }

    #[test]
    fn summary_drops_author_when_empty() {
        let c = Comment::new(Address::new(SheetId::A, 0, 0), "", "anonymous note");
        assert_eq!(c.summary(), "anonymous note");
    }

    #[test]
    fn summary_flattens_newlines_and_collapses_whitespace() {
        let c = Comment::new(
            Address::new(SheetId::A, 0, 0),
            "Bob",
            "line one\nline two\n\n  three",
        );
        assert_eq!(c.summary(), "Bob: line one line two three");
    }
}
