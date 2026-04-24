//! Encode a [`PageGrid`] to CUPS `lp`.
//!
//! [`LpOptions`], [`lp_args`], and [`write_lp_stream`] are pure,
//! testable, and have no process side effects — available on every
//! platform. [`to_lp`] spawns `lp` and pipes the stream to its stdin;
//! it is unix-only because the `lp` binary ships on macOS / Linux but
//! not Windows.

use std::io::{self, Write};

use crate::encode::ascii::to_ascii;
use crate::grid::PageGrid;

/// CUPS `lp` invocation knobs.
#[derive(Debug, Clone, Default)]
pub struct LpOptions {
    /// `-d <destination>` — CUPS printer name. `None` uses the system
    /// default printer.
    pub destination: Option<String>,
    /// `-n <copies>`. `0` is coerced to `1`.
    pub copies: u16,
    /// Printer escape sequence prepended to the stream (1-2-3's
    /// `/Print Options Setup` string).
    pub setup_string: Option<String>,
}

/// Build the `lp` argv (without the leading `"lp"` binary name).
pub fn lp_args(opts: &LpOptions) -> Vec<String> {
    let mut out = Vec::new();
    if let Some(dest) = &opts.destination {
        out.push("-d".into());
        out.push(dest.clone());
    }
    if opts.copies > 1 {
        out.push("-n".into());
        out.push(opts.copies.to_string());
    }
    out
}

/// Write the bytes that go to `lp`'s stdin: optional setup string,
/// then the ASCII rendering of `grid`.
pub fn write_lp_stream<W: Write>(grid: &PageGrid, opts: &LpOptions, w: &mut W) -> io::Result<()> {
    if let Some(s) = &opts.setup_string {
        w.write_all(s.as_bytes())?;
    }
    w.write_all(to_ascii(grid).as_bytes())?;
    Ok(())
}

/// Spawn `lp` and pipe the stream to its stdin. Returns the child's
/// exit status. Unix-only; the `lp` binary is not present on Windows.
#[cfg(unix)]
pub fn to_lp(grid: &PageGrid, opts: &LpOptions) -> io::Result<std::process::ExitStatus> {
    use std::process::{Command, Stdio};
    let mut child = Command::new("lp")
        .args(lp_args(opts))
        .stdin(Stdio::piped())
        .spawn()?;
    {
        let stdin = child
            .stdin
            .as_mut()
            .ok_or_else(|| io::Error::other("lp child has no stdin"))?;
        write_lp_stream(grid, opts, stdin)?;
    }
    child.wait()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::{Page, PageGrid};

    fn one_page_grid() -> PageGrid {
        PageGrid {
            pages: vec![Page {
                number: 1,
                header: None,
                footer: None,
                rows: vec!["hello\n".into()],
                top_blank: 0,
                bottom_blank: 0,
            }],
            page_width: 5,
        }
    }

    #[test]
    fn lp_args_empty_by_default() {
        assert!(lp_args(&LpOptions::default()).is_empty());
    }

    #[test]
    fn lp_args_emits_destination() {
        let opts = LpOptions {
            destination: Some("Office_HP".into()),
            ..Default::default()
        };
        assert_eq!(lp_args(&opts), vec!["-d", "Office_HP"]);
    }

    #[test]
    fn lp_args_emits_copies_when_gt_one() {
        let opts = LpOptions {
            copies: 3,
            ..Default::default()
        };
        assert_eq!(lp_args(&opts), vec!["-n", "3"]);
    }

    #[test]
    fn lp_args_suppresses_copies_at_one_or_zero() {
        let one = LpOptions {
            copies: 1,
            ..Default::default()
        };
        let zero = LpOptions {
            copies: 0,
            ..Default::default()
        };
        assert!(lp_args(&one).is_empty());
        assert!(lp_args(&zero).is_empty());
    }

    #[test]
    fn lp_args_combines_destination_and_copies() {
        let opts = LpOptions {
            destination: Some("Lab".into()),
            copies: 2,
            setup_string: None,
        };
        assert_eq!(lp_args(&opts), vec!["-d", "Lab", "-n", "2"]);
    }

    #[test]
    fn stream_is_ascii_when_no_setup() {
        let grid = one_page_grid();
        let mut buf = Vec::new();
        write_lp_stream(&grid, &LpOptions::default(), &mut buf).unwrap();
        assert_eq!(buf, b"hello\n");
    }

    #[test]
    fn stream_prepends_setup_string() {
        let grid = one_page_grid();
        let opts = LpOptions {
            setup_string: Some("\x1bE".into()),
            ..Default::default()
        };
        let mut buf = Vec::new();
        write_lp_stream(&grid, &opts, &mut buf).unwrap();
        assert_eq!(buf, b"\x1bEhello\n");
    }
}
