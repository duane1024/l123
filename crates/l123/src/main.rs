//! L123 binary entry point.

use std::ffi::OsString;
use std::path::PathBuf;
use std::process::ExitCode;

use anyhow::Result;

/// Decoded command-line action.
#[derive(Debug, PartialEq, Eq)]
enum Action {
    /// Run the TUI, optionally opening `path`.
    Run(Option<PathBuf>),
    /// Print --help to stdout and exit 0.
    Help,
    /// Print --version to stdout and exit 0.
    Version,
    /// Usage error; print `msg` to stderr and exit 2.
    Usage(String),
}

const HELP: &str = "\
l123 — a TUI clone of Lotus 1-2-3 Release 3.4a for DOS

USAGE:
    l123 [OPTIONS] [FILE]

ARGS:
    <FILE>    Workbook to open (.xlsx, .wk3). If omitted, starts empty.

OPTIONS:
    -h, --help       Print this help and exit
    -V, --version    Print version and exit

ENVIRONMENT:
    L123_LOG    Path to a log file. When set, events are appended to
                this file; when unset, no logging is performed.
    RUST_LOG    Standard tracing env filter (e.g. `l123=debug`).
                Defaults to `info` when L123_LOG is set.

Inside the program:
    /         Open the 1-2-3 slash menu
    F1        Help
    F10       Graph view
    /QY       Quit
";

fn main() -> ExitCode {
    let _log_guard = init_tracing();
    let args: Vec<OsString> = std::env::args_os().skip(1).collect();
    match parse(&args) {
        Action::Run(path) => match run(path) {
            Ok(()) => ExitCode::SUCCESS,
            Err(e) => {
                tracing::error!(error = %e, "l123 run failed");
                eprintln!("l123: {e:#}");
                ExitCode::FAILURE
            }
        },
        Action::Help => {
            print!("{HELP}");
            ExitCode::SUCCESS
        }
        Action::Version => {
            println!("l123 {}", env!("CARGO_PKG_VERSION"));
            ExitCode::SUCCESS
        }
        Action::Usage(msg) => {
            eprintln!("{msg}");
            eprintln!("Try 'l123 --help' for more information.");
            ExitCode::from(2)
        }
    }
}

fn parse(args: &[OsString]) -> Action {
    let mut positional: Option<PathBuf> = None;
    for arg in args {
        let s = arg.to_string_lossy();
        match s.as_ref() {
            "-h" | "--help" => return Action::Help,
            "-V" | "--version" => return Action::Version,
            flag if flag.starts_with('-') && flag != "-" => {
                return Action::Usage(format!("l123: unknown option '{flag}'"));
            }
            _ => {
                if positional.is_some() {
                    return Action::Usage(
                        "l123: too many arguments (expected at most one FILE)".into(),
                    );
                }
                positional = Some(PathBuf::from(arg));
            }
        }
    }
    Action::Run(positional)
}

fn run(path: Option<PathBuf>) -> Result<()> {
    l123_ui::App::run_with_file(path)
}

/// Install a tracing subscriber that appends to the file named by
/// `L123_LOG`. When the env var is unset, no subscriber is installed
/// and `tracing::*!` macros compile down to no-ops — zero overhead for
/// users who don't opt in.
///
/// The returned `WorkerGuard` flushes the async appender on drop; keep
/// it alive until the process exits.
fn init_tracing() -> Option<tracing_appender::non_blocking::WorkerGuard> {
    let path = PathBuf::from(std::env::var_os("L123_LOG")?);
    let file_name = path.file_name()?.to_os_string();
    let dir = path
        .parent()
        .map(PathBuf::from)
        .unwrap_or_else(PathBuf::new);
    if !dir.as_os_str().is_empty() {
        let _ = std::fs::create_dir_all(&dir);
    }
    let appender = tracing_appender::rolling::never(
        if dir.as_os_str().is_empty() {
            PathBuf::from(".")
        } else {
            dir
        },
        file_name,
    );
    let (nb, guard) = tracing_appender::non_blocking(appender);
    let filter = tracing_subscriber::EnvFilter::try_from_default_env()
        .unwrap_or_else(|_| tracing_subscriber::EnvFilter::new("info"));
    let _ = tracing_subscriber::fmt()
        .with_env_filter(filter)
        .with_writer(nb)
        .with_ansi(false)
        .try_init();
    Some(guard)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn osv(args: &[&str]) -> Vec<OsString> {
        args.iter().map(|s| OsString::from(*s)).collect()
    }

    #[test]
    fn no_args_runs_empty() {
        assert_eq!(parse(&osv(&[])), Action::Run(None));
    }

    #[test]
    fn single_positional_opens_that_file() {
        assert_eq!(
            parse(&osv(&["sheet.xlsx"])),
            Action::Run(Some(PathBuf::from("sheet.xlsx")))
        );
    }

    #[test]
    fn short_help_flag() {
        assert_eq!(parse(&osv(&["-h"])), Action::Help);
    }

    #[test]
    fn long_help_flag() {
        assert_eq!(parse(&osv(&["--help"])), Action::Help);
    }

    #[test]
    fn short_version_flag() {
        assert_eq!(parse(&osv(&["-V"])), Action::Version);
    }

    #[test]
    fn long_version_flag() {
        assert_eq!(parse(&osv(&["--version"])), Action::Version);
    }

    #[test]
    fn unknown_option_is_usage_error() {
        match parse(&osv(&["--nope"])) {
            Action::Usage(m) => assert!(m.contains("--nope"), "msg should mention flag: {m}"),
            other => panic!("expected Usage, got {other:?}"),
        }
    }

    #[test]
    fn two_positionals_is_usage_error() {
        match parse(&osv(&["a.xlsx", "b.xlsx"])) {
            Action::Usage(_) => {}
            other => panic!("expected Usage, got {other:?}"),
        }
    }

    #[test]
    fn help_wins_over_positional_regardless_of_order() {
        assert_eq!(parse(&osv(&["sheet.xlsx", "--help"])), Action::Help);
    }
}
