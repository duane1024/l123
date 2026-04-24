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
    /// `l123 config [--init [--force]]` — show or initialize config.
    Config(ConfigAction),
    /// Usage error; print `msg` to stderr and exit 2.
    Usage(String),
}

#[derive(Debug, PartialEq, Eq)]
enum ConfigAction {
    /// Print effective configuration.
    Show,
    /// Write a sample L123.CNF at the default path.
    Init { force: bool },
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

SUBCOMMANDS:
    config                 Show effective configuration and sources
    config --init          Write a sample ~/.l123/L123.CNF
    config --init --force  Overwrite an existing L123.CNF

ENVIRONMENT:
    L123_USER   Name shown on the startup splash.
    L123_ORG    Organization shown on the startup splash.
    L123_LOG    Path to a log file. When set, events are appended to
                this file; when unset, no logging is performed.
    RUST_LOG    Standard tracing env filter (e.g. `l123=debug`).
                Defaults to `info` when L123_LOG is set.

CONFIG FILE:
    ~/.l123/L123.CNF    Optional. Run `l123 config --init` to create
                        an annotated sample. See docs/CONFIG.md for
                        the full reference.

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
        Action::Config(ConfigAction::Show) => {
            print!("{}", l123_ui::Config::resolve().render_table());
            ExitCode::SUCCESS
        }
        Action::Config(ConfigAction::Init { force }) => match init_config(force) {
            Ok(msg) => {
                println!("{msg}");
                ExitCode::SUCCESS
            }
            Err(e) => {
                eprintln!("l123 config --init: {e}");
                ExitCode::FAILURE
            }
        },
        Action::Usage(msg) => {
            eprintln!("{msg}");
            eprintln!("Try 'l123 --help' for more information.");
            ExitCode::from(2)
        }
    }
}

/// Write the sample `L123.CNF` at `~/.l123/L123.CNF`. Returns a
/// status string on success, or an `anyhow` error with a message
/// suitable for showing to the user.
fn init_config(force: bool) -> Result<String> {
    let path = l123_ui::config::default_config_path()
        .ok_or_else(|| anyhow::anyhow!("$HOME is not set; cannot determine config path"))?;
    if path.exists() && !force {
        anyhow::bail!(
            "{} already exists (use --force to overwrite)",
            path.display()
        );
    }
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| anyhow::anyhow!("creating {}: {e}", parent.display()))?;
    }
    std::fs::write(&path, l123_ui::config::SAMPLE_CNF)
        .map_err(|e| anyhow::anyhow!("writing {}: {e}", path.display()))?;
    Ok(format!("Wrote {}", path.display()))
}

fn parse(args: &[OsString]) -> Action {
    // `l123 --help` / `l123 --version` are honored no matter where
    // they appear, including before a subcommand (`l123 --help config`).
    for arg in args {
        let s = arg.to_string_lossy();
        match s.as_ref() {
            "-h" | "--help" => return Action::Help,
            "-V" | "--version" => return Action::Version,
            _ => {}
        }
    }

    // Subcommand dispatch: `config` must be the first non-flag arg.
    let first_nonflag = args.iter().find(|a| {
        let s = a.to_string_lossy();
        !s.starts_with('-') || s.as_ref() == "-"
    });
    if let Some(first) = first_nonflag {
        if first.to_string_lossy() == "config" {
            return parse_config_subcommand(args);
        }
    }

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

fn parse_config_subcommand(args: &[OsString]) -> Action {
    // Skip up to and including the `config` token. Any leading flags
    // before `config` (e.g. `l123 --help config`) were already handled
    // in the global branch; here we only see flags that belong to the
    // subcommand itself.
    let mut seen_config = false;
    let mut init = false;
    let mut force = false;
    for arg in args {
        let s = arg.to_string_lossy();
        if !seen_config {
            if s == "config" {
                seen_config = true;
            }
            continue;
        }
        match s.as_ref() {
            "--init" => init = true,
            "--force" | "-f" => force = true,
            "-h" | "--help" => return Action::Help,
            flag if flag.starts_with('-') => {
                return Action::Usage(format!("l123 config: unknown option '{flag}'"));
            }
            _ => {
                return Action::Usage(format!(
                    "l123 config: unexpected argument '{s}' (try --init)"
                ));
            }
        }
    }
    if force && !init {
        return Action::Usage("l123 config: --force only applies with --init".into());
    }
    Action::Config(if init {
        ConfigAction::Init { force }
    } else {
        ConfigAction::Show
    })
}

fn run(path: Option<PathBuf>) -> Result<()> {
    l123_ui::App::run_with_file(path)
}

/// Install a tracing subscriber that appends to `log_file` from the
/// resolved configuration. When `log_file` is empty (the default),
/// no subscriber is installed and `tracing::*!` macros compile down
/// to no-ops — zero overhead for users who don't opt in.
///
/// The returned `WorkerGuard` flushes the async appender on drop; keep
/// it alive until the process exits.
fn init_tracing() -> Option<tracing_appender::non_blocking::WorkerGuard> {
    let cfg = l123_ui::Config::resolve();
    let path = cfg.log_file_path()?;
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
    let filter = if !cfg.log_filter.value.is_empty() {
        tracing_subscriber::EnvFilter::try_new(&cfg.log_filter.value)
            .unwrap_or_else(|_| tracing_subscriber::EnvFilter::new("info"))
    } else {
        tracing_subscriber::EnvFilter::new("info")
    };
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

    #[test]
    fn config_alone_is_show() {
        assert_eq!(parse(&osv(&["config"])), Action::Config(ConfigAction::Show),);
    }

    #[test]
    fn config_init_without_force() {
        assert_eq!(
            parse(&osv(&["config", "--init"])),
            Action::Config(ConfigAction::Init { force: false }),
        );
    }

    #[test]
    fn config_init_with_force_flag() {
        assert_eq!(
            parse(&osv(&["config", "--init", "--force"])),
            Action::Config(ConfigAction::Init { force: true }),
        );
        assert_eq!(
            parse(&osv(&["config", "--init", "-f"])),
            Action::Config(ConfigAction::Init { force: true }),
        );
    }

    #[test]
    fn config_force_without_init_is_usage_error() {
        match parse(&osv(&["config", "--force"])) {
            Action::Usage(m) => assert!(m.contains("--force"), "msg: {m}"),
            other => panic!("expected Usage, got {other:?}"),
        }
    }

    #[test]
    fn config_unknown_option_is_usage_error() {
        match parse(&osv(&["config", "--bogus"])) {
            Action::Usage(m) => assert!(m.contains("--bogus"), "msg: {m}"),
            other => panic!("expected Usage, got {other:?}"),
        }
    }

    #[test]
    fn config_positional_is_usage_error() {
        match parse(&osv(&["config", "random"])) {
            Action::Usage(m) => assert!(m.contains("random"), "msg: {m}"),
            other => panic!("expected Usage, got {other:?}"),
        }
    }

    #[test]
    fn file_named_config_still_opens_as_subcommand() {
        // Documented behavior: `config` is reserved as a subcommand.
        // Users wanting to open a file literally named `config` must
        // qualify with a path, e.g. `./config`.
        assert_eq!(parse(&osv(&["config"])), Action::Config(ConfigAction::Show),);
    }

    #[test]
    fn file_path_ending_in_config_is_not_subcommand() {
        assert_eq!(
            parse(&osv(&["./config"])),
            Action::Run(Some(PathBuf::from("./config"))),
        );
    }

    #[test]
    fn global_help_before_config_still_wins() {
        assert_eq!(parse(&osv(&["--help", "config"])), Action::Help);
    }
}
