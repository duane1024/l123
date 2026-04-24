//! L123 binary entry point.

use std::path::PathBuf;
use std::process::ExitCode;

use anyhow::Result;

fn main() -> ExitCode {
    let mut args = std::env::args_os().skip(1);
    let first = args.next();
    if args.next().is_some() {
        eprintln!("usage: l123 [file.xlsx]");
        return ExitCode::from(2);
    }
    let path = first.map(PathBuf::from);
    match run(path) {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("l123: {e:#}");
            ExitCode::FAILURE
        }
    }
}

fn run(path: Option<PathBuf>) -> Result<()> {
    l123_ui::App::run_with_file(path)
}
