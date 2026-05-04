//! Process entry points and the event loop.
//!
//! `App::run`, `run_with_file`, `run_with` set up crossterm and the
//! tokio runtime, then hand the alt-screen Terminal to `event_loop`.
//! `/System` interruption (the `pending_system_suspend` flag) drops
//! into `suspend_to_shell` between iterations.

use std::io;
use std::path::PathBuf;
use std::time::Duration;

use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event, KeyEventKind,
    },
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{backend::CrosstermBackend, Terminal};

use super::{emit_bell, App};

impl App {
    pub fn run() -> anyhow::Result<()> {
        Self::run_with(None, None)
    }

    /// CLI entry point. When `path` is set the app opens that workbook
    /// and skips the splash; when `None` it greets the user with the
    /// licensing block until the first keypress.
    pub fn run_with_file(path: Option<PathBuf>) -> anyhow::Result<()> {
        Self::run_with(path, None)
    }

    /// Full entry point — accepts an optional file and an optional
    /// `--theme` override that wins over env / config-file. The
    /// override is `Some` when the user passed `--theme`; `None` lets
    /// the resolved [`crate::Config`] decide.
    pub fn run_with(
        path: Option<PathBuf>,
        theme_override: Option<crate::Theme>,
    ) -> anyhow::Result<()> {
        let mut stdout = io::stdout();
        enable_raw_mode()?;
        execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        let cfg = crate::Config::resolve();
        let mut app = match path {
            Some(p) => App::new_with_file(p),
            None => App::new_with_splash(cfg.user.value.clone(), cfg.organization.value.clone()),
        };
        app.set_beep_enabled(cfg.error_beep_enabled());
        app.set_theme(theme_override.unwrap_or_else(|| cfg.theme()));
        app.probe_image_picker();
        let result = app.event_loop(&mut terminal);

        disable_raw_mode()?;
        execute!(
            terminal.backend_mut(),
            LeaveAlternateScreen,
            DisableMouseCapture
        )?;
        terminal.show_cursor()?;

        result
    }

    fn event_loop<B: ratatui::backend::Backend>(
        &mut self,
        terminal: &mut Terminal<B>,
    ) -> anyhow::Result<()>
    where
        B::Error: Send + Sync + 'static,
    {
        while self.running {
            terminal.draw(|f| self.render(f.area(), f.buffer_mut()))?;
            if self.take_pending_beep() {
                emit_bell();
            }
            if self.pending_system_suspend {
                self.pending_system_suspend = false;
                suspend_to_shell(terminal)?;
                continue;
            }
            // §4.7 — drain any queued long-running op. The render
            // above already showed the WAIT frame for this iteration;
            // the next render after the drain shows READY.
            self.tick();
            if event::poll(Duration::from_millis(100))? {
                match event::read()? {
                    Event::Key(k) if k.kind == KeyEventKind::Press => self.handle_key(k),
                    Event::Mouse(m) => self.handle_mouse(m),
                    _ => {}
                }
            }
        }
        Ok(())
    }
}

/// `/System` — leave the alt screen + raw mode, run an interactive
/// shell, and on its exit restore the TUI. Mirrors the original 1-2-3
/// R3.4a behavior of suspending to a DOS shell ("Type EXIT to return
/// to 1-2-3"). Errors from the spawn are printed to the underlying
/// terminal and otherwise swallowed; we always try to restore the TUI.
fn suspend_to_shell<B: ratatui::backend::Backend>(terminal: &mut Terminal<B>) -> anyhow::Result<()>
where
    B::Error: Send + Sync + 'static,
{
    let mut stdout = io::stdout();
    disable_raw_mode()?;
    execute!(stdout, LeaveAlternateScreen, DisableMouseCapture)?;
    terminal.show_cursor()?;

    println!();
    println!("(Type 'exit' to return to 1-2-3.)");

    #[cfg(windows)]
    let shell = std::env::var("COMSPEC").unwrap_or_else(|_| "cmd.exe".to_string());
    #[cfg(not(windows))]
    let shell = std::env::var("SHELL").unwrap_or_else(|_| "/bin/sh".to_string());

    if let Err(e) = std::process::Command::new(&shell).status() {
        eprintln!("l123: /System: failed to launch {shell}: {e}");
    }

    enable_raw_mode()?;
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    terminal.clear()?;
    Ok(())
}
