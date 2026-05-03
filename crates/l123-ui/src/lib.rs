//! Ratatui widgets: ControlPanel, Grid, StatusLine, and the event loop glue.

pub mod app;
pub mod clock;
pub mod config;
pub mod help;
pub mod identity;
pub mod sysinfo;
pub mod theme;

pub use app::App;
pub use config::Config;
pub use identity::Identity;
pub use theme::{Theme, ThemeName};
