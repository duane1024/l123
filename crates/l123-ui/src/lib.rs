//! Ratatui widgets: ControlPanel, Grid, StatusLine, and the event loop glue.

pub mod app;
pub mod clock;
pub mod config;
pub mod identity;
pub mod sysinfo;

pub use app::App;
pub use config::Config;
pub use identity::Identity;
