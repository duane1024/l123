//! Ratatui widgets: ControlPanel, Grid, StatusLine, and the event loop glue.

pub mod app;
pub mod clock;
pub mod config;
pub mod help;
pub mod help_topics_decoded;
pub mod identity;
pub mod sysinfo;

pub use app::App;
pub use config::Config;
pub use identity::Identity;
