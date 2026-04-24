//! Session modes. Every mode has a visible indicator in the control panel.
//! See SPEC §5.

#[derive(Copy, Clone, Debug, Default, PartialEq, Eq)]
pub enum Mode {
    #[default]
    Ready,
    Label,
    Value,
    Edit,
    Point,
    Menu,
    Files,
    Names,
    Help,
    Error,
    Wait,
    Find,
    Stat,
    /// Full-screen graph view entered by F10 or `/Graph View`. Esc
    /// returns to READY without mutating any graph state.
    Graph,
}

impl Mode {
    /// Right-justified indicator string as displayed on control-panel line 1.
    pub fn indicator(self) -> &'static str {
        match self {
            Mode::Ready => "READY",
            Mode::Label => "LABEL",
            Mode::Value => "VALUE",
            Mode::Edit => "EDIT",
            Mode::Point => "POINT",
            Mode::Menu => "MENU",
            Mode::Files => "FILES",
            Mode::Names => "NAMES",
            Mode::Help => "HELP",
            Mode::Error => "ERROR",
            Mode::Wait => "WAIT",
            Mode::Find => "FIND",
            Mode::Stat => "STAT",
            Mode::Graph => "GRAPH",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn all_indicators_ascii_upper() {
        for m in [
            Mode::Ready,
            Mode::Label,
            Mode::Value,
            Mode::Edit,
            Mode::Point,
            Mode::Menu,
            Mode::Files,
            Mode::Names,
            Mode::Help,
            Mode::Error,
            Mode::Wait,
            Mode::Find,
            Mode::Stat,
            Mode::Graph,
        ] {
            let s = m.indicator();
            assert!(s.chars().all(|c| c.is_ascii_uppercase()), "{m:?} → {s}");
            assert!(s.len() <= 5, "{m:?} indicator too long: {s}");
        }
    }
}
