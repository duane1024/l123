//! Resolve the user name and organization shown on the startup splash.
//!
//! Thin view over [`crate::config::Config`] — every source of truth
//! and fallback ordering lives there; this module just projects the
//! two string fields the splash cares about.

use crate::config::{Config, ConfigSource, StdSource};

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Identity {
    pub user: String,
    pub organization: String,
}

impl Identity {
    pub fn resolve() -> Self {
        Self::resolve_with(&StdSource)
    }

    pub fn resolve_with(src: &dyn ConfigSource) -> Self {
        let cfg = Config::resolve_with(src);
        Identity {
            user: cfg.user.value,
            organization: cfg.organization.value,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::ConfigSource;
    use std::cell::RefCell;
    use std::collections::HashMap;
    use std::path::{Path, PathBuf};

    struct MockSource {
        vars: RefCell<HashMap<String, String>>,
        git_user: Option<String>,
        hostname: Option<String>,
    }

    impl MockSource {
        fn new() -> Self {
            Self {
                vars: RefCell::new(HashMap::new()),
                git_user: None,
                hostname: None,
            }
        }
        fn with_var(self, k: &str, v: &str) -> Self {
            self.vars.borrow_mut().insert(k.into(), v.into());
            self
        }
    }

    impl ConfigSource for MockSource {
        fn var(&self, key: &str) -> Option<String> {
            self.vars.borrow().get(key).cloned()
        }
        fn config_path(&self) -> Option<PathBuf> {
            None
        }
        fn read_file(&self, _path: &Path) -> Option<String> {
            None
        }
        fn git_user_name(&self) -> Option<String> {
            self.git_user.clone()
        }
        fn hostname(&self) -> Option<String> {
            self.hostname.clone()
        }
    }

    #[test]
    fn env_vars_win_over_everything() {
        let src = MockSource::new()
            .with_var("L123_USER", "Env User")
            .with_var("L123_ORG", "Env Org")
            .with_var("USER", "shell_user");
        let id = Identity::resolve_with(&src);
        assert_eq!(id.user, "Env User");
        assert_eq!(id.organization, "Env Org");
    }

    #[test]
    fn derives_from_git_and_hostname() {
        let src = MockSource {
            vars: RefCell::new(HashMap::new()),
            git_user: Some("Git Identity".to_string()),
            hostname: Some("host.local".to_string()),
        };
        let id = Identity::resolve_with(&src);
        assert_eq!(id.user, "Git Identity");
        assert_eq!(id.organization, "host.local");
    }

    #[test]
    fn placeholder_when_nothing_available() {
        let src = MockSource::new();
        let id = Identity::resolve_with(&src);
        assert_eq!(id.user, "l123 User");
        assert_eq!(id.organization, "l123");
    }
}
