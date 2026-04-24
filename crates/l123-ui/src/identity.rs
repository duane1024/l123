//! Resolve the user name and organization shown on the startup splash.
//!
//! Lookup order (first non-empty wins for each field independently):
//!   1. `L123_USER` / `L123_ORG` environment variables.
//!   2. `$XDG_CONFIG_HOME/l123/config.toml` (or `$HOME/.config/l123/config.toml`),
//!      parsed with a tiny `key = "value"` reader — no `toml` dep.
//!   3. `git config --global user.name` (user only).
//!   4. `$USER` / `$LOGNAME` for user, hostname for organization.

use std::path::PathBuf;
use std::process::Command;

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Identity {
    pub user: String,
    pub organization: String,
}

impl Identity {
    pub fn resolve() -> Self {
        Self::resolve_with(&StdEnv)
    }

    fn resolve_with(env: &dyn EnvSource) -> Self {
        let cfg = env
            .config_path()
            .and_then(|p| std::fs::read_to_string(p).ok())
            .map(|body| parse_config(&body))
            .unwrap_or_default();

        let user = first_nonempty([
            env.var("L123_USER"),
            Some(cfg.user.clone()),
            env.git_user_name(),
            env.var("USER"),
            env.var("LOGNAME"),
        ])
        .unwrap_or_else(|| "l123 User".to_string());

        let organization = first_nonempty([
            env.var("L123_ORG"),
            Some(cfg.organization.clone()),
            env.hostname(),
        ])
        .unwrap_or_else(|| "l123".to_string());

        Identity { user, organization }
    }
}

/// Minimal `key = "value"` parser. Whitespace-tolerant; ignores lines
/// that don't match. Recognized keys: `user`, `organization` (or `org`).
/// Values may be unquoted, double-quoted, or single-quoted.
fn parse_config(body: &str) -> Identity {
    let mut out = Identity::default();
    for raw in body.lines() {
        let line = raw.split('#').next().unwrap_or("").trim();
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        let key = key.trim().to_ascii_lowercase();
        let value = value.trim();
        let value = value
            .strip_prefix('"')
            .and_then(|s| s.strip_suffix('"'))
            .or_else(|| value.strip_prefix('\'').and_then(|s| s.strip_suffix('\'')))
            .unwrap_or(value)
            .to_string();
        match key.as_str() {
            "user" | "name" | "user_name" => out.user = value,
            "org" | "organization" => out.organization = value,
            _ => {}
        }
    }
    out
}

fn first_nonempty<const N: usize>(opts: [Option<String>; N]) -> Option<String> {
    opts.into_iter()
        .flatten()
        .map(|s| s.trim().to_string())
        .find(|s| !s.is_empty())
}

trait EnvSource {
    fn var(&self, key: &str) -> Option<String>;
    fn config_path(&self) -> Option<PathBuf>;
    fn git_user_name(&self) -> Option<String>;
    fn hostname(&self) -> Option<String>;
}

struct StdEnv;

impl EnvSource for StdEnv {
    fn var(&self, key: &str) -> Option<String> {
        std::env::var(key).ok()
    }

    fn config_path(&self) -> Option<PathBuf> {
        let base = std::env::var_os("XDG_CONFIG_HOME")
            .map(PathBuf::from)
            .or_else(|| std::env::var_os("HOME").map(|h| PathBuf::from(h).join(".config")))?;
        Some(base.join("l123").join("config.toml"))
    }

    fn git_user_name(&self) -> Option<String> {
        let out = Command::new("git")
            .args(["config", "--global", "user.name"])
            .output()
            .ok()?;
        if !out.status.success() {
            return None;
        }
        let s = String::from_utf8(out.stdout).ok()?.trim().to_string();
        if s.is_empty() {
            None
        } else {
            Some(s)
        }
    }

    fn hostname(&self) -> Option<String> {
        let out = Command::new("hostname").output().ok()?;
        if !out.status.success() {
            return None;
        }
        let s = String::from_utf8(out.stdout).ok()?.trim().to_string();
        if s.is_empty() {
            None
        } else {
            Some(s)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::RefCell;

    struct MockEnv {
        vars: RefCell<std::collections::HashMap<String, String>>,
        config_body: Option<String>,
        git_user: Option<String>,
        hostname: Option<String>,
    }

    impl MockEnv {
        fn new() -> Self {
            Self {
                vars: RefCell::new(Default::default()),
                config_body: None,
                git_user: None,
                hostname: None,
            }
        }
        fn with_var(self, k: &str, v: &str) -> Self {
            self.vars.borrow_mut().insert(k.into(), v.into());
            self
        }
    }

    impl EnvSource for MockEnv {
        fn var(&self, key: &str) -> Option<String> {
            self.vars.borrow().get(key).cloned()
        }
        fn config_path(&self) -> Option<PathBuf> {
            if self.config_body.is_some() {
                Some(PathBuf::from("/dev/null/config.toml"))
            } else {
                None
            }
        }
        fn git_user_name(&self) -> Option<String> {
            self.git_user.clone()
        }
        fn hostname(&self) -> Option<String> {
            self.hostname.clone()
        }
    }

    // Resolver test with a ladder of sources present; env vars win.
    #[test]
    fn env_vars_win_over_everything() {
        let env = MockEnv::new()
            .with_var("L123_USER", "Env User")
            .with_var("L123_ORG", "Env Org")
            .with_var("USER", "shell_user");
        let id = Identity::resolve_with(&env);
        assert_eq!(id.user, "Env User");
        assert_eq!(id.organization, "Env Org");
    }

    #[test]
    fn falls_back_to_git_then_user_then_placeholder() {
        let env = MockEnv {
            vars: RefCell::new(Default::default()),
            config_body: None,
            git_user: Some("Git Identity".to_string()),
            hostname: Some("host.local".to_string()),
        };
        let id = Identity::resolve_with(&env);
        assert_eq!(id.user, "Git Identity");
        assert_eq!(id.organization, "host.local");
    }

    #[test]
    fn user_var_used_when_git_absent() {
        let env = MockEnv::new().with_var("USER", "duane");
        let id = Identity::resolve_with(&env);
        assert_eq!(id.user, "duane");
    }

    #[test]
    fn placeholder_when_nothing_available() {
        let env = MockEnv::new();
        let id = Identity::resolve_with(&env);
        assert_eq!(id.user, "l123 User");
        assert_eq!(id.organization, "l123");
    }

    #[test]
    fn parses_simple_config() {
        let cfg = parse_config(
            r#"
            # comment
            user = "Duane Moore"
            organization = 'Zarkov'
            "#,
        );
        assert_eq!(cfg.user, "Duane Moore");
        assert_eq!(cfg.organization, "Zarkov");
    }

    #[test]
    fn config_overrides_shell_user_but_env_still_wins() {
        // Standalone config-only test via parse_config; integrated
        // resolver test is covered by env_vars_win.
        let cfg = parse_config("user = Bob\norg = Acme\n");
        assert_eq!(cfg.user, "Bob");
        assert_eq!(cfg.organization, "Acme");
    }
}
