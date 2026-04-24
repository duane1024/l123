//! Runtime configuration resolver.
//!
//! Lookup order for each key (first non-empty wins):
//!   1. Environment variable.
//!   2. `~/.l123/L123.CNF` config file.
//!   3. Built-in default.
//!
//! The config file uses the same minimal `key = "value"` syntax as
//! historical DOS CNFs: one key per line, `#` introduces a comment,
//! quotes are optional. No nested tables, no arrays — keep it boring
//! so users can edit it with any editor.

use std::path::{Path, PathBuf};

/// Origin of a resolved setting — shown by `l123 config` so users can
/// tell *why* a value is what it is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Source {
    Default,
    File,
    Env,
    /// Derived from an OS query — `git config user.name`, the `$USER`
    /// shell var, or `hostname`. Only `user` and `organization` can
    /// end up with this source; `log_file` / `log_filter` cannot.
    Derived,
}

impl Source {
    pub fn label(self) -> &'static str {
        match self {
            Source::Default => "default",
            Source::File => "file",
            Source::Env => "env",
            Source::Derived => "derived",
        }
    }
}

/// A single resolved setting with provenance.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Setting {
    pub value: String,
    pub source: Source,
}

/// Full resolved configuration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Config {
    pub user: Setting,
    pub organization: Setting,
    pub log_file: Setting,
    pub log_filter: Setting,
    /// Path that was searched for the config file (always reported,
    /// even if the file doesn't exist).
    pub file_path: Option<PathBuf>,
    /// Whether `file_path` existed and was readable.
    pub file_found: bool,
}

/// Key metadata used to render `l123 config` and the sample file.
pub struct KeyInfo {
    pub key: &'static str,
    pub env_var: Option<&'static str>,
    pub description: &'static str,
}

pub const KEYS: &[KeyInfo] = &[
    KeyInfo {
        key: "user",
        env_var: Some("L123_USER"),
        description: "Name shown on the startup splash and in headers.",
    },
    KeyInfo {
        key: "organization",
        env_var: Some("L123_ORG"),
        description: "Organization shown on the startup splash (alias: org).",
    },
    KeyInfo {
        key: "log_file",
        env_var: Some("L123_LOG"),
        description: "Path to append tracing logs to. Empty = no logging.",
    },
    KeyInfo {
        key: "log_filter",
        env_var: Some("RUST_LOG"),
        description: "tracing EnvFilter directive (e.g. `l123=debug`).",
    },
];

/// Things the resolver needs from the outside world. Abstracted so
/// tests can drive it without touching real env vars or filesystem.
pub trait ConfigSource {
    fn var(&self, key: &str) -> Option<String>;
    /// Default config file path. `None` if `$HOME` is unset.
    fn config_path(&self) -> Option<PathBuf>;
    fn read_file(&self, path: &Path) -> Option<String>;
    /// `git config --global user.name` output; `None` if unavailable.
    fn git_user_name(&self) -> Option<String>;
    /// System hostname; `None` if unavailable.
    fn hostname(&self) -> Option<String>;
}

pub struct StdSource;

impl ConfigSource for StdSource {
    fn var(&self, key: &str) -> Option<String> {
        std::env::var(key).ok()
    }
    fn config_path(&self) -> Option<PathBuf> {
        default_config_path()
    }
    fn read_file(&self, path: &Path) -> Option<String> {
        std::fs::read_to_string(path).ok()
    }
    fn git_user_name(&self) -> Option<String> {
        let out = std::process::Command::new("git")
            .args(["config", "--global", "user.name"])
            .output()
            .ok()?;
        if !out.status.success() {
            return None;
        }
        let s = String::from_utf8(out.stdout).ok()?.trim().to_string();
        (!s.is_empty()).then_some(s)
    }
    fn hostname(&self) -> Option<String> {
        let out = std::process::Command::new("hostname").output().ok()?;
        if !out.status.success() {
            return None;
        }
        let s = String::from_utf8(out.stdout).ok()?.trim().to_string();
        (!s.is_empty()).then_some(s)
    }
}

/// `~/.l123/L123.CNF`, or `None` if `$HOME` is not set.
pub fn default_config_path() -> Option<PathBuf> {
    let home = std::env::var_os("HOME")?;
    Some(PathBuf::from(home).join(".l123").join("L123.CNF"))
}

impl Config {
    pub fn resolve() -> Self {
        Self::resolve_with(&StdSource)
    }

    pub fn resolve_with(src: &dyn ConfigSource) -> Self {
        let file_path = src.config_path();
        let file_body = file_path.as_deref().and_then(|p| src.read_file(p));
        let file_found = file_body.is_some();
        let file = file_body
            .as_deref()
            .map(parse_config_body)
            .unwrap_or_default();

        let user = pick_with_derive(src, "L123_USER", file.user.as_deref(), "l123 User", || {
            src.git_user_name()
                .or_else(|| src.var("USER"))
                .or_else(|| src.var("LOGNAME"))
        });
        let organization = pick_with_derive(
            src,
            "L123_ORG",
            file.organization.as_deref(),
            "l123",
            || src.hostname(),
        );
        let log_file = pick(src, "L123_LOG", file.log_file.as_deref(), "");
        let log_filter = pick(src, "RUST_LOG", file.log_filter.as_deref(), "");

        Config {
            user,
            organization,
            log_file,
            log_filter,
            file_path,
            file_found,
        }
    }

    /// Effective value for `log_file`, or `None` when unset.
    pub fn log_file_path(&self) -> Option<PathBuf> {
        if self.log_file.value.is_empty() {
            None
        } else {
            Some(PathBuf::from(&self.log_file.value))
        }
    }

    /// Render the human-facing `l123 config` table.
    pub fn render_table(&self) -> String {
        let mut out = String::new();
        match &self.file_path {
            Some(p) if self.file_found => {
                out.push_str(&format!("Config file: {} (loaded)\n", p.display()))
            }
            Some(p) => out.push_str(&format!("Config file: {} (not found)\n", p.display())),
            None => out.push_str("Config file: <$HOME not set>\n"),
        }
        out.push('\n');
        for info in KEYS {
            let setting = self.get(info.key).expect("KEYS matches struct fields");
            let display = if setting.value.is_empty() {
                "<unset>".to_string()
            } else {
                setting.value.clone()
            };
            let env = info.env_var.unwrap_or("");
            out.push_str(&format!(
                "  {:<14} = {:<32} [{}]  env: {}\n",
                info.key,
                display,
                setting.source.label(),
                env,
            ));
        }
        out
    }

    fn get(&self, key: &str) -> Option<&Setting> {
        match key {
            "user" => Some(&self.user),
            "organization" => Some(&self.organization),
            "log_file" => Some(&self.log_file),
            "log_filter" => Some(&self.log_filter),
            _ => None,
        }
    }
}

fn pick(src: &dyn ConfigSource, env_var: &str, file: Option<&str>, default: &str) -> Setting {
    if let Some(v) = src.var(env_var).filter(|s| !s.is_empty()) {
        return Setting {
            value: v,
            source: Source::Env,
        };
    }
    if let Some(v) = file.filter(|s| !s.is_empty()) {
        return Setting {
            value: v.to_string(),
            source: Source::File,
        };
    }
    Setting {
        value: default.to_string(),
        source: Source::Default,
    }
}

/// Like `pick`, but if env + file are both empty, run `derive` before
/// falling back to the hard-coded default. Used for `user` / `org` so
/// an unconfigured user sees their git identity / hostname rather
/// than the bland "l123 User" / "l123" placeholders.
fn pick_with_derive(
    src: &dyn ConfigSource,
    env_var: &str,
    file: Option<&str>,
    default: &str,
    derive: impl FnOnce() -> Option<String>,
) -> Setting {
    if let Some(v) = src.var(env_var).filter(|s| !s.is_empty()) {
        return Setting {
            value: v,
            source: Source::Env,
        };
    }
    if let Some(v) = file.filter(|s| !s.is_empty()) {
        return Setting {
            value: v.to_string(),
            source: Source::File,
        };
    }
    if let Some(v) = derive().filter(|s| !s.is_empty()) {
        return Setting {
            value: v,
            source: Source::Derived,
        };
    }
    Setting {
        value: default.to_string(),
        source: Source::Default,
    }
}

/// Plain-data view of what a config file contained. Fields are
/// `None` when the key wasn't present (distinct from `Some("")`,
/// which we treat as a user explicitly clearing a value).
#[derive(Debug, Default, PartialEq, Eq)]
pub struct ConfigFile {
    pub user: Option<String>,
    pub organization: Option<String>,
    pub log_file: Option<String>,
    pub log_filter: Option<String>,
}

/// Parse `key = value` lines. Accepts `"..."`, `'...'`, or bare values.
/// Lines starting with `#` or empty lines are ignored. Unknown keys
/// are silently skipped so old/new configs can coexist.
pub fn parse_config_body(body: &str) -> ConfigFile {
    let mut out = ConfigFile::default();
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
            "user" | "name" | "user_name" => out.user = Some(value),
            "org" | "organization" => out.organization = Some(value),
            "log_file" | "log" => out.log_file = Some(value),
            "log_filter" | "rust_log" => out.log_filter = Some(value),
            _ => {}
        }
    }
    out
}

/// Annotated sample file written by `l123 config --init`.
pub const SAMPLE_CNF: &str = "\
# L123.CNF — configuration for l123
#
# Syntax: key = value. Values may be bare, \"double-quoted\", or
# 'single-quoted'. Lines starting with # are comments. Unknown
# keys are ignored, so it's safe to leave notes here.
#
# Every key may be overridden by the matching environment variable;
# see `l123 config` for the current effective value and source.

# Name shown on the startup splash. Env: L123_USER.
# user = \"Your Name\"

# Organization shown on the startup splash (alias: org). Env: L123_ORG.
# organization = \"Your Company\"

# Path to append tracing logs to. Empty or unset disables logging
# entirely (zero overhead). Env: L123_LOG.
# log_file = \"~/.l123/l123.log\"

# tracing EnvFilter directive, e.g. `l123=debug,ironcalc=info`.
# Defaults to `info` when log_file is set. Env: RUST_LOG.
# log_filter = \"info\"
";

#[cfg(test)]
mod tests {
    use super::*;
    use std::cell::RefCell;
    use std::collections::HashMap;

    struct MockSource {
        vars: RefCell<HashMap<String, String>>,
        config_body: Option<String>,
        config_path: Option<PathBuf>,
        git_user: Option<String>,
        hostname: Option<String>,
    }

    impl MockSource {
        fn new() -> Self {
            Self {
                vars: RefCell::new(HashMap::new()),
                config_body: None,
                config_path: Some(PathBuf::from("/fake/.l123/L123.CNF")),
                git_user: None,
                hostname: None,
            }
        }
        fn with_var(self, k: &str, v: &str) -> Self {
            self.vars.borrow_mut().insert(k.into(), v.into());
            self
        }
        fn with_file(mut self, body: &str) -> Self {
            self.config_body = Some(body.into());
            self
        }
        fn without_home(mut self) -> Self {
            self.config_path = None;
            self
        }
        fn with_git_user(mut self, name: &str) -> Self {
            self.git_user = Some(name.into());
            self
        }
        fn with_hostname(mut self, h: &str) -> Self {
            self.hostname = Some(h.into());
            self
        }
    }

    impl ConfigSource for MockSource {
        fn var(&self, key: &str) -> Option<String> {
            self.vars.borrow().get(key).cloned()
        }
        fn config_path(&self) -> Option<PathBuf> {
            self.config_path.clone()
        }
        fn read_file(&self, _path: &Path) -> Option<String> {
            self.config_body.clone()
        }
        fn git_user_name(&self) -> Option<String> {
            self.git_user.clone()
        }
        fn hostname(&self) -> Option<String> {
            self.hostname.clone()
        }
    }

    #[test]
    fn defaults_when_no_env_and_no_file() {
        let src = MockSource::new().without_home();
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "l123 User");
        assert_eq!(cfg.user.source, Source::Default);
        assert_eq!(cfg.organization.value, "l123");
        assert_eq!(cfg.log_file.value, "");
        assert_eq!(cfg.log_file.source, Source::Default);
        assert!(!cfg.file_found);
    }

    #[test]
    fn env_beats_file_beats_default() {
        let src = MockSource::new()
            .with_file("user = \"From File\"\norganization = FileOrg\n")
            .with_var("L123_USER", "From Env");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "From Env");
        assert_eq!(cfg.user.source, Source::Env);
        assert_eq!(cfg.organization.value, "FileOrg");
        assert_eq!(cfg.organization.source, Source::File);
    }

    #[test]
    fn log_file_and_filter_round_trip() {
        let src = MockSource::new().with_file("log_file = /tmp/l.log\nlog_filter = l123=trace\n");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.log_file.value, "/tmp/l.log");
        assert_eq!(cfg.log_file.source, Source::File);
        assert_eq!(
            cfg.log_file_path().as_deref(),
            Some(Path::new("/tmp/l.log")),
        );
        assert_eq!(cfg.log_filter.value, "l123=trace");
    }

    #[test]
    fn empty_env_var_does_not_shadow_file() {
        let src = MockSource::new()
            .with_file("user = Duane\n")
            .with_var("L123_USER", "");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "Duane");
        assert_eq!(cfg.user.source, Source::File);
    }

    #[test]
    fn parse_accepts_quoted_and_bare_and_comments() {
        let f = parse_config_body(
            r#"
            # this is a comment
            user = "Duane Moore"   # trailing comment
            org  = 'Zarkov'
            log_file = /var/log/l.log
            "#,
        );
        assert_eq!(f.user.as_deref(), Some("Duane Moore"));
        assert_eq!(f.organization.as_deref(), Some("Zarkov"));
        assert_eq!(f.log_file.as_deref(), Some("/var/log/l.log"));
        assert_eq!(f.log_filter, None);
    }

    #[test]
    fn parse_ignores_unknown_keys() {
        let f = parse_config_body("mystery = 42\nuser = Bob\n");
        assert_eq!(f.user.as_deref(), Some("Bob"));
    }

    #[test]
    fn log_filter_env_takes_precedence() {
        let src = MockSource::new()
            .with_file("log_filter = file_filter\n")
            .with_var("RUST_LOG", "env_filter");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.log_filter.value, "env_filter");
        assert_eq!(cfg.log_filter.source, Source::Env);
    }

    #[test]
    fn render_table_shows_path_and_every_key() {
        let src = MockSource::new()
            .with_file("user = Bob\n")
            .with_var("L123_LOG", "/tmp/x.log");
        let cfg = Config::resolve_with(&src);
        let out = cfg.render_table();
        assert!(out.contains("/fake/.l123/L123.CNF"), "path missing: {out}");
        assert!(out.contains("user"));
        assert!(out.contains("organization"));
        assert!(out.contains("log_file"));
        assert!(out.contains("log_filter"));
        assert!(out.contains("Bob"));
        assert!(out.contains("/tmp/x.log"));
        assert!(out.contains("[env]"));
        assert!(out.contains("[file]"));
    }

    #[test]
    fn render_table_notes_missing_home() {
        let src = MockSource::new().without_home();
        let out = Config::resolve_with(&src).render_table();
        assert!(out.contains("$HOME not set"), "got: {out}");
    }

    #[test]
    fn render_table_notes_missing_file() {
        let src = MockSource::new(); // path set but no body -> not found
        let out = Config::resolve_with(&src).render_table();
        assert!(out.contains("not found"), "got: {out}");
    }

    #[test]
    fn user_derives_from_git_when_env_and_file_empty() {
        let src = MockSource::new()
            .with_git_user("Git Person")
            .with_hostname("lab.local");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "Git Person");
        assert_eq!(cfg.user.source, Source::Derived);
        assert_eq!(cfg.organization.value, "lab.local");
        assert_eq!(cfg.organization.source, Source::Derived);
    }

    #[test]
    fn user_falls_back_to_user_env_when_no_git() {
        let src = MockSource::new().with_var("USER", "duane");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "duane");
        assert_eq!(cfg.user.source, Source::Derived);
    }

    #[test]
    fn file_beats_derive() {
        let src = MockSource::new()
            .with_file("user = From File\n")
            .with_git_user("Git Person");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.user.value, "From File");
        assert_eq!(cfg.user.source, Source::File);
    }

    #[test]
    fn log_file_has_no_derive_tier() {
        // If no env + no file, log_file is always Default (empty string),
        // never Derived — there's no sensible "derive" for it.
        let src = MockSource::new()
            .with_git_user("Git Person")
            .with_hostname("lab.local");
        let cfg = Config::resolve_with(&src);
        assert_eq!(cfg.log_file.value, "");
        assert_eq!(cfg.log_file.source, Source::Default);
    }

    #[test]
    fn sample_cnf_parses_back_as_all_commented() {
        // Every non-comment line in SAMPLE_CNF is actually commented,
        // so the parsed config should be entirely empty.
        let f = parse_config_body(SAMPLE_CNF);
        assert_eq!(f, ConfigFile::default());
    }
}
