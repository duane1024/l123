# L123.CNF — configuration reference

l123 reads its runtime settings from four sources, in this order of
precedence (first non-empty wins for each key):

1. **Environment variables** — quick per-invocation overrides.
2. **`~/.l123/L123.CNF`** — per-user defaults. Optional.
3. **Derived** — for `user` and `organization` only: `git config --global
   user.name` / `$USER` / `$LOGNAME` for user, and `hostname` for org.
4. **Built-in defaults** — `"l123 User"` and `"l123"`.

The `log_file` and `log_filter` keys skip step 3 — they have no sensible
derivation, so they go straight from env/file to the built-in default
(empty string, which means "don't install a logging subscriber at all").

---

## ✦ Seeing what's effective

```bash
l123 config
```

Prints every key with its current value, the source it came from, and
the matching environment variable. Example output:

```
Config file: /Users/you/.l123/L123.CNF (loaded)

  user           = Duane Moore                      [file]     env: L123_USER
  organization   = Acme                             [file]     env: L123_ORG
  log_file       = /tmp/l123.log                    [env]      env: L123_LOG
  log_filter     = <unset>                          [default]  env: RUST_LOG
```

Source labels:

| Label       | Meaning                                             |
|-------------|-----------------------------------------------------|
| `env`       | Read from the environment variable.                 |
| `file`      | Read from `~/.l123/L123.CNF`.                       |
| `derived`   | Computed from `git`, `$USER`, or `hostname`.        |
| `default`   | Falling back to the built-in placeholder.           |

---

## ✦ Creating the file

```bash
l123 config --init
```

Writes an annotated sample `~/.l123/L123.CNF` if one doesn't already
exist. Every key is commented-out — uncomment the ones you want to set.
To overwrite an existing file, add `--force`.

---

## ✦ Syntax

`L123.CNF` is a simple `key = value` file. No nested tables, no arrays.

- One key per line.
- Values may be bare, `"double-quoted"`, or `'single-quoted'`.
- `#` introduces a comment; trailing comments on a value line are also
  stripped.
- Unknown keys are silently ignored, so it's safe to leave notes in the
  file.

Example:

```
user         = "Duane Moore"
organization = Acme                 # trailing comment ok
log_file     = /var/log/l123.log
log_filter   = l123=debug,ironcalc=info
```

---

## ✦ Keys

### `user`

Name shown on the startup splash and in any "user"-labeled fields.

- **Env:** `L123_USER`
- **Aliases in file:** `name`, `user_name`
- **Derived from:** `git config --global user.name`, then `$USER`, then
  `$LOGNAME`
- **Default:** `l123 User`

### `organization`

Organization shown on the startup splash.

- **Env:** `L123_ORG`
- **Aliases in file:** `org`
- **Derived from:** `hostname`
- **Default:** `l123`

### `log_file`

Path to append tracing logs to. When unset (empty string), no logging
subscriber is installed — `tracing::*!` macros compile down to no-ops,
so there is zero runtime overhead.

- **Env:** `L123_LOG`
- **Aliases in file:** `log`
- **Default:** *unset*

Logs are written in plain text, no ANSI colors. The parent directory
is created if missing.

### `log_filter`

[`tracing_subscriber::EnvFilter`][env-filter] directive. Only applied
when `log_file` is set.

- **Env:** `RUST_LOG`
- **Aliases in file:** `rust_log`
- **Default:** `info` (when `log_file` is set)

Examples:

- `info` — everything at info level and above
- `l123=debug` — debug for l123 crates, default for others
- `l123=trace,ironcalc=info` — mix levels per crate

[env-filter]: https://docs.rs/tracing-subscriber/latest/tracing_subscriber/filter/struct.EnvFilter.html

---

## ✦ File location

| Path                      | Meaning                                   |
|---------------------------|-------------------------------------------|
| `~/.l123/L123.CNF`        | The one and only config path l123 reads.  |

If `$HOME` is unset, no config file is consulted — l123 relies entirely
on env vars and defaults. `l123 config` will print `<$HOME not set>` in
that case.

There is intentionally no per-directory `L123.CNF` lookup: running the
same binary in different directories should behave the same.

---

## ✦ Overriding per-invocation

Any env var beats the file, so one-off runs are easy:

```bash
L123_LOG=/tmp/debug.log RUST_LOG=l123=trace l123 sheet.xlsx
L123_USER="Demo Account" l123
```
