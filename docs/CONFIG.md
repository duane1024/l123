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

### `error_beep`

Soft terminal bell (the ASCII BEL character, `\x07`) fired when the
pointer hits an edge of the sheet — going up from row 1, left from
column A, and the equivalent on the trailing edges. Terminal
preferences decide whether that rings, flashes, or is silently
ignored, so "soft" here just means we defer to the user's terminal.

Toggle at runtime with `/Worksheet Global Default Other Beep
Enable|Disable`.

- **Env:** `L123_BEEP`
- **Aliases in file:** `beep`
- **Accepted values:** `true/false`, `on/off`, `yes/no`, `1/0`
  (case-insensitive)
- **Default:** `true`

### `theme`

Chrome theme. Affects the status line, menu / F1-help selection
highlights, splash field, mode indicator, and the dim gridline glyph
painted by `:Display Options Grid Yes`. Cell colors that come from
the document — `:Format Color`, xlsx fills/fonts, sheet tab tints —
are unaffected; the theme only paints chrome, never data.

- **Env:** `L123_THEME`
- **CLI:** `--theme <name>` (overrides env / file for the run)
- **Accepted values:**
  - `dos` (default) — classic 1-2-3 R3.4a DOS look.
  - `wysiwyg` — R3.4a WYSIWYG paper look (magenta gridlines, blue
    selection).
  - `amber` — CRT amber phosphor.
  - `green` — CRT green phosphor.
  - Aliases: `default` / `classic` → `dos`; `paper` → `wysiwyg`.
- **Bad values:** silently ignored in `L123_THEME` and the config
  file (the next-lower tier applies). The CLI surfaces typos as a
  usage error and exits 2.
- **Default:** `dos`

There is no runtime command for switching themes mid-session — set
the value before launching.

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
l123 --theme amber sheet.xlsx
```
