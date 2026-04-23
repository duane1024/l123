# AGENTS.md — L123 project guide

Session-level conventions for working on L123 (a Lotus 1-2-3 R3.1 TUI clone).
Living document — edit freely as conventions emerge.

## Canonical docs — read these before making decisions

- `docs/SPEC.md` — what L123 *is*. Scope, modes, menu tree, authenticity contract.
- `docs/PLAN.md` — implementation milestones, risk register, test strategy.
- `docs/MENU.md` — complete menu tree; source of truth for `l123-menu`.

If code and doc disagree, fix the doc first, then write code to match.

---

## How we develop: red / green / refactor (strict TDD)

**Test-first, always.** Every change to shipping code follows this loop:

### 1. Red
Write the failing test(s) that describe the behavior you want. Run them.
They must **fail for the right reason** — not a compile error, not a
panic from unrelated setup. `cargo test -p <crate>` shows a clean
failure in the new test, pointing at the behavior being specified.

If you can't figure out what the test should assert, the behavior isn't
specified enough yet. Re-read SPEC, or write the acceptance transcript
first, then come back.

### 2. Green
Write the **minimum** code to make the tests pass. No refactoring, no
extra features, no speculative abstractions, no helpers you "might need
later." If a test is green before you meant it to be, delete the extra
code until only the test's claim is covered.

### 3. Refactor
Now, and only now, clean up. Re-run tests after every small refactor
step. Never refactor with red tests. Clippy must stay clean:
`cargo clippy --workspace --all-targets -- -D warnings`.

### Exceptions (be honest — they should be rare)
- Pure docs/comment changes.
- Throwaway spikes — must be marked `WIP: spike` and deleted before
  landing. Any code you keep must have tests written before it.

---

## Test tiers

Every feature lands with tests at one or more of:

1. **Unit tests** — colocated with source (`#[cfg(test)] mod tests`).
   Pure-function tests; no ratatui, no IronCalc.
2. **Engine integration tests** — `crates/l123-engine/src/ironcalc_adapter.rs`
   under `#[cfg(test)]`. Drive the `Engine` trait against real IronCalc.
3. **Acceptance transcripts** — `tests/acceptance/*.tsv`, run by
   `crates/l123-ui/tests/acceptance.rs`. Keystrokes in, on-screen state
   out. **Every item in SPEC §20 "Authenticity Contract" gets ≥1
   transcript** before the claim can be called done.

UI-touching features default to **both** a unit test on the state
transition **and** an acceptance transcript covering the keystroke
experience.

### Adding an acceptance transcript (red step)

1. Write the `.tsv` under `tests/acceptance/` — see the directive syntax
   in `tests/acceptance/README.md`.
2. Register it in the `transcripts!` macro at the bottom of
   `crates/l123-ui/tests/acceptance.rs`.
3. `cargo test -p l123-ui --test acceptance` → the new test fails.
4. Implement until green.

---

## Common commands

| Task | Command |
|---|---|
| Full build | `cargo build` |
| Full tests | `cargo test --workspace` |
| One crate | `cargo test -p l123-core` |
| Acceptance only | `cargo test -p l123-ui --test acceptance` |
| Lint (warnings = errors) | `cargo clippy --workspace --all-targets -- -D warnings` |
| Format | `cargo fmt --all` |
| Run the binary | `cargo run -p l123` |

CI gate (once added): fmt clean + clippy -D warnings + all tests pass.

---

## Crate layering — do not violate

```
l123-core  ← zero external deps (types only)
  ↑
l123-parse, l123-menu    (pure layers on core)
  ↑
l123-engine             (core + IronCalc)
  ↑
l123-cmd, l123-io        (core + engine)
  ↑
l123-macro              (core + cmd)
  ↑
l123-ui                 (core + menu; no direct engine dep — UI is engine-agnostic)
  ↑
l123 (binary)           (wires ui + engine + cmd + io)
```

No upward edges. If a type needs to cross layers, it probably belongs in
`l123-core`. IronCalc types never leak above `l123-engine`.

---

## Project-specific conventions

- `thiserror` for library error enums; `anyhow` only at binary edges.
- **0-based addressing** internally (`Address { col: u16, row: u32, sheet }`).
  IronCalc's 1-based coords live inside the adapter only.
- **1-2-3 syntax → Excel syntax** translation happens in `l123-parse`
  before the engine is called. The engine sees `SUM(A1:B2)`, never
  `@SUM(A1..B2)`.
- When an enum has an obvious "zero" value, derive `Default` with
  `#[default]` on the variant — no manual `impl Default`.
- No comments that restate the code. SPEC explains *why*; code explains
  *what*. A comment earns its place when it documents a non-obvious
  constraint or a workaround.
- **`.WK3` / `.wk3`** refers to the **legacy Lotus 1-2-3 R3 file format**
  (an external format the project reads/writes). Do not rename these
  references when refactoring — they are not project-name references.

---

## What this file is NOT

- A duplicate of SPEC.md. For *what* L123 does, read SPEC.
- A to-do list. Use the Task tool and PLAN.md milestones.

---

## Updating AGENTS.md

Grow this file when:
- A convention emerges you had to re-learn twice.
- A pitfall bites you (IronCalc quirk, ratatui gotcha, terminal edge case).
- A milestone lands that changes how the test tiers work.

Keep it under ~200 lines. Prune aggressively. Terse beats complete.
