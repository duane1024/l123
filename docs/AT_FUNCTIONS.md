# @Function catalog

Source of truth for the `@name` → IronCalc-name translation table in
`l123-parse`. Pairs with SPEC §15 (function surface) and §16
(authenticity). When SPEC and this file disagree, fix SPEC first.

IronCalc target names are the canonical Excel names IronCalc 0.7.1
exposes — verified against `ironcalc_base/src/language/language.json`
and the `Function` enum in `ironcalc_base/src/functions/mod.rs`.

## Status legend

| Status   | Meaning                                                                 |
|----------|-------------------------------------------------------------------------|
| 1:1      | `@NAME` and Excel `NAME` are identical; the parser's `@`-strip is enough. |
| rename   | Pure name swap in the translation table (no semantic change).           |
| arg-fix  | Same engine name but argument shape differs (e.g. 0-based vs 1-based).  |
| emulate  | No IronCalc equivalent — rewrite at parse time to a composition, or hook a custom evaluator. |

## Implementation status

Shipped in `l123-parse` (covered by parse unit tests + acceptance
transcripts under `tests/acceptance/function_*.tsv`):

- **Renames**: `@AVG @COUNT @STD @STDS @VAR @VARS @ISSTRING @LENGTH
  @REPEAT @COLS @CODE`, the database family
  `@DAVG @DSTD @DSTDS @DVAR @DVARS`, and `@@`.
- **Niladic-paren completion** for `@PI @NOW @TODAY @RAND @NA @TRUE
  @FALSE` (1-2-3 omits `()`; Excel requires it).
- **Arg-fix**: `@MID` 0→1-based start, `@FIND` 0→1-based start with
  result-minus-1, `@STRING` → `TEXT` with a built format string
  (static for literal `decimals`, dynamic via `IF`/`REPT` otherwise).
- **Emulations**: `@ERR` → `#VALUE!` literal, `@CTERM` and `@TERM` to
  log compositions, `@SUMPRODUCT` → `SUM` with array broadcast
  (verified IronCalc 0.7.1 `SUM` accepts `(A1:A3)*(B1:B3)`).

Deferred, with rationale:

- `@PROPER` — IronCalc 0.7.1 lacks `PROPER` and word-boundary
  capitalization isn't expressible in stock Excel functions. Needs
  an upstream PR or an engine-side shim.
- `@CHAR` — IronCalc lacks both `CHAR` and `UNICHAR`. Engine shim
  required.
- `@CELL` argument keyword map — keyword sets overlap heavily but
  not entirely (e.g. `"sheetname"` is R3-only). Will be a small
  per-keyword dispatch when prioritized.
- `@CELLPOINTER` — needs cursor-context the engine doesn't carry;
  belongs at the `l123-cmd` layer.
- `@COORD` — would need `ADDRESS`, which IronCalc 0.7.1 lacks.
- `@VDB @S @DQUERY @ISRANGE` — post-MVP niche; no clean
  composition.
- `@N` argument-shape rewrite — IronCalc's `N` accepts ranges and
  silently picks the first cell, matching 1-2-3 behavior closely
  enough that the rewrite is a no-op in practice. Revisit if a
  divergence shows up in a transcript.

Engine-adapter gap surfaced during PR3: IronCalc 0.7.1's `CellValue`
enum has no `Error` variant — errors come back as
`CellValue::String("#VALUE!")`. The current `IronCalcEngine` adapter
maps that to `Value::Text("#VALUE!")` instead of
`Value::Error(ErrKind::Value)`. As a result the cell renders the
literal `#VALUE!` text instead of the Lotus-style `ERR` tag. This
predates the function work; tracked as a follow-up on the engine
crate, not a function-translation bug.

## Reverse translation (engine → 1-2-3 source form)

`l123_parse::to_lotus_source` reverses the cosmetic subset so the
control panel and the cells cache stay authentic across save +
reload. Wired into `cell_view_to_contents` in `l123-ui`.

Reversed:

- All entries from `FN_RENAMES_BACK` (the inverse of the forward
  rename table where the swap is canonical — `LEN`/`LENGTH` is *not*
  reversed because `@LEN` is the canonical 1-2-3 form).
- Niladic-paren elision: `PI()` → `@PI`, `NOW()` → `@NOW`, etc.
- Range separator: `:` → `..`.
- Sheet refs: bare `Sheet1!A1` and quoted `'Q1 Sales'!B5` → `A:A1`
  / `A:B5` based on the engine's sheet-name list.
- `INDIRECT(...)` → `@@(...)`.
- `#VALUE!` literal → `@ERR`. Other error literals
  (`#DIV/0!`, `#REF!`, ...) pass through.
- The leading `@` IronCalc's xlsx codec adds for implicit
  intersection (e.g. `=@INDIRECT(...)` after a save round-trip) is
  treated as a no-op so it doesn't get re-emitted on top of the
  `@@` substitution.

Output gets a `@` prefix when the body starts with a function call,
or a `+` prefix otherwise (`A1+B1` → `+A1+B1`, `42` → `+42`).

Not reversed (irreversibly destroyed by the forward rewrite):

- Arg-fix: `MID(s, (b)+1, c)` does not reverse to `@MID(s, b, c)`.
  The `(...)+1` wrapper is detectable in principle but the round
  trip only works for L123-native files, and we don't currently
  carry sidecar metadata to know which `+1` came from us.
- Emulations: `LN(fv/pv)/LN(1+rate)` does not reverse to
  `@CTERM(...)`; `TEXT(n,"0.00")` does not reverse to
  `@STRING(n,2)`; `SUM((a)*(b))` does not reverse to `@SUMPRODUCT`.
  These show in their decomposed form on the panel after reload.
- 3D ranges: a `@SUM(A:B3..C:B3)` saved as
  `=SUM(Sheet1!B3:B3,Sheet2!B3:B3,Sheet3!B3:B3)` reloads as the
  comma-expanded form rather than collapsing back to the 3D
  shorthand.

Restoring the irreversible cases would require sidecar metadata
(e.g. an xlsx custom property keyed by cell address storing the
original Lotus source). Out of scope for the cosmetic round-trip.

---

## MVP set (SPEC §15) — must ship in M2

### Math

| @-name   | IronCalc | Status | Notes |
|----------|----------|--------|-------|
| `@ABS`   | `ABS`    | 1:1    |       |
| `@INT`   | `INT`    | 1:1    | Both truncate toward −∞. |
| `@MOD`   | `MOD`    | 1:1    |       |
| `@ROUND` | `ROUND`  | 1:1    |       |
| `@SQRT`  | `SQRT`   | 1:1    |       |
| `@EXP`   | `EXP`    | 1:1    |       |
| `@LN`    | `LN`     | 1:1    |       |
| `@LOG`   | `LOG`    | 1:1    | 1-2-3 `@LOG` is base-10; Excel `LOG` defaults to base-10 with one arg. |
| `@RAND`  | `RAND`   | 1:1    |       |
| `@PI`    | `PI`     | 1:1    | 1-2-3 `@PI` takes no parens; parser already strips `@`. Verify `=PI()` is what reaches the engine, not `=PI`. |

### Trig

| @-name   | IronCalc | Status | Notes |
|----------|----------|--------|-------|
| `@SIN`   | `SIN`    | 1:1    |       |
| `@COS`   | `COS`    | 1:1    |       |
| `@TAN`   | `TAN`    | 1:1    |       |
| `@ASIN`  | `ASIN`   | 1:1    |       |
| `@ACOS`  | `ACOS`   | 1:1    |       |
| `@ATAN`  | `ATAN`   | 1:1    |       |
| `@ATAN2` | `ATAN2`  | arg-fix | 1-2-3 is `@ATAN2(x, y)`; Excel is `ATAN2(x, y)` — same order, same semantics. Confirmed 1:1. |

### Stats

| @-name   | IronCalc   | Status  | Notes |
|----------|------------|---------|-------|
| `@SUM`   | `SUM`      | 1:1     |       |
| `@AVG`   | `AVERAGE`  | rename  |       |
| `@COUNT` | `COUNTA`   | rename  | 1-2-3 `@COUNT` counts non-empty cells (incl. labels); Excel `COUNT` only counts numbers. `COUNTA` matches 1-2-3 semantics. |
| `@MAX`   | `MAX`      | 1:1     |       |
| `@MIN`   | `MIN`      | 1:1     |       |
| `@STD`   | `STDEV.P`  | rename  | 1-2-3 `@STD` is population stdev. |
| `@STDS`  | `STDEV.S`  | rename  | Sample stdev (R3+ only).         |
| `@VAR`   | `VAR.P`    | rename  | Population variance.             |
| `@VARS`  | `VAR.S`    | rename  | Sample variance (R3+ only).      |

### Logical

| @-name      | IronCalc   | Status  | Notes |
|-------------|------------|---------|-------|
| `@IF`       | `IF`       | 1:1     |       |
| `@TRUE`     | `TRUE`     | 1:1     | Returns 1 in 1-2-3, but `TRUE`'s coercion to 1 in arithmetic matches. |
| `@FALSE`    | `FALSE`    | 1:1     |       |
| `@NA`       | `NA`       | 1:1     |       |
| `@ERR`      | —          | emulate | No Excel equivalent. Rewrite to a literal `#VALUE!` error or hook a custom evaluator that returns `Value::Error(ErrKind::Err)`. |
| `@ISERR`    | `ISERR`    | 1:1     |       |
| `@ISNA`     | `ISNA`     | 1:1     |       |
| `@ISNUMBER` | `ISNUMBER` | 1:1     |       |
| `@ISSTRING` | `ISTEXT`   | rename  |       |

### String

| @-name        | IronCalc | Status  | Notes |
|---------------|----------|---------|-------|
| `@LEN`        | `LEN`    | 1:1     |       |
| `@LENGTH`     | `LEN`    | rename  | Synonym for `@LEN` in some R3 docs. |
| `@LEFT`       | `LEFT`   | 1:1     |       |
| `@RIGHT`      | `RIGHT`  | 1:1     |       |
| `@MID`        | `MID`    | arg-fix | 1-2-3 `@MID(s, start, count)` is **0-based**; Excel `MID` is **1-based**. Translator must add 1 to the start arg. |
| `@UPPER`      | `UPPER`  | 1:1     |       |
| `@LOWER`      | `LOWER`  | 1:1     |       |
| `@PROPER`     | —        | emulate | IronCalc 0.7.1 has no `PROPER`. Either land an upstream PR, or compose: `UPPER(LEFT(s,1)) & LOWER(MID(s,2,LEN(s)-1))` doesn't handle word boundaries — needs a real shim. |
| `@TRIM`       | `TRIM`   | 1:1     | Both collapse runs of internal whitespace. |
| `@REPEAT`     | `REPT`   | rename  |       |
| `@FIND`       | `FIND`   | arg-fix | 1-2-3 `@FIND(needle, haystack, start)` is **0-based** and returns **0-based** index. Excel is **1-based** in both directions. Adjust args + result. |
| `@EXACT`      | `EXACT`  | 1:1     |       |
| `@STRING`     | `TEXT`   | rename  | 1-2-3 `@STRING(n, decimals)` → `TEXT(n, "0.000…")` — must build the Excel format string from the decimal count. |
| `@VALUE`      | `VALUE`  | 1:1     |       |
| `@CHAR`       | —        | emulate | IronCalc 0.7.1 lacks `CHAR`/`UNICHAR`. Shim required. |
| `@CODE`       | —        | emulate | IronCalc has `UNICODE` (returns codepoint of first char) — close, but `@CODE` is LICS in R3 and ASCII elsewhere; pick one. |

### Date/Time — all 1:1

`@DATE @DATEVALUE @DAY @MONTH @YEAR @NOW @TODAY @TIME @TIMEVALUE
@HOUR @MINUTE @SECOND` all map to identically named IronCalc functions.

Caveat: 1-2-3 epoch is **1900-01-01 = 1**; Excel epoch is the same but
treats 1900 as a leap year (the legacy `Feb 29 1900` bug). IronCalc
inherits Excel's behavior. Round-trip serial dates ≥ 1900-03-01 match;
dates before need a transcript test before claiming parity.

### Financial

| @-name   | IronCalc | Status  | Notes |
|----------|----------|---------|-------|
| `@PMT`   | `PMT`    | 1:1     |       |
| `@PV`    | `PV`     | 1:1     |       |
| `@FV`    | `FV`     | 1:1     |       |
| `@NPV`   | `NPV`    | 1:1     |       |
| `@IRR`   | `IRR`    | 1:1     |       |
| `@RATE`  | `RATE`   | 1:1     |       |
| `@CTERM` | —        | emulate | Periods to grow PV→FV at fixed rate: `LN(fv/pv)/LN(1+rate)`. Pure rewrite. |
| `@TERM`  | —        | emulate | Periods of payments: `LN(1+(fv*rate)/pmt)/LN(1+rate)`. Pure rewrite. |
| `@SLN`   | `SLN`    | 1:1     |       |
| `@SYD`   | `SYD`    | 1:1     |       |
| `@DDB`   | `DDB`    | 1:1     |       |

### Lookup — all 1:1

`@VLOOKUP @HLOOKUP @INDEX @CHOOSE` map cleanly. Verify
`@VLOOKUP`'s match-type default (1-2-3 = exact for text, range-match
for numbers) against Excel's `range_lookup` arg in transcripts.

### Reference

| @-name         | IronCalc   | Status  | Notes |
|----------------|------------|---------|-------|
| `@CELL`        | `CELL`     | arg-fix | Argument keywords differ (`"contents"` vs `"value"`, etc.). Build a small dispatch in the parser. |
| `@CELLPOINTER` | —          | emulate | Same args as `@CELL` but acts on the current cursor — needs UI-layer context the engine doesn't have. Resolve at the `l123-cmd` layer, not in the parser. |
| `@@`           | `INDIRECT` | rename  | The translator must recognize `@@(ref)` as a function call, not strip both `@`s. |
| `@ROWS`        | `ROWS`     | 1:1     |       |
| `@COLS`        | `COLUMNS`  | rename  |       |
| `@SHEETS`      | `SHEETS`   | 1:1     |       |
| `@COORD`       | —          | emulate | Builds an address string from `(sheet, col, row, abs)` ints. Pure rewrite to `ADDRESS` + sheet-name lookup, but Excel's `ADDRESS` doesn't take a sheet *index*. |

---

## Post-MVP — named in SPEC §15

| @-name        | IronCalc      | Status  | Notes |
|---------------|---------------|---------|-------|
| `@DSUM`       | `DSUM`        | 1:1     |       |
| `@DCOUNT`     | `DCOUNT`      | 1:1     |       |
| `@DAVG`       | `DAVERAGE`    | rename  |       |
| `@DMAX`       | `DMAX`        | 1:1     |       |
| `@DMIN`       | `DMIN`        | 1:1     |       |
| `@DSTD`       | `DSTDEVP`     | rename  | Population. |
| `@DSTDS`      | `DSTDEV`      | rename  | Sample (R3+).  |
| `@DVAR`       | `DVARP`       | rename  | Population. |
| `@DVARS`      | `DVAR`        | rename  | Sample. |
| `@SUMPRODUCT` | —             | emulate | Not in IronCalc 0.7.1. |
| `@VDB`        | —             | emulate | Not in IronCalc. |
| `@INFO`       | `INFO`        | 1:1     | Only a subset of `info_type` keywords are implemented in IronCalc — verify each. |
| `@N`          | `N`           | arg-fix | 1-2-3 `@N(range)` returns the number in the **top-left cell** of the range; Excel `N(value)` coerces a single value. Wrap arg with `INDEX(.., 1, 1)` if it's a range. |
| `@S`          | —             | emulate | String form of `@N`; same top-left semantics. No Excel equivalent. |
| `@DQUERY`     | —             | emulate | External-DB hook; out of scope for L123. Emit `#NAME?`. |
| `@ISRANGE`    | —             | emulate | `ISREF` is close but checks for any reference, not range-shape specifically. |

---

## Open arg-shape questions (resolve before M2)

These need acceptance transcripts before claiming parity:

1. **`@MID` / `@FIND` 0-vs-1-based** — confirm with R3.4a docs/screenshots
   (we've been working from secondary sources). If confirmed, the
   translator must rewrite numeric *literal* args; rewriting cell-ref
   args means generating an `(arg+1)` expression at translation time.
2. **`@COUNT` semantics** — does R3.4a count labels? `1-2-3 R3 Reference`
   p. 4-50 says yes; verify. If yes, `COUNTA` is correct.
3. **`@LOG` base** — single-arg `@LOG(x)` is base-10 in R3; verify R3
   doesn't accept a 2-arg form. If it does, drop it (Excel `LOG(x, b)`
   is identical).
4. **`@STRING` format** — `@STRING(1234.5, 2)` → `"1234.50"` in 1-2-3
   (uses International punct). Confirm Excel `TEXT(n, "0.00")` produces
   identical output under each Punctuation A-H setting.
5. **`@CELL` keyword set** — enumerate which 1-2-3 keywords map to which
   Excel ones; document the gaps as `#VALUE!` for now.

---

## Implementation phasing

The renames split cleanly into three PRs, each with its own red/green/refactor cycle:

1. **Renames + `@@`** — pure name-swap table in `l123-parse`. Covers
   the bulk of the work (~10 functions). Tests: extend
   `crates/l123-parse/src/lib.rs` `tests` module with one assertion
   per rename.
2. **Arg-fix translations** — `@MID`, `@FIND`, `@N`, `@CELL`,
   `@COUNT`, `@STRING`. Each needs a small per-function rewriter.
   Tests: parser unit tests + at least one acceptance transcript per
   function (MID/FIND off-by-one is exactly the kind of bug a unit
   test misses but a transcript catches).
3. **Emulations** — `@ERR`, `@PROPER`, `@CHAR`, `@CODE`, `@CTERM`,
   `@TERM`, `@CELLPOINTER`, `@COORD`, `@@` (if the rename approach
   doesn't work). Each one is its own design decision: rewrite-time
   composition vs. engine-side hook. `@CELLPOINTER` definitely needs
   the engine hook (UI context); the rest are probably rewrites.
