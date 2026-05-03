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
  @REPEAT @COLS @CODE @D360 @ISRANGE`, the database family
  `@DAVG @DSTD @DSTDS @DVAR @DVARS`, and `@@`.
- **Niladic-paren completion** for `@PI @NOW @TODAY @RAND @NA @TRUE
  @FALSE` (1-2-3 omits `()`; Excel requires it).
- **Arg-fix**: `@MID` 0→1-based start, `@FIND` 0→1-based start with
  result-minus-1, `@STRING` → `TEXT` with a built format string
  (static for literal `decimals`, dynamic via `IF`/`REPT` otherwise).
- **Emulations**: `@ERR` → `#VALUE!` literal, `@CTERM` and `@TERM` to
  log compositions, `@SUMPRODUCT` → `SUM` with array broadcast
  (verified IronCalc 0.7.1 `SUM` accepts `(A1:A3)*(B1:B3)`),
  `@REPLACE` → `LEFT(s,start) & new & MID(s,start+n+1,LEN(s))`
  composition (IronCalc has `SUBSTITUTE` but not `REPLACE`),
  `@DQUERY(...)` → `#NAME?` literal (out-of-scope external-DB hook —
  any args swallowed; sidecar preserves the source form),
  `@S(range)` → `IF(ISTEXT(<topleft>),<topleft>,"")` where `<topleft>`
  is `INDEX(range,1,1)` for explicit ranges and the arg itself for
  single-cell args (IronCalc's `INDEX` rejects non-range args).
- **`@CELL` keyword dispatch**: keywords pass through unchanged for
  the implemented set (`address`, `col`, `contents`, `row`, `type`)
  and the IronCalc-stubbed set (`color`, `filename`, `format`,
  `parentheses`, `prefix`, `protect`, `width` — render as `ERR` until
  IronCalc lands real implementations). R3-only `"sheet"` rewrites
  to `SHEET(ref)` (1-based sheet number). `"coord"` emits `#VALUE!`
  because IronCalc 0.7.1 lacks `ADDRESS`.
- **Passthroughs** that need no rename or rewrite (the `@`-strip is
  enough): `@DGET` (R3 multi-input form deviates from Excel's
  single-range `DGET` — same caveat as the rest of the database
  family).
- **`@CELLPOINTER` pre-pass**: `l123_parse::expand_cellpointer(lotus,
  cursor)` rewrites `@CELLPOINTER(attr)` → `@CELL(attr, <cursor>)`
  before the main translator runs. The cmd-layer formula commit path
  (`l123-ui/src/app.rs`) supplies the cell address as cursor. Special
  case: `@CELLPOINTER("sheet")` rewrites directly to `@SHEET()` (the
  niladic form) so we don't generate a `SHEET(<self>)` self-reference
  that IronCalc flags as `#CIRC!`.

Deferred, with rationale:

- `@PROPER` — IronCalc 0.7.1 lacks `PROPER` and word-boundary
  capitalization isn't expressible in stock Excel functions. Needs
  an upstream PR or an engine-side shim.
- `@CHAR` — IronCalc lacks both `CHAR` and `UNICHAR`. Engine shim
  required.
- `@COORD` — would need `ADDRESS`, which IronCalc 0.7.1 lacks.
- `@VDB` — IronCalc 0.7.1 lacks `VDB`. The Excel formula is identical
  to 1-2-3's; landing `VDB` upstream in IronCalc is the natural fix.
  Until then, the call passes through and IronCalc returns `#NAME?`.
- `@N` argument-shape rewrite — IronCalc's `N` accepts ranges and
  silently picks the first cell, matching 1-2-3 behavior closely
  enough that the rewrite is a no-op in practice. Revisit if a
  divergence shows up in a transcript.

Boolean-rendering note: `@ISERR @ISNA @ISNUMBER @ISSTRING @ISRANGE
@TRUE @FALSE` all return Excel booleans which the renderer currently
paints as `TRUE`/`FALSE`. Authentic 1-2-3 displays booleans as `1`/`0`.
Cleanup is project-wide (a `cell_render.rs` change), not per-function.

Engine-adapter error mapping (resolved): IronCalc 0.7.1's `CellValue`
enum still has no `Error` variant — errors come back as
`CellValue::String("#VALUE!")`, `"#DIV/0!"`, etc. The
`IronCalcEngine::get_cell` adapter now inverts the six standard Excel
error codes (`#VALUE!`, `#DIV/0!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`)
to `Value::Error(ErrKind::*)`, gated on `formula.is_some()` so that
user-typed labels like `'#VALUE!` still pass through as text. As a
result the renderer paints the Lotus-style `ERR`/`NA` tag for
formula-derived errors.

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

Cosmetic reversal handles cases where forward+reverse compose
losslessly. The remaining cases — arg-fix wrappers (`MID(s,(b)+1,c)`,
`FIND(...)-1`), emulations (`LN(fv/pv)/LN(1+rate)`,
`TEXT(n,"0.00")`, `SUM((a)*(b))`), and 3D ranges expanded to a
comma-list — are handled via the **formula-sources sidecar**.

## Formula-sources sidecar

`l123-io::formula_sources` embeds the user-typed Lotus source per
formula cell as `l123/sources.tsv` inside the .xlsx zip. On save the
UI dumps every `CellContents::Formula.expr` to the sidecar; on load
the sidecar's `expr` overrides the cosmetic reverse-translated one.

This makes the irreversible cases round-trip too:

- `@CTERM(0.05,1000,500)` survives save → reload as
  `@CTERM(0.05,1000,500)`, even though the engine stores
  `LN(1000/500)/LN(1+0.05)`.
- `@MID("hello",2,3)` round-trips as `@MID("hello",2,3)`, not
  `MID("hello",(2)+1,3)`.
- `@SUMPRODUCT(B1..B3,C1..C3)`, `@TERM(...)`, `@STRING(n,2)`, and
  the `@FIND` arg-fix wrapper all round-trip exactly.

Limitations:

- The sidecar only survives if the file stays in L123. Vanilla Excel
  ignores unknown zip parts on load but drops them on resave; a file
  round-tripped through Excel falls back to the cosmetic reverse
  translator.
- 3D ranges still reload comma-expanded for *cells without a
  sidecar entry* (e.g. files originating from Excel). L123-native
  files preserve the `A:B3..C:B3` shorthand.
- The sidecar is trusted on load — if an external editor changes
  formulas without updating the sidecar, the panel will show a
  stale source. Editing the cell in L123 (F2) re-derives the
  sidecar entry on commit, fixing the staleness for that cell.

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
| `@CODE`       | `UNICODE` | rename | 1-2-3 `@CODE` is LICS in R3 and ASCII elsewhere; IronCalc only ships `UNICODE`, which agrees for ASCII input. Revisit if a transcript surfaces a divergence. |
| `@REPLACE`    | —        | emulate | IronCalc 0.7.1 has `SUBSTITUTE` (substring-match) but not `REPLACE` (fixed-position). Composed as `LEFT(s,p) & new & MID(s,p+n+1,LEN(s))` with 0-based `p`. |

### Date/Time

`@DATE @DATEVALUE @DAY @MONTH @YEAR @NOW @TODAY @TIME @TIMEVALUE
@HOUR @MINUTE @SECOND` all map to identically named IronCalc functions
(1:1).

| @-name   | IronCalc   | Status  | Notes |
|----------|------------|---------|-------|
| `@D360`  | `DAYS360`  | rename  | Day count between two dates on a 360-day-year basis (12 × 30-day months). |

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
| `@CELL`        | `CELL`     | arg-fix | Most keywords share the same name as Excel and pass through (`address`, `col`, `contents`, `row`, `type`, plus the IronCalc-stubbed `color`/`filename`/`format`/`parentheses`/`prefix`/`protect`/`width`). R3-only `"sheet"` rewrites to `SHEET(ref)`; `"coord"` emits `#VALUE!` (IronCalc lacks `ADDRESS`). Non-literal keyword args pass through verbatim. |
| `@CELLPOINTER` | —          | pre-pass | `expand_cellpointer(lotus, cursor)` rewrites `@CELLPOINTER(attr)` → `@CELL(attr, <cursor>)` before the main translator runs. The cmd-layer formula commit path supplies the cursor. `@CELLPOINTER("sheet")` is special-cased to `@SHEET()` (niladic) to avoid a `SHEET(<self>)` self-reference that IronCalc would flag as `#CIRC!`. |
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
| `@DGET`       | `DGET`        | 1:1     | Verified IronCalc 0.7.1 ships `DGET`. R3's multi-input form `@DGET(input1,...,inputn,field,criteria)` follows the same divergence as the rest of the database family — single-input usage matches Excel. |
| `@DMAX`       | `DMAX`        | 1:1     |       |
| `@DMIN`       | `DMIN`        | 1:1     |       |
| `@DSTD`       | `DSTDEVP`     | rename  | Population. |
| `@DSTDS`      | `DSTDEV`      | rename  | Sample (R3+).  |
| `@DVAR`       | `DVARP`       | rename  | Population. |
| `@DVARS`      | `DVAR`        | rename  | Sample. |
| `@SUMPRODUCT` | —             | emulate | Not in IronCalc 0.7.1. |
| `@VDB`        | —             | deferred | IronCalc 0.7.1 lacks `VDB`. Algorithm matches Excel's `VDB` exactly — landing it upstream in IronCalc is the natural fix; until then the bare name passes through and IronCalc returns `#NAME?`. |
| `@INFO`       | `INFO`        | 1:1     | Only a subset of `info_type` keywords are implemented in IronCalc — verify each. |
| `@N`          | `N`           | 1:1\*   | \*Doc-bug deferred: 1-2-3 `@N(range)` returns the number in the **top-left cell** of the range; Excel `N(value)` coerces a single value. In practice IronCalc's `N` accepts a range and silently picks the first cell, so the no-op translation matches 1-2-3 closely. Revisit (rewrite to `INDEX(.., 1, 1)`) if a transcript surfaces a divergence. |
| `@S`          | —             | emulate | `IF(ISTEXT(<topleft>),<topleft>,"")` where `<topleft>` is `INDEX(range,1,1)` for ranges and the arg itself for single cells (IronCalc's `INDEX` rejects non-range args). |
| `@DQUERY`     | —             | emulate | External-DB hook; out of scope for L123. Rewrites to `#NAME?` literal regardless of args. Sidecar preserves the original Lotus form for round-trip. |
| `@ISRANGE`    | `ISREF`       | rename  | 1-2-3 `@ISRANGE` checks "is this a defined range or valid range address?"; Excel's `ISREF` is broader (any reference). Agree on the cases users actually write; accept the small divergence. |

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
5. **`@CELL` keyword set** (resolved) — surveyed against IronCalc
   0.7.1's `fn_cell` (`information.rs:378`). Most 1-2-3 keywords share
   their name with Excel and pass through. R3-only `"sheet"` rewrites
   to `SHEET(ref)`; `"coord"` emits `#VALUE!` until IronCalc gains
   `ADDRESS`. The IronCalc-stubbed keywords (`color`, `filename`,
   `format`, `parentheses`, `prefix`, `protect`, `width`) render as
   `ERR` from the engine side — adequate until each is implemented.

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
