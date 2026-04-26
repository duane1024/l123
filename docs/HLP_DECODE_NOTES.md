# 123.HLP decoding — reconnaissance notes

Working file: `~/Documents/dosbox-cdrive/123R34/123.HLP`, 457302 bytes,
dated Mar 17 1993.

This document captures what's known about the format so far. The
runtime help system in `crates/l123-ui/src/help.rs` is wired with
hand-authored topics; replacing those with `.HLP`-decoded content is
this document's eventual target.

## TL;DR (session 6 — solved)

**Compression**: static Huffman, MSB-first bits with **inverted** branch
sense. Tree is 106 nodes at file `0x10..0x1b8`, each two `i16` (left,
right children). Non-negative child = literal byte to emit (then reset
to root); negative child = internal-node reference, jump to node index
`-value`.

**Tree length** is stored explicitly at file offset `0x0e` as a `u16`
(`0x01a8` = 424 bytes = 106 × 4).

**Offset table** lives at the dword stored at file offset `0x0a`
(`0x01b8`). It runs from there until the first record; values are
absolute file offsets (`u32` LE) clustered in monotonically-increasing
groups. The first record begins at `0x0e98`.

**Records** are individual Huffman-encoded blobs. After Huffman decode
each record is mostly readable text plus an unsolved second layer of
**renderer control bytes** — `0x00` is a line break, `0xC4` is layout
spacing/fill, and various other low/high bytes are commands the help
renderer interprets (cross-reference targets, attribute switches).
Body text decodes cleanly; cross-reference rows on index pages still
have noise around the topic titles.

Implemented in `crates/l123-help/src/{dict,huffman}.rs` with passing
tests against three real records (`Help Index`, the `ALT-F2 (RECORD)`
description, and the `/Print Margins` line).

External reference (used to crack the Huffman layer — credit to whoever
wrote it): `~/Library/Mobile Documents/com~apple~CloudDocs/l123/
lotus_hlp_extract.py` plus `lotus_hlp_findings.md`.

## Renderer-control byte stream — partial reverse (session 6, cont.)

After Huffman decode, each body topic record follows the layout

```
<renderer noise> [0x00 or 0xC4 boundary] <noise letters glued><Title> -- <body>
```

where:

- `0x00` is a hard line break.
- `0xC4` is a layout-fill byte (renders as a soft space) and also acts
  as a strong boundary between trailing renderer-noise and the title.
- The ` -- ` delimiter (with real `0x20` spaces) separates the title
  from the body in body topic records. Index pages don't carry it.

Title extraction (`crates/l123-help/src/renderer.rs::extract_title`):

1. Walk back from ` -- ` until hitting `0x00` / `0xC4` / non-printable.
2. The candidate region is a stretch of plain ASCII that looks like
   `<noise letters>[whitespace]<Title>` (the noise letters frequently
   abut the title with no separating space — e.g. `nnAbout 1-2-3 Help`,
   `cnn/Copy`, `nengntnnPointer-Movement Keys`).
3. Take everything from the first title-start character (uppercase
   letter, `@`, `/`, `(`, `$`, `+`) to the end. That cleanly recovers
   `About 1-2-3 Help`, `/Copy`, `Pointer-Movement Keys`,
   `F3 (NAME) Key`, all `/Data ...`/`/Print ...`/`/Worksheet ...`
   command paths, etc.

Body extraction (`renderer::clean_body`):

- Split raw body bytes on `0x00` (line breaks).
- For each line, drop bytes outside `0x20..=0x7E` (renderer commands),
  then strip the per-line noise prefix by skipping characters until
  reaching an 8-character-or-longer clean run that begins with a
  letter/digit/sentence-starter.
- Collapse whitespace and rejoin lines with `\n`.

### Survey across the whole file

| Records total | With ` -- ` | Title extracted | "Looks clean" titles | Unique clean titles |
|--:|--:|--:|--:|--:|
| 799 | 353 | 340 | 289 | **268** |

The 268 unique clean titles cover every Lotus 1-2-3 R3.4a command
path (`/File Save`, `/Worksheet Global Default Other Beep`, `/Data
Query Modify`, …), every function-key topic (`F1` through `F10`,
`(COMPOSE)`, `(GRAPH)`, `(QUERY)`), the help-system metatopics
(`About 1-2-3 Help`, `Help (continued)`), and most concept topics
(`Control Panel`, `Available Memory`, `Pointer-Movement Keys`,
`Types of Cell and Range References`, …). Run
`cargo run -p l123-help --example dump_topics` to see all 340.

### What's still rough

- **Body text has interleaved mid-line noise.** The renderer-control
  stream uses bytes like `n n 4 c SP n r ) i u s` mid-sentence that
  appear to be cursor-positioning / column-target commands. They look
  like printable ASCII, so the cleanup heuristic can't drop them
  perfectly. The noise is generally short and surrounded by clean
  text, so meaning is recoverable but the result is not byte-perfect.
- **Index-page cross-reference rows** still don't yield topic graph
  cleanly. Each cross-reference has a longer noise prefix that
  encodes the target topic ID; we haven't decoded that prefix yet.
- A handful of body titles miss their first word (e.g. `BPrint
  (continued)` should be `Print (continued)`, `CnRemove…` should be
  `Remove…`) because an uppercase letter inside the noise prefix gets
  picked as the title-start character.

For the L123 project's purposes — populating `HELP_TOPICS` with real
R3.4a content — what we have is enough to ship a topic-by-topic help
system that's substantially correct in title, mostly correct in body
text, and missing only the precise cross-reference graph.

### Wired into `l123-ui`

`l123-help` now exposes `pub fn topics(&[u8]) -> Result<Vec<Topic>>`
which walks the offset table, Huffman-decodes each record, runs the
renderer extractor, and dedupes by title. `l123-ui` depends on
`l123-help` and uses it for an authenticity audit:
`help::tests::help_topics_agree_with_decoded_source` confirms that
distinctive phrases from the hand-authored `HELP_TOPICS` curation
("context-sensitive", "ALT-F4", "Copies a range of data", "top three
lines") also appear in the decoded `123.HLP` source. The test skips
gracefully when the `.HLP` file isn't present (CI / contributors).

### All 321 topics shipped as committed source

`crates/l123-ui/src/help_topics_decoded.rs` is a generated 270 KB
Rust source file holding **all 321 unique decoded topics** as a
`pub static HELP_TOPICS_DECODED: &[(&str, &str)]`. Bodies still
contain the unreversed mid-line renderer noise — see "What's still
rough" above — but every topic the binary file holds is now part of
the L123 source tree and accessible without a runtime dependency on
`123.HLP`.

The generator at `crates/l123-help/examples/generate_topics_rs.rs`
re-emits the file from a fresh `123.HLP` decode:

```
cargo run -p l123-help --example generate_topics_rs
```

Tests in `help.rs` (`decoded_topics_table_present_and_nonempty`,
`decoded_topics_carry_authentic_phrases`) lock in the count, the
non-empty invariant, and the presence of well-known titles plus the
known-clean opening of `About 1-2-3 Help`.

The hand-authored `HELP_TOPICS` (32 curated topics) provides the
canonical clean bodies and complete cross-reference graph for the
top-level topics. `HELP_TOPICS_DECODED` covers the remaining ~289
specific commands and subtopics.

### F1 overlay navigates all 321 topics

`crate::help::all_topics()` returns a merged static slice that the
F1 overlay's Tab/Shift-Tab walker uses: curated 32 first, then every
decoded entry whose title isn't already curated. Decoded entries get
synthetic `decoded-<slug>` ids and empty `cross_refs` (the cross-ref
graph for non-curated topics isn't recovered yet). The header reads
`[topic N/319]` where `N` is the user's position; hand-curated
content is shown when available, otherwise the decoded body (with
its residual mid-line noise) is rendered.

Tests in `help.rs` (`all_topics_merges_curated_and_decoded`,
`all_topics_includes_decoded_only_entries`, `slugify_examples`)
lock in the merge invariants: curated entries appear first in
order, no duplicate ids/titles, and decoded-only entries like
`File Combine`, `/File Erase`, `File Admin` are reachable.

## Findings from session 8 (2026-04-26) — cross-reference pages cracked

Goal: interpret the non-ASCII characters that decoded as `Cn` before
"Remove an active file from memory" and `nn ng tsno·a)ge` before
"data in a worksheet file" — the user hypothesized these were
hyperlink markers. Confirm or refute, ideally using the live DOSBox-X
session for ground truth.

**Confirmed**: those bytes *are* cross-reference markers. Specifically,
each row of an index/cross-ref page (Task Index, Function Index,
Keyboard Index, etc.) has the on-byte-stream layout

```
<topic-id-bytes><description text><spaces>-- see\0
```

The bold cross-reference target the user sees on screen at the *end*
of the line ( `/File Close`, `/File Save`, etc.) is **not** in the
byte stream. The help renderer reads the leading topic-id bytes and
renders the target name from a topic-name table at runtime.

This explains the longstanding mystery from sessions 6 and 7:

- The "noise" before each visible description is a variable-length
  topic-id encoding. Length varies from 2 bytes (`Cn` for /File
  Close) to ~15 bytes (`nn ng tsno·a)ge` for /File Save) — consistent
  with a Huffman-style prefix code over ~824 topics.
- Index pages also have a *header* section at the start of the record
  with each cross-ref target's name in a "render this title bold"
  form. That section was decoding as "noise plus topic name" and got
  confused with the body content.
- Body topic records share the same format: their footer contains the
  three cross-references (`Continued`, `Specifying Ranges`, `Help
  Index`) the user sees at the bottom of every page, encoded the
  same way.

### Verified against the running DOSBox-X session

Captured Task Index page 1 live (`screencapture` + osascript-driven
keystrokes through DOSBox-X) and matched it byte-for-byte against
the decoded record at `0x01cbd`:

```
Task Index                                           ← title (encoded
                                                       in noise; not
                                                       in raw text)
Change column width                  -- see /Worksheet Column   ← row 1
Change display of data               -- see /Range Format       ← row 2
…
Continued    [...]    Help Index                                 ← footer
```

Each visible row's byte slice begins with topic-id bytes the renderer
uses to draw the bold target at the right edge. Stripping the prefix
yields the clean description on the left.

### What we shipped in `crates/l123-help/src/renderer.rs`

`extract_topic` now branches on a `count_see_markers` signal: any
record with ≥ 3 ` -- see` markers is treated as a cross-ref page.
For those records:

- `extract_crossref_page` splits on `0x00`, runs each row through
  `clean_crossref_row`, and returns a `Topic` whose body is the
  cleaned cross-ref entries one per line.
- `clean_crossref_row` calls `find_description_start`, which scans
  for the earliest character position whose suffix tokenizes into
  *all* English-looking tokens (using `looks_english` — vowel ratio
  ≥ 25 %, ≤ 2 leading consonants, no embedded capitals, ≤ 4
  consecutive consonants, plus a small allow-list of common
  short words). That position is the start of the description.
- If the row contains ` -- see`, we re-attach `  -- see` after
  trimming so the cleaned output reads as a complete cross-ref line.
- `guess_index_title` recovers the title (Task Index / @Function
  Index / Keyboard Index / etc.) by string-search against the full
  decoded byte stream.

Body topic records (the majority — `/Copy`, `About 1-2-3 Help`, etc.)
keep the existing `extract_topic` behavior. The mid-body noise on
those records is the same renderer-control issue from session 7 and
remains structurally unrecovered (see *What's still rough* below).

### Result on real records

Before:

```
("CnRemove an active file from memory",
 "see\nnn ng tsno…")
```

After:

```
("Task Index (continued)",
 "Remove an active file from memory  -- see\n
  data in a worksheet file  -- see\n
  Share files on a network  -- see\n
  …")
```

(Some entries still carry a 2–4 char residual where the renderer's
cursor encoding ate the description's first character — `gepart` for
"Save part", `tindthare` for "Share". The fix for that is the same
one needed for body-topic noise: simulate the renderer's cell-by-cell
draw with the full opcode table, which still requires the runtime
trace described in session 7.)

### What's still rough

- ~30 of the ~321 decoded topics have the "first char eaten" issue
  in cross-ref descriptions (`gedata`, `gepart`, `tindthare`, etc.).
- Body topic records still have mid-line noise on most lines (`nn4c
  nr)i` between "protection" and "us, to a range").
- The exact topic-id-byte encoding (variable-length prefix code over
  topic IDs) is still not pinned down. We don't need it for clean
  description text, but we'd need it to wire up the cross-reference
  navigation graph.

The cleanest next step for closing the remaining gap is the
DOSBox-X runtime trace from session 7's recommendation, with a
breakpoint on the help-render entry rather than `int 21h`.

## Findings from session 7 (2026-04-26) — noise prefix is structurally elusive

Goal: pin down the per-line "noise prefix" rule so decoded body text
comes out byte-perfect, then ship the full 340-topic decoded set as
user-facing help.

Outcome: **bailed**. After ~50 minutes of structural analysis no clean
rule was found. The curated `HELP_TOPICS` stays in place; no code
changes shipped. Recommended next step: DOSBox-X runtime tracing
(session 5 identified that as the most viable path).

### Survey of the noise prefix across 348 records / 3801 body lines

Used the existing Huffman decoder to dump every body record (where the
record contains ` -- `), split bodies on `0x00`, and detected the
"first English word" position to estimate prefix length. Histogram:

| Prefix length | Share | Cumulative |
|---:|---:|---:|
| 0–8   | 8.2 % | 8.2 % |
| 9–12  | 8.4 % | 16.6 % |
| 13–16 | 22.5 % | 39.1 % |
| 17–20 | 17.3 % | 56.4 % |
| 21–25 | 15.6 % | 72.0 % |
| 26–32 | 13.8 % | 85.8 % |
| 33–50 | 11.6 % | 97.5 % |
| 51–74 |  2.5 % | 100 % |

Modal prefix length is 16, but the spread is wide (1 → 74 bytes) and
no single length dominates. **Prefix length within a single record is
not constant** — body lines of the same record show prefixes of e.g.
`[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 12, 14, 0, 0]` (a heuristic
artifact: lines whose first word matches the dictionary look like
"prefix=0").

### Hypotheses tested and rejected

1. **Fixed-length header** — rejected. Lengths span 0–74 bytes per
   line; modal value 16 covers <8 % of samples.
2. **Leading length byte** — rejected. The first prefix byte is most
   commonly `0x6e` (`'n'`, 25 % of all prefix bytes). Its value does
   not correlate with prefix length, and `'n'` itself appears as a
   literal text byte at high frequency too.
3. **(column, attr, …) tuple** — rejected. Pair-aligned analysis
   (even/odd byte streams) yielded no readable text in either stream.
4. **Distinct command-byte / data-byte byte ranges** — rejected. The
   prefix byte distribution is *the same shape* as the body text byte
   distribution: heavy in `'n' 'e' 'o' 't' 'r' 'i' 'a' 's' 'l'`,
   plus `0xC4` (fill) and `0x20` (space). There is no escape byte
   (e.g. high-bit-set, control range) that delimits prefix from text.
5. **Boundary marker** — rejected. The last 1–3 bytes before the
   visible English text don't cluster on any one byte value or
   sequence; `c4` and `0x20` precede text only ~13 % of the time.
6. **Prefix encodes leading whitespace + bold attribute** — partial
   match for some lines, but inconsistent. Prefix length does not
   correlate with the OCR-confirmed leading-space count of the
   visible line.

### What we did learn

- The first 1–5 chars of each visible line are routinely **missing**
  from the decoded byte stream. e.g. `/Copy` line 2 is `" status,
  to a range…"` in OCR but decodes to `"…ius, to a range…"` —
  `" stat"` (5 chars) is gone, with no obvious encoding of those bytes
  anywhere in the prefix or surrounding bytes.
- Suggests the renderer's encoding is **not a simple prefix-then-data
  layout**. More likely it's a Microsoft-style help-compiler scheme
  where each Huffman *literal* byte carries dual meaning — sometimes a
  display char, sometimes a renderer command (cursor advance, attr
  switch, "draw next char with attribute X"). Distinguishing them
  requires the dispatch table that lives inside `123DOS.EXE`'s help
  overlay, which session 5 confirmed is unreachable via static
  analysis without first cracking Lotus's overlay format.

### Why this isn't a heuristic-only problem

The current `clean_body` heuristic ("strip until first ≥8-char clean
run") leaves visible noise inside body lines because the noise bytes
*are* clean — they're letters and punctuation, just not in
English-word order. A stricter "find first English word" heuristic
fixes this for ~70 % of lines but consistently truncates leading
words (`"depending"` → `"pending"`, `"information"` → `"formation"`,
`"CAUTION"` → `"TION"`) because the first-word characters genuinely
aren't in the byte stream we have. Stripping the prefix doesn't
recover them; only a real renderer simulation would.

### Decision

Per the original prompt's guidance ("better to leave the curated set
in place and document what was tried"), `crates/l123-ui/src/help.rs`
is unchanged: `HELP_TOPICS` stays curated; the `help_topics_decoded.rs`
artifact was not generated. The acceptance transcript
`m11_f1_help_open_close.tsv` is unchanged. The
`help::tests::help_topics_agree_with_decoded_source` provenance check
continues to assert curated-vs-decoded phrase agreement at the
substring level, which the decoded stream supports today.

### Recommended next step (session 8)

DOSBox-X runtime tracing — session 5 identified this as the highest-
likelihood path. The plan:

1. Boot 1-2-3 R3.4a in DOSBox-X with the built-in debugger enabled
   (`debug=heavy` in the config).
2. Set a breakpoint on `int 21h` with `AH=3Dh` and `DS:DX → "123.HLP"`.
3. Press F1; when the breakpoint trips, you're at the file-open call
   site of the help overlay.
4. Step through `read` → decoder loop. The loop's branch on each
   decoded byte (`if (b < THRESHOLD) emit; else dispatch_command`)
   is the rule we're missing.
5. Reproduce the dispatch table in `crates/l123-help/src/renderer.rs`.

The structural-only path is exhausted; further static work without
runtime data is unlikely to produce results.

## Header (offsets 0x00..0x10)

| Offset | Value | LE u16 | Notes |
|---|---|---|---|
| 0x00 | `00 10 00 00` | `0x1000`, `0x0000` | Probable magic / version |
| 0x04 | `38 03 0e 00` | `824`, `14` | Field counts (function below) |
| 0x08 | `00 00 b8 01` | `0`, `440` | |
| 0x0c | `00 00 a8 01` | `0`, `424` | |

`0x1a8 = 424` matches the byte length of the dictionary block from
`0x10..0x1b8`, so the value at `0x0e` (424) appears to be the
dictionary byte size.

## Dictionary block (0x10..0x1b8, 424 bytes = 212 LE16 = 106 pairs)

Each entry is a pair of `i16`:

```
(-1, -26)   (-2, 196)   (-3, -8)    (-4, 97)    (104, -5)
(-6, 112)   (-7, 119)   (84, 107)   (-9, 116)   (-10, 108)
...
```

This is the classic **byte-pair encoding (BPE) dictionary** Lotus used
in their help compiler. Each pair at slot `N` (counting from 1)
defines an expansion: code → `(left, right)`. Codes with negative
values point recursively to other dictionary slots; positive values
are literal bytes (typically ASCII characters — note 84='T', 76='L',
83='S' all show up).

The 824 (0x338) at offset 0x04 may be the *number of distinct symbols*
the dictionary covers, or the maximum code value used. 824 > 256 so
the codec is not byte-aligned — it almost certainly uses a variable-
or short-bit code, possibly 10-bit or 12-bit.

## Topic offset table (starts 0x1bc)

A run of monotonically-increasing `u32` LE offsets:

```
0x01bc: 0x00001205   (first content chunk)
0x01c0: 0x00001548
0x01c4: 0x000018d4
...
```

53 monotonic entries before non-monotonic data appears around 0x290.
This is likely just the *first* offset table; more index structure
sits between 0x290 and the first content chunk at 0x1205.

## Content section (starts 0x1205)

Begins with `24 00 07 00 03 01 10 0f 00 1f 00 15 ef 09 f8 b6 …`.
The leading `0x24 = 36` could be a chunk length, `0x07` a
sub-flag/count, then encoded body bytes.

No ASCII runs of 12+ chars appear *anywhere* in the 457KB file. Byte
0xaa is the most frequent at 6.9% — typical of a BPE escape or pad
byte. `the`, `and`, `cell`, `menu` all yield zero literal matches.

## Findings from session 2 (2026-04-25)

Ground truth captured via `screencapture -R` against DOSBox-X (window
position scraped via System Events) — see `tests/help_groundtruth/`
for index, About, and Function-Keys screens with OCR text.

### Dictionary slot values are byte pairs, not slot refs

Re-interpreting each `i16` as its low byte (mask `& 0xff`) makes every
slot a `(u8, u8)` pair — sign-extended bytes, not recursive references:

- slot 1: `(0xff, 0xe6)` — 256-1 = 0xff; not "ref to slot 1"
- slot 5: `('h', 0xfb)` — `('h' = 0x68, 256-5 = 0xfb)`
- slot 8: `('T', 'k')`
- slot 26: `('O', 'N')`
- slot 50: `(0xd9, 0xbf)` — CP437 line-drawing chars `┘┐`
- slot 105: `('o', 'n')`
- slot 106: `('o', 'n')` (also "on")

Pattern: high-frequency CP437 attribute/extended bytes (0xE0..0xFF)
that paired-up with letters during BPE training got their u8 bit-7
sign-extended into i16 storage. The "self-reference" pattern is just
because slot N stores byte (256 − N) on one side; coincidental, not
semantic.

So the dictionary is **flat** — 106 codes, each emits exactly two
bytes — not the recursive expansion the recon assumed.

### Topic offset table is 824 entries of `(segment, offset)`

The header value `824` at byte 0x04 matches the *number of topics*.
The table runs `0x1bc..0xe98` (3292 bytes ≈ 823 × 4). Each entry is
two LE u16s — `segment, offset_within_segment` — yielding a 32-bit
file offset of `(segment << 16) | offset`. This explains the apparent
"non-monotonic" jumps at 0x290, 0x440, 0x4e0: topics aren't sorted by
file offset, they're sorted by topic ID. Topics live in segments
0, 1, 2, 5 of the 7-segment file (file size 0x6FA56 = ~7×64KB).

### Per-chunk record format

Each topic chunk begins with a 5-byte header
`<u16 record_count> <u16 ?> <u8 attr_or_color>` then a series of
records. Records have header `2c <cmp_len:u8> <unc_len:u8> <pad:u8>`
followed by `cmp_len` bytes of payload.

Records frequently end with the byte triplet `04 10 NN` where `NN`
appears to be a screen row number — increments record-by-record
(`04, 05, 06, ...`). For some records the row byte sits *inside* the
`cmp_len` content; for others it's an out-of-band byte between the
record and the next `2c`. The exact rule isn't pinned down yet — the
most likely explanation is that "advance to next row" can be either a
trailing escape inside the bit stream or a record separator outside
it.

Chunk types vary: the first u16 of chunks 0–3 looks like a record
count (17, 19, 19, 13) but chunks 4–7 have different leading bytes
(`0b 00 0b 00 03 0a`, `03 00 12 00 14 00`, etc.). Index pages,
text-flow topics, and tabular topics likely use different layouts.

### Bit-stream payload encoding is unknown

The chunk payloads contain a mix of:

- Plain ASCII letters and punctuation (likely literal codes 0x20..0x7F)
- Control bytes 0x01, 0x04, 0x10, 0x15, 0x1d, 0x1e (probable cursor /
  attribute commands)
- Extended bytes 0x80..0xFF (some are dict refs, some are CP437
  literals — discriminator unclear)

Tried decoding record 0 of chunk 0 with every combination of
{8, 9, 10, 11, 12}-bit codes × {LSB, MSB} bit order, treating
codes 0..255 as literals and codes 256..361 as dict slots. None
produced text containing OCR-confirmed strings ("1-2-3 Help Index",
"About 1-2-3 Help", "press F1") at any starting bit offset.

The payload looks **byte-aligned** rather than bit-packed: `cmp_len`
counts whole bytes, the long `0xAA` runs that pad chunk *boundaries*
sit *between* the offset table and content (not inside chunks), and
records have clean byte-granular structure. So the encoding is more
likely a per-byte command/data scheme than a bit-packed BPE stream.

### Realistic next step: disassemble `123DOS.EXE`

`123DOS.EXE` (956 KB DOS MZ) is the binary that opens `123.HLP`
(strings: `123.HLP`, `_HELP_SWT`, `_HELP_STR_SWTD`, `_HELP_ERR_SWT`).
Disassembling the help-render routine is the only realistic way to
nail the encoding down. Tools needed: `nasm`/`ndisasm` for raw 16-bit
disasm, or a Ghidra session with the DOS-MZ loader.

The ground-truth captures + OCR text in `tests/help_groundtruth/` are
sufficient validation cribs for whatever decoder we end up building.

## Findings from session 3 (2026-04-25 cont.) — Ghidra hit overlay wall

Ran Ghidra 12.0.4 headless against `123DOS.EXE`. Three blockers, all
caused by the same root cause:

### `123DOS.EXE` is a DOS overlay binary; Ghidra MZ loader sees only 7 % of it

Parsing the MZ header (`pages=112, last_page=99, header_paragraphs=32`)
yields an MZ image of **56 931 bytes** out of 956 499 — the remaining
~900 KB at file offset 0x0DE63 onward is overlay data the MZ loader
never maps. After auto-analysis, Ghidra exposes only three code blocks
totalling ~56 KB:

| Block   | Range                  | Size  |
|---------|-----------------------|-------|
| CODE_0  | 1000:0000..1000:1baf  | 7 088 |
| CODE_1  | 11bb:0000..11bb:7b2f  | 31 536|
| CODE_2  | 196e:0000..196e:4582  | 17 795|

`grep --byte-offset 123.HLP` against the file finds the literal at
**file offset 0x07e906** — well inside the overlay region. So
`findBytes("123.HLP")` across all of Ghidra's blocks returns zero
hits, and `getReferencesTo` of the (non-existent) hit address returns
zero xrefs. The help-decode routine itself is also almost certainly
in overlay land, not in the loaded 56 KB.

### What that means for the next attempt

Three viable paths, in roughly increasing complexity:

1. **Manually map the overlay region as additional Ghidra memory
   blocks** and re-analyze. A 64K-paragraph-aware Java script that
   reads bytes 56931..EOF from disk and stuffs them into synthetic
   `0x4000:0000`, `0x5000:0000`, ... blocks is the right shape; my
   first attempt failed compilation (`createInitializedBlock`
   signature drift) and I bailed before debugging. Once the bytes are
   resident, the byte search will hit and we can sweep instructions
   for far-pointer loads of `(seg=overlay_seg, off=str_off)` followed
   by `int 21h, ah=3D` (DOS open) to find the file-open call site.
   From there, walk callers to the chunk-decode entry point.

2. **Drive Ghidra interactively from the GUI.** Auto-analyze the MZ
   image, then `Window → Memory Map → +` to add the overlay region as
   one or more new blocks (file-bytes-backed). The GUI string search
   + xref view is dramatically faster than scripted iteration for
   reverse engineering, and the decompiler can be steered to the right
   function without us having to encode a search policy in Java.

3. **Skip Ghidra; raw-disasm the overlay range** with `ndisasm` (16-
   bit), search for the byte sequence `b8 .. .. ba 06 e9` (mov ax,…,
   mov dx, 0xe906) or similar load-and-open patterns near the help
   string offset.

### Empirical findings that survive the Ghidra dead end

These are still useful for whoever resumes this:

#### Byte-frequency analysis of `123.HLP`

```
0xaa: 31540  6.90%  ← top, ~10× any plausible BPE bigram
0x00: 27509  6.02%
0x01:  8741  1.91%
0x04:  6668  1.46%
0x08:  6619  1.45%
0x10:  6035  1.32%
0x41:  5155  1.13%  ← ASCII 'A'
0x20:  4947  1.08%  ← ASCII space
0x80:  4250  0.93%
0x06:  4149  0.91%
```

`0xAA`'s 6.9 % share is too high to be a normal BPE-slot reference —
the highest-entropy English bigram caps near 3.5 %. It's much more
likely a structural marker: an escape byte that introduces a dict
slot index, a single-byte SPACE substitute, or chunk padding. Any
viable decoder will treat 0xAA as special, not as another dict slot.

ASCII 'A' (1.13 %) and SPACE (1.08 %) are present at low rates
consistent with the literal-byte fallback path of a dictionary-driven
encoder (most letters and spaces are absorbed into 2-byte slots).

#### Manual chunk parsing breaks the 4-byte-header model

Walking chunk 0 (file 0x1205, size 835) byte-by-byte against the
recon's `2c <cmp_len> <unc_len> <pad>` rule:

```
rec  cmp_len  body-bytes-before-next-2c  trailing
 0     17     17                         (next 2c immediate)
 1     13     13                         (one floating byte 0x05 then 2c)
 2     12     12                         (two floating bytes 0x20 0x06 then 2c)
 3     25     22  ← cmp_len > body       (next 2c three bytes too early)
 4     17     16  ← cmp_len > body       (next 2c one byte too early)
```

So `cmp_len` matches the body-to-next-`2c` distance for some records
but overshoots for others, by varying amounts. Either (a) `0x2c` is
not a record separator (some records use a different start byte and
the 2c we see is content), (b) `cmp_len` actually counts something
other than body bytes for non-text records, or (c) chunks contain
multiple record *types* with different framing. Whatever the answer,
the rule isn't "count cmp_len bytes and find another 2c."

The 1-/2-byte gaps between records (the `0x05` and `0x20 0x06`
floating bytes above) are likely the row-advance opcode the recon
identified: most records emit `04 10 NN` inline at end-of-line, but
some emit only `04 10` with the row number `NN` written *outside* the
record as a structural break. That's still consistent.

#### Patterns visible inside record bodies (not yet decoded)

The first 5 bytes of most rec-0 bodies in chunks 0 and 3+ follow the
same template `<u16 X> 15 38 41`:

```
chunk 0 rec 0:  dd 01  15  38  41 1f b3 7e 45 89 c0 1d 1e f8 04 10 04
chunk 0 rec 3:  bf 01  15  38  41 1f b3 7e 45 89 c8 d0 77 f6 90 a0 9e e3 f5 82 08 07
chunk 0 rec 4:  e4 01  15  38  41 1f b3 7e 45 89 d0 06 90 f0 41 08
chunk 0 rec 7:  e9 01  15  38  41 1f b3 7e 45 89 c2 60 0d 0f a1 04
chunk 3 rec 0:  e0 01  15  38  41 1f b3 7e 45 88 50 5a 2c 21 04 04
```

The shared prefix `15 38 41 1f b3 7e 45 89` is too long to be
coincidence — it's likely a fixed header (cursor reset, attribute
init, leading dictionary refs, or a topic-link binding). The
varying `<u16 X>` first field plausibly encodes the next-line offset
or topic-cross-ref id (the field is 0x01dd, 0x01bf, 0x01e4, 0x01e9 —
all near 0x01e0, suggesting they're indices into one of the upper
tables in the file header's 14-entry index region).

### Status of `crates/l123-help/`

Unchanged from session 2: `dict.rs` parses the 106-pair dictionary
correctly; `examples/explore.rs` dumps chunks/records and runs
speculative decoders. No production decoder modules exist yet — we
intentionally stopped before introducing speculative code, since the
encoding rule needs to come from disassembly.

`HELP_TOPICS` in `crates/l123-ui/src/help.rs` is still the 8 hand-
authored stubs from Phase 1.

### Anchors found in the overlay near the `123.HLP` literal

Even though Ghidra didn't load the overlay, the bytes around the
`123.HLP` string (file offset `0x7e986`) are usefully readable as raw
x86-16. The 0x66 bytes before the string at `0x7e920..0x7e985` are a
small **filename-dispatch routine**:

```
0x7e920: 90 55 8b ec 56 c4 76 08 8b 46 06 48 74 07 48 74 34 48 eb 3b
         nop / push bp / mov bp,sp / push si / les si,[bp+8] / mov ax,[bp+6]
         dec ax / jz +7 / dec ax / jz +0x34 / dec ax / jmp +0x3b
0x7e934: 90 26 80 3c 31 75 0a 1e 68 9e 56 0e 68 e3 63 eb 32
         cmp es:[si], '1'  / jnz +10 / push ds / push 0x569e / push cs / push 0x63e3 / jmp +0x32
0x7e945: 26 80 3c 50 75 0a 1e 68 9e 56 0e 68 cb 63 eb 22
         cmp es:[si], 'P'  / jnz +10 / push ds / push 0x569e / push cs / push 0x63cb / jmp +0x22
0x7e955: 26 80 3c 4c 75 14 1e 68 9e 56 0e 68 ef 63 eb 12
         cmp es:[si], 'L'  / jnz +0x14 / ... push 0x63ef / jmp +0x12
0x7e965: 1e 68 9e 56 0e 68 c3 63 eb 08
                          ... push 0x63c3 / jmp +8
0x7e96f: 1e 68 9e 56 0e 68 d7 63 9a 56 39 08 01
                          push ds / push 0x569e / push cs / push 0x63d7
                          call far 0x108:0x3956
0x7e97c: b8 9e 56 8c da 5e c9 ca 06 00
         mov ax,0x569e / mov dx,ds / pop si / leave / retf 6
0x7e986: "123.HLP\0"
0x7e98e: "L123SMP3.RI\0"
0x7e99a: "UNKNOWN FIL\0"
0x7e9a6: "L123TXT3.RI\0"
0x7e9b2: "LTSADDIN.RI\0"
```

What this tells us:

1. **`call far 0x108:0x3956` is the file-open helper.** Every dispatch
   case ends with this same far-call target. So the chunk-decode work
   does not happen here; this routine just chooses *which* file to
   open.

2. **The five filenames live in one segment whose runtime base puts
   the string offsets at `0x569e` (= `0x7e986 - 0x792e8`).** Whatever
   overlay segment hosts this code+data block, its data offsets begin
   at `0x569e` for `123.HLP` and increment from there. Mapping that
   segment in Ghidra at any base — e.g. set the block start so that
   the bytes for `123.HLP` land at offset `0x569e` — should let
   Ghidra's auto-resolver follow the `push 0x569e + push 0x63xx + call
   far 0x108:0x3956` arguments cleanly.

3. **The chunk-decode routine is reached *after* this dispatch, via
   the file-open helper at `0x108:0x3956`.** Segment 0x108 is small
   enough to plausibly live inside the resident MZ image (Ghidra's
   `CODE_0`/`CODE_1` blocks). If the load-time relocation puts segment
   0x108 inside `CODE_0`/`CODE_1`, you can find the help-rendering
   code by following calls *out* of the routine at `0x108:0x3956` —
   that's where the chunk decoder lives. (If it lives in another
   overlay, you're back to mapping more overlay region, but at least
   the dispatch + file-open chain is now anchored.)

4. **No `b8 00 b8`/`b8 00 b0`/`push 0xb800`/`push 0xb000` anywhere in
   the binary.** Lotus 1-2-3 doesn't poke video memory directly — it
   goes through the `123VIDEO.*` driver layer. So you can't find the
   decoder by tracing video-mem writes; trace from `0x108:0x3956`
   forward instead.

### What's still unknown

- The runtime base of the overlay segment hosting `123.HLP` and the
  dispatch code (need to find Lotus's overlay table — typically a list
  of `(file_offset, segment_id, length)` records somewhere in the
  resident image).
- Whether `0x108:0x3956` is the actual file-open or just an
  intermediate trampoline. The body at file-paragraph `0x108`*16 +
  `0x3956` = file `0x49d6` *is* in the resident image; sample at
  `0x49d6..+64` looks like normal x86-16 prologue/epilogue glue, so
  worth decompiling.
- The encoding rule itself. Nothing in this stretch reduced
  uncertainty about how chunk bytes map to characters.

## Findings from session 5 (2026-04-25 cont.) — Ghidra GUI scripts hit wall

Drove Ghidra GUI via three scripts (`MapOverlayAndDumpHelp`,
`DumpHelpThunkAndCallers`, `FindDispatchCallers`) — all under
`~/ghidra_scripts/` and `crates/.../ghidra-projects/l123/scripts/`.
Mapped the 900 KB overlay as 14 synthetic blocks at segments
`4000:0..f000:fefff`, force-disassembled the help dispatch at
file `0x7e920` (= synthetic `b000:11bd`), force-created a function
there, ran auto-analysis on the new bytes, and walked references.

### What we confirmed

1. The dispatch is real and decompiles cleanly. Listing:

   ```
   90 55 8b ec 56 c4 76 08    nop; push bp; mov bp,sp; push si; les si,[bp+8]
   8b 46 06                   mov ax, [bp+6]            ; selector
   48 74 07 48 74 34 48 eb 3b ; switch ax: 1→case_1, 2→case_2, default→case_3
   26 80 3c 31 75 0a          ; case_1: cmp [es:si], '1'; jnz next
   1e 68 9e 56 0e 68 e3 63    ; push ds, 0x569e, cs, 0x63e3
   eb 32                      ; jmp common
   ... (similar cases for 'P', 'L', default)
   9a 56 39 08 01             ; common: call far 0x108:0x3956
   b8 9e 56 8c da             ; mov ax, 0x569e; mov dx, ds  (return filename)
   5e c9 ca 06 00             ; pop si; leave; retf 6
   ```

2. **`0x108:0x3956` is not the file-open helper** — it's a generic
   varargs runtime utility. Across 156 call sites (in 74 distinct
   functions), it's invoked with arg counts ranging 0 → 9 with
   inconsistent shapes. Decompiled callers do `func_0x000049d6(0x108,
   uVar3, uVar4, puVar3)`, `func_0x000049d6(0x150, *(int*)0x7b6a)`,
   `func_0x000049d6(0x110, uVar1, uVar2, puVar3)`, etc. — different
   "first segment" values (0x108, 0x110, 0x150, 0x178, 0x200, 0x400,
   …). This is the shape of a `printf`-family or message-formatter
   accepting `(format_seg, format_off, ...args)` pairs, not a file
   primitive.

3. **The dispatch has zero static callers.** Both Ghidra's
   `getReferencesTo` and a brute byte-search for `9a bd 11 ?? ??`
   (`call far ??:0x11bd`) returned 0 hits across the whole 956 KB
   binary. Even broader — searching for the literal `0x569e` (the
   filename offset the dispatch returns) — yields **only six hits,
   all inside the dispatch itself**:

   ```
   push 0x569e:    file 7e93c, 7e94c, 7e95c, 7e966, 7e970
   mov ax, 0x569e: file 7e97c
   ```

   No other code in the binary loads or references the filename.

### What that almost certainly means

The dispatch is reached via a **runtime indirection** whose fixed-up
target leaves no static byte trace:

- Lotus's overlay manager registers `(overlay_id, entry_offset)`
  pairs at startup, then resolves them at runtime — the `call`
  instruction reads the function pointer from a table populated by
  the overlay manager, not from a literal in the code stream.
- Or the dispatch is exposed as an exported entry of the help
  overlay (Lotus's overlay format has its own export tables that
  the manager loads), and the F1-key handler calls it via that
  manager.

Either way, **static analysis of `123DOS.EXE` cannot find the
caller** without first reverse-engineering Lotus's overlay format.
The decompiled functions we *did* get (30+ in the report) are all
either `0x108:0x3956`-callers (printf-family wrappers) or unrelated
utility code.

### Practical implications

Three forward paths, ranked by likelihood of success:

1. **Runtime tracing in DOSBox-X** (recommended). DOSBox-X has a
   built-in debugger (Ctrl-F1 by default; toggleable via config).
   Set a breakpoint on `int 21h` with `AH=3Dh` (DOS open) when DS:DX
   points to a buffer ending `"123.HLP\0"`, run 1-2-3, press F1.
   When the breakpoint hits, you're at the actual help-render call
   chain; step through the file-read loop to find the chunk decoder
   live. Bypasses the entire static-analysis problem because the
   overlay manager has already done its job by that point.

2. **Manual transcription of help screens.** The L123 project's
   ground-truth captures (`tests/help_groundtruth/`) already cover
   3 screens. Capturing the rest — ~30 screens × ~5 minutes of OCR
   each = ~3 hours — gets us a hand-authored HELP_TOPICS table that
   *is* the real R3.4a help text, no decoder needed. Falls short of
   the original "decode `123.HLP` end-to-end" goal but ships
   authentic help content fast.

3. **Reverse-engineer Lotus's overlay table** in the resident image.
   The MZ image holds the overlay table somewhere (paragraph
   pointers / file offsets per overlay segment). Finding *that*
   table would let us resolve `0x108:0x3956` and segment `b000`
   properly, after which Ghidra's xrefs would actually populate.
   Open-ended work — likely several days, no guarantee.

### Files & state

- Ghidra project: `~/ghidra-projects/l123/`
- Scripts (idempotent, runnable from GUI Script Manager):
  `~/ghidra_scripts/MapOverlayAndDumpHelp.java`,
  `DumpHelpThunkAndCallers.java`, `FindDispatchCallers.java`
- Latest report (~45 KB): `~/ghidra-projects/l123/scripts/help_decompile_report.txt`
- Workspace still green: `cargo build/test/clippy --workspace` clean.
- `crates/l123-help/` unchanged from session 2; `HELP_TOPICS` still the
  Phase-1 stubs.

## Findings from session 4 (2026-04-25 cont.) — ndisasm + far-call census

Installed `nasm` (gives `ndisasm`) and disassembled all 956 KB. The
output is at `/tmp/123DOS.disasm` (~383 K lines) and is grep-friendly.

### Segment `0x108` is Lotus's runtime-library segment

Counted every `call far <seg>:<off>` opcode in the binary
(`9a <off:LE16> <seg:LE16>`). **Every single far-call segment value in
the overlay is `0x108`** — 2 970 distinct call sites, all targeting
that one segment.

The most-popular targets — likely the C-runtime entry points Lotus
links against — are:

| Target offset | Sites | Plausible role |
|--------------:|------:|---|
| `0x386e` | 484 | (something universal — `memcpy`/`malloc` shape?) |
| `0x3939` | 289 | |
| `0x4690` | 189 | |
| `0x3858` | 123 | |
| `0x490c` | 148 | |
| `0x475c` | 140 | |
| `0x4c22` |  77 | |
| `0x3956` | **156** | **file-open helper** (the one the help dispatch calls) |

Only *some* of these are actually defined in the resident MZ image at
the address you'd compute with the standard formula
`file_off = 512 + seg*16 + off`. For `0x108:0x3956` that gives file
`0x4BD6`, where ndisasm shows three bytes of garbage
(`38 5b c3 = cmp [bp+di-0x3d], bl`) followed by tiny stub functions
and a clean prologue at file `0x4BE0`. Either:

- segment `0x108` is **not** literally "image paragraph 0x108"
  (Lotus's loader may relocate it differently — maybe to a
  trampoline/thunk segment installed at runtime), or
- `0x3956` lands inside a thunk table whose entries are
  3–9-byte fragments (matches what we see — a run of single-
  instruction thunks at file `0x4BCF..0x4BDF`).

The thunk-segment theory fits Lotus's design: a fixed runtime segment
exposed at a stable address, with a table of jump-stubs at fixed
offsets. Every overlay sees the same `0x108:<off>` mnemonic for the
same library function regardless of where the overlay itself loads.

### Three actionable consequences

1. **Disassembling at file `0x4BD6` and following calls outward
   probably won't lead to the decoder directly.** That address is in a
   thunk table; the actual functions live elsewhere in the resident
   image and the thunk just `jmp far`s to them. We'd need to read the
   3–9 bytes at `0x108:0x3956` as a thunk and follow *its* jump.

2. **The help-decode routine is one of the seg=0x108 entries
   *with a small caller count*.** Generic runtime functions (open,
   read, malloc, printf) get called everywhere. A help-specific entry
   point will have only a few callers, all in the help-rendering path.
   The 484 / 289 / 156 hot spots are not the decoder. Look at low-
   count entries (1–3 callers) whose call sites are clustered near
   the `123.HLP` literal at `0x7e986`.

3. **The dispatch caller is somewhere we can find.** The dispatch
   function ends with `retf 6`, so its caller did `call far <seg>:<off>`
   passing a far pointer + word arg. Searching the overlay for
   `9a <off> <seg>` where `<seg>:<off>` decodes to file offset
   `0x7E920` would pinpoint the caller — but we don't know the
   overlay's runtime segment yet, so this is a chicken-and-egg.

### No video segment immediates anywhere

Confirmed: zero hits for `b8 00 b8` / `b8 00 b0` / `push 0xb800` /
`push 0xb000` across the whole 956 KB. Lotus does *not* poke video
memory directly — output goes through the `123VIDEO.*` driver. So the
decoder's "where do I write characters" target is an indirect call
through a video driver vtable, not a fixed `0xB800:0` write. That
rules out chasing video writes as a way to find the decoder.

### Where this leaves the project

We now know roughly *what* the call graph around the help system
looks like, but not where the decode routine lives. Three viable next
steps remain, in increasing cost:

1. **Disassemble the resident image with proper analysis** (Ghidra
   GUI, or IDA / radare2 with manual segment hints). With the
   `0x108`-thunk-segment hypothesis above, you can manually map
   segment `0x108` to a working memory block in Ghidra and the
   decompiler will follow the thunks. That's the fastest path to the
   decoder.

2. **Search the overlay for far calls to the dispatch.** Pick
   plausible runtime segments for the dispatch overlay (any segment
   value such that `<seg>*16 + 0x?? == 0x7E920`), grep for
   `9a 20 e9 <seg>` patterns, look at the bytes preceding each match.
   Cheaper but tedious.

3. **Skip the call graph; pattern-match the decoder by shape.** A BPE
   decoder's inner loop reads bytes, branches on a threshold, and
   indexes a 4-byte table. Look for sequences like
   `mov bx, ax; shl bx, 2; mov ax, [bx+<dict_base>]` in the overlay.
   Speculative but cheap to grep.

I tried option 3 lightly and didn't get a clear hit; the binary is
huge enough that pattern-matching is noisy without a tighter signal.

## Phase 2 plan

1. Capture ground truth: in DOSBox, run 1-2-3 R3.4a, F1 through every
   help screen, transcribe text to `tests/help_groundtruth/<topic>.txt`.
   At least cover the 8 topics already in `crate::help::HELP_TOPICS`.
2. Parse the dictionary at 0x10..0x1b8 into `Vec<(i16, i16)>`.
3. Parse the index/offset structure beyond 0x1bc to map *topic id →
   chunk offset* (need to scan past 0x290 first to understand the
   middle section).
4. Implement a BPE decoder: given a starting offset, a bit reader that
   consumes 10/12-bit codes (try both), and the dictionary, emit
   expanded byte stream.
5. Validate against ground truth; iterate on bit-width and dictionary
   index direction.
6. Wrap as `l123-help` crate (probably) that exposes `topics() ->
   Vec<HelpTopic>` for runtime consumption. The `&'static str` body
   shape in `help.rs` is intentionally compatible with this.

## Risks / unknowns

- The dictionary index direction: are negative slot numbers `-N`
  resolved by indexing `dict[N-1]` from the start, or `dict[len-N]`
  from the end?  Validate both with one decoded sample.
- The bit-width of codes in the content stream (8? 9? 10? 12? variable
  length?). Frequency analysis of the content section may reveal it.
- Non-BPE structure in the index (cross-references, hyperlinks,
  category tree) is not yet identified.
- 1-2-3 R3.x had localized help variants — the file dated 1993 is the
  US English version; other locales would need their own decoders.

Once decoding works, the `crate::help::HELP_TOPICS` table becomes
auto-generated at build time from a small `build.rs` that runs the
decoder over `~/Documents/dosbox-cdrive/123R34/123.HLP` (or a redacted
copy committed alongside, license permitting).
