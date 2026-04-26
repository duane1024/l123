# 123.HLP decoding — reconnaissance notes

Working file: `~/Documents/dosbox-cdrive/123R34/123.HLP`, 457302 bytes,
dated Mar 17 1993.

This document captures what's known about the format so far. The
runtime help system in `crates/l123-ui/src/help.rs` is wired with
hand-authored topics; replacing those with `.HLP`-decoded content is
this document's eventual target.

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
