# Findings

## Current run

| | |
|---|---|
| Commit | `9474a170` (main) |
| Date | 2026-09-03 |
| Runtime | .NET SDK 10.0.400, Linux x64, Release build |
| Input | `TestFiles/NVCA-Model-COI.docx`, 147,622 bytes |
| Input SHA-256 | `d75600769c12724990de48149d7a2bb161f3522daa54b1783672f93697d87d29` |
| Edit script | `edits/nvca-coi.json` (8 edits) |
| Command | `dotnet run --project benchmarks/complex-form-doc -c Release -- TestFiles/NVCA-Model-COI.docx --out <dir>` |
| Exit code | 0 (`ALL CHECKS PASSED`) |

### Stable assertions

These are the harness's contract. They are expected to hold on every run of any
comparable form document, and a change that breaks one is a regression.

| Check | Result |
|---|---|
| `round-trip text exact` — open/save with no edits preserves every text run | PASS |
| `round-trip parts identical` — open/save with no edits preserves the part inventory | PASS |
| `accept single revision` — one tracked revision accepts in isolation | PASS |
| `reject single revision` — one tracked revision rejects in isolation | PASS |
| `tracked save adds no schema findings` — output validator count ≤ source baseline | PASS |
| `accept-all == modified text` — accepting the whole redline reproduces the revised document | PASS |
| `reject-all == baseline text` — rejecting the whole redline reproduces the original | PASS |
| `redline adds no schema findings` — output validator count ≤ source baseline | PASS |

The source document's own validator baseline is **80 findings**; every output matched it
exactly, so the toolchain added none.

### Indicative measurements

Single run, cold process, one machine. These are *not* benchmark-grade numbers and no
threshold is asserted on them — they exist to catch order-of-magnitude drift.

| Stage | Time | Output |
|---|---|---|
| `html (footnotes+headers rendered)` | 1,469 ms | — |
| `markdown projection` | 384 ms | 214,469 chars, 462 anchors, all 94 footnotes projected inline |
| `DocxDiff compatibility probe` | 71 ms | 0 warnings |
| `no-edit session round-trip` | 365 ms | — |
| `tracked session: full edit script` | 3,723 ms | 8/8 edits applied, 11 native revisions, author `Series A Counsel` |
| `clean session: same edit script untracked` | 1,182 ms | — |
| `DocxDiff.Compare` | 667 ms | — |
| `DocxDiff.GetRevisions` | 285 ms | 20 revisions |
| `DocxDiff edit script + semantic changes` | 1,397 ms | 46,460 / 67,164 chars |
| `DocxDiff round-trip invariants` | 779 ms | — |
| `redline -> HTML with tracked-change markup` | 499 ms | — |

The harness prints a stage's own detail line before its timing line, so read the two
together per row above rather than in program output order — the earlier table here paired
the markdown projection's size with the HTML stage. Times are from a different machine than
the 2026-09-02 run and are uniformly lower; nothing changed by an order of magnitude, which
is all these numbers are for.

Revision counts are one-run observations of a specific edit script against a specific
document; they are reported, not asserted. The 8-edit script yields 11 tracked-session
revisions (some edits split across runs) and 20 `DocxDiff` revisions (word-level
granularity over the same edits).

## Open issues this benchmark surfaced

1. **Formatting-only edits that cross a field envelope surface as delete+insert pairs
   in the native redline.** Italicizing a span that intersects a cross-reference field
   produces a full del+ins of the field region rather than a formatting revision —
   correct OOXML, but noisy for a human reviewer. The semantic changeset already
   classifies it precisely (`run_formatting` modify on the exact token span plus a
   `field` envelope change), so this is a markup-shaping improvement, not a diff bug.

2. **`DocxDiffRevision` has no useful `ToString()`** — logging a revision prints the type
   name. Minor, but it makes agent traces harder to read.

3. **Two tracked-changes knobs are easy to conflate.** The projection-side knob
   (`ProjectionSettings.TrackedChanges`) is separate from the mutation-recording knob
   (`SetTrackedChanges`): an agent that records tracked edits and then projects sees
   clean text unless it also sets the projection mode.

## Fixed since this benchmark reported them

**`DocxSession.DeleteBlock` left orphaned footnote definitions** (fixed by #591). Deleting a
paragraph whose text carried a footnote reference removed the reference but left the
footnote body in `word/footnotes.xml`. Word rendered nothing, the note being unreferenced,
but the text still shipped inside the file — for legal workflows a confidentiality-adjacent
leak, where "deleted" drafting commentary survives in the package.

`Docxodus/Internal/NoteReferenceOps.cs` now prunes definitions whose last reference an op
removed, through the same Word-faithful pruner revision resolution has used since #516. It
is revision-aware, as the finding asked: a tracked delete keeps the note until the revision
is accepted. `Docxodus.Tests/DocxSessionNotePruneTests.cs` holds the regression net —
`DS640`–`DS649` cover `DeleteBlock`, shared references, pre-existing danglers, endnotes,
tracked deletion, range and section deletes, table row/column deletes and undo on built
fixtures, and `DS650` re-proves the case as this benchmark reported it, by deleting a
footnote-bearing paragraph of `TestFiles/NVCA-Model-COI.docx` and asserting the charter's
94 note definitions become 93 with no new validator finding.

## Historical: the 2026-08-26 run (superseded)

The first run of this harness, at commit `c8e13d2`, additionally measured the legacy
`WmlComparer` engine alongside `DocxDiff`. **That stage no longer exists** — `WmlComparer`
was removed from the library in v11.0.0 (#643), and the harness's legacy phase went with
it. The observations below are kept only as the record of why that engine was retired;
they do not describe any behaviour of current `main`.

- `WmlComparer` failed both round-trip invariants on this document: accept-all and
  reject-all each diverged from the expected text, because the engine silently dropped an
  empty paragraph inside a footnote.
- It rewrote the package wholesale — validator findings dropped 80 → 29, i.e. it
  normalized markup it should have preserved.
- It was roughly 3× slower than `DocxDiff` (6.3–7.5 s vs 1.5–2.4 s).
- `DocxDiff` passed both invariants exactly on the identical inputs, which it still does.

Two further findings from that run have since been fixed and are not repeated above:
`@docxodus/export`'s Chromium sandbox posture (now covered by `docxodus doctor`), and the
absence of a preflight environment check.
