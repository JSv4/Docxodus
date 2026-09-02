# Findings

## Current run

| | |
|---|---|
| Commit | `18eb50ea` (main) |
| Date | 2026-09-02 |
| Runtime | .NET SDK 10.0.301, Linux x64, Release build |
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
| `html (footnotes+headers rendered)` | 1,569 ms | 214,469 chars |
| `markdown projection` | 542 ms | 462 anchors, all 94 footnotes projected inline |
| `DocxDiff compatibility probe` | 108 ms | 0 warnings |
| `no-edit session round-trip` | 725 ms | — |
| `tracked session: full edit script` | 3,814 ms | 8/8 edits applied, 11 native revisions, author `Series A Counsel` |
| `clean session: same edit script untracked` | 1,793 ms | — |
| `DocxDiff.Compare` | 1,397 ms | — |
| `DocxDiff.GetRevisions` | 721 ms | 20 revisions |
| `DocxDiff edit script + semantic changes` | 2,659 ms | 46,460 / 67,164 chars |
| `DocxDiff round-trip invariants` | 1,960 ms | — |
| `redline -> HTML with tracked-change markup` | 1,095 ms | — |

Revision counts are one-run observations of a specific edit script against a specific
document; they are reported, not asserted. The 8-edit script yields 11 tracked-session
revisions (some edits split across runs) and 20 `DocxDiff` revisions (word-level
granularity over the same edits).

## Open issues this benchmark surfaced

1. **`DocxSession.DeleteBlock` leaves orphaned footnote definitions.** Deleting a
   paragraph whose text carries a footnote reference removes the reference but leaves
   the footnote body in `word/footnotes.xml`. Word renders nothing (the note is
   unreferenced), but the text still ships inside the file — for legal workflows this is
   a confidentiality-adjacent leak: "deleted" drafting commentary survives in the
   package. Options: prune unreferenced notes on delete, prune on `Save()`, or expose a
   `CompactFootnotes`-style op; whichever is chosen should be revision-aware (a tracked
   delete must keep the note until the revision is accepted).

2. **Formatting-only edits that cross a field envelope surface as delete+insert pairs
   in the native redline.** Italicizing a span that intersects a cross-reference field
   produces a full del+ins of the field region rather than a formatting revision —
   correct OOXML, but noisy for a human reviewer. The semantic changeset already
   classifies it precisely (`run_formatting` modify on the exact token span plus a
   `field` envelope change), so this is a markup-shaping improvement, not a diff bug.

3. **`DocxDiffRevision` has no useful `ToString()`** — logging a revision prints the type
   name. Minor, but it makes agent traces harder to read.

4. **Two tracked-changes knobs are easy to conflate.** The projection-side knob
   (`ProjectionSettings.TrackedChanges`) is separate from the mutation-recording knob
   (`SetTrackedChanges`): an agent that records tracked edits and then projects sees
   clean text unless it also sets the projection mode.

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
