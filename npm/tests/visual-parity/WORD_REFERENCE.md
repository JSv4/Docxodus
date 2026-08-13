# Word-reference capture procedure

This is the issue-#402 procedure for obtaining **Microsoft Word evidence** for the visual-parity
corpus. LibreOffice is a comparison implementation, not the correctness oracle; when Docxodus and
LibreOffice disagree and the OOXML is ambiguous, Word's rendering is the tie-breaker. This
procedure turns that tie-breaker from a one-off anecdote into recorded, reproducible data.

## What is committed, and what is not

`word-reference.json` is a committed, numbers-only record mirroring `ratchet.json`'s philosophy:

- **Committed:** per-case page counts, page dimensions, ink extents, named per-case
  measurements (`keyMeasurements`), operator notes, the fixture's SHA-256, and the Word/OS
  versions used.
- **Never committed:** the Word-exported PDFs, rasterized PNGs, or any other binary. No
  proprietary rendering enters the repository; only facts about it.

Every corpus case has a row. `pending` rows are the coverage ledger — a new corpus case gets a
pending row (the pure spec `visual-parity-word-reference.spec.ts` fails on every PR until it
does), so unmeasured evidence is visible instead of forgotten.

## What Word evidence can and cannot decide

Word renders with the **genuine Office fonts**. The benchmark's two engines share license-safe
substitutes under the font-substitution contract (issue #379), so glyph-level pixel scores
against Word are depressed by font differences that have nothing to do with correctness.

- **Decisive:** structural questions. Page counts. Whether Word suppresses a declared
  `w:spacing w:after` before a nested table. Whether heading space-before is dropped at a page
  top. Where a table's first ink row sits relative to the margin. These survive font
  substitution because they are geometry of the layout model, not of glyphs.
- **Advisory:** SSIM / ink-F1 against Word (recorded under `comparisons` when a benchmark run
  is supplied). Useful for spotting gross divergence; never a gate and never sufficient
  evidence by itself.

A `reference-deviation` disposition should cite recorded Word data via the corpus entry's
`disposition.wordEvidence` field once its fixture is measured. The pure spec enforces the
converse: a `wordEvidence` citation whose row is not `measured` fails, so a rationale can never
claim Word data that was never captured.

## Prerequisites

- A licensed Microsoft Word. Record the exact version (File → Account → About Word) — e.g.
  "Word for Microsoft 365 MSO (16.0.18129.20158) 64-bit" — and the OS (e.g. "Windows 11 24H2").
- A clean checkout of this repository at a known commit (the capture records each fixture's
  SHA-256, so a modified fixture is detectable).
- On the machine that runs the capture command (need not be the Word machine): Node + `npm ci`
  in `npm/`, and `poppler-utils` (`pdftoppm`). LibreOffice is **not** required.

## Procedure

1. **Export one PDF per corpus case** in Word, named `<case-id>.pdf` (the ids are the `id`
   fields in `corpus.ts`, e.g. `nested-table.pdf`). For each fixture:
   - Open the tracked fixture path from `corpus.ts` (e.g. `TestFiles/WC/WC043-Nested-Table.docx`).
     Do not edit, repaginate in a different view, or save the DOCX.
   - For cases with `revisionMode: 'accepted'` (currently `tracked-deletion`): set
     **Review → Tracking → No Markup** first, so Word renders the final view the benchmark
     compares. Note this in the case's `notes`.
   - **File → Save As → PDF** (standard/print quality). Word's PDF export renders the current
     document view; do not use a third-party PDF printer, which would re-rasterize through a
     different engine.
   - If Word prompts about missing fonts or converts anything on open, record it in `notes`.
2. **Run the capture** on the directory of PDFs:

   ```bash
   cd npm
   DOCXODUS_WORD_REFERENCE_PDFS=/path/to/word-pdfs \
   DOCXODUS_WORD_VERSION="Word for Microsoft 365 MSO (16.0.18129.20158) 64-bit" \
   DOCXODUS_WORD_OS="Windows 11 24H2" \
   npm run capture:word-reference
   ```

   The capture rasterizes each PDF with the benchmark's own contract (Poppler at exactly
   96 DPI, `C.UTF-8`, UTC), measures pages with the same ink model as the pairwise metrics,
   and updates `word-reference.json` in place. A partial directory is fine — only the supplied
   cases are re-measured; operator-authored `notes`/`keyMeasurements` survive re-capture.
3. **Optionally record the three-way comparison** by also passing a completed benchmark run:

   ```bash
   DOCXODUS_WORD_REFERENCE_RUN=/tmp/docxodus-visual-parity ... npm run capture:word-reference
   ```

   This adds per-case advisory `comparisons` (Docxodus-vs-Word, LibreOffice-vs-Word) tied to
   the run's commit.
4. **Add the measurements that decide the open question.** The automatic extraction records
   page geometry and ink extents; the semantically named coordinates a disposition needs
   (e.g. "gap between the `Before.` paragraph baseline and the nested table's top border, in
   px") are read off the rasterized pages by the operator and added to the case's
   `keyMeasurements` in `word-reference.json` by hand. Name them in px at 96 DPI.
5. **Annotate the disposition.** In `corpus.ts`, set `disposition.wordEvidence` to a one-line
   citation of the recorded data and what it decides, and update the `rationale`/`kind` if the
   evidence changed the attribution. Update BASELINE.md's narrative for any disposition the
   evidence changed.
6. **Commit the diff** (`word-reference.json`, `corpus.ts`, BASELINE.md). The pure spec
   validates the record's consistency on every PR.

## Open questions this data exists to decide

Maintained as dispositions change; see each case's `rationale` in `corpus.ts` for the live list.

- `nested-table` — does Word paint or suppress the document-default `w:spacing w:after="160"`
  on the paragraph preceding a nested table? (Docxodus paints; LibreOffice suppresses.)
- `legal-contract` — does Word drop heading space-before at the top of a page? (LibreOffice
  drops; Docxodus paints the declared value.)
- `fields-and-tabs` — **resolved by issue #427:** Word records 0 blue pixels in the cached TOC
  result, so field context suppresses hyperlink presentation; an ordinary same-style link remains
  blue and underlined.
- `shape` — does the `a:ln` outline enlarge an auto-fit shape in Word? (DrawingML says no;
  CSS borders do.)
- The #404 reductions (`landscape-section`, `inline-image`, `tracked-deletion`) — whether
  Docxodus's same-font layout choices (paragraph spacing, wrap points, heading metrics) are
  also Word-correct.
