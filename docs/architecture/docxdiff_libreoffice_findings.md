# DocxDiff vs LibreOffice — parity findings

Verification of the `DocxDiff` engine against LibreOffice's *Compare Documents*
on a real contract (`NVCA-Model-SPA-10-28-2025-1.docx`: 405 body paragraphs,
3 tables, 6 headers, 8 footers, 111 footnotes, custom styles & numbering).

Harness: `tools/diffharness/` (see
`docs/superpowers/specs/2026-06-23-docxdiff-libreoffice-parity-design.md`).
Oracle stance: **correctness-first** — LibreOffice is a reference; we fix ours
only when ours is genuinely wrong (loss / invalid output / semantic error /
divergence from the blessed `WmlComparer` oracle). Where LibreOffice is merely
cruder, we keep ours and document.

## Round-1 survey (22 scenarios)

Legend: round-trip `acc/rej` = `accept==right` / `reject==left`. `LO` = redline
count from LibreOffice's own compare. `seenByLO` = redlines LibreOffice
recognizes in **our** output (rendering proxy). `hf#` = header/footer part count
ours/original (>orig ⇒ duplicate-part bloat).

| scenario | body | notes | hdrftr | fine | LO | seenByLO | hf# | verdict |
|---|---|---|---|---|---|---|---|---|
| body-replace-word | Y/Y | Y/Y | Y/Y | 2 | 2 | 2 | 26/14 | ✅ match |
| body-insert-word | Y/Y | Y/Y | Y/Y | 1 | 1 | 1 | 26/14 | ✅ match |
| body-delete-word | Y/Y | Y/Y | Y/Y | 1 | 1 | 1 | 26/14 | ✅ match |
| body-replace-phrase | Y/Y | **n/n** | Y/Y | 4 | 2 | 4 | 26/14 | 🐞 footnote reorder |
| body-insert-paragraph | Y/Y | Y/Y | Y/Y | 1 | 1 | 1 | 26/14 | ✅ match |
| body-delete-paragraph | Y/Y | Y/Y | Y/Y | 1 | 2 | 1 | 26/14 | 📘 LO coarser |
| body-move-paragraph | Y/Y | Y/Y | Y/Y | 2 | 2 | 4 | 26/14 | 📘 LO no move support |
| body-split-paragraph | Y/Y | Y/Y | Y/Y | 1 | 2 | 1 | 26/14 | 📘 granularity |
| format-bold-run | Y/Y | Y/Y | Y/Y | 1 | 0 | 2 | 26/14 | 📘 LO ignores format |
| format-italic-run | Y/Y | Y/Y | Y/Y | 1 | 0 | 10 | 26/14 | 📘 LO ignores format |
| format-fontsize-run | Y/Y | Y/Y | Y/Y | 1 | 0 | 2 | 26/14 | 📘 LO ignores format |
| format-color-run | Y/Y | Y/Y | Y/Y | 1 | 0 | 10 | 26/14 | 📘 LO ignores format |
| format-underline-run | Y/Y | Y/Y | Y/Y | 0 | 0 | 0 | 26/14 | ⚠️ test artifact (anchor already underlined) |
| style-change-paragraph | Y/Y | Y/Y | Y/Y | 0 | 0 | 0 | 26/14 | 📘 pStyle not a tracked rev (matches oracle) |
| table-cell-edit | Y/Y | Y/Y | Y/Y | 2 | 2 | 2 | 26/14 | ✅ match |
| table-cell-insert-word | Y/Y | Y/Y | Y/Y | 1 | 2 | 1 | 26/14 | 📘 LO coarser |
| table-insert-row | Y/Y | Y/Y | Y/Y | 1 | 2 | 1 | 26/14 | 📘 LO coarser |
| table-delete-row | Y/Y | Y/Y | Y/Y | 1 | 2 | 1 | 26/14 | 📘 LO coarser |
| header-edit | Y/Y | Y/Y | **n/n** | 0 | 0 | 0 | 26/14 | 🐞 hdr/ftr dup; hdr not diffed (matches oracle) |
| footer-edit | Y/Y | Y/Y | **Y/n** | 0 | 0 | 0 | 26/14 | 🐞 hdr/ftr dup |
| footnote-edit | Y/Y | Y/Y | Y/Y | 1 | 0 | 1 | 26/14 | ✅ ours detects, LO ignores footnotes |
| multi-edit | Y/Y | Y/Y | Y/Y | 5 | 4 | 6 | 26/14 | ✅ ours finer |

**Body content round-trips perfectly in all 22.** All 22 outputs open clean in
LibreOffice and our redline markup is recognized.

## Classified findings

### 🐞 FIX — genuine defects (ours wrong vs the blessed oracle)

- **F1. Header/footer parts are duplicated in every `DocxDiff.Compare` output.**
  Output carries the LEFT package's header/footer parts (`header1.xml`…) **plus**
  the RIGHT document's, re-imported as `P<guid>.xml` (26 vs 14 parts). When a
  header/footer differs between sides, the output contains **both** versions.
  The `WmlComparer` oracle is clean (14 parts, left's content only).
  - **Root cause:** `IrMarkupRenderer` clones `EqualBlock`s from the RIGHT
    document (by design — right carries accepted-state rsid/format). The base's
    section-break paragraphs carry an inner `w:sectPr` with header/footer
    references; cloned from the right they reference RIGHT's header/footer parts,
    and `ImportRightSourcedMedia → WmlComparer.MoveRelatedPartsToDestination`
    then copies those parts into the left-based package as `P<guid>` duplicates.
  - **Fix direction:** header/footer scopes are deliberately NOT diffed, so the
    LEFT package's header/footer parts are authoritative. The renderer must not
    import RIGHT header/footer parts; section-break references on right-sourced
    Equal blocks should resolve to the LEFT package's existing parts.

- **F2. Footnotes can be reordered when a body edit touches a footnote-bearing
  paragraph.** `body-replace-phrase` — footnote store same length (43060) and
  count (111) but two footnotes swapped order; round-trip notes check fails.
  Same root-cause class as F1 (right-sourced clones import/remap RIGHT note
  content). No content loss observed, but order/anchoring must be verified.

### 📘 DOCUMENT — LibreOffice is cruder (keep ours, no fix)

- LibreOffice **ignores format-only changes** (bold/italic/size/color → 0
  redlines). Ours detects them as `w:rPrChange`. Word agrees with ours.
- LibreOffice **ignores footnote changes** (footnote-edit → 0). Ours detects.
- LibreOffice has **no move detection** (move → 12 del + 12 ins, or whole-para
  del+ins). Ours emits native `w:moveFrom`/`w:moveTo`.
- LibreOffice does **whole-region replacement** in tables — a 2-word cell edit
  produced **27 del + 27 ins** vs ours' 2. Ours is far more precise.
- Granularity differences (delete-paragraph, split, table edits): LibreOffice
  often reports 2 where ours reports 1; both are correct, different atomization.

### ⚠️ TEST ARTIFACTS (harness, not engine)

- `format-underline-run`: the heading anchor already carries `<w:u>`, so adding
  underline is a no-op → fine=0. Underline **is** modeled
  (`IrModeledFormat.cs:46`). Fix: target a non-underlined run.

## Fix log

_(updated as fixes land)_
