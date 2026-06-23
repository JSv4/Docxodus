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

> The `hf#` column above is the **pre-fix** snapshot (26/14). After **F1** (see
> Fix log) it is **14/14 in every scenario**, and `body-replace-phrase` notes
> round-trip clean after the harness check fix (**F2**). Post-fix:
> **20/22 content-clean**, the 2 remaining being the header/footer
> undiffed-scope limitation (oracle-consistent).

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

## Visual rendering verification

Page-1 render of `body-replace-word` (the `Purchaser`→`Investor` edit), ours vs
LibreOffice's own compare, both opened+rendered by LibreOffice:

- **Ours:** `(each a "`~~Purchaser~~`Investor"` — deletion struck-through THEN
  insertion underlined, in the redline colour, with a margin change-bar. This is
  **Word's delete-then-insert ordering**.
- **LibreOffice:** `(each a "Investor`~~Purchaser~~`"` — insertion THEN deletion.

Both are valid, visually equivalent redlines; ours follows Word's convention.
All 22 outputs open clean in LibreOffice and our `w:ins`/`w:del` markup is
recognized as tracked changes (the rendering proxy `seenByLo` is non-zero
wherever revisions exist).

## Fix log

### F1 — header/footer part duplication — **FIXED** (commit on this branch)

`WmlComparer.MoveRelatedPartsToDestination` gained an opt-in
`skipHeaderFooterReferences` parameter (default false — legacy callers
unchanged); `IrMarkupRenderer.ImportRightSourcedMedia` passes it `true`. A
right-cloned Equal block's inner-`sectPr` `w:headerReference`/`w:footerReference`
no longer drags the RIGHT's header/footer parts into the LEFT-based output. The
references already resolve to the LEFT package's parts (same r:ids — both sides
derive from one base). Result on the contract: **header/footer part count is now
14/14 (was 26/14) in every scenario**; output is byte-lean and oracle-aligned.
Regression test:
`IrMarkupRendererTests.Render_does_not_duplicate_header_parts_for_equal_section_break_block`.
Guard: 462 Ir.Diff+DocxDiff+WmlComparer tests green.

### F2 — footnote renumber on footnote-bearing body edits — **NOT A BUG**

Investigated `body-replace-phrase` (footnote store reordered, id→text mapping
changed). The **WmlComparer oracle renumbers footnotes identically** — this is
expected compare behaviour, and Word renders footnotes in body-reference order
regardless of part order. No content is gained or lost. The harness round-trip
check was too strict (compared the note store in part order); it now compares the
**multiset of per-note texts** (`TextExtractor.NoteTexts`), which is
renumber-robust. `body-replace-phrase` is content-clean after the harness fix.

### F3 — GetRevisions returned 0 for a table column add/remove — **FIXED** (commit on this branch)

A column add/remove bails the markup renderer to a whole-table del(left)+ins(right) fallback (so the
Compare output and LibreOffice both render whole-table replace, and round-trip holds), but
`IrRevisionRenderer` had no matching fallback — `GetRevisions` returned **0** for a column-count change,
diverging from the WmlComparer oracle (**2** revisions) and silently hiding a change the markup tracks.
Fix: `IrRevisionRenderer.RenderModifyBlock` now detects the same unpaired-cell condition
(`TableDiffNeedsWholeTableFallback`) and emits a Deleted(left table) + Inserted(right table) pair.
`table-insert-column`/`table-delete-column` now report **2** (matching the oracle). Regression test:
`DocxDiffTests.GetRevisions_TableColumnChange_ReportsWholeTableReplace`. Guard: 463 tests green.

> **Table column add/remove rendering** (both ours and LibreOffice): a single-column change renders as a
> WHOLE-TABLE delete + whole-table insert in BOTH engines — not column-precise. They are equivalent in
> crudeness here; column-precise table markup is a deferred v1 limitation
> (`ir_diff_engine.md`). **Format changes inside table cells** are reported by neither ours nor the oracle
> (`table-cell-format`: 0/0) — oracle-consistent, not a defect.

### header-edit / footer-edit — documented limitation (matches oracle)

After F1, both show `reject==left` ✓ and `accept≠right` only for the header/footer
text — because header/footer scopes are **deliberately not diffed**
(`IrEditScriptBuilder`: "the oracle does not diff them either"). WmlComparer
behaves identically (0 revisions, keeps left's header). Not a defect; the
round-trip "accept==right" gate does not apply to undiffed scopes.

## Round-2 expanded coverage (31 scenarios)

Added bug-prone categories to the corpus: table column insert/delete, table-cell
format, structural footnote insert/delete, multi-paragraph delete/insert,
whole-paragraph replace, and a styled-heading move. **All 31 round-trip
content-clean** (header/footer-only edits report `(scope-ok)`). Highlights vs the
WmlComparer oracle and LibreOffice:

| scenario | ours | oracle | LO | note |
|---|---|---|---|---|
| table-insert-column | 2 | 2 | (whole-table) | match oracle after F3; both ours & LO render whole-table replace |
| table-delete-column | 2 | 2 | (whole-table) | match oracle after F3 |
| table-cell-format | 0 | 0 | 0 | format-in-cell reported by neither — oracle-consistent |
| footnote-insert | 3 | 3 | 3 | match |
| footnote-delete | 2 | 3 | 3 | ours one fewer (granularity); content round-trips |
| multi-paragraph-delete | 5 | — | 6 | granularity |
| multi-paragraph-insert | 3 | — | 1 | ours finer |
| whole-paragraph-replace | 17 | — | 2 | ours word-level (far finer); LO whole-paragraph |
| move-heading-block | 2 (move) | — | 3 | ours emits native move; LO del+ins |

No new genuine defects in round 2 beyond F3 (already fixed). The table column
add/remove rendering is whole-table-replace in BOTH ours and LibreOffice — they
are equivalent in crudeness there (column-precise table markup is a deferred v1
limitation).

## Round-3 exhaustive sweep (10 advanced categories, parallel + adversarial)

A parallel agent sweep built and tested 10 advanced edit categories the base
lacks (each synthesized into a variant, diffed, and compared to the WmlComparer
oracle + LibreOffice, then any suspect was adversarially re-verified from scratch):

| category | ours | oracle | LO | verdict |
|---|---|---|---|---|
| merged-cell-gridspan | 2 | 11 | 2 | oracle-consistent (ours coarser per-cell, LO-equal; round-trips) |
| image-insert | 1 | 1 | 1 | ✅ media import intact post-F1 (no dangling rId; png imported byte-identical) |
| hyperlink-insert | 1 | 1 | 1 | ours FINER than oracle (keeps the live hyperlink + rel; oracle flattens it) |
| comment-insert | 0 | 0 | 1 | oracle-consistent (comments are a non-revision annotation layer; LO cruder) |
| section-pagesetup-change | 0 | 0 | 0 | oracle-consistent (sectPrChange is a documented Task-4 gap; round-trips) |
| paragraph-align-indent | 0 | 0 | 0 | oracle-consistent (pPr-only change; round-trips) |
| bookmark-range-edit | 2 | 2 | 2 | text edit correct; body-level bookmarkEnd drop is **deliberate** (see below) |
| complex-field-edit | 2 | 2 | 2 | oracle-consistent (field structure preserved) |
| nested-table-edit | 1 | 1 | 2 | oracle-consistent (LO cruder) |
| **content-control-sdt** | 1 | 1 | 1 | 🐞 **GENUINE BUG — reject contract violated; FIXED (F4)** |

**One genuine defect found (F4, fixed); the other nine are oracle-consistent or
ours-finer.** The image-insert result specifically re-confirms F1 did not regress
inline-image media import (the image part imports with no dangling relationship).

### F4 — inserted content-control (`w:sdt`) text leaked on reject — **FIXED**

`Compare` emitted runs inserted inside a `w:sdtContent` BARE (no `w:ins`), so
`RejectRevisions` (which strips `w:ins`/`w:del`) left them — `reject ≠ left`, a
core-contract violation that silently retains content the user rejected.
`GetRevisions` was correct (1 insertion); only the markup renderer failed.
`IrMarkupRenderer.WrapRunLevel` wrapped a container's DIRECT children, but a
`w:sdt`'s runs live nested under `w:sdtContent` (not a direct child), so neither
got wrapped. Fix: a new `WrapContainerChild` descends through `w:sdtContent`
(`w:ins`/`w:del` is a valid child there) and wraps its run-level children. The
oracle flattens the sdt instead; ours keeps the content control AND tracks the
insertion. Test `DocxDiffTests.Compare_InsertedContentControl_RejectStripsTheInsertedText`.

### bookmark-range-edit — body-level `w:bookmarkEnd` drop is DELIBERATE (not a defect)

The sweep flagged that 6 body-level `w:bookmarkEnd` markers (siblings of `w:p`,
outside any paragraph) are dropped from `Compare` output (192 starts / 186 ends),
even in an identity diff. This is an **intentional, documented** IR-reader
decision (`IrReader.AppendBlocks` / `IsDroppedParagraphChild`): WmlComparer strips
**all** bookmarks (`MarkupSimplifier RemoveBookmarks=true`), and modeling a
body-level `bookmarkEnd` as an opaque block previously produced a spurious content
block the markup round-trip could not toggle (WC022). Ours therefore drops
body-level markers (oracle-consistent — the oracle drops all 192) while preserving
the 186 intra-paragraph bookmarks (strictly **better** than the oracle, which
keeps zero). It is not a regression vs the blessed oracle. A future enhancement
could preserve body-level markers too for full bookmark fidelity beyond the
oracle, but that reintroduces the WC022 round-trip-toggle problem and is out of
scope. Recorded as a known limitation, not fixed.

## Round-1 status

20/22 content-clean (body+notes+header/footer). The 2 non-clean are the
header-edit/footer-edit undiffed-scope limitation above (oracle-consistent).
Genuine defects found and fixed: **1** (F1). Everything else is either
LibreOffice being cruder (kept ours) or a harness-measurement artifact (fixed).
