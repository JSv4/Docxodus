# Visual parity baseline and remediation log — 2026-08-09

This is the first stratified Docxodus-versus-LibreOffice full-page baseline. It is diagnostic, not a
claim that LibreOffice is always correct and not yet a release gate.

## Relationship to existing tests

The repository already had two narrower forms of visual coverage:

- `tabs-visual.spec.ts` compares five browser renderings with committed Playwright snapshots. The
  adjacent `__reference__` directory also contains LibreOffice-derived tab PDFs/PNGs and a manual
  root-cause analysis, but the test does not dynamically compare against them.
- `demo-arcade-libreoffice.spec.ts` dynamically compares one exported frame from each arcade game
  with LibreOffice at 96 DPI, checking canvas geometry and tolerant ink F1.

This benchmark reuses the proven headless-LibreOffice approach but adds all-page rendering, a
stratified document corpus, exact and perceptual metrics, bounded alignment, heatmaps, portable JSON,
artifact hashes, page-count checks, revision-view normalization, and a scheduled/manual workflow. It
does not replace the fast committed snapshots or the arcade export test.

## Baseline contract

- Source base: `ee44c351fabc6a2f23589eb78d74864c72013f5f` plus this branch's renderer and
  benchmark changes. A clean CI run will record the eventual PR commit instead.
- Corpus: 12 existing committed fixtures, covering text, tables, lists, sections,
  headers/footers, images, charts, shapes, fields, footnotes, tracked changes, and page geometry.
- Licensing boundary: every fixture is a regular file whose worktree blob must exactly match
  `HEAD:<path>`. No untracked/ignored corpus or harness can enter through the manifest.
- Raster contract: Chromium device scale 1 and pagination scale 1; LibreOffice PDF plus Poppler at
  96 DPI; `C.UTF-8`; UTC; fresh LibreOffice profiles; all pages paired by ordinal.
- Revision contract: cases needing final view are accepted once into a temporary DOCX outside the
  checkout, and identical accepted bytes go to both renderers.
- False-positive controls: wait for fonts/images/two animation frames; disable animation, carets,
  shadows, labels, and page gaps; search only a +/-2 px translation; report page geometry
  independently; use no masks.
- Environment: Chromium 143.0.7499.4; LibreOffice 25.8.7.3; Poppler 25.03.0; Calibri mapped by
  fontconfig to Carlito, Calibri Light to Noto Sans, and Times New Roman to Liberation Serif.

Two clean full render passes produced identical normalized metrics and all 60 Docxodus,
LibreOffice, and overlay PNG SHA-256 hashes. The final ink-F1 edge-case correction changed no image
hashes. Generated reports and images remained under `/tmp` and were not copied into the repository.

## Initial aggregate result

| Signal | Result |
|---|---:|
| Cases | 12 |
| Paired pages | 20 |
| Conversion errors | 0 |
| Page-count mismatches | 1 |
| Case severity | 1 close, 1 minor, 0 major, 10 severe |
| Page severity | 1 close, 1 minor, 2 major, 16 severe |
| Mean SSIM | 0.974586 |
| Mean tolerant ink F1 | 0.394106 |

The low mean ink F1 is intentional: absent or displaced ink scores zero even when most of a white
page is identical and SSIM remains high. A benchmark bug originally returned F1=1 when precision and
recall were both zero; the generated one-sided-content regression now pins the correct zero score.

## Remediation rerun

The baseline's two pending pull requests are now on `main`: pagination remediation PR #372 at
`7199e54` and this benchmark itself in PR #373 at `404d9da`. The branch from that exact fresh main
also repairs the integration-only duplicate `HCO085` test identifier by assigning the footnote and
chart regressions unique `HCO086`/`HCO087` identifiers.

The complete 12-case corpus was rerun after rebuilding the production WASM bundle. The generated
report remained outside the checkout under `/tmp`; no benchmark artifacts or additional corpus files
were added to the repository.

`Current` is one clean-worktree rerun of the whole corpus on this branch after it merged `main`, so
it measures the running-content fix (issue #377) and the DrawingML anchor fix (PR #381) TOGETHER
rather than splicing two separately measured runs. The two fixes touch disjoint cases, and every
per-case figure below reproduced its own branch's measurement exactly; only the corpus-wide means
move.

| Signal | Initial | After PRs #372–#374 | Current |
|---|---:|---:|---:|
| Cases | 12 | 12 | 12 |
| Paired pages | 20 | 21 | 21 |
| Conversion errors | 0 | 0 | 0 |
| Page-count mismatches | 1 | 0 | 0 |
| Case severity | 1 close, 1 minor, 0 major, 10 severe | 2 close, 1 minor, 0 major, 9 severe | 4 close, 1 minor, 0 major, 7 severe |
| Page severity | 1 close, 1 minor, 2 major, 16 severe | 2 close, 1 minor, 1 major, 17 severe | 13 close, 1 minor, 0 major, 7 severe |
| Mean SSIM | 0.974586 | 0.978298 | 0.980074 |
| Mean tolerant ink F1 | 0.394106 | 0.412753 | 0.841237 |

Resolved items:

1. **Pagination count and story inheritance (PR #372).** `running-content` now produces the same
   five pages as LibreOffice instead of four, and both `running-content` and `multi-section` retain
   inherited odd/even/default header and footer stories. Residual header/body/footer vertical
   placement still scores severe and remains a distinct geometry problem.
2. **Executable benchmark baseline (PR #373).** The corpus guard, all-page LibreOffice comparison,
   perceptual/ink metrics, artifacts, and scheduled/manual workflow are now part of `main`.
3. **Cached clustered bar/column rendering (this branch).** `HC043-Chart.docx` moves from a blank,
   severe result (SSIM 0.92933, ink F1 0) to a populated, **close** result (SSIM 0.98687, ink F1
   0.96817, perceptual diff 0.01436). Page count, page size, and bounded alignment all match exactly.
   The renderer consumes the portable chart caches and stored drawing extent, so it does not need
   the optional embedded workbook, a JavaScript chart library, or an Office process.
4. **Running-content vertical placement (issue #377).** `w:header`/`w:footer` are distances from
   the paper edge to the top of the header story and the bottom of the footer story; the paginator
   ignored both and anchored the stories to the MARGINS, pulling each toward the body by exactly
   `margin − distance`. Measured against LibreOffice on `running-content` page 2, the header ink sat
   at y 77–86 against LibreOffice's 52–61 and the footer at 969–978 against 994–1003 — 25 px in
   each direction on a page whose `w:header`/`w:footer` are both 720 twips and whose margins are
   1440. Both tracked cases become **close**: `multi-section` SSIM 0.99586 → 0.99934 with worst ink
   F1 0.00000 → 0.99797, `running-content` 0.99773 → 0.99921 and 0.00000 → 0.96486.
   `landscape-section` improves as a side effect (0.91746 → 0.92607, 0.50174 → 0.60042). Story
   inheritance and page counts are unchanged, which the generated regression asserts alongside the
   coordinates.
5. **DrawingML textbox anchor geometry.** `DB011-Body-With-Shape.docx` now honors its
   column-centered, margin-relative-width and paragraph-offset geometry. The rendered box moves
   from `(273,113)–(520,202)` to `(268,133)–(524,222)`, versus LibreOffice's
   `(268,132)–(524,237)`: horizontal bounds match exactly and the vertical origin is within 1 px.
   SSIM improves from 0.96967 to 0.97428 and ink F1 from 0.50719 to 0.63049. The remaining height
   difference is auto-fit line geometry, independent of the now-correct anchor origin and width.

## Attribution dispositions and the issue #378 rerun — 2026-08-11

Severity measures how different two renderings are; it cannot say whose difference it is. Every
corpus entry now carries a reviewed `disposition` (`renderer-bug` / `environment` /
`reference-deviation` / `unsupported-feature` / `unattributed`, each with a mandatory rationale and
an optional tracking reference — see README). Dispositions flow into `metrics.json` and
`summary.json` (`aggregate.severeByDisposition`, `aggregate.strictGatingCases`), and strict mode now
gates only on severe cases the renderer owns (`renderer-bug`/`unattributed`) plus conversion
errors. New corpus entries default to `unattributed`, which gates, so an untriaged severe case
cannot hide behind a non-gating label.

The whole corpus was rerun twice on this branch (once before, once after the footnote fix) in a
second, honestly-different environment: Chromium 141.0.7390.37, LibreOffice 24.2.7.2, Poppler
24.02.0, Calibri→Carlito, Calibri Light→DejaVu Sans, Times New Roman→Liberation Serif. Every
case reproduced the recorded CI baseline within environment noise, and the cases that moved most
(`footnote`, `landscape-section`, `tracked-deletion`) are exactly the ones attributed
`environment` — the reproduction itself corroborates the attributions. One reference-version
finding: LibreOffice 24.2 draws its legacy 25%-of-column footnote separator where 25.8 draws the
two-inch default, so the separator-width comparison is only meaningful against LibreOffice ≥ 25.8.

**Footnote block vertical placement is fixed (issue #378).** The note area was already anchored at
the correct edge — the bottom margin line — but the paginated container carried web chrome inside
the bottom-anchored box: a 1.4 line-height (Word's FootnoteText is single-spaced), 4pt inter-note
margins (Word stacks notes with zero spacing of their own), a 6pt separator gap (about double
Word's one-empty-line model), an `N.`-style number label, and a trailing `↩` back-reference link
that no print renderer shows. Together they lifted the visible ink ~13px off the margin and drew
false-positive ink. Measured on the tracked case at 96 DPI (bottom margin line at row 960):

| Signal | Before | After | LibreOffice |
|---|---:|---:|---:|
| Note ink rows | 935–948 | 942–955 | 946–955 |
| Separator row | 922 | 935 | 939–940 |
| Note ink x-extent | 97–201 | 97–176 | 96–171 |
| Case SSIM | 0.99247 | 0.99277 | — |
| Case ink F1 | 0.66696 | 0.73700 | — |

The note-text bottom is now flush with LibreOffice's (row 955 exactly), the separator-to-note gap
matches (6–7px both), and the backref ink is gone. The generated regression
`npm/tests/pagination-footnote-geometry.spec.ts` pins the note block and the separator separately:
area bottom on the margin line, last line flush, single-spaced line box, `line-height: normal`,
2in separator one line above the first note, zero inter-note spacing, bare superscript number, no
backref. The case's disposition moves to `environment` on this evidence: the residual is line
metrics of the substituted font differing between Chromium and LibreOffice (issue #379) plus the
LibreOffice-24.2 separator width, and the case's remaining severe score tracks the same body-text
glyph differences as the other environment cases.

| Signal | Current (2026-08-09, CI env) | This branch (local env) |
|---|---:|---:|
| Cases | 12 | 12 |
| Paired pages | 21 | 21 |
| Conversion errors | 0 | 0 |
| Case severity | 4 close, 1 minor, 0 major, 7 severe | 4 close, 1 minor, 0 major, 7 severe |
| Severe by disposition | — | 2 renderer-bug, 4 environment, 1 reference-deviation |
| Strict-gating cases | — | `shape`, `fields-and-tabs` |
| Mean SSIM | 0.980074 | 0.980203 |
| Mean tolerant ink F1 | 0.841237 | 0.857746 |

The aggregate means are not comparable across the environment change (different LibreOffice,
Chromium, and Calibri Light substitutes); the per-case table below carries the current local
figures. The headline number is no longer the raw severe count but the strict-gating set: two
renderer-attributable severe cases remain.

## Font-substitution contract — 2026-08-11 (issue #379)

Font policy is now a shared contract, not a host observation: `fonts.conf` pins each declared
Office family to a license-safe metric-compatible substitute (Calibri/Calibri Light → Carlito,
Cambria → Caladea, Times New Roman/Arial/Courier New → Liberation), and both renderers load it via
`FONTCONFIG_FILE` — LibreOffice per subprocess, Chromium at launch through `playwright.config.ts`.
Enforcement is layered (fc-match assertion with install hints, in-browser canvas-width check,
cross-renderer wrapping probe), and the resolved family/file/version set plus the contract file's
SHA-256 are recorded in every `summary.json`. The probe is negatively validated: without the
contract, Calibri Light wraps 5 lines instead of 4 on a stock Ubuntu host (DejaVu Sans vs
Carlito). See the README's contract section for the full mechanism.

Corpus effect, against the immediately preceding run in the same environment:

- **Nine of twelve cases byte-identical** — this host's defaults already matched the contract for
  their families, confirming the pinning changes nothing except what it claims to pin.
- **`tracked-deletion` improved** (SSIM 0.94833, +0.01688; ink F1 0.58271, +0.04299). The
  renderer-only `Calibri Light → Carlito` fallback was rejected in the initial baseline because it
  worsened this exact case; the SAME mapping shared by both engines improves it — the contract's
  thesis, demonstrated.
- **`fields-and-tabs` measured lower** (ink F1 0.15994, −0.19971): with Calibri Light pinned in
  both engines the TOC line-height mismatch no longer partially overlaps by accident. The case was
  already strict-gating `renderer-bug`; the pinned number is the honest one.

Severity counts, strict gating (`shape`, `fields-and-tabs`), and every disposition are unchanged.
`environment` now has a sharper meaning: the engines lay out the SAME fonts differently
(line breaking, justification, rasterization) — never that they picked different fonts.

## Regression ratchet — 2026-08-11 (issue #395)

Every measurement above was a snapshot nothing defended. The scheduled run uploaded an artifact
that expired in 14 days, and no run was compared against the previous one, so a renderer
regression was caught only if a human downloaded and eyeballed the report in time.

`ratchet.json` is now a committed, numbers-only record — one row per case carrying page counts,
severity, mean SSIM, worst ink F1, and the disposition — and every run compares against it. It is
deliberately broader than strict mode: strict gates only severe cases the renderer owns, whereas
"no case may get worse than recorded" covers all twelve at every severity. Full-strict remains
unreachable while two renderer-attributable severe cases stand; this is the part that is
enforceable today.

The record is seeded from a clean-worktree full-corpus run at `c4bf105` (the merge of issue #379)
in the environment the baseline contract names: LibreOffice 25.8.7.3, Chromium 143.0.7499.4,
Poppler 25.03.0. The `shape` and `chart` figures reproduce that run's recorded values to five
decimal places; `footnote` and `tracked-deletion` differ from the 2026-08-11 *local* table below
exactly as documented, because that table was measured under LibreOffice 24.2 with its legacy
footnote separator.

Two properties set the tolerances. Within one environment the benchmark is deterministic — two
clean passes produced identical normalized metrics and identical SHA-256s for all 60 images — so
0.0005 SSIM and 0.001 ink F1 are an order of magnitude below the smallest movement any recorded
fix produced (the two-inch footnote separator, +0.000128 SSIM and +0.003537 ink F1). Across
environments the numbers move materially, and CI's LibreOffice comes from unpinned
`ubuntu-latest` apt (issue #403), so the record carries an environment fingerprint. A fingerprint
mismatch reports `environment-changed` and demands a deliberate refresh — it is never reported as
a renderer regression, because attributing a LibreOffice release to Docxodus is precisely the
false accusation the font contract and the disposition field were built to prevent.

The alarm itself is verified continuously rather than by anecdote: the comparison layer is pure,
so `visual-parity-ratchet.spec.ts` feeds it deliberately worsened summaries on every pull request
— no LibreOffice, no renderer, and no need to break rendering on purpose to prove the gate fires.

## Current case results and triage

`SSIM` is the mean over paired pages. `Ink F1` is the worst paired-page value, so it exposes a
single blank or disjoint page rather than averaging it away. Figures are from the 2026-08-11 local
rerun (environment above); `Disposition` is the corpus attribution the strict gate reads.

| Case | Pages D/L | Severity | Disposition | SSIM | Ink F1 | Triage |
|---|---:|---|---|---:|---:|---|
| text-formatting | 1/1 | close | environment | 0.99730 | 0.96575 | Control case; fonts and small caps are close. |
| merged-table | 1/1 | minor | unattributed | 0.96340 | 1.00000 | Ink geometry aligns; fill/border color dominates the perceptual delta and has not been reduced to a minimal case. |
| numbered-lists | 1/1 | severe | reference-deviation | 0.99426 | 0.55580 | Whole content is about 28 px lower in Docxodus. The OOXML top margin is 1701 twips (113.4 px at 96 DPI), which matches Docxodus; LibreOffice appears to import it differently. Treat as a reference-specific deviation unless Word evidence says otherwise. |
| multi-section | 6/6 | close | environment | 0.99932 | 0.99796 | Header/body/footer bands now sit at the distances `w:pgMar` declares (issue #377), across the landscape/portrait section transition. |
| landscape-section | 1/1 | severe | environment | 0.92525 | 0.64655 | Page dimensions match; paragraph spacing/font wrapping differs. Re-triage after issue #379. |
| running-content | 5/5 | close | environment | 0.99919 | 0.96486 | PR #372 resolved the missing page and inherited story semantics; issue #377 resolved the vertical placement of the inherited stories themselves. |
| inline-image | 1/1 | severe | environment | 0.93585 | 0.64828 | Image and text are separate source paragraphs; the discrepancy is indentation/font/wrapping, not an inline-flow failure. Re-triage after issue #379. |
| chart | 1/1 | close | environment | 0.98651 | 0.96816 | Cached clustered column data now renders as accessible inline SVG at the stored extent. Other chart families and stacked groupings remain unsupported and are not yet in the corpus. |
| shape | 1/1 | severe | renderer-bug | 0.97418 | 0.63170 | Column centering, paragraph offset, and 40%-of-margin relative width now match: horizontal bounds are exact and the top is within 1 px. The residual is auto-fit text/line height (the Docxodus box is 15 px shorter). |
| fields-and-tabs | 1/1 | severe | renderer-bug | 0.89412 | 0.15994 | Right-tab page numbers now reach the declared 9350-twip target and leaders fill the rendered remainder. TOC line height and hyperlink styling remain renderer-side; the font contract made this measurement honest and lower. |
| footnote | 1/1 | severe | environment | 0.99277 | 0.73700 | Note placement fixed (issue #378): note-text bottom flush with LibreOffice's, separator-to-note gap matches. Residual is same-font line-box metrics and LibreOffice-24.2's legacy separator width. |
| tracked-deletion | 1/1 | severe | environment | 0.94833 | 0.58271 | Identical accepted-revision bytes are now compared. Improved by the shared Calibri Light pinning (issue #379); the residual is heading metrics and wrapping of the same fonts. |

## Fixes justified by the baseline

1. **Print-layout `w:webHidden`.** Word uses this on cached TOC leader/page-number runs. They should
   be hidden in Web view but remain visible in paginated Print layout. A generated DOCX regression
   now checks both modes. Restoring the content slightly worsens the current pixel score because the
   former tab calculation placed it incorrectly; semantic correctness was retained while that
   geometry was corrected separately in item 5.
2. **Two-inch footnote separator.** The previous `width: 33%` produced 2.145 inches in a standard
   6.5-inch text area. `width: 2in; max-width: 100%` improves SSIM from 0.992055 to 0.992183 and
   tolerant ink F1 from 0.562434 to 0.565971. A generated regression pins the CSS.
3. **One-sided ink metric.** A zero-precision/zero-recall case incorrectly returned F1=1. It now
   returns zero, has a synthetic regression, and correctly upgrades the blank chart result from major
   to severe.
4. **Cached chart projection.** Standard clustered column and horizontal bar charts now project
   cached series, categories, title, legend, axis scale, theme/default series colors, overlap/gap,
   DrawingML font sizes, and stored extent into accessible inline SVG. An independently generated
   DOCX regression deliberately omits an embedded workbook and checks semantic SVG output; the
   tracked `HC043` fixture supplies the separate pixel-level LibreOffice validation.
5. **Aligned tab-stop geometry.** Right, center, and decimal tabs now measure only authored text,
   normalize unavailable-font estimates to CSS pixels, and use a flexible inline tab remainder to
   absorb native/browser metric drift. Dot, hyphen, and underscore leaders are CSS rules across that
   exact remainder. Generated DOCX and Chromium checks pin the right edge, midpoint, decimal point,
   current position, following-run width, and leader/no-leader variants. On tracked `HC022`, the page
   number lands within 1.5 px of the 9350-twip target and the leader spans the full gap. The global
   pixel scores above remain dominated by line-height, hyperlink-style, and font differences;
   LibreOffice is retained as a comparison implementation rather than used to hide valid print ink.
6. **DrawingML textbox anchor geometry.** The converter preserves positioning bases, offsets,
   alignments, extents, relative sizes, wrap clearances, and `wps:bodyPr` insets independently.
   Pagination resolves page/margin/column coordinates against the final page box and
   paragraph/line/character coordinates against the laid-out anchor, then promotes the object out
   of the clipped text column. Generated DOCX browser tests pin outer and inner coordinates for all
   supported bases; the tracked shape fixture supplies the separate pixel comparison above.

An attempted explicit `Calibri Light -> Carlito` browser fallback was rejected: it worsened the
accepted-revision case (SSIM 0.93177 to 0.92905; ink F1 0.46817 to 0.42477) despite a small SSIM gain
on the TOC. Font substitution therefore remains an observed environment variable, not a hidden
renderer heuristic.

## Prioritized next work

The former first priority (blank charts), the PR #372 rerun, aligned tab geometry,
header/body/footer vertical placement (issue #377), DrawingML textbox anchor geometry, footnote
block vertical placement (issue #378), the font-substitution contract (issue #379), and the
regression ratchet (issue #395) are resolved above. The remaining order is:

1. Model auto-fit text/line height for DrawingML textboxes (`shape`, strict-gating).
2. TOC line height and hyperlink styling (`fields-and-tabs`, strict-gating — now honestly
   measured under the pinned contract).
3. Reduce the `merged-table` fill/border color delta to a minimal case and attribute it.
4. Obtain Word evidence for the `numbered-lists` top margin to close or reclassify the
   reference-deviation.

The list-margin discrepancy should not be changed merely to imitate LibreOffice: current evidence
supports Docxodus's use of the declared OOXML margin. LibreOffice is a comparison implementation, not
the correctness oracle.
