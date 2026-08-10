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
- Environment: Chromium 143.0.7499.4; LibreOffice 25.8.7.3; Poppler 25.03.0. Fonts are governed by
  the shared substitution contract (issue #379): one fontconfig fragment both engines read, mapping
  Calibri and Calibri Light to Carlito, Cambria to Caladea, and Times New Roman / Arial / Courier
  New to the matching Liberation faces. The run resolves every declared family before either engine
  starts, records the exact font file each one used, and skips (or fails, under strict mode) rather
  than reporting numbers from an unknown font environment.

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

| Signal | Initial | After PRs #372–#374 | Current |
|---|---:|---:|---:|
| Cases | 12 | 12 | 12 |
| Paired pages | 20 | 21 | 21 |
| Conversion errors | 0 | 0 | 0 |
| Page-count mismatches | 1 | 0 | 0 |
| Case severity | 1 close, 1 minor, 0 major, 10 severe | 2 close, 1 minor, 0 major, 9 severe | 5 close, 1 minor, 0 major, 6 severe |
| Page severity | 1 close, 1 minor, 2 major, 16 severe | 2 close, 1 minor, 1 major, 17 severe | 14 close, 1 minor, 0 major, 6 severe |
| Mean SSIM | 0.974586 | 0.978298 | 0.981207 |
| Mean tolerant ink F1 | 0.394106 | 0.412753 | 0.854815 |

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
5. **Footnote block placement (issue #378).** The note area is bottom-aligned to the body band, so
   the block's HEIGHT is its position — which is why the note text and the already-correct
   two-inch separator were both about 13 px too high, and why fixing the separator's width could
   not move it. Removing the block's own slack (a trailing margin below the last note, a bare rule
   with a tuned gap instead of the separator paragraph's line, a line-box-inflating superscript
   mark, and a fixed 1.4 line-height overriding the note style's single spacing) lands the
   separator within 2 px of LibreOffice's and the note text's bottom edge exactly on it. The case
   moves from severe (SSIM 0.99218, ink F1 0.56597) to **close** (0.99481 / 0.95875).
6. **Shared font-substitution contract (issue #379).** One fontconfig fragment now governs both
   engines, so a family cannot be set in different faces on the two sides of the comparison. The
   run resolves every declared family before starting, records the font file each one used, and
   skips or fails rather than reporting numbers from an unknown environment; a generated probe
   compares the engines' advances and wrapped line counts and labels a mismatch as font-environment
   drift rather than a renderer regression. `tracked-deletion` improves markedly (SSIM 0.93177 →
   0.95011, ink F1 0.46817 → 0.62634) because Chromium and LibreOffice previously disagreed about
   Calibri Light. `fields-and-tabs` moves the other way on ink F1 (0.30164 → 0.15913, SSIM
   0.88672 → 0.89415): the substitution had been masking a real TOC line-height difference, which
   is now visible as the renderer signal it always was.

## Current case results and triage

`SSIM` is the mean over paired pages. `Ink F1` is the worst paired-page value, so it exposes a
single blank or disjoint page rather than averaging it away.

| Case | Pages D/L | Severity | SSIM | Ink F1 | Triage |
|---|---:|---|---:|---:|---|
| text-formatting | 1/1 | close | 0.99725 | 0.96770 | Control case; fonts and small caps are close. |
| merged-table | 1/1 | minor | 0.96348 | 1.00000 | Ink geometry aligns; fill/border color dominates the perceptual delta. |
| numbered-lists | 1/1 | severe | 0.99426 | 0.55580 | Whole content is about 28 px lower in Docxodus. The OOXML top margin is 1701 twips (113.4 px at 96 DPI), which matches Docxodus; LibreOffice appears to import it differently. Treat as a reference-specific deviation unless Word evidence says otherwise. |
| multi-section | 6/6 | close | 0.99934 | 0.99797 | Header/body/footer bands now sit at the distances `w:pgMar` declares (issue #377), across the landscape/portrait section transition. |
| landscape-section | 1/1 | severe | 0.92607 | 0.60042 | Page dimensions match; paragraph spacing/font wrapping differs. Improved as a side effect of issue #377. |
| running-content | 5/5 | close | 0.99921 | 0.96486 | PR #372 resolved the missing page and inherited story semantics; issue #377 resolved the vertical placement of the inherited stories themselves. |
| inline-image | 1/1 | severe | 0.93660 | 0.64760 | Image and text are separate source paragraphs; the discrepancy is indentation/font/wrapping, not an inline-flow failure. |
| chart | 1/1 | close | 0.98687 | 0.96817 | Cached clustered column data now renders as accessible inline SVG at the stored extent; bars, grid, colors, labels, title, and bottom legend align closely. Other chart families and stacked groupings remain unsupported. |
| shape | 1/1 | severe | 0.96967 | 0.50719 | Textbox content exists but is roughly 5 px right and 13 px down with a small size difference: drawing-anchor geometry. |
| fields-and-tabs | 1/1 | severe | 0.89415 | 0.15913 | Right-tab page numbers now reach the declared 9350-twip target and leaders fill the rendered remainder. The whole-page score falls slightly because the corrected leader adds ink at the still-mismatched TOC line height; hyperlink styling, font metrics, and paragraph spacing remain separate differences. |
| footnote | 1/1 | close | 0.99481 | 0.95875 | Note block composition corrected (issue #378): the separator is a line with the rule on its baseline, spacing falls between notes, and the last note ends on the body band's bottom edge. |
| tracked-deletion | 1/1 | severe | 0.95011 | 0.62634 | Identical accepted-revision bytes, and both engines now resolve Calibri Light the same way (issue #379). Remaining differences are heading metrics and wrapping, not revision semantics. |

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

An attempted explicit `Calibri Light -> Carlito` browser fallback was rejected: it worsened the
accepted-revision case (SSIM 0.93177 to 0.92905; ink F1 0.46817 to 0.42477) despite a small SSIM gain
on the TOC. That result was the diagnosis, not a dead end — changing ONE engine's mind made the two
disagree more, because LibreOffice was already resolving Calibri Light through its own substitution
table. Issue #379 resolves it where it belongs: a fontconfig fragment both engines read, which moves
the same case to SSIM 0.95011 / ink F1 0.62634. Font substitution is now a declared, verified
contract rather than either an observed variable or a hidden renderer heuristic.

## Prioritized next work

The former first priority (blank charts), the PR #372 rerun, aligned tab geometry,
header/body/footer vertical placement (issue #377), footnote block placement (issue #378), and the
shared font-substitution contract (issue #379) are resolved above. The remaining order is:

1. Correct drawing-anchor offsets with a generated textbox geometry regression.
2. Investigate the paragraph line-height differences the font contract now exposes on
   `fields-and-tabs`, `landscape-section`, and `inline-image` — with the fonts pinned, these are
   renderer signals rather than environment noise.
3. Make the paginator's 60% note-area cap defer rather than clip. A page whose citations would
   fill more than that share admits more notes than it leaves room for, and the surplus is cut off
   by the note block's `overflow: hidden` — 187 px on a generated twelve-note page, reproduced
   identically before issue #378's changes, so it is a question about the cap and not about the
   note block's composition.

The list-margin discrepancy should not be changed merely to imitate LibreOffice: current evidence
supports Docxodus's use of the declared OOXML margin. LibreOffice is a comparison implementation, not
the correctness oracle.
