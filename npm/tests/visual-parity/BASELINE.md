# Visual parity baseline — 2026-08-09

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

## Aggregate result

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

## Case results and triage

`SSIM` is the mean over paired pages. `Ink F1` is the worst paired-page value, so it exposes a
single blank or disjoint page rather than averaging it away.

| Case | Pages D/L | Severity | SSIM | Ink F1 | Triage |
|---|---:|---|---:|---:|---|
| text-formatting | 1/1 | close | 0.99725 | 0.96770 | Control case; fonts and small caps are close. |
| merged-table | 1/1 | minor | 0.95958 | 0.99968 | Ink geometry aligns; fill/border color dominates the perceptual delta. |
| numbered-lists | 1/1 | severe | 0.99405 | 0.55098 | Whole content is about 28 px lower in Docxodus. The OOXML top margin is 1701 twips (113.4 px at 96 DPI), which matches Docxodus; LibreOffice appears to import it differently. Treat as a reference-specific deviation unless Word evidence says otherwise. |
| multi-section | 6/6 | severe | 0.99677 | 0.00000 | Four pages are nearly blank but have disjoint running content; page 4 lacks an inherited story. First-page header/body/footer offsets also differ. This overlaps the isolated pagination work in PR #372 and should be remeasured after that lands. |
| landscape-section | 1/1 | severe | 0.91746 | 0.50174 | Page dimensions match; paragraph spacing/font wrapping differs. |
| running-content | 4/5 | severe | 0.99779 | 0.00000 | One page is missing and inherited header/footer stories diverge. This is also in PR #372's pagination cluster. |
| inline-image | 1/1 | severe | 0.93660 | 0.64760 | Image and text are separate source paragraphs; the discrepancy is indentation/font/wrapping, not an inline-flow failure. |
| chart | 1/1 | severe | 0.92933 | 0.00000 | Docxodus emits a blank page where LibreOffice renders the chart: a clear unsupported-content gap. |
| shape | 1/1 | severe | 0.96967 | 0.50719 | Textbox content exists but is roughly 5 px right and 13 px down with a small size difference: drawing-anchor geometry. |
| fields-and-tabs | 1/1 | severe | 0.89207 | 0.34460 | Cached TOC leaders/page numbers are now present in print layout, but the right tab/leader span is too short and displaced; hyperlink styling and font metrics also differ. |
| footnote | 1/1 | severe | 0.99218 | 0.56597 | Separator width now matches Word/LibreOffice's two-inch default; the note block remains about 13 px too high. |
| tracked-deletion | 1/1 | severe | 0.93177 | 0.46817 | Identical accepted-revision bytes are now compared. Remaining differences cluster around Calibri Light substitution, heading metrics, and wrapping rather than revision semantics. |

## Fixes justified by the baseline

1. **Print-layout `w:webHidden`.** Word uses this on cached TOC leader/page-number runs. They should
   be hidden in Web view but remain visible in paginated Print layout. A generated DOCX regression
   now checks both modes. Restoring the content slightly worsens the current pixel score because the
   existing tab calculation places it incorrectly; semantic correctness is retained and tab geometry
   remains a separate defect.
2. **Two-inch footnote separator.** The previous `width: 33%` produced 2.145 inches in a standard
   6.5-inch text area. `width: 2in; max-width: 100%` improves SSIM from 0.992055 to 0.992183 and
   tolerant ink F1 from 0.562434 to 0.565971. A generated regression pins the CSS.
3. **One-sided ink metric.** A zero-precision/zero-recall case incorrectly returned F1=1. It now
   returns zero, has a synthetic regression, and correctly upgrades the blank chart result from major
   to severe.

An attempted explicit `Calibri Light -> Carlito` browser fallback was rejected: it worsened the
accepted-revision case (SSIM 0.93177 to 0.92905; ink F1 0.46817 to 0.42477) despite a small SSIM gain
on the TOC. Font substitution therefore remains an observed environment variable, not a hidden
renderer heuristic.

## Prioritized next work

1. Implement a chart fallback/rendering path; blank content is the most unambiguous user-visible
   failure.
2. Reduce right/center/decimal tab-stop layout to generated documents and fix leader width plus
   post-tab alignment. The existing tab reference analysis already identifies `ProcessTab` and
   `CalcWidthOfRunInTwips` as the likely path; browser geometry assertions should accompany the fix.
3. Land or rebase PR #372, then rerun `multi-section` and `running-content` before doing overlapping
   pagination work here.
4. Correct drawing-anchor offsets with a generated textbox geometry regression.
5. Correct footnote block vertical placement independently of the now-fixed separator width.
6. Define and provision one font-substitution contract shared by Chromium and LibreOffice CI before
   treating paragraph-wrap differences as renderer regressions.

The list-margin discrepancy should not be changed merely to imitate LibreOffice: current evidence
supports Docxodus's use of the declared OOXML margin. LibreOffice is a comparison implementation, not
the correctness oracle.
