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

## Automatic line spacing — 2026-08-11 (issue #396, and #397's line-box half)

The `shape` case was filed as "auto-fit height is not modeled". It is: the converter already drops
the CSS height for a `spAutoFit` textbox, so the box already grew to its content. What was wrong
was the CONTENT height, and the cause was neither auto-fit nor textboxes.

OOXML's `w:lineRule="auto"` gives line spacing in 240ths of a **line**, where a line is the font's
own single-line height. CSS percentages and unitless line-heights both resolve against
**font-size**, so translating 259/240 to `line-height: 107.9%` under-measures every line by the
ratio between the font's natural line box and its em square — about 19% for Calibri/Carlito, and
`w:line="259"` is Word's own default for documents created since Word 2013. PR #372 had already
built the correct model (`line-height: normal` on the paragraph plus `calc(1lh * multiplier)` on
its inline children, so nothing font-specific is hard-coded) but enabled it only for EMPTY
paragraph marks, which is what pagination parity needed at the time. Populated paragraphs kept the
percentage fallback.

Measured at 11pt against LibreOffice, both engines under the pinned font contract:

| Model | Line advance |
|---|---:|
| `line-height: 107.9%` (percentage of the em square) | 15.69px |
| 1.0792 × the font's line box (≈1.22 em) | **19.33px** |
| LibreOffice, from PDF text extents at 96 DPI | **19.33px** |

Two dimensions the error was directly visible in:

- **`shape` (issue #396).** The auto-fit box is its laid-out text plus the `bodyPr` insets, so a
  text error IS a box error. The box was `(268,133)–(525,223)`, 90.0px tall against LibreOffice's
  106; it is now 108.7px — the height error falls from **−16.0px to +2.7px**, with left, top, and
  width already exact. The residual is the shape outline: CSS draws a border inside the border box
  and adds it to an auto height, while DrawingML strokes `a:ln` centered on the shape boundary, so
  it does not enlarge the shape. Changing that would move the inner coordinates PR #381
  deliberately pinned, so it is recorded rather than adjusted to fit the oracle.
- **`fields-and-tabs` (issue #397's line-box half).** TOC entries were displaced by a growing
  amount down the page — 7.3px, 11.0px, 14.7px for the first three. They now land at 140.05,
  166.13, 192.20 against LibreOffice's 140.17, 166.17, 192.17: **within 0.12px**. Ink F1 goes from
  0.15913 to 0.99775. The remaining difference is hyperlink styling, which issue #397 attributes.

The corpus rerun also dissolved a standing attribution. `numbered-lists` was `reference-deviation`
on the reading that "the whole content is about 28px lower, so LibreOffice must import the 1701-twip
top margin differently". It was the accumulated line-spacing error: the case is now close with
**exact ink geometry (F1 1.00000)**, so the two engines never disagreed about the margin. Issue #398
obtained the independent Word evidence: the Microsoft Graph conversion's first ink is row 117 at
96 DPI, exactly matching LibreOffice and within one raster row of Docxodus (118). The premise is
gone, and the case is correctly attributed to the substituted-font environment rather than a
reference deviation.

**One regression, found and fixed by the change itself.** With taller line boxes, `footnote`'s ink
F1 fell 0.70330 → 0.58697: CA008's single body line sat 2px lower, and on a page carrying ~1230 ink
pixels total that dominates the metric. The cause was the superscript footnote reference. A
`vertical-align: super` inline box is counted when CSS sizes the line box, so the raised box made
the line 25.31px instead of 19.42px and pushed every glyph on it down; Word instead rides a
superscript inside the existing line, and a paragraph does not open up because it carries a
footnote reference. `sup, sub { line-height: 0 }` joins the document-layout stylesheet as a third
Word layout invariant CSS defaults do not match — the same device the converter already used to
stop the two compacted pieces of a `w:br` contributing a line box. `footnote` then goes to **close**
(ink F1 0.99064), better than before the change.

Corpus effect, one clean-worktree full rerun against the immediately preceding run in the same
environment:

| Signal | Before | After |
|---|---:|---:|
| Case severity | 4 close, 1 minor, 0 major, 7 severe | 7 close, 3 minor, 2 major, **0 severe** |
| Strict-gating cases | `shape`, `fields-and-tabs` | **none** |
| Mean SSIM | 0.981316 | 0.987810 |
| Mean tolerant ink F1 | 0.848523 | **0.981644** |

Nine of twelve cases improved, three of them out of `severe` entirely (`numbered-lists`, `shape`,
`footnote` to close; `fields-and-tabs` and `tracked-deletion` to minor; `landscape-section` and
`inline-image` to major). `merged-table`, `chart`, and `running-content` are byte-identical.
`inline-image` is the one case whose SSIM fell (0.93660 → 0.93255) while its ink F1 rose
0.64760 → 0.77340 and its severity improved — its glyphs land better, and the remaining
same-font rasterization difference is what SSIM's block statistics see.

No severe case remains and the strict-gating set is empty, but strict mode is NOT enabled here:
that is a separate decision, and two `major` cases still sit above the threshold a strict run
would eventually want.

## TOC hyperlink styling — 2026-08-11 (issue #397)

Issue #397 named two residuals in the `fields-and-tabs` case, and they turned out to have opposite
answers. Pinning them separately is what the issue asked for, and it is what the evidence supports.

**Line-box height was ours.** TOC entries drifted further down the page with every entry — 7.3px,
11.0px, 14.7px for the first three — because automatic line spacing was measured against font-size
instead of the font's own line box. Fixed in issue #396 above; entries now land within **0.12px**
of LibreOffice and the case's ink F1 is 0.99775.

**Hyperlink styling is not.** The entry run in `HC022` carries `<w:rStyle w:val="Hyperlink"/>`, and
the document's `Hyperlink` character style declares `w:color="0563C1"` with `w:u w:val="single"`.
Sampling the two renders over the TOC entry rows:

| Renderer | Dominant entry-text colour |
|---|---|
| Docxodus | **`#0563C1`** — the declared value, byte for byte |
| LibreOffice | **`#000000`** — no hyperlink colour at all |

Docxodus renders what the file declares. LibreOffice drops a character style that the run
explicitly references, which is a deviation from the OOXML, not a Docxodus defect — and "the TOC
came out blue and underlined" is the well-known consequence of Word's own `\h` table of contents
applying this style. LibreOffice is a comparison implementation, not the correctness oracle, so the
output is **not** changed to match it and the disposition moves `renderer-bug` →
`reference-deviation`. This is the same standard applied when the `numbered-lists` margin was
kept over LibreOffice's import.

The claim only means something if the renderer is reading the style rather than decorating every
hyperlink it sees, so `npm/tests/toc-line-geometry.spec.ts` generates a TOC containing both kinds
of entry: one with `w:rStyle` and an otherwise identical one without. The styled entry must carry
the declared colour and underline; the unstyled one must carry neither. Line geometry is pinned by
separate assertions — the entry line box equals the OOXML multiple of the font's line box, and the
entries are evenly spaced so displacement cannot accumulate — and one assertion pins that the two
are independent, i.e. that applying the character style does not change the line box.

With this, the corpus's remaining non-`environment` work is the `merged-table` colour delta (issue
#399), the last `unattributed` case.

## The merged-table colour delta — 2026-08-11 (issue #399)

The corpus's last `unattributed` case. The issue recorded that `merged-table` has perfect ink
geometry (F1 1.00000) and that "the whole perceptual delta is fill/border *color*". **That premise
was wrong**, and the measurement is easy once taken: sampling every non-white pixel of both renders,

| Renderer | Header fill | Band fill | Border |
|---|---|---|---|
| LibreOffice | `#4472C4` | `#D9E2F3` | `#8EAADB` |
| Docxodus | `#4472C4` | `#D9E2F3` | `#8EAADB` |

The two engines paint the **same three colours**, and they are the theme-derived values: `HC029`
declares no `w:shd` or border of its own, taking everything from the `Grid Table 4 Accent 5` style,
whose cached literals equal `accent5` (`4472C4`) under Word's tint formula exactly — `tint 0x99` →
`8EAADB`, `tint 0x33` → `D9E2F3`. There is nothing for the engines to disagree about.

What differs is **geometry**. Horizontal extents are identical to the pixel (fills span x 96–719 and
97–718 in both). Vertically the rows are about 1px taller in Docxodus, accumulating to ~3px by the
second header band (LibreOffice 200–236, Docxodus 203–239). Because the fills are large solid
areas, a one-pixel edge offset moves a great many pixels past the ΔE threshold, which is exactly how
a case can hold ink F1 1.00000 — the ink masks overlap within the tolerance — while its SSIM sits at
0.96348.

That row-height difference isolates to font line metrics by elimination rather than by assumption:
the table declares no `w:trHeight`, its `w:tblCellMar` top and bottom are both **0**, and its
paragraphs are `w:line="240"` — exactly single spacing, which issue #396 does not touch. A row is
therefore exactly one font line box, and the 1px is the two engines measuring that box differently
for the same substituted font. The disposition moves `unattributed` → **`environment`**, the same
residual the other environment cases carry. Note this changes no gating: `gatesStrictRun` fires
only on *severe* cases, and `merged-table` is `minor`.

**A real inconsistency found while reducing it.** The tracked fixture cannot say whether a renderer
resolves `w:themeFill`/`w:themeColor` or just paints the cached `w:fill`/`w:color` literal, because
Word keeps the two in sync. A generated table can, by making them disagree — and that exposed
Docxodus resolving the *same* accent colour from two different sources within one style: shading
resolved the theme, while border colour read the cache. ECMA-376 makes the theme reference the
authority and the literal a cache of the last resolution, so border colour now resolves the theme
like shading always did. For any Word-written file this changes nothing (which is why no corpus
number moves); it changes documents whose theme was swapped without a cache rewrite.

With this the corpus has **no `unattributed` case left**, and every remaining residual is either
`environment` or a recorded `reference-deviation`.

## Second-wave corpus — 2026-08-11 (issue #400)

One fixture per category means a fixed case hides a reopened feature gap: the `chart` case reads
close while every chart family except clustered bar/column is unsupported — the gap became
invisible the moment the one fixture was fixed. Nine cases were added covering the missing
shapes; five reference existing tracked fixtures in place, four are authored deterministically by
`TestFiles/VP/make-vp-fixtures.py` and committed under the same blob-hash guard (no third-party
corpus). All nine entered `unattributed`, which strict-gates, and were triaged from the first
measured run below.

| Case | Fixture | Provenance |
|---|---|---|
| chart-stacked | `TestFiles/VP/VP001-Chart-Stacked-Column.docx` | authored — HC043's clustered chart regrouped stacked |
| chart-pie | `TestFiles/CU002-Chart-Cached-Data-02.docx` | existing tracked fixture |
| chart-line | `TestFiles/CU004-Chart-Cached-Data-04.docx` | existing tracked fixture |
| wrapped-image-square | `TestFiles/DB007-WhitePaper.docx` | existing tracked fixture |
| wrapped-image-tight | `TestFiles/VP/VP002-Image-Wrap-Tight.docx` | authored — HC042's picture re-anchored wrapTight |
| nested-table | `TestFiles/WC/WC043-Nested-Table.docx` | existing tracked fixture |
| two-column-section | `TestFiles/VP/VP003-Two-Column-Section.docx` | authored |
| endnote | `TestFiles/WC/WC036-Endnote-With-Table-Before.docx` | existing tracked fixture |
| legal-contract | `TestFiles/VP/VP004-Legal-Contract.docx` | authored |

**The first measured run was taken in the record's own environment, reconstructed exactly**:
LibreOffice 25.8.7.3 (TDF debs), Chromium 143.0.7499.4 (the repository's pinned Playwright), and
Poppler 25.03.0 — the same fingerprint `ratchet.json` already carried. The proof the
reconstruction is faithful: the record refresh reproduced **all twelve existing cases'
SSIM and ink-F1 values to five decimal places** — the update diff touches only `sourceCommit`
and the nine added rows.

One environment finding with its own lesson: TDF-packaged LibreOffice **bundles its own
Caladea/Carlito/Liberation fonts**, which silently override the font-substitution contract inside
LibreOffice only — and the issue-#379 wrapping probe caught it exactly as designed (Cambria→
Caladea wrapped 4 lines in Chromium vs 5 in LibreOffice against the bundled copy). Removing the
bundled duplicates restored identical wrapping, and only the one Cambria-declaring case moved.
Distro LibreOffice packages resolve system fonts and do not hit this.

First measured results (figures are what `ratchet.json` records):

| Case | Pages D/L | Severity | Disposition | SSIM | Ink F1 |
|---|---:|---|---|---:|---:|
| chart-stacked | 1/1 | severe | unsupported-feature | 0.94151 | 0.00000 |
| chart-pie | 1/1 | severe | unsupported-feature | 0.95467 | 0.00000 |
| chart-line | 1/1 | severe | unsupported-feature | 0.93289 | 0.00000 |
| wrapped-image-square | 1/1 | severe | unsupported-feature | 0.74176 | 0.39352 |
| wrapped-image-tight | 1/1 | severe | unsupported-feature | 0.73466 | 0.51179 |
| nested-table | 1/1 | severe | reference-deviation | 0.96773 | 0.40867 |
| two-column-section | 2/1 | severe | renderer-bug | 0.80838 | 0.07207 |
| endnote | 1/1 | severe | renderer-bug | 0.84081 | 0.93747 |
| legal-contract | 3/3 | severe | renderer-bug | 0.69140 | 0.52148 |

Triage, per case — each disposition is the reviewed claim in `corpus.ts`, with the tracking
issue where one exists:

1. **Chart families (issue #411, `unsupported-feature` ×3).** The cached-data SVG projection
   gates on `c:barChart` + `grouping="clustered"` (`WmlToHtmlConverter.Charts.cs`); a stacked
   grouping, `pie3DChart`, and `lineChart` all return null and the extent renders **blank** —
   ink F1 0.00000 on all three. Exactly the reopened gap this wave existed to make visible.
2. **Floating-image text wrap (issue #412, `unsupported-feature` ×2).** The picture is
   positioned, but text does not flow around it: LibreOffice wraps five lines beside
   `DB007-WhitePaper.docx`'s square-wrapped picture, Docxodus resumes the text below the image,
   displacing the page's lower half. Same mechanism on the authored wrapTight case.
3. **Two-column section (issue #413, `renderer-bug`).** Two failures: the
   `w:type="continuous"` section start renders as a **page break** (2/1 pages — the only
   page-count mismatch in the corpus), and `w:cols w:num="2"` is **ignored** (the body renders
   one full-width column, ink F1 0.07207). General multi-section support is fine —
   `multi-section` is close at 6/6 — it is specifically the continuous start type and column
   geometry.
4. **Endnotes (issue #414, `renderer-bug`).** The converter emits the endnotes section
   (`docx2html --render-footnotes` output contains `class="endnotes"` and the note's table),
   but `pagination.ts` has footnote handling only — the endnotes section never reaches a page.
   The citation marker also renders decimal `1` where Word's default endnote numbering is
   lowercase roman (`i`).
5. **Legal contract (issue #415, `renderer-bug`).** The dominant residual is measured, not
   assumed: the `(a)` list marker lands at 145 px in **both** engines, but the following text
   starts at 193 px in Docxodus — the next 720-twip default tab stop — versus 169 px in
   LibreOffice, where the declared `w:ind w:left="1080"` is 168 px. The list-number suffix tab
   advances to the next default stop instead of the declared text indent, ~25 px right on every
   numbered clause — the heavy-numbering legal shape the library's positioning makes central.
   Secondary residuals: LibreOffice drops heading space-before at the top of a page where
   Docxodus paints the declared `w:spacing w:before` (16 px at each page top), and the cached
   TOC carries the recorded issue-#397 hyperlink-style deviation. Everything else about the
   case — cached TOC entries, leaders, PAGEREF results, multilevel heading numbers, cached REF
   cross-references, the signature table — renders with content parity.
6. **Nested table (`reference-deviation`).** Both engines nest correctly and start the outer
   table at the same margin row (96). The outer ink band is 96–163 (Docxodus) vs 96–152
   (LibreOffice), and the ~11 px difference is the document default `w:spacing w:after="160"`
   (10.7 px) on the paragraph preceding the nested table: Docxodus paints the declared spacing,
   LibreOffice suppresses it. Nothing in the OOXML licenses the suppression; Word-behavior
   evidence for this exact shape is the open question, as with the `numbered-lists` history.

Corpus aggregate after the wave (21 cases, 32 paired pages, 0 conversion errors): 7 close,
3 minor, 2 major, 9 severe; mean SSIM 0.929394; mean ink F1 0.770255. The nine severe are
5 `unsupported-feature`, 3 `renderer-bug`, 1 `reference-deviation`; the strict-gating set is
`two-column-section`, `endnote`, `legal-contract` — renderer-owned work now visible instead of
hidden, which is the acceptance criterion of issue #400.

## The evidence contracts — 2026-08-13 (issues #402, #403, #404)

The disposition system separated "our bug" from "system difference", but its evidence chain had
three gaps, closed together because they are one design: the LibreOffice side of the comparison
was not contractually pinned (#403), Word evidence had nowhere to live (#402), and the
`environment` attributions rested on whole-fixture impressions rather than reduced cases (#404).

**Reference-version contract (#403).** The benchmark is contracted to LibreOffice 25.8, declared
once in `environment-contract.ts` and enforced twice: `assertLibreOfficeContract()` fails the run
at start with install guidance (the TDF 25.8.7.3 archive, the bundled-font removal step, the known
cross-version differences), and the ratchet fingerprint — which now also carries the Poppler
major.minor, record schema 2 — refuses to compare numbers across reference changes. CI installs
the exact TDF build instead of `ubuntu-latest` apt, which carries 24.2: **every scheduled run
since the record was seeded under 25.8 would have rendered for twenty minutes and then died at
the fingerprint check**; now it fails in the first second or not at all. The failure mode is
proven purely on every pull request, alongside new assertions that the declared contract, the
committed record, and the corpus stay mutually consistent.

**Word-reference evidence store (#402).** `word-reference.json` is the committed, numbers-only
answer to "unless Word evidence says otherwise": page counts, page geometry, ink extents, named
per-case measurements, and the Word/OS versions used — never binaries, never an image corpus.
The manual step shrinks to exporting each fixture to PDF with a licensed Word;
`npm run capture:word-reference` does everything downstream under the benchmark's own contract
(Poppler at 96 DPI, the shared ink model) and can record advisory three-way comparisons against
a benchmark run. The honesty boundary is explicit: Word renders with genuine Office fonts, not
the contract substitutes, so Word evidence decides STRUCTURAL questions (spacing suppressed or
painted, page counts, block positions); pixel scores against Word stay advisory. Dispositions
cite recorded data via `wordEvidence`, which the pure spec refuses unless the cited case is
measured. All 21 cases are seeded `pending` — **the measurements themselves require a Word
license and are the open half of issue #402**; WORD_REFERENCE.md lists the open questions the
first capture should decide (the `nested-table` spacing suppression first among them).

**Reduced environment cases (#404).** The issue's premise had gone stale — it names three severe
cases at ink F1 0.58–0.65, but issue #396 had since moved `tracked-deletion` to minor (0.99667)
and the other two to major. What remained owed was the reduction: `visual-parity-reductions.spec.ts`
(now part of the benchmark run) reduces each case to a minimal generated document measured
identically in both engines, and the dispositions now cite those numbers instead of impressions:

| Reduction | Observable | Docxodus | LibreOffice | Residual isolated |
|---|---|---:|---:|---|
| `landscape-spacing` | pitch of four identical single-line Calibri paragraphs, landscape section | 29 px/paragraph, uniform | 30 px/paragraph, uniform | 1 px/line same-font line-box delta; ink bands 12 vs 14 rows (rasterization spread) |
| `inline-image` | 150x75 px `wp:extent` between text paragraphs | exactly 150x75 at the margin | 151x75, 1 px offset | extent and flow agree; text resumes 20 vs 16 px below the image |
| `heading-metrics` | Calibri Light (Carlito) heading over Calibri body | 24 px advance | 24 px advance | advance IDENTICAL; heading ink 17 vs 19 rows — rasterization, not layout |

Structure agrees completely in every reduction (band counts, x-extents, page geometry), so the
`environment` disposition now means something falsifiable: the engines lay the same fonts out
with sub-pixel-accumulating metric differences, and nothing else. Whether Docxodus's choices are
also Word-correct is exactly what the #402 capture will decide; no renderer change was made,
matching the issue's "no change solely to imitate LibreOffice".

**The record refresh, and what it caught.** The contract environment was reconstructed for this
work (TDF LibreOffice 25.8.7.3 with bundled fonts removed, Chromium 143.0.7499.4, Poppler
24.02.0 — the version CI's `ubuntu-latest` actually has, unlike the 25.03 the previous record
was measured under). A clean full-corpus run reproduced **13 of 21 cases to five decimal
places**, which both proves the reconstruction faithful and shows Poppler 24.02-vs-25.03 benign
for unchanged cases — the new fingerprint field guards the boundary anyway. The other eight
cases had moved because renderer PRs #417–#421 landed real fixes without refreshing the record;
this refresh banks them, and their dispositions were re-triaged from the new evidence:

| Case | Movement | Re-triage |
|---|---|---|
| chart-stacked | severe 0.94151/0.00000 → minor 0.97053/0.97696 | `unsupported-feature` → `environment`: stacked charts project since PR #417; residual mirrors the clustered `chart` control case |
| chart-pie | 0.95467/0.00000 → 0.94359/0.67052 | → `renderer-bug`: the pie renders; slice angle/ordering, 3-D perspective, and label placement are renderer-owned geometry |
| chart-line | 0.93289/0.00000 → 0.90389/0.50465 | → `renderer-bug`: the dominant residual is measured — cached date-serial categories print raw (`41518`) where LibreOffice formats dates (`9/1/2013`) |
| wrapped-image-square | 0.74176/0.39352 → 0.75039/0.53248 | `unsupported-feature` → `unattributed`: wrap works since PR #418; the remaining severe residual is un-triaged and gates until it is |
| wrapped-image-tight | 0.73466/0.51179 → 0.84881/0.94860 | → `unattributed`: nearly major-boundary; same re-triage debt |
| two-column-section | 2/1 pages 0.80838/0.07207 → 1/1 0.70903/0.45891 | stays `renderer-bug`: #413's two failures fixed by PR #419; residual is the column fill/split point |
| endnote | 0.84081/0.93747 → 0.82594/0.91258 | `renderer-bug` → `unattributed`: PR #420 changed the case materially (endnotes now flow); the systematic body shift needs fresh triage |
| legal-contract | 0.69140/0.52148 → 0.69413/0.52440 | stays `renderer-bug`: PR #421's tab fix moved this fixture barely — the severity was never that one defect alone |

Corpus aggregate after the refresh (21 cases, 32 paired pages, 0 conversion errors; mean SSIM
0.929572, mean ink F1 0.867257): 7 close, 4 minor, 2 major, 8 severe; the strict-gating set is
now the honest seven (`chart-pie`,
`chart-line`, `two-column-section`, `legal-contract` as renderer-bug; `endnote`,
`wrapped-image-square`, `wrapped-image-tight` as unattributed re-triage debt). The lesson the
refresh teaches: a renderer PR that changes corpus numbers must refresh the record in the same
diff — the improvements list printed by every passing run announces exactly when this is owed.

## Current case results and triage

`SSIM` is the mean over paired pages. `Ink F1` is the worst paired-page value, so it exposes a
single blank or disjoint page rather than averaging it away. Figures are from the 2026-08-11 rerun after issue #396
(LibreOffice 25.8.7.3, Chromium 143.0.7499.4, Poppler 25.03.0); `Disposition` is the corpus
attribution the strict gate reads, and these numbers are what `ratchet.json` records. The
second-wave cases (issue #400) are tabulated in their own section above; `ratchet.json` records
all 21 together.

| Case | Pages D/L | Severity | Disposition | SSIM | Ink F1 | Triage |
|---|---:|---|---|---:|---:|---|
| text-formatting | 1/1 | close | environment | 0.99708 | 0.96770 | Control case; fonts and small caps are close. |
| merged-table | 1/1 | minor | environment | 0.96348 | 1.00000 | Both engines paint the SAME theme-derived colours (issue #399); the residual is row height (~1 px/row) that large solid fills amplify perceptually. |
| numbered-lists | 1/1 | close | reference-deviation | 0.99898 | 1.00000 | The "28 px lower" reading was the accumulated auto-line-spacing error (issue #396), not a top-margin deviation: ink geometry is now exact. |
| multi-section | 6/6 | close | environment | 0.99986 | 1.00000 | Header/body/footer bands sit at the distances `w:pgMar` declares (issue #377); ink geometry now exact. |
| landscape-section | 1/1 | major | environment | 0.94878 | 0.95533 | Page dimensions match; improved by issue #396. Residual is same-font wrapping and rasterization. |
| running-content | 5/5 | close | environment | 0.99927 | 0.96486 | PR #372 resolved the missing page and inherited story semantics; issue #377 resolved their vertical placement. |
| inline-image | 1/1 | major | environment | 0.93255 | 0.77340 | Improved by issue #396 (ink F1 0.64760 to 0.77340). Residual is indentation/wrapping of the same fonts, not an inline-flow failure. |
| chart | 1/1 | close | environment | 0.98687 | 0.96817 | Cached clustered column data renders as accessible inline SVG at the stored extent. Other chart families remain unsupported. |
| shape | 1/1 | close | renderer-bug | 0.98599 | 1.00000 | Auto-fit height now follows the laid-out text (issue #396): height error -16.0 px to +2.7 px, ink geometry exact. Residual is the CSS border adding to an auto height where DrawingML strokes `a:ln` on the shape boundary. |
| fields-and-tabs | 1/1 | minor | reference-deviation | 0.96026 | 0.99775 | Tab targets and leaders correct since PR #380; entry line height correct since issue #396 (within 0.12 px). The residual is hyperlink styling: the run references the `Hyperlink` character style, Docxodus paints its declared `#0563C1`, LibreOffice paints black (issue #397). |
| footnote | 1/1 | close | environment | 0.99502 | 0.99064 | Note placement fixed (issue #378); issue #396 stopped the superscript reference inflating its line box. Residual is substituted-font rasterization and LibreOffice-24.2 separator width. |
| tracked-deletion | 1/1 | minor | environment | 0.97950 | 0.99667 | Identical accepted-revision bytes are compared. Improved by the font contract (issue #379) and issue #396; residual is same-font heading metrics and wrapping. |

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
block vertical placement (issue #378), the font-substitution contract (issue #379), the
regression ratchet (issue #395), automatic line spacing (issue #396, which also resolved
#397's line-box half and dissolved #398's premise), the second-wave corpus (issue #400), the
first implementations behind #411/#412/#413/#414/#415 (PRs #417–#421, whose numbers the
2026-08-13 refresh banked), the reference-version contract (issue #403), the Word-reference
evidence framework (issue #402 — tooling and procedure; measurements await a Word license),
and the #404 reductions are resolved above. The remaining order is:

1. The re-triage debt the 2026-08-13 refresh made explicit: `endnote`'s systematic body shift
   and the two wrapped-image residuals (all `unattributed`, all strict-gating until triaged).
2. The measured renderer-bug severes: `chart-line`'s date-serial axis labels and `chart-pie`'s
   slice geometry (issue #411 follow-on), `two-column-section`'s column fill/split point
   (issue #413 follow-on), and `legal-contract`'s remaining per-clause offsets (issue #415).
3. The first Word-reference capture (issue #402's open half): export the corpus from a licensed
   Word and run `npm run capture:word-reference` — deciding at least the `nested-table` spacing
   suppression, the `legal-contract` heading space-before, and whether the #404 reductions'
   same-font layout choices are Word-correct.

The list-margin discrepancy should not be changed merely to imitate LibreOffice: current evidence
supports Docxodus's use of the declared OOXML margin. LibreOffice is a comparison implementation, not
the correctness oracle.
