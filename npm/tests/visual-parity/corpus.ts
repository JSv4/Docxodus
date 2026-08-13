export const REQUIRED_VISUAL_CATEGORIES = [
  'text',
  'tables',
  'lists',
  'multi-section',
  'headers-footers',
  'images',
  'charts',
  'shapes',
  'fields',
  'footnotes',
  'tracked-changes',
  'page-geometry',
  // Second-wave categories (issue #400): shapes the first corpus left invisible —
  // one fixture per category had let a fixed `chart` case hide every unsupported
  // chart family, and the library's legal-document positioning was uncovered.
  'image-wrap',
  'nested-tables',
  'columns',
  'endnotes',
  'contract',
] as const;

export type VisualCategory = typeof REQUIRED_VISUAL_CATEGORIES[number];

/**
 * Attribution of a case's dominant residual discrepancy. Severity alone conflates "our bug",
 * "different comparison environment", and "LibreOffice deviates from the OOXML evidence";
 * the disposition is the reviewed claim that separates them, so trend runs and strict mode
 * can gate on the renderer-attributable subset instead of the raw severity count.
 */
export const VISUAL_DISPOSITION_KINDS = [
  // The discrepancy is an established Docxodus rendering defect.
  'renderer-bug',
  // The discrepancy is dominated by the comparison environment (font substitution, wrapping
  // and line metrics that differ between Chromium and LibreOffice), not by OOXML geometry.
  'environment',
  // Docxodus follows the OOXML evidence and LibreOffice deviates from it.
  'reference-deviation',
  // The content is a known unimplemented feature, tracked as a feature gap rather than a bug.
  'unsupported-feature',
  // Not yet triaged. New corpus entries start here; strict mode treats it as gating so an
  // untriaged severe case cannot hide behind a non-gating label.
  'unattributed',
] as const;

export type VisualDispositionKind = typeof VISUAL_DISPOSITION_KINDS[number];

/** Disposition kinds whose severe cases fail a strict run. */
export const GATING_DISPOSITION_KINDS: readonly VisualDispositionKind[] =
  ['renderer-bug', 'unattributed'];

export interface VisualDisposition {
  kind: VisualDispositionKind;
  /** The evidence for the attribution — a disposition is a reviewed claim, never a bare label. */
  rationale: string;
  /** Tracking issue or PR for the residual discrepancy, when one exists. */
  reference?: string;
  /**
   * Citation of recorded Word evidence (issue #402): what `word-reference.json` measured that
   * supports this attribution — e.g. which key measurement decides the question. Validated by
   * the pure spec: a citation requires the cited case to be `measured`, so a rationale cannot
   * claim Word data that was never captured. `reference-deviation` claims should carry one as
   * soon as the corresponding fixture has been measured.
   */
  wordEvidence?: string;
}

export interface VisualCorpusEntry {
  id: string;
  path: string;
  categories: VisualCategory[];
  rationale: string;
  disposition: VisualDisposition;
  revisionMode?: 'source' | 'accepted';
}

/**
 * A deliberately small, stratified corpus. Every path is enforced as Git-tracked by the runner;
 * documents are referenced in place and never copied into a new committed corpus.
 */
export const VISUAL_PARITY_CORPUS: VisualCorpusEntry[] = [
  {
    id: 'text-formatting',
    path: 'TestFiles/HC020-Small-Caps.docx',
    categories: ['text'],
    rationale: 'Compact font metrics and character-formatting control.',
    disposition: {
      kind: 'environment',
      rationale: 'Close control case; the residual is glyph antialiasing and substituted-font metrics.',
    },
  },
  {
    id: 'merged-table',
    path: 'TestFiles/HC029-Table-Merged-Cells.docx',
    categories: ['tables'],
    rationale: 'Merged-cell borders, widths, and vertical geometry.',
    disposition: {
      kind: 'environment',
      rationale: 'Colour is NOT the delta (issue #399). Sampling both renders, the two engines paint ' +
        'the identical theme-derived values — #4472C4 header, #D9E2F3 bands, #8EAADB borders — and ' +
        'the style\'s cached literals equal those values exactly under Word\'s tint formula, so ' +
        'there is nothing to disagree about. Horizontal extents match exactly; the residual is row ' +
        'HEIGHT, about 1 px per row accumulating to ~3 px over the table. That isolates to font ' +
        'line metrics by elimination: the table declares no `w:trHeight`, its `w:tblCellMar` top ' +
        'and bottom are both 0, and its paragraphs are `w:line="240"` single spacing — so a row IS ' +
        'exactly one font line box, and the 1 px is the two engines measuring that box differently ' +
        'for the same substituted font. Large solid fills make a one-pixel edge offset dominate ' +
        'the perceptual signal while ink F1 stays 1.00000.',
      reference: 'https://github.com/JSv4/Docxodus/issues/399',
    },
  },
  {
    id: 'numbered-lists',
    path: 'TestFiles/DB012-Lists-With-Different-Numberings.docx',
    categories: ['lists'],
    rationale: 'Multiple numbering definitions and indentation behavior.',
    disposition: {
      kind: 'environment',
      rationale: 'The "content sits 28 px lower" reading was never a top-margin deviation: it was the ' +
        'accumulated automatic-line-spacing error (issue #396), which displaced every line by a ' +
        'growing amount down the page. With auto spacing measured against the font line box the ' +
        'case is close and ink geometry is exact (F1 1.00000), so the margin was never in ' +
        'disagreement. The residual is substituted-font rasterization.',
      reference: 'https://github.com/JSv4/Docxodus/issues/398',
      wordEvidence: 'At 96 DPI, the Microsoft Graph conversion starts first-line ink at row 117, ' +
        'matching LibreOffice exactly and lying within one raster row of Docxodus (118).',
    },
  },
  {
    id: 'multi-section',
    path: 'TestFiles/DB001-Sections.docx',
    categories: ['multi-section', 'page-geometry'],
    rationale: 'Multiple sections and page-size transitions across six pages.',
    disposition: {
      kind: 'environment',
      rationale: 'Close since issue #377; the residual is substituted-font line metrics.',
    },
  },
  {
    id: 'landscape-section',
    path: 'TestFiles/DB002-Landscape-Section.docx',
    categories: ['page-geometry'],
    rationale: 'Landscape dimensions and margin placement.',
    disposition: {
      kind: 'environment',
      rationale: 'Page dimensions match exactly; unchanged by the font contract, so the residual is ' +
        'paragraph spacing and line breaking of the SAME fonts differing between the two engines. ' +
        'Reduced (issue #404, visual-parity-reductions.spec.ts `landscape-spacing`): four identical ' +
        'single-line Calibri paragraphs in a landscape section — both engines lay them out with ' +
        'UNIFORM pitch and identical x-extents, Docxodus at 29 px/paragraph vs LibreOffice at 30 px ' +
        '(the auto-line-spacing line box of the same Carlito differing by 1 px/line), ink bands 12 ' +
        'vs 14 rows tall (glyph rasterization spread). The residual is that per-line metric, isolated.',
      reference: 'https://github.com/JSv4/Docxodus/issues/404',
    },
  },
  {
    id: 'running-content',
    path: 'TestFiles/DB005-Headers-With-Images.docx',
    categories: ['headers-footers', 'images', 'multi-section'],
    rationale: 'Default/first/even running stories, embedded images, and section inheritance.',
    disposition: {
      kind: 'environment',
      rationale: 'Close since PR #372 and issue #377; the residual is substituted-font line metrics.',
    },
  },
  {
    id: 'inline-image',
    path: 'TestFiles/HC042-Image-Png.docx',
    categories: ['images'],
    rationale: 'Simple image sizing and placement control.',
    disposition: {
      kind: 'environment',
      rationale: 'The image and text are separate source paragraphs; the discrepancy is ' +
        'indentation/font/wrapping, not an inline-flow failure. Unchanged by the font contract, ' +
        'so the engines lay out the same fonts differently. Reduced (issue #404, ' +
        'visual-parity-reductions.spec.ts `inline-image`): a generated 150x75 px inline picture ' +
        'between two text paragraphs renders at EXACTLY the declared wp:extent at the margin in ' +
        'both engines (1 px origin offset), and the following text resumes 20 px below the image ' +
        'in Docxodus vs 16 px in LibreOffice — extent and flow agree; the residual is the line ' +
        'box/baseline gap of the same fonts.',
      reference: 'https://github.com/JSv4/Docxodus/issues/404',
    },
  },
  {
    id: 'chart',
    path: 'TestFiles/HC043-Chart.docx',
    categories: ['charts'],
    rationale: 'Chart fallback and drawing geometry.',
    disposition: {
      kind: 'environment',
      rationale: 'Close since PR #374 for this clustered-column fixture; the residual is text ' +
        'antialiasing. Other chart families remain an unsupported feature the corpus does not yet ' +
        'exercise.',
    },
  },
  {
    id: 'shape',
    path: 'TestFiles/DB011-Body-With-Shape.docx',
    categories: ['shapes'],
    rationale: 'Body-level shape rendering and anchoring.',
    disposition: {
      kind: 'reference-deviation',
      rationale: 'Anchor origin and width match exactly since PR #381, and the auto-fit height now ' +
        'follows the laid-out text since issue #396 (the box was 16 px short, it is now 2.7 px ' +
        'tall against LibreOffice). The residual is the shape outline: CSS draws a border INSIDE ' +
        'the border box and adds it to an auto height, while DrawingML strokes `a:ln` centered on ' +
        'the shape boundary, so it does not enlarge the shape. Docxodus follows the declared ' +
        'insets, which PR #381 pinned deliberately.',
      reference: 'https://github.com/JSv4/Docxodus/issues/396',
    },
  },
  {
    id: 'fields-and-tabs',
    path: 'TestFiles/HC022-Table-Of-Contents.docx',
    categories: ['fields', 'text'],
    rationale: 'Field results, tab leaders, and right-aligned page numbers.',
    disposition: {
      kind: 'reference-deviation',
      rationale: 'Tab targets and leaders are correct since PR #380, and TOC entry line height is ' +
        'correct since issue #396 — entries land within 0.12 px of LibreOffice, where they were ' +
        'displaced by up to 14.7 px, and ink F1 is 0.99775. The residual is hyperlink styling, ' +
        'attributed in issue #397: the entry runs carry `w:rStyle w:val="Hyperlink"`, and that ' +
        'character style declares `w:color 0563C1` and `w:u single`. Docxodus paints exactly the ' +
        'declared colour; LibreOffice paints the entries black, ignoring the style the file ' +
        'applies. Docxodus follows the OOXML, so the output is not changed to match.',
      reference: 'https://github.com/JSv4/Docxodus/issues/397',
    },
  },
  {
    id: 'footnote',
    path: 'TestFiles/CA/CA008-Footnote-Reference.docx',
    categories: ['footnotes'],
    rationale: 'Footnote reference, separator, note text, and bottom-of-page placement.',
    disposition: {
      kind: 'environment',
      rationale: 'Note placement is fixed (issue #378): the last note line ends on the bottom margin ' +
        'in both engines and the separator-to-note gap matches. Issue #396 additionally stopped ' +
        'the superscript reference from inflating its line box, which had pushed the body line ' +
        'down the page; the case is now close. The residual is substituted-font rasterization, ' +
        'and older LibreOffice drawing its 25%-column separator instead of the two-inch default.',
    },
  },
  // --- Second wave (issue #400). Entries entered as `unattributed` (which strict-gates) and
  // --- were triaged from the first measured run — see BASELINE.md's second-wave section for
  // --- the measurements each disposition below cites.
  {
    id: 'chart-stacked',
    path: 'TestFiles/VP/VP001-Chart-Stacked-Column.docx',
    categories: ['charts'],
    rationale: 'Stacked column chart family — HC043\'s clustered fixture regrouped as stacked, ' +
      'so segment stacking and axis rescaling are exercised rather than cluster offsets.',
    disposition: {
      kind: 'environment',
      rationale: 'Projected since PR #417 (stacked grouping from cached data): the case moved ' +
        'severe/blank (ink F1 0.00000) to minor with near-exact ink (F1 0.97696, SSIM 0.97053) ' +
        'in the 2026-08-13 record refresh. The residual now mirrors the clustered `chart` ' +
        'control case: label antialiasing and segment-edge rounding of the same substituted ' +
        'fonts and theme colors.',
      reference: 'https://github.com/JSv4/Docxodus/pull/417',
    },
  },
  {
    id: 'chart-pie',
    path: 'TestFiles/CU002-Chart-Cached-Data-02.docx',
    categories: ['charts'],
    rationale: 'Pie chart family (3-D pie with cached data, chart style/colors parts).',
    disposition: {
      kind: 'renderer-bug',
      rationale: 'pie3DChart projects since PR #417 and both engines draw the full pie (ink F1 ' +
        '0.00000 to 0.67052), so "unsupported" is obsolete. The residual is renderer-owned ' +
        'drawing geometry: slice start angle/ordering and 3-D perspective differ from ' +
        'LibreOffice\'s, and the title/legend text sits offset — large filled slices turn small ' +
        'angular differences into many perceptually-different pixels.',
      reference: 'https://github.com/JSv4/Docxodus/issues/411',
    },
  },
  {
    id: 'chart-line',
    path: 'TestFiles/CU004-Chart-Cached-Data-04.docx',
    categories: ['charts'],
    rationale: 'Line chart family (standard grouping, cached data).',
    disposition: {
      kind: 'renderer-bug',
      rationale: 'lineChart projects since PR #417 (ink F1 0.00000 to 0.50465). The dominant ' +
        'measured residual is renderer-owned: the cached categories are date SERIALS, which ' +
        'LibreOffice renders through the axis number format as dates ("9/1/2013") while ' +
        'Docxodus prints the raw serial ("41518") — visibly doubled axis labels in the overlay; ' +
        'the polylines also sit a few pixels off vertically.',
      reference: 'https://github.com/JSv4/Docxodus/issues/411',
    },
  },
  {
    id: 'wrapped-image-square',
    path: 'TestFiles/DB007-WhitePaper.docx',
    categories: ['images', 'image-wrap'],
    rationale: 'Floating picture with square wrap inside a real multi-paragraph document.',
    disposition: {
      kind: 'unattributed',
      rationale: 'Text wraps beside the picture since PR #418, dissolving the unsupported-feature ' +
        'attribution (ink F1 0.39352 to 0.53248 in the 2026-08-13 refresh). The remaining severe ' +
        'residual mixes wrap-pocket geometry with a page-wide vertical drift that starts at the ' +
        'title, and has not been re-triaged — unattributed so it gates until it is.',
      reference: 'https://github.com/JSv4/Docxodus/issues/412',
    },
  },
  {
    id: 'wrapped-image-tight',
    path: 'TestFiles/VP/VP002-Image-Wrap-Tight.docx',
    categories: ['images', 'image-wrap'],
    rationale: 'Floating picture with tight wrap (wrapTight + wrapPolygon) and enough text to wrap.',
    disposition: {
      kind: 'unattributed',
      rationale: 'Wrapping implemented by PR #418: ink F1 0.51179 to 0.94860 in the 2026-08-13 ' +
        'refresh, with SSIM 0.84881 sitting just under the major boundary (0.85). The prior ' +
        'attribution is obsolete and the small remaining residual (wrap-pocket line breaks vs ' +
        'the polygon, same-font line metrics) has not been re-triaged — unattributed so it ' +
        'gates until it is.',
      reference: 'https://github.com/JSv4/Docxodus/issues/412',
    },
  },
  {
    id: 'nested-table',
    path: 'TestFiles/WC/WC043-Nested-Table.docx',
    categories: ['tables', 'nested-tables'],
    rationale: 'A table nested inside another table\'s cell — border collapse and width inheritance.',
    disposition: {
      kind: 'reference-deviation',
      rationale: 'Both engines nest the table correctly and start it at the same margin row (96). ' +
        'The outer ink band is 96–163 in Docxodus vs 96–152 in LibreOffice, and the ~11 px is the ' +
        'document default `w:spacing w:after="160"` (10.7 px) on the "Before." paragraph that ' +
        'precedes the nested table: Docxodus paints the declared spacing, LibreOffice suppresses ' +
        'it. Nothing in the OOXML licenses the suppression; Word-behavior evidence for this exact ' +
        'shape is the open question, as with the `numbered-lists` margin history.',
    },
  },
  {
    id: 'two-column-section',
    path: 'TestFiles/VP/VP003-Two-Column-Section.docx',
    categories: ['columns', 'page-geometry'],
    rationale: 'Single-column title section followed by a continuous two-column (`w:cols`) section.',
    disposition: {
      kind: 'renderer-bug',
      rationale: 'The two issue-#413 failures are fixed by PR #419: the continuous section stays ' +
        'on one page (1/1, the corpus\'s only page-count mismatch gone) and `w:cols w:num="2"` ' +
        'renders two columns (ink F1 0.07207 to 0.45891). The remaining severe residual is ' +
        'renderer-owned column geometry: the column fill/split point differs — Docxodus breaks ' +
        'to the second column at a different paragraph than LibreOffice, displacing the right ' +
        'column\'s content, with per-line vertical offsets inside each column.',
      reference: 'https://github.com/JSv4/Docxodus/issues/413',
    },
  },
  {
    id: 'endnote',
    path: 'TestFiles/WC/WC036-Endnote-With-Table-Before.docx',
    categories: ['endnotes'],
    rationale: 'Endnote reference, end-of-document note placement, and a table inside the note.',
    disposition: {
      kind: 'unattributed',
      rationale: 'PR #420 flowed endnotes onto the page and fixed the roman citation numbering, ' +
        'making the prior rationale obsolete and moving the numbers (ink F1 0.93747 to 0.91258, ' +
        'SSIM 0.84081 to 0.82594 in the 2026-08-13 refresh — the note now occupies page space ' +
        'the old render left empty). The remaining dominant residual — a systematic ~one-line ' +
        'downward shift of the whole body plus note-table row offsets — has not been re-triaged ' +
        'since the flow change; unattributed so it gates until it is.',
      reference: 'https://github.com/JSv4/Docxodus/issues/414',
    },
  },
  {
    id: 'legal-contract',
    path: 'TestFiles/VP/VP004-Legal-Contract.docx',
    categories: ['contract', 'lists', 'fields', 'text'],
    rationale: 'Realistic services agreement: cached TOC with hyperlink entries and PAGEREF ' +
      'fields, multilevel heading numbering bound to Heading1/Heading2, (a)/(i) sub-clause ' +
      'lists, cached REF cross-references, and a borderless signature table — the heavy-' +
      'numbering legal shape the library\'s positioning makes central.',
    disposition: {
      kind: 'renderer-bug',
      rationale: 'The formerly dominant residual — the list-number suffix tab advancing to the ' +
        'next default stop instead of the declared text indent — was fixed by PR #421, with a ' +
        'small corpus effect (ink F1 0.52148 to 0.52440), so the case\'s severity was never ' +
        'that one defect alone. Remaining measured residuals, renderer side: heading ' +
        'space-before painted at page tops where LibreOffice drops it (16 px per page top), ' +
        'and accumulated per-clause offsets across the three pages; the cached TOC also ' +
        'carries the recorded issue-#397 hyperlink-style deviation (not gating).',
      reference: 'https://github.com/JSv4/Docxodus/issues/415',
    },
  },
  {
    id: 'tracked-deletion',
    path: 'TestFiles/FA/RevTracking/001-DeletedRun.docx',
    categories: ['tracked-changes'],
    rationale: 'Final-view handling of a tracked deletion.',
    revisionMode: 'accepted',
    disposition: {
      kind: 'environment',
      rationale: 'Identical accepted bytes go to both engines; the differences cluster around heading ' +
        'metrics and wrapping rather than revision semantics. The font contract improved this case ' +
        '(shared Calibri Light pinning) — the residual is the engines laying out the same fonts ' +
        'differently. Reduced (issue #404, visual-parity-reductions.spec.ts `heading-metrics`): a ' +
        'Calibri Light (Carlito) heading over a Calibri body line advances IDENTICALLY in both ' +
        'engines (24 px heading-to-body in each); what differs is the heading\'s ink height, 17 ' +
        'vs 19 rows — glyph rasterization spread of the same substituted face, not layout.',
      reference: 'https://github.com/JSv4/Docxodus/issues/404',
    },
  },
];
