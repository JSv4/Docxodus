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
        'paragraph spacing and line breaking of the SAME fonts differing between the two engines.',
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
        'so the engines lay out the same fonts differently.',
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
      kind: 'unsupported-feature',
      rationale: 'The cached-data SVG projection covers only clustered bar/column ' +
        '(WmlToHtmlConverter.Charts.cs gates on grouping "clustered"); a stacked grouping ' +
        'returns null and the chart extent renders blank — ink F1 0.00000.',
      reference: 'https://github.com/JSv4/Docxodus/issues/411',
    },
  },
  {
    id: 'chart-pie',
    path: 'TestFiles/CU002-Chart-Cached-Data-02.docx',
    categories: ['charts'],
    rationale: 'Pie chart family (3-D pie with cached data, chart style/colors parts).',
    disposition: {
      kind: 'unsupported-feature',
      rationale: 'pie3DChart is not projected (only clustered barChart is); the chart extent ' +
        'renders blank — ink F1 0.00000.',
      reference: 'https://github.com/JSv4/Docxodus/issues/411',
    },
  },
  {
    id: 'chart-line',
    path: 'TestFiles/CU004-Chart-Cached-Data-04.docx',
    categories: ['charts'],
    rationale: 'Line chart family (standard grouping, cached data).',
    disposition: {
      kind: 'unsupported-feature',
      rationale: 'lineChart is not projected (only clustered barChart is); the chart extent ' +
        'renders blank — ink F1 0.00000.',
      reference: 'https://github.com/JSv4/Docxodus/issues/411',
    },
  },
  {
    id: 'wrapped-image-square',
    path: 'TestFiles/DB007-WhitePaper.docx',
    categories: ['images', 'image-wrap'],
    rationale: 'Floating picture with square wrap inside a real multi-paragraph document.',
    disposition: {
      kind: 'unsupported-feature',
      rationale: 'Anchored-object text wrap is not implemented: the picture is placed, but ' +
        'LibreOffice wraps five lines of the paragraph beside it while Docxodus resumes the ' +
        'text below the image, displacing the page\'s lower half (ink F1 0.394).',
      reference: 'https://github.com/JSv4/Docxodus/issues/412',
    },
  },
  {
    id: 'wrapped-image-tight',
    path: 'TestFiles/VP/VP002-Image-Wrap-Tight.docx',
    categories: ['images', 'image-wrap'],
    rationale: 'Floating picture with tight wrap (wrapTight + wrapPolygon) and enough text to wrap.',
    disposition: {
      kind: 'unsupported-feature',
      rationale: 'Same text-wrap gap as wrapped-image-square, on the wrapTight/wrapPolygon ' +
        'shape: text resumes below the anchored picture instead of flowing around it ' +
        '(ink F1 0.512).',
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
      rationale: 'Two renderer-owned failures: the `w:type="continuous"` section start is ' +
        'rendered as a page break (2/1 pages — LibreOffice keeps title and body on one page), ' +
        'and `w:cols w:num="2"` is ignored entirely (the body renders as one full-width column, ' +
        'ink F1 0.072). General multi-section support is fine (multi-section is close at 6/6); ' +
        'it is specifically the continuous start type and column geometry.',
      reference: 'https://github.com/JSv4/Docxodus/issues/413',
    },
  },
  {
    id: 'endnote',
    path: 'TestFiles/WC/WC036-Endnote-With-Table-Before.docx',
    categories: ['endnotes'],
    rationale: 'Endnote reference, end-of-document note placement, and a table inside the note.',
    disposition: {
      kind: 'renderer-bug',
      rationale: 'The converter emits the endnotes section (docx2html output contains ' +
        'class="endnotes" and the note\'s table), but pagination.ts has footnote handling only ' +
        'and the section never reaches a page; the citation marker also renders decimal "1" ' +
        'where Word\'s default endnote numbering is lowercase roman ("i").',
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
      rationale: 'Dominant measured residual: the list-number suffix tab advances to the next ' +
        'default tab stop instead of the declared text indent — the (a) marker lands at 145 px ' +
        'in both engines, but the following text starts at 193 px (the next 720-twip default ' +
        'stop) in Docxodus vs 169 px in LibreOffice, where the declared `w:ind w:left="1080"` ' +
        'is 168 px — ~25 px right on every numbered clause. Secondary residuals: LibreOffice ' +
        'drops heading space-before at the top of a page where Docxodus paints the declared ' +
        'value, and the cached TOC carries the recorded issue-#397 hyperlink-style deviation.',
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
        'differently.',
    },
  },
];
