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
      kind: 'unattributed',
      rationale: 'Ink geometry aligns; the perceptual delta is dominated by fill/border color that has ' +
        'not been reduced to a minimal case deciding which engine matches the OOXML color semantics.',
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
