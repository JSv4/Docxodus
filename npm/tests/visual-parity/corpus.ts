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

export interface VisualCorpusEntry {
  id: string;
  path: string;
  categories: VisualCategory[];
  rationale: string;
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
  },
  {
    id: 'merged-table',
    path: 'TestFiles/HC029-Table-Merged-Cells.docx',
    categories: ['tables'],
    rationale: 'Merged-cell borders, widths, and vertical geometry.',
  },
  {
    id: 'numbered-lists',
    path: 'TestFiles/DB012-Lists-With-Different-Numberings.docx',
    categories: ['lists'],
    rationale: 'Multiple numbering definitions and indentation behavior.',
  },
  {
    id: 'multi-section',
    path: 'TestFiles/DB001-Sections.docx',
    categories: ['multi-section', 'page-geometry'],
    rationale: 'Multiple sections and page-size transitions across six pages.',
  },
  {
    id: 'landscape-section',
    path: 'TestFiles/DB002-Landscape-Section.docx',
    categories: ['page-geometry'],
    rationale: 'Landscape dimensions and margin placement.',
  },
  {
    id: 'running-content',
    path: 'TestFiles/DB005-Headers-With-Images.docx',
    categories: ['headers-footers', 'images', 'multi-section'],
    rationale: 'Default/first/even running stories, embedded images, and section inheritance.',
  },
  {
    id: 'inline-image',
    path: 'TestFiles/HC042-Image-Png.docx',
    categories: ['images'],
    rationale: 'Simple image sizing and placement control.',
  },
  {
    id: 'chart',
    path: 'TestFiles/HC043-Chart.docx',
    categories: ['charts'],
    rationale: 'Chart fallback and drawing geometry.',
  },
  {
    id: 'shape',
    path: 'TestFiles/DB011-Body-With-Shape.docx',
    categories: ['shapes'],
    rationale: 'Body-level shape rendering and anchoring.',
  },
  {
    id: 'fields-and-tabs',
    path: 'TestFiles/HC022-Table-Of-Contents.docx',
    categories: ['fields', 'text'],
    rationale: 'Field results, tab leaders, and right-aligned page numbers.',
  },
  {
    id: 'footnote',
    path: 'TestFiles/CA/CA008-Footnote-Reference.docx',
    categories: ['footnotes'],
    rationale: 'Footnote reference, separator, note text, and bottom-of-page placement.',
  },
  {
    id: 'tracked-deletion',
    path: 'TestFiles/FA/RevTracking/001-DeletedRun.docx',
    categories: ['tracked-changes'],
    rationale: 'Final-view handling of a tracked deletion.',
    revisionMode: 'accepted',
  },
];
