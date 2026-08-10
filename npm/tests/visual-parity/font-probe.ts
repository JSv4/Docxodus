import { storedZip, xml, W_NS } from '../docx-zip.js';
import type { RgbaImage } from './png.js';
import { FONT_SUBSTITUTION_CONTRACT } from './fonts.js';

/**
 * The synthetic probe that tells a renderer regression apart from a change in the font
 * environment.
 *
 * Both quantities it measures are functions of the glyph advances each engine actually got, so
 * if one of them silently starts using a different face — a package removed, the fontconfig
 * fragment not applied, a new distro default — the probe moves, and every corpus score that moved
 * with it did not move because of this repository. It is deliberately NOT compared against a
 * stored expectation: it compares the two engines to EACH OTHER on a document whose only variable
 * is the font.
 *
 * Two halves, because they fail differently:
 *
 * - a short line that cannot wrap, compared by its ADVANCE. This is pure font resolution: no
 *   line-breaking policy is involved, so the tolerance can be tight.
 * - a long paragraph that wraps, compared by its LINE COUNT. Substituting a different face over
 *   a paragraph this long changes how many lines it takes.
 *
 * The wrapping half is deliberately not compared by break POSITION. Measured on this host with
 * the contract satisfied and both engines confirmed on Caladea, the Cambria paragraph's widest
 * line still ended 34 px apart — the engines break identically-measured text differently. That is
 * a real difference, and one this benchmark exists to surface, but it is not font drift, and a
 * probe that cannot tell them apart is no better than the corpus scores it is meant to explain.
 */

/** Fixed prose: no numerals or punctuation clusters, so the metrics come from the letters. */
const WRAPPING_TEXT = [
  'The quick brown fox jumps over the lazy dog while the mist settles on the far river bank',
  'and every careful compositor watches the measure fill and wonders where the line will break',
  'before the paragraph reaches its final word.',
].join(' ');

/** Short enough to never reach the measure, in any of the declared families. */
const ADVANCE_TEXT = 'Docxodus font substitution contract';

const PAGE_WIDTH_TWIPS = 12240;
const PAGE_HEIGHT_TWIPS = 15840;
const MARGIN_TWIPS = 1440;
/** Half-point units, i.e. 11 pt — the size Word's own defaults use. */
const PROBE_SIZE_HALF_POINTS = 22;
/**
 * 1.5 line spacing. Single-spaced serif lines can leave no fully blank raster row between them,
 * which merges two lines into one ink band — and a band count that depends on the font's
 * ascenders is not a measurement of anything.
 */
const PROBE_LINE_TWENTIETHS = 360;

/** The families the probe exercises, in the order their paragraphs appear. */
export const PROBE_FAMILIES = FONT_SUBSTITUTION_CONTRACT.map(entry => entry.family);

function paragraph(family: string, text: string): string {
  return `<w:p><w:pPr>` +
    `<w:spacing w:before="0" w:after="240" w:line="${PROBE_LINE_TWENTIETHS}" w:lineRule="auto"/>` +
    `</w:pPr><w:r><w:rPr>` +
    `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>` +
    `<w:sz w:val="${PROBE_SIZE_HALF_POINTS}"/><w:szCs w:val="${PROBE_SIZE_HALF_POINTS}"/>` +
    `</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/**
 * The advance lines first, one per family, then the wrapping paragraphs. The order matters: the
 * first {@link PROBE_FAMILIES}`.length` ink bands are the advance lines, which is how the
 * comparison finds them without having to guess from their width.
 */
export function generateFontProbeDocx(): Uint8Array {
  const body =
    PROBE_FAMILIES.map(family => paragraph(family, ADVANCE_TEXT)).join('') +
    PROBE_FAMILIES.map(family => paragraph(family, WRAPPING_TEXT)).join('');

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:sz w:val="${PROBE_SIZE_HALF_POINTS}"/><w:szCs w:val="${PROBE_SIZE_HALF_POINTS}"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  ${body}
  <w:sectPr>
    <w:pgSz w:w="${PAGE_WIDTH_TWIPS}" w:h="${PAGE_HEIGHT_TWIPS}"/>
    <w:pgMar w:top="${MARGIN_TWIPS}" w:right="${MARGIN_TWIPS}" w:bottom="${MARGIN_TWIPS}"
             w:left="${MARGIN_TWIPS}" w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

/** One rendered line of text: the extent of its ink. */
export interface InkLine {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

const INK_LUMINANCE_THRESHOLD = 200;
/** Rows this close together belong to the same line; a wider gap starts a new one. */
const LINE_GAP_ROWS = 3;

/** Every horizontal band of ink in `image`, top to bottom. */
export function inkLines(image: RgbaImage): InkLine[] {
  const { width, height, data } = image;
  const rows: Array<{ left: number; right: number } | null> = [];
  for (let y = 0; y < height; y++) {
    let left = -1;
    let right = -1;
    for (let x = 0; x < width; x++) {
      const i = (y * width + x) * 4;
      const luminance = (data[i] * 299 + data[i + 1] * 587 + data[i + 2] * 114) / 1000;
      if (data[i + 3] > 16 && luminance < INK_LUMINANCE_THRESHOLD) {
        if (left < 0) left = x;
        right = x;
      }
    }
    rows.push(left < 0 ? null : { left, right });
  }

  const lines: InkLine[] = [];
  let y = 0;
  while (y < height) {
    if (!rows[y]) { y++; continue; }
    let end = y;
    let cursor = y;
    while (cursor < height) {
      if (rows[cursor]) { end = cursor; cursor++; continue; }
      let gap = cursor;
      while (gap < height && !rows[gap]) gap++;
      if (gap - cursor <= LINE_GAP_ROWS && gap < height) { cursor = gap; continue; }
      break;
    }
    const spanned = rows.slice(y, end + 1).filter((row): row is { left: number; right: number } => !!row);
    lines.push({
      top: y,
      bottom: end,
      left: Math.min(...spanned.map(row => row.left)),
      right: Math.max(...spanned.map(row => row.right)),
    });
    y = end + 1;
  }
  return lines;
}

export interface FontProbeAdvance {
  family: string;
  docxodusPx: number;
  libreofficePx: number;
  deltaPx: number;
}

export interface FontProbeResult {
  /** Ink lines each engine rendered across the whole probe page. */
  docxodusLines: number;
  libreofficeLines: number;
  /** Per-family advance of the non-wrapping line, in pixels at 96 DPI. */
  advances: FontProbeAdvance[];
  maxAdvanceDeltaPx: number;
  agreed: boolean;
  problem: string;
}

/**
 * Rasteriser and hinting differences move a short line's end by a pixel or two; a different face
 * moves it by tens. The tolerance sits between them.
 */
export const PROBE_ADVANCE_TOLERANCE_PX = 3;

export function compareProbeLines(docxodus: InkLine[], libreoffice: InkLine[]): FontProbeResult {
  const families = PROBE_FAMILIES;
  const base: FontProbeResult = {
    docxodusLines: docxodus.length,
    libreofficeLines: libreoffice.length,
    advances: [],
    maxAdvanceDeltaPx: 0,
    agreed: true,
    problem: '',
  };

  // The wrapping half: substituting a different face over paragraphs this long changes how many
  // lines they take. Compared as a count, not as break positions — see the note at the top.
  if (docxodus.length !== libreoffice.length) {
    return {
      ...base,
      agreed: false,
      problem: `the engines wrapped the probe into different line counts ` +
        `(Docxodus ${docxodus.length}, LibreOffice ${libreoffice.length}); ` +
        `they are not using the same faces`,
    };
  }
  if (docxodus.length < families.length) {
    return {
      ...base,
      agreed: false,
      problem: `the probe rendered only ${docxodus.length} lines, fewer than its ` +
        `${families.length} advance lines; the probe document did not render as authored`,
    };
  }

  // The advance half: the first line per family cannot wrap, so its width is the face's own.
  base.advances = families.map((family, index) => {
    const a = docxodus[index];
    const b = libreoffice[index];
    const advance = {
      family,
      docxodusPx: a.right - a.left,
      libreofficePx: b.right - b.left,
      deltaPx: Math.abs((a.right - a.left) - (b.right - b.left)),
    };
    base.maxAdvanceDeltaPx = Math.max(base.maxAdvanceDeltaPx, advance.deltaPx);
    return advance;
  });

  const drifted = base.advances.filter(a => a.deltaPx > PROBE_ADVANCE_TOLERANCE_PX);
  if (drifted.length) {
    return {
      ...base,
      agreed: false,
      problem: `the engines measured ${drifted.map(a =>
        `${a.family} ${a.deltaPx} px apart (${a.docxodusPx} vs ${a.libreofficePx})`).join(', ')} ` +
        `against a ${PROBE_ADVANCE_TOLERANCE_PX} px tolerance; they resolved different faces`,
    };
  }
  return base;
}
