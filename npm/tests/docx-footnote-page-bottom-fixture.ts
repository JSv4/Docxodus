import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export interface FootnotePageBottomOptions {
  /** `w:line` of the body paragraphs (480 is double spacing). */
  line?: number;
  paragraphs?: number;
  /** How many notes each paragraph cites, about 70% of the way through it (default one each). */
  cites?: number[];
  widowControl?: boolean;
  /** The note that runs to four lines instead of one. */
  longNote?: number;
  /** `w:compatSetting w:name="compatibilityMode"`; unstated by default, which reads as mode 14 or earlier. */
  compatibilityMode?: number;
}

/**
 * Issue #1738: justified paragraphs, each citing one-line notes about 70% of the way through.
 * At double spacing the fifth paragraph runs across the bottom of page 1. The defaults are the
 * converter's for the issue's unstyled document (10pt, notes at w:line=278), in Liberation Serif,
 * which has Times New Roman's metrics.
 */
export function footnotePageBottomDocx(options: FootnotePageBottomOptions = {}): Uint8Array {
  const { line = 480, paragraphs = 5, widowControl = true } = options;
  const cites = options.cites ?? Array.from({ length: paragraphs }, () => 1);
  const superscript = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>';
  const body: string[] = [];
  const notes: string[] = [];
  for (let paragraph = 1; paragraph <= paragraphs; paragraph++) {
    const text = [1, 2, 3, 4].map(row => `In season ${paragraph} the orchard crew pruned row ${row}, logged the yield by hand, `
      + 'and compared it with the estimate kept in the shed.').join(' ');
    const cut = text.indexOf(' ', Math.floor(text.length * 0.7));
    let references = '';
    for (let cite = 0; cite < cites[paragraph - 1]; cite++) {
      const id = notes.length + 1;
      references += `<w:r>${superscript}<w:footnoteReference w:id="${id}"/></w:r>`;
      const more = id === options.longNote ? (' The sheet records the crew, the rows pruned, the hand count of the yield'
        + ' and the estimate kept in the shed, and the cooperative files it with the other sheets of the season.').repeat(2) : '';
      notes.push(`<w:footnote w:id="${id}"><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r>${superscript}<w:footnoteRef/></w:r>`
        + `<w:r><w:t xml:space="preserve"> Yield sheet ${id} is filed with the cooperative.${more}</w:t></w:r></w:p></w:footnote>`);
    }
    body.push(`<w:p><w:pPr>${widowControl ? '' : '<w:widowControl w:val="0"/>'}`
      + `<w:spacing w:after="0" w:line="${line}" w:lineRule="auto"/><w:ind w:firstLine="720"/><w:jc w:val="both"/></w:pPr>`
      + `<w:r><w:t xml:space="preserve">${text.slice(0, cut)}</w:t></w:r>${references}`
      + `<w:r><w:t xml:space="preserve">${text.slice(cut)}</w:t></w:r></w:p>`);
  }
  const separator = (type: string, id: number, mark: string) => `<w:footnote w:type="${type}" w:id="${id}"><w:p><w:pPr>`
    + `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r>${mark}</w:r></w:p></w:footnote>`;
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>
      <Relationship Id="notes" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
      <Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/>`),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="278" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`
      + (options.compatibilityMode === undefined ? '' : '<w:compat><w:compatSetting w:name="compatibilityMode"'
        + ` w:uri="http://schemas.microsoft.com/office/word" w:val="${options.compatibilityMode}"/></w:compat>`)
      + '</w:settings>',
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}">${separator('separator', -1, '<w:separator/>')}`
      + `${separator('continuationSeparator', 0, '<w:continuationSeparator/>')}${notes.join('')}</w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
