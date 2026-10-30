import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type UnpaintedSourceKind = 'rowless-table' | 'empty-cell' | 'joined-hidden' | 'joined-visible' | 'standalone-hidden' | 'cell-hidden' | 'joined-hidden-empty';

/** Independently authored package containing a structural object with no visible content. */
export function unpaintedSourceDocx(kind: UnpaintedSourceKind): Uint8Array {
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const hidden = kind === 'joined-hidden' || kind === 'joined-hidden-empty';
  const hiddenParagraph = '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr></w:p>';
  const content = kind === 'standalone-hidden' ? paragraph('Comet observation.') + hiddenParagraph
    : kind === 'cell-hidden' ? `<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc>${paragraph('Comet observation.')}${hiddenParagraph}</w:tc></w:tr></w:tbl>`
    : kind === 'rowless-table' || kind === 'empty-cell'
    ? `<w:tbl><w:tblPr/><w:tblGrid/>${kind === 'empty-cell' ? '<w:tr><w:tc><w:p/></w:tc></w:tr>' : ''}</w:tbl>`
      + paragraph('Comet observation.')
    : '<w:p><w:pPr><w:rPr><w:specVanish/></w:rPr></w:pPr><w:r><w:t>Observation:</w:t></w:r></w:p>'
      + `<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r>${hidden ? '<w:rPr><w:vanish/></w:rPr>' : ''}
        <w:t>Internal telescope note.</w:t></w:r>${kind === 'joined-hidden-empty' ? '<w:r><w:t/></w:r>' : ''}</w:p>`
      + paragraph('Orbit verified.');
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
    </Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${content}
      <w:sectPr><w:pgSz w:w="7000" w:h="9000"/>
        <w:pgMar w:top="900" w:right="900" w:bottom="900" w:left="900" w:header="300" w:footer="300"/>
      </w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
