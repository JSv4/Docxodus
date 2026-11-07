import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export type TableRemainderKind = 'rows' | 'oversized' | 'row-cant-split' | 'cell' | 'kept' | 'kept-bookmarks' | 'kept-borders' | 'cell-cant-split' | 'columns' | 'columns-fit';

export function tableRemainderDocx(kind: TableRemainderKind, mainPart = 'word/document.xml'): Uint8Array {
  const columns = kind.startsWith('columns');
  const count = kind === 'oversized' ? 18 : kind === 'columns' ? 12 : kind === 'columns-fit' ? 4 : 8;
  const width = columns ? 2800 : 6000;
  const geometry = '<w:pgSz w:w="6800" w:h="3600"/><w:pgMar w:top="400" w:right="400" w:bottom="400" w:left="400" w:header="0" w:footer="0"/>';
  const paragraph = (text: string, extra = '') => `<w:p><w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/>
    <w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/>${extra}</w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="24"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const leads = Array.from({ length: 5 }, (_, i) => paragraph(`Lead${i + 1}`,
    columns && i === 4 ? `<w:sectPr>${geometry}</w:sectPr>` : '')).join('');
  const entries = Array.from({ length: count }, (_, i) => paragraph(`Cedar${i + 1}`))
    .map((p, i) => (kind.startsWith('kept') && i < count - 1 ? p.replace('<w:keepNext w:val="0"/>', '<w:keepNext/>') : p)
      + (kind === 'kept-bookmarks' ? `<w:bookmarkStart w:id="${i}" w:name="RowEnd${i}"/><w:bookmarkEnd w:id="${i}"/>` : ''));
  if (kind === 'kept-borders') entries.forEach((p, i) => {
    entries[i] = p.replace('<w:spacing', '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="0"/></w:pBdr><w:spacing');
  });
  const row = (content: string) => `<w:tr>${kind.includes('cant-split') ? '<w:trPr><w:cantSplit/></w:trPr>' : ''}
    <w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${content}</w:tc></w:tr>`;
  const table = `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblLayout w:type="fixed"/>
    <w:tblCellMar>${['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="0" w:type="dxa"/>`).join('')}</w:tblCellMar></w:tblPr>
    <w:tblGrid><w:gridCol w:w="${width}"/></w:tblGrid>${kind.startsWith('cell') ? row(entries.join('')) : entries.map(row).join('')}</w:tbl>`;
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/${mainPart}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="${mainPart}"/></Relationships>`,
    [mainPart]: `<w:document xmlns:w="${W_NS}"><w:body>${leads}${table}${paragraph('Complete.')}
      <w:sectPr>${columns ? '<w:type w:val="continuous"/>' : ''}${geometry}${columns ? '<w:cols w:num="2" w:space="400"/>' : ''}</w:sectPr></w:body></w:document>`,
  }).map(([name, content]) => ({ name, data: xml(content) })));
}
