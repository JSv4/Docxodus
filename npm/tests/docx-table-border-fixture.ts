import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

export type TableBorderCase = 'ordinary' | 'style' | 'merged' | 'asymmetric' | 'none' | 'wrap' | 'exact' | 'auto' | 'repeated';

/** Authored two-column grid: widths are 120pt and margins are 4.5pt. */
export function tableBorderDocx(size: number, kind: TableBorderCase = 'ordinary', margin = 90,
  placement: 'legacy' | 'text' | 'modern' | 'end' | 'rtl-end' = 'legacy'): Uint8Array {
  const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const borders = (sz: number, tag = 'tcBorders') => `<w:${tag}>` +
    ['top', 'left', 'bottom', 'right', ...(tag === 'tblBorders' ? ['insideH', 'insideV'] : [])]
      .map(side => `<w:${side} w:val="single" w:sz="${sz}" w:color="205080"/>`).join('') + `</w:${tag}>`;
  const cell = (text: string, sz: number, span = 1) => `<w:tc><w:tcPr><w:tcW w:w="${2400*span}" w:type="dxa"/>`
    + (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '')
    + (kind === 'none' || kind === 'style' ? '' : borders(sz)) + '<w:tcMar>'
    + ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="${margin}" w:type="dxa"/>`).join('')
    + `</w:tcMar></w:tcPr>${paragraph(text)}</w:tc>`;
  const wrapText = 'MMMMMMMMMM iii';
  const row = `<w:tr>${kind === 'exact' ? '<w:trPr><w:trHeight w:val="500" w:hRule="exact"/></w:trPr>' : kind === 'auto' ? '<w:trPr><w:trHeight w:val="1440" w:hRule="auto"/></w:trPr>' : ''}`
    + cell(kind === 'wrap' ? wrapText : kind === 'exact' ? 'Cypress first second third fourth' : 'CypressNorth', size)
    + cell('CypressSouth', kind === 'asymmetric' ? 8 : size) + '</w:tr>';
  const table = '<w:tbl><w:tblPr>' + (kind === 'style' ? '<w:tblStyle w:val="ThinGrid"/>' : '')
    + '<w:tblW w:w="4800" w:type="dxa"/>' + (placement === 'text' ? '<w:tblInd w:w="0" w:type="dxa"/>' : '')
    + (placement.endsWith('end') ? '<w:jc w:val="right"/>' : '') + (placement === 'rtl-end' ? '<w:bidiVisual/>' : '')
    + '<w:tblLayout w:type="fixed"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>'
    + (kind === 'repeated'
      ? `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('CypressHeader',24)}${cell('CypressLabel',24)}</w:tr>`
        + Array.from({length:6},(_,i)=>`<w:tr><w:trPr><w:trHeight w:val="3200" w:hRule="atLeast"/></w:trPr>`
          + cell(`CypressRow${i+1}`,size)+cell(`CypressValue${i+1}`,size)+'</w:tr>').join('')
      : (kind === 'merged' ? `<w:tr>${cell('CypressMerged', size, 2)}</w:tr>` : '') + row) + '</w:tbl>';
  const parts = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings']].map(([name,type]) =>
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles','settings'].map(name => `<Relationship Id="${name}" Type="${r}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>',
    'word/document.xml': `<w:document xmlns:w="${w}"><w:body>${table}${paragraph('CypressAfter')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:left="1440" w:right="1440" w:bottom="1440" w:header="360" w:footer="360"/></w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>`
      + '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>'
      + (kind === 'style' ? `<w:style w:type="table" w:styleId="ThinGrid"><w:name w:val="Thin grid"/><w:tblPr>${borders(size,'tblBorders')}</w:tblPr></w:style>` : '') + '</w:styles>',
    'word/settings.xml': `<w:settings xmlns:w="${w}">` + (placement === 'modern'
      ? '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' : '') + '</w:settings>',
  };
  return storedZip(Object.entries(parts).map(([name,value]) => ({ name, data: xml(value) })));
}
