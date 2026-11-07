import { fittingImageRun, fittingTilePng } from './docx-page-fitting-fixture.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export type BlockHeightOptions = {
  kind: 'image' | 'row'; multiple?: number; mark?: number; imageHeight?: number; mixed?: boolean;
  top?: number; bottom?: number; rule?: 'atLeast' | 'exact' | 'auto'; rows?: number;
  borders?: number; lines?: number;
};

export function blockHeightDocx(options: BlockHeightOptions): Uint8Array {
  const single = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
  const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const paragraph = (text: string) => `<w:p><w:pPr>${single}</w:pPr>${run(text)}</w:p>`;
  let body: string;
  if (options.kind === 'image') {
    const picture = fittingImageRun(1, 100, options.imageHeight ?? 100);
    body = paragraph('PineAbove') + `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${240 * (options.multiple ?? 1)}" w:lineRule="auto"/>
      <w:rPr><w:sz w:val="${2 * (options.mark ?? 12)}"/></w:rPr></w:pPr>${picture}${options.mixed ? run('PineCaption') : ''}</w:p>` + paragraph('PineBelow');
  } else {
    const borders = options.borders ? `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side =>
      `<w:${side} w:val="single" w:sz="${options.borders! * 8}" w:color="444444"/>`).join('')}</w:tblBorders>` : '';
    body = `<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="8000"/>${borders}<w:tblLayout w:type="fixed"/>
      <w:tblCellMar><w:top w:type="dxa" w:w="${20 * (options.top ?? 0)}"/><w:left w:type="dxa" w:w="108"/>
      <w:bottom w:type="dxa" w:w="${20 * (options.bottom ?? 0)}"/><w:right w:type="dxa" w:w="108"/></w:tblCellMar></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${Array.from({ length: options.rows ?? 5 }, (_, row) =>
      `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="${options.rule ?? 'atLeast'}"/></w:trPr>${['A', 'B'].map(column =>
        `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr>${Array.from({ length: options.lines ?? 1 }, (_, line) =>
          paragraph(`Pine${row + 1}${column}${line + 1}`)).join('')}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>` + paragraph('PineAfter');
  }
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>${['document', 'styles', 'settings'].map(name =>
      `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name === 'document' ? 'document.main' : name}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(name =>
      `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}<Relationship Id="tile" Type="${R_NS}/image" Target="media/tile.png"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
      <w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${body}
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip([...Object.entries(parts).map(([name, data]) => ({ name, data: xml(data) })),
    { name: 'word/media/tile.png', data: fittingTilePng() }]);
}
