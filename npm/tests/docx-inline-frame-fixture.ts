import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original Sage paragraphs, with native Word captures for each placement/wrapping mode. */
export function inlineFrameFixture(options: {
  story?: 'body' | 'header' | 'footer' | 'table';
  wrap?: 'around' | 'notBeside' | 'none';
  normal?: boolean;
  floating?: boolean;
  spill?: boolean;
  fillerCount?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  horizontalTextDistance?: number;
  verticalTextDistance?: number;
  horizontalOffset?: string;
  horizontalAnchor?: 'text' | 'page' | 'margin';
  horizontalAlign?: 'left' | 'center' | 'right';
  secondOffset?: string;
  separate?: boolean;
  heightRule?: 'exact' | 'atLeast' | 'auto';
  tailWords?: number;
  tailPrefix?: string;
} = {}): Uint8Array {
  const { story = 'body', wrap = 'notBeside', normal = false, floating = false, spill = false,
    fillerCount = spill ? 52 : 0, spaceBefore = 0, spaceAfter = 0, horizontalOffset = '240' } = options;
  const paragraph = (text: string, properties = '') => `<w:p><w:pPr>${properties}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const distances = (options.horizontalTextDistance === undefined ? '' : ` w:hSpace="${options.horizontalTextDistance * 20}"`)
    + (options.verticalTextDistance === undefined ? '' : ` w:vSpace="${options.verticalTextDistance * 20}"`);
  const horizontal = options.horizontalAlign ? `w:xAlign="${options.horizontalAlign}"` : `w:x="${horizontalOffset}"`;
  const frame = normal ? '' : `<w:framePr w:w="3600" ${options.heightRule ? `w:h="800" w:hRule="${options.heightRule}" ` : ''}w:hAnchor="${options.horizontalAnchor ?? 'text'}" w:vAnchor="text" ${horizontal} ${floating ? 'w:y="0"' : 'w:yAlign="inline"'} w:wrap="${wrap}"${distances}/>`;
  const secondFrame = options.secondOffset === undefined ? frame : frame.replace(`w:x="${horizontalOffset}"`, `w:x="${options.secondOffset}"`);
  const content = (fillerCount ? Array.from({ length: fillerCount }, (_, i) => paragraph(`SageFiller${String(i).padStart(3, '0')}`)).join('') : paragraph('SageBefore'))
    + paragraph('SageFrameFirst', frame) + (options.separate ? paragraph('SageBetween') : '') + paragraph('SageFrameLast', secondFrame)
    + paragraph((options.tailPrefix ?? '') + (options.tailWords ? Array.from({ length: options.tailWords }, (_, i) => `SageTail${String(i).padStart(3, '0')}`).join(' ') : 'SageAfter'));
  const running = story === 'header' || story === 'footer';
  const part = story + '1';
  const names = ['document', 'styles', 'settings', ...(running ? [part] : [])];
  const body = running ? paragraph('SageDocument') : story === 'table'
    ? '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblCellMar>'
      + ['top', 'left', 'bottom', 'right'].map(side => `<w:${side} w:w="0" w:type="dxa"/>`).join('')
      + '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>'
      + content + '</w:tc></w:tr></w:tbl>' + paragraph('SageTableReceipt') : content;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${names.map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name === 'document' ? 'document.main' : name === part ? story : name}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.filter(name => name !== 'document').map(name => `<Relationship Id="${name}" Type="${R_NS}/${name === part ? story : name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}<w:sectPr>${running ? `<w:${story}Reference w:type="default" r:id="${part}"/>` : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="${spaceBefore * 20}" w:after="${spaceAfter * 20}" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  if (running) parts[`word/${part}.xml`] = `<w:${story === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W_NS}">${content}</w:${story === 'header' ? 'hdr' : 'ftr'}>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
