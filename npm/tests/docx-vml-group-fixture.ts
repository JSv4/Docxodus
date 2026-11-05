import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

export function vmlGroupDocx(header: boolean, kind: 'group' | 'nested' | 'separate'): Uint8Array {
  const placement = 'position:absolute;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;';
  const box = (id: string, text: string, style: string) => `<v:shape id="${id}" style="${style}" stroked="f">
    <v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>`;
  const first = box('alder', 'ALDER BOX NORTH', kind === 'separate'
    ? placement + 'margin-left:72pt;margin-top:100pt;width:120pt;height:28pt'
    : 'position:absolute;left:100;top:200;width:240;height:56');
  let second = box('birch', 'BIRCH BOX SOUTH', kind === 'separate'
    ? placement + 'margin-left:222pt;margin-top:150pt;width:120pt;height:28pt'
    : kind === 'nested' ? 'position:absolute;left:10;top:20;width:240;height:56'
      : 'position:absolute;left:400;top:300;width:240;height:56');
  if (kind === 'nested') second = `<v:group style="position:absolute;left:400;top:300;width:240;height:100"
    coordorigin="10,20" coordsize="240,100">${second}</v:group>`;
  const drawing = kind === 'separate' ? `<w:pict>${first}</w:pict><w:pict>${second}</w:pict>`
    : `<w:pict><v:group id="grove" style="${placement}margin-left:72pt;margin-top:100pt;width:300pt;height:100pt"
      coordorigin="100,200" coordsize="600,200">${first}${second}</v:group></w:pict>`;
  const carrier = `<w:p><w:r>${drawing}</w:r></w:p>`;
  const namespaces = `xmlns:w="${w}" xmlns:r="${r}" xmlns:v="urn:schemas-microsoft-com:vml"`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings'], ...(header ? [['header', 'header']] : [])]
        .map(([name, type]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')
      + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings', ...(header ? ['header'] : [])].map(name => `<Relationship Id="${name}" Type="${r}/${name}" Target="${name}.xml"/>`).join('')
      + '</Relationships>',
    'word/document.xml': `<w:document ${namespaces}><w:body><w:p><w:r><w:t>Grove register control</w:t></w:r></w:p>
      ${header ? '' : carrier}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="header"/>' : ''}
      <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440" w:header="360" w:footer="360"/>
      </w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr></w:pPrDefault>
      </w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${w}"/>`,
  };
  if (header) parts['word/header.xml'] = `<w:hdr ${namespaces}>${carrier}</w:hdr>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
