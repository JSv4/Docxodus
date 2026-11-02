import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

export const drawingGroupLabels = ['CEDAR_WINDOW', 'MOSS_WINDOW', 'SKY_WINDOW'];

export interface DrawingGroupOrientation {
  groupFlip?: 'H' | 'V';
  nestedFlip?: 'H' | 'V';
  nestedRotation?: number;
  leafFlip?: 'H' | 'V';
  leafRotation?: number;
  upright?: boolean;
}

export function drawingGroupDocx(header: boolean, kind: 'group' | 'nested' | 'separate', inserted = false,
  rotation = 0, orientation: DrawingGroupOrientation = {}): Uint8Array {
  const shapeNs = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const groupNs = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
  const transform = (x: number, y: number, width: number, height: number, extra = '', rotate = 0, flip?: 'H' | 'V') =>
    `<a:xfrm rot="${rotate}" ${flip ? `flip${flip}="1"` : ''}><a:off x="${x*12700}" y="${y*12700}"/><a:ext cx="${width*12700}" cy="${height*12700}"/>${extra}</a:xfrm>`;
  const box = (label: string, x: number, y = 0, width = 100, height = 30) =>
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>${transform(x,y,width,height, '', label === drawingGroupLabels[1] ? (orientation.leafRotation ?? 0)*60000 : 0, orientation.leafFlip)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${label}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
      <wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" upright="${orientation.upright ? 1 : 0}"><a:noAutofit/></wps:bodyPr></wps:wsp>`;
  const group = (members: string, nested = false) => {
    const tag = nested ? 'grpSp' : 'wgp';
    const children = nested ? '<a:chOff x="127000" y="254000"/><a:chExt cx="2540000" cy="762000"/>'
      : '<a:chOff x="0" y="0"/><a:chExt cx="4318000" cy="381000"/>';
    return `<wpg:${tag}><wpg:cNvGrpSpPr/><wpg:grpSpPr>${nested
      ? transform(120,0,100,30,children,(orientation.nestedRotation ?? 0)*60000,orientation.nestedFlip) : transform(0,0,340,30,children,rotation,orientation.groupFlip)}</wpg:grpSpPr>${members}</wpg:${tag}>`;
  };
  const drawing = (payload: string, grouped: boolean, x = 72, id = 1) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0"
    simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:posOffset>${x*12700}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>2540000</wp:posOffset></wp:positionV>
    <wp:extent cx="${(grouped ? 340 : 100)*12700}" cy="381000"/><wp:wrapNone/><wp:docPr id="${id}" name="Window group"/>
    <a:graphic><a:graphicData uri="${grouped ? groupNs : shapeNs}">${payload}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const members = drawingGroupLabels.map((label,i) => kind === 'nested' && i === 1
    ? group(box(label,10,20,200,60),true) : box(label,i*120)).join('');
  let content = kind === 'separate' ? drawingGroupLabels.map((label,i) => drawing(box(label,0),false,72+i*120,i+1)).join('')
    : drawing(group(members),true);
  if (inserted) content = `<w:ins w:id="11" w:author="Fixture Author" w:date="2026-02-03T10:00:00Z">${content}</w:ins>`;
  const carrier = `<w:p>${content}</w:p>`;
  const namespaces = `xmlns:w="${w}" xmlns:r="${r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="${shapeNs}" xmlns:wpg="${groupNs}"`;
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
    'word/document.xml': `<w:document ${namespaces}><w:body><w:p><w:r><w:t>Harbor register control</w:t></w:r></w:p>
      ${header ? '' : carrier}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="header"/>' : ''}
      <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440" w:header="360" w:footer="360"/>
      </w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Docxodus Canvas Mono" w:hAnsi="Docxodus Canvas Mono"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr></w:pPrDefault>
      </w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${w}"/>`,
  };
  if (header) parts['word/header.xml'] = `<w:hdr ${namespaces}>${carrier}</w:hdr>`;
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
