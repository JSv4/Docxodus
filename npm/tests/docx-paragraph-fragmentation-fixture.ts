import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/** Original 42-line paragraph, with independently selectable revision and floating textbox. */
export function paragraphFragmentationDocx(tracked: boolean, anchored: boolean, lineCount = 42): Uint8Array {
  const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
  const drawing = `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>4800000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
    <wp:extent cx="1400000" cy="400000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Fixture drawing"/><wp:cNvGraphicFramePr/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>
    <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1400000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>
    <wps:txbx><w:txbxContent><w:p>${run('ORANGE MARKER')}</w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr>
    </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  let contents = (anchored ? drawing : '') + Array.from({ length: lineCount }, (_, i) =>
    (i ? '<w:r><w:br/></w:r>' : '') + run(`Workshop row ${String(i + 1).padStart(2, '0')} carries a copper token.`)).join('');
  if (tracked) contents = `<w:ins w:id="7" w:author="Fixture Author" w:date="2026-02-03T10:00:00Z">${contents}</w:ins>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}${name === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="wps"><w:body><w:p>${contents}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="9000"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="480" w:footer="480"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
