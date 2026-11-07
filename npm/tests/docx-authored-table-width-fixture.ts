import { fittingImageRun, fittingTilePng } from './docx-page-fitting-fixture.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

export function tableDocx(width: number, type: string, centered: boolean, rows = 1, indent = 0,
  options: { layout?: 'fixed' | 'autofit' | 'inferred'; text?: string; imageWidth?: number; anchored?: boolean; textbox?: boolean } = {}): Uint8Array {
  const picture = (id: number) => {
    if (!options.imageWidth) return '';
    let image = fittingImageRun(id, options.imageWidth, options.textbox ? 48 : 24);
    if (options.textbox) image = image.replace(/<a:graphicData[\s\S]*?<\/a:graphicData>/,
      `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>The ferry crews repaint the harbor signs every autumn before the storms arrive at the pier</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData>`);
    return options.anchored ? image.replace('<wp:inline>', '<wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="character"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')
      .replace('<wp:docPr ', '<wp:wrapNone/><wp:docPr ').replace('</wp:inline>', '</wp:anchor>') : image;
  };
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}<Relationship Id="tile" Type="${R_NS}/image" Target="media/tile.png"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:tbl><w:tblPr>
      <w:tblW w:type="${type}" w:w="${type === 'dxa' ? width * 20 : type === 'pct' ? 5000 : 0}"/>
      <w:jc w:val="${centered ? 'center' : 'left'}"/><w:tblInd w:type="dxa" w:w="${indent * 20}"/>
      <w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="8" w:color="222222"/>`).join('')}</w:tblBorders>
      ${(options.layout ?? (type === 'dxa' ? 'fixed' : 'inferred')) === 'inferred' ? '' : `<w:tblLayout w:type="${options.layout ?? 'fixed'}"/>`}<w:tblCellMar><w:left w:type="dxa" w:w="0"/><w:right w:type="dxa" w:w="0"/></w:tblCellMar>
      </w:tblPr><w:tblGrid><w:gridCol w:w="${width * 10}"/><w:gridCol w:w="${width * 10}"/></w:tblGrid>
      ${Array.from({ length: rows }, (_, row) => `<w:tr><w:trPr><w:trHeight w:val="640" w:hRule="atLeast"/></w:trPr>${['A', 'B'].map(column => `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="${width * 10}"/></w:tcPr><w:p>${column === 'B' ? picture(row + 1) : ''}<w:r><w:t>${options.text ?? `Quill${row + 1}${column}`}</w:t></w:r></w:p></w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>
      <w:p><w:r><w:t>Finish</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  return storedZip([...Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })),
    { name: 'word/media/tile.png', data: fittingTilePng() }]);
}

