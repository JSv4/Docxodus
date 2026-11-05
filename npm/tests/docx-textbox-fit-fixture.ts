import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

export type TextBoxFit = 'implicit' | 'fixed' | 'shrink' | 'auto';
export const textboxFitLabels = ['Birch01', 'Birch02', 'Birch03'];
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const fitXml = (fit: TextBoxFit) => fit === 'implicit' ? '' : `<a:${{ fixed: 'noAutofit', shrink: 'normAutofit', auto: 'spAutoFit' }[fit]}/>`;

/** Original fixed-position story: all three labels fit on paper, but its final deleted mark does not. */
export function textboxFitDocx(fit: TextBoxFit, placement: 'edge' | 'onpage' | 'outside', inherited?: TextBoxFit,
  story: 'body' | 'header' | 'footer' = 'body'): Uint8Array {
  const content = textboxFitLabels.map(text => '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/>' +
    `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`).join('') +
    '<w:p><w:pPr><w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="exact"/>' +
    '<w:rPr><w:del w:id="9" w:author="Birch reviewer" w:date="2026-10-31T00:00:00Z"/></w:rPr></w:pPr></w:p>';
  const top = placement === 'edge' ? 750 : placement === 'outside' ? 820 : 60;
  const drawing = `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>762000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="page"><wp:posOffset>${top * 12700}</wp:posOffset></wp:positionV>` +
    '<wp:extent cx="2032000" cy="1016000"/><wp:wrapNone/><wp:docPr id="1" name="Birch edge box"/>' +
    `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
    `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx><wps:bodyPr>${fitXml(fit)}</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  const namespaces = `xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:a="${A}" xmlns:wp="${WP}" xmlns:wps="${WPS}"`;
  const host = `<w:p><w:r><w:t>Birch host complete</w:t></w:r><w:r>${drawing}</w:r></w:p>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      ['document', 'styles', 'settings', ...(story === 'body' ? [] : [story])].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('') +
      (inherited ? '<Override PartName="/word/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' : '') + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      ['styles', 'settings', ...(story === 'body' ? [] : [story]), ...(inherited ? ['theme'] : [])].map(p => `<Relationship Id="${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('') + '</Relationships>',
    'word/document.xml': `<w:document ${namespaces}><w:body>${story === 'body' ? host : '<w:p><w:r><w:t>Birch body complete</w:t></w:r></w:p>'}` +
      '<w:sectPr>' + (story === 'body' ? '' : `<w:${story}Reference w:type="default" r:id="${story}"/>`) + '<w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  if (story !== 'body') parts[`word/${story}.xml`] = `<w:${story === 'header' ? 'hdr' : 'ftr'} ${namespaces}>${host}</w:${story === 'header' ? 'hdr' : 'ftr'}>`;
  if (inherited) {
    const fill = '<a:solidFill><a:srgbClr val="202020"/></a:solidFill>';
    const font = '<a:latin typeface="Liberation Serif"/><a:ea typeface=""/><a:cs typeface=""/>';
    parts['word/theme.xml'] = `<a:theme xmlns:a="${A}" name="Birch theme"><a:themeElements><a:clrScheme name="Birch colors">` +
      ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(n => `<a:${n}><a:srgbClr val="202020"/></a:${n}>`).join('') +
      `</a:clrScheme><a:fontScheme name="Birch fonts"><a:majorFont>${font}</a:majorFont><a:minorFont>${font}</a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Birch format"><a:fillStyleLst>${fill.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${(`<a:ln>${fill}</a:ln>`).repeat(3)}</a:lnStyleLst>` +
      `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${fill.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements>` +
      `<a:objectDefaults><a:txDef><a:spPr/><a:bodyPr>${fitXml(inherited)}</a:bodyPr><a:lstStyle/></a:txDef></a:objectDefaults></a:theme>`;
  }
  return storedZip(Object.entries(parts).map(([name, data]) => ({ name, data: xml(data) })));
}
