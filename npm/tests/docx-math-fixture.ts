import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

const m = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const term = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`;
const fraction = `<m:f><m:num>${term('x')}</m:num><m:den>${term('31.4')}</m:den></m:f>`;
export const MATH_FORMULAS = {
  fraction,
  root: `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${term('x')}</m:e></m:rad>`,
  indexed: `<m:rad><m:deg>${term('3')}</m:deg><m:e>${term('x')}</m:e></m:rad>`,
  accent: `<m:acc><m:e>${term('x')}</m:e></m:acc>`,
  styled: `<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>x</m:t></m:r>`,
};
export type MathLayout = 'inline' | 'display' | 'mixed' | 'edge' | 'long';

export function mathDocx(kind: keyof typeof MATH_FORMULAS, layout: MathLayout,
  mathProperties = ''): Uint8Array {
  const display = layout === 'display' || layout === 'edge';
  const formula = `<m:oMath>${MATH_FORMULAS[kind]}</m:oMath>`;
  const words = Array.from({ length: 100 }, (_, i) => `Birch${String(i).padStart(3, '0')}`);
  const math = display ? `<m:oMathPara>${formula}</m:oMathPara>` : formula;
  const body = layout === 'long' ? run(words.slice(0, 35).join(' ') + ' ') + math + run(' ' + words.slice(35).join(' '))
    : layout === 'mixed' ? run('MathAnchor ') + math + run(' MathTail') : math;
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings']].map(([name, type]) =>
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + ['styles', 'settings'].map(name => `<Relationship Id="${name}" Type="${r}/${name}" Target="${name}.xml"/>`).join('') + '</Relationships>',
    'word/document.xml': `<w:document xmlns:w="${w}" xmlns:m="${m}"><w:body><w:p>${run('MathBefore')}</w:p>`
      + (layout === 'edge' ? Array.from({ length: 8 }, (_, i) => `<w:p>${run('Edge' + i)}</w:p>`).join('') : '')
      + `<w:p>${body}</w:p><w:p>${run('MathAfter')}</w:p><w:sectPr><w:pgSz w:w="6000" w:h="4000"/>`
      + '<w:pgMar w:top="600" w:left="600" w:bottom="600" w:right="600" w:header="300" w:footer="300"/></w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>`
      + '<w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
      + '</w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
    'word/settings.xml': `<w:settings xmlns:w="${w}" xmlns:m="${m}">${mathProperties ? `<m:mathPr>${mathProperties}</m:mathPr>` : ''}</w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
