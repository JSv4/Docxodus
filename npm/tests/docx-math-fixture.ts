import { storedZip, xml, W_NS as w, R_NS as r } from './docx-zip.js';

const m = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const term = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`;
const fraction = `<m:f><m:num>${term('x')}</m:num><m:den>${term('31.4')}</m:den></m:f>`;
export const MATH_FORMULAS = {
  fraction,
  root: `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${term('x')}</m:e></m:rad>`,
  indexed: `<m:rad><m:deg>${term('3')}</m:deg><m:e>${term('x')}</m:e></m:rad>`,
  structured: `<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>${term('k=1')}</m:sub><m:sup>${term('n')}</m:sup>`
    + `<m:e><m:d><m:e><m:sSup><m:e>${term('x')}</m:e><m:sup>${term('k')}</m:sup></m:sSup></m:e></m:d></m:e></m:nary>`,
  // Aligned scripts stack the sub- and superscript flush, which MathML does not draw; still a placeholder.
  alignedScripts: `<m:sSubSup><m:sSubSupPr><m:alnScr/></m:sSubSupPr><m:e>${term('x')}</m:e><m:sub>${term('i')}</m:sub><m:sup>${term('2')}</m:sup></m:sSubSup>`,
  normalText: `<m:r><m:rPr><m:nor/></m:rPr><m:t>x</m:t></m:r>`,
  // Phantom, skewed fraction, a run of spaces, a struck border box and a bold run (#1705).
  remaining: `<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e>${term('w')}</m:e></m:phant>`
    + `<m:f><m:fPr><m:type m:val="skw"/></m:fPr><m:num>${term('1')}</m:num><m:den>${term('2')}</m:den></m:f>`
    + `<m:r><m:t xml:space="preserve">  </m:t></m:r><m:borderBox><m:borderBoxPr><m:strikeBLTR/></m:borderBoxPr><m:e>${term('x')}</m:e></m:borderBox>`
    + `<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>+y</m:t></m:r>`,
};
export type MathLayout = 'inline' | 'display' | 'mixed' | 'edge' | 'long';

export function mathDocx(kind: keyof typeof MATH_FORMULAS, layout: MathLayout,
  mathProperties = '', options: { largeProse?: boolean; markSize?: number; paragraphSize?: number; repeats?: number } = {}): Uint8Array {
  const display = layout === 'display' || layout === 'edge';
  const formula = `<m:oMath>${MATH_FORMULAS[kind]}</m:oMath>`;
  const words = Array.from({ length: 100 }, (_, i) => `Birch${String(i).padStart(3, '0')}`);
  const math = display ? `<m:oMathPara>${formula}</m:oMathPara>` : formula.repeat(options.repeats ?? 1);
  const body = layout === 'long' ? run(words.slice(0, 35).join(' ') + ' ') + math + run(' ' + words.slice(35).join(' '))
    : layout === 'mixed' ? (options.largeProse ? '<w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t>MathAnchor </w:t></w:r>' : run('MathAnchor '))
      + math + run(' MathTail') : math;
  const pPr = '<w:pPr>' + (options.paragraphSize ? '<w:pStyle w:val="MathBody"/>' : '')
    + (options.markSize ? `<w:rPr><w:sz w:val="${options.markSize}"/></w:rPr>` : '') + '</w:pPr>';
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
      + `<w:p>${pPr}${body}</w:p><w:p>${run('MathAfter')}</w:p><w:sectPr><w:pgSz w:w="6000" w:h="4000"/>`
      + '<w:pgMar w:top="600" w:left="600" w:bottom="600" w:right="600" w:header="300" w:footer="300"/></w:sectPr></w:body></w:document>',
    'word/styles.xml': `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>`
      + '<w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
      + '</w:pPr></w:pPrDefault></w:docDefaults>'
      + (options.paragraphSize ? `<w:style w:type="paragraph" w:styleId="MathBody"><w:name w:val="Math Body"/><w:rPr><w:sz w:val="${options.paragraphSize}"/></w:rPr></w:style>` : '') + '</w:styles>',
    'word/settings.xml': `<w:settings xmlns:w="${w}" xmlns:m="${m}">${mathProperties ? `<m:mathPr>${mathProperties}</m:mathPr>` : ''}</w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}
