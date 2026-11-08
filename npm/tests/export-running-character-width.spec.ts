import { expect, test } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox:true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const words = Array.from({ length:12 }, (_, i) => `W${String(i).padStart(4, '0')}`);
function source(owner:'header' | 'footer', percent:number, variants:boolean) {
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const rels = (content:string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const section = (width:number, reference:boolean) => '<w:sectPr>'
    + (reference ? `<w:${owner}Reference w:type="default" r:id="story"/>`
      + (variants ? ['first', 'even'].map(kind => `<w:${owner}Reference w:type="${kind}" r:id="${kind}"/>`).join('') : '') : '')
    + `<w:pgSz w:w="${width}" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>`
    + (reference && variants ? '<w:titlePg/>' : '') + '</w:sectPr>';
  const styled = variants ? '<w:pStyle w:val="RunningTitle"/>' : '';
  const body = `<w:p><w:pPr>${styled}</w:pPr><w:r><w:t>Body0</w:t></w:r></w:p>`
    + `<w:p><w:pPr>${styled}<w:pageBreakBefore/>${section(5760, true)}</w:pPr><w:r><w:t>Body1</w:t></w:r></w:p>`
    + `<w:p><w:pPr>${styled}</w:pPr><w:r><w:t>Body2</w:t></w:r></w:p><w:p><w:pPr>${styled}<w:pageBreakBefore/></w:pPr><w:r><w:t>Body3</w:t></w:r></w:p>`
    + section(7200, false);
  const tag = owner === 'header' ? 'hdr' : 'ftr';
  const stories = variants ? [[owner, 'story', 'W'], [owner + '-first', 'first', 'F'], [owner + '-even', 'even', 'E']] : [[owner, 'story', 'W']];
  const parts:Record<string,string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ['settings', 'settings'], ...stories.map(([name]) => [name, owner])].map(([name, suffix]) =>
        `<Override PartName="/word/${name}.xml" ContentType="${type}${suffix}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': rels(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': rels(`<Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/>`
      + `<Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/>`
      + stories.map(([name, id]) => `<Relationship Id="${id}" Type="${R_NS}/${owner}" Target="${name}.xml"/>`).join('')),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>`
      + (variants ? '<w:style w:type="paragraph" w:styleId="RunningTitle"><w:name w:val="RunningTitle"/></w:style>' : '') + '</w:styles>',
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}">${variants ? '<w:evenAndOddHeaders/>' : ''}</w:settings>`,
    ...Object.fromEntries(stories.map(([name, , prefix]) => [`word/${name}.xml`, `<w:${tag} xmlns:w="${W_NS}"><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`
      + `<w:r><w:rPr><w:w w:val="${percent}"/></w:rPr><w:t>${words.map(w => prefix + w.slice(1)).join(' ')}</w:t></w:r></w:p>`
      + (variants ? '<w:p><w:fldSimple w:instr=" STYLEREF RunningTitle "><w:r><w:t>Cached</w:t></w:r></w:fldSimple></w:p>' : '') + `</w:${tag}>`])),
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}</w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data:xml(value) })));
}

for (const owner of ['header', 'footer'] as const) for (const [percent, variants] of [[80, false], [100, false], [150, false], [80, true]] as const)
  test(`${owner} at ${percent}% ${variants ? 'resolves first/even stories and fields' : 'repeats'} at each inherited section width`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = source(owner, percent, variants);
    writeFileSync(testInfo.outputPath('source.docx'), bytes);
    const result = await renderDocxArtifacts(bytes, { browser, outputs:['pdf', 'html'], reviewProfile:'final', commentProfile:'hidden' });
    expect(result.pageCount).toBe(4);
    const task = getDocument({ data:new Uint8Array(result.pdf!), verbosity:0 });
    try {
      const pdf = await task.promise;
      for (let n = 1; n <= pdf.numPages; n++) {
        const rows = new Map<number,string[]>();
        const items = (await (await pdf.getPage(n)).getTextContent()).items;
        const prefix = variants ? ['F', 'E', 'W', 'E'][n - 1] : 'W';
        const expected = words.map(w => prefix + w.slice(1));
        for (const item of items) if ('str' in item && /^[WEF]\d/.test(item.str.trim())) {
          const y = Math.round(item.transform[5] * 100);
          rows.set(y, [...(rows.get(y) ?? []), ...item.str.trim().split(/\s+/)]);
          expect(item.width / item.str.length).toBeCloseTo(1229 / 2048 * 12 * percent / 100, 2);
          expect(item.height).toBeCloseTo(12, 2);
        }
        const band = n <= 2 ? 144 : 216;
        const perLine = Math.floor((band / (1229 / 2048 * 12 * percent / 100) + 1) / 6);
        expect([...rows.values()]).toEqual(Array.from({ length:Math.ceil(words.length / perLine) }, (_, i) =>
          expected.slice(i * perLine, (i + 1) * perLine)));
        const tokens = items.flatMap(i => 'str' in i ? i.str.trim().split(/\s+/) : []);
        expect(tokens.filter(t => t === `Body${n - 1}`)).toHaveLength(variants ? 2 : 1);
        expect(tokens).not.toContain('Cached');
      }
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const fragments = await page.locator(`.page-${owner} p[data-source-anchor-id]`).evaluateAll(ps => ps.map(p => {
      const box = p.getBoundingClientRect(), page = p.closest<HTMLElement>('.page-box')!, paper = page.getBoundingClientRect();
      return { anchor:p.getAttribute('data-source-anchor-id'), page:Number(page.dataset.pageNumber), text:p.textContent!.trim(),
        geometry:{ x:(box.left - paper.left) * .75, y:(box.top - paper.top) * .75, width:box.width * .75, height:box.height * .75 } };
    }).filter(p => /^[WEF]\d/.test(p.text)));
    expect(fragments).toHaveLength(4);
    for (const fragment of fragments) {
      const prefix = variants ? ['F', 'E', 'W', 'E'][fragment.page - 1] : 'W';
      expect(fragment.text.split(/\s+/)).toEqual(words.map(w => prefix + w.slice(1)));
      const mapped = result.pageMap.fragments.find(f => f.anchorId === fragment.anchor && f.pageNumber === fragment.page && f.story === owner)!;
      expect(mapped).toBeDefined();
      for (const key of ['x', 'y', 'width', 'height'] as const)
        expect(Math.abs(mapped.geometry[key] - fragment.geometry[key])).toBeLessThan(.1);
    }
  });
