import { expect, test } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox:true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const pack = (parts:Record<string,string>) => storedZip(Object.entries(parts).map(([name,value]) => ({ name,data:xml(value) })));
const section = (width:number, type:string) => `<w:sectPr><w:type w:val="${type}"/><w:pgSz w:w="12240" w:h="15840"/>`
  + `<w:pgMar w:top="1440" w:right="${(612 - 72 - width) * 20}" w:bottom="1440" w:left="1440"/></w:sectPr>`;

function source(percent:number, count:number) {
  const labels = Array.from({ length:count }, (_, i) => `N${String(i).padStart(4,'0')}`);
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const rels = (content:string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts:Record<string,string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['document','document.main'],['styles','styles'],['footnotes','footnotes']].map(([name,suffix]) =>
        `<Override PartName="/word/${name}.xml" ContentType="${type}${suffix}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': rels(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': rels(['styles','footnotes'].map(name =>
      `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Mono" w:hAnsi="Liberation Mono"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/footnotes.xml': `<w:footnotes xmlns:w="${W_NS}"><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r></w:p>`
      + '<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:right="6480"/></w:pPr>'
      + `<w:r><w:rPr><w:w w:val="${percent}"/></w:rPr><w:t>${labels.join(' ')}</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Citation</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
  };
  return { labels, parts, bytes:pack(parts) };
}

async function inspectNotes(bytes:Uint8Array, percent:number) {
  const task = getDocument({ data:new Uint8Array(bytes),verbosity:0 });
  try {
    const pdf = await task.promise, lines:string[][] = [], tokens:string[] = [], pitches:number[] = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const rows = new Map<number,string[]>();
      for (const item of (await (await pdf.getPage(n)).getTextContent()).items) if ('str' in item && item.str.trim()) {
        tokens.push(...item.str.trim().split(/\s+/));
        if (!item.str.trim().startsWith('N')) continue;
        const y = Math.round(item.transform[5] * 100);
        rows.set(y,[...(rows.get(y) ?? []),...item.str.trim().split(/\s+/)]);
        expect(item.width / item.str.length).toBeCloseTo(1229 / 2048 * 12 * percent / 100,2);
        expect(item.height).toBeCloseTo(12,2);
      }
      lines.push(...rows.values());
      const baselines = Array.from(rows.keys());
      pitches.push(...baselines.slice(1).map((y,i) => (baselines[i] - y) / 100));
    }
    return { lines,tokens,pitches };
  } finally { await task.destroy(); }
}

for (const count of [12,216]) for (const [percent, perLine] of [[80,4],[100,3],[150,2]])
  test(`${count} footnote words at ${percent}% preserve fitted lines, glyphs and canonical identities`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const { labels,bytes } = source(percent,count);
    writeFileSync(testInfo.outputPath('source.docx'),bytes);
    const result = await renderDocxArtifacts(bytes,{ browser,outputs:['pdf','html'],reviewProfile:'final',commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');
    expect(result.pageCount > 1).toBe(count > 12);
    const { lines,tokens } = await inspectNotes(result.pdf!,percent);
    expect(lines).toEqual(Array.from({ length:Math.ceil(count / perLine) },(_,i) => labels.slice(i * perLine,(i + 1) * perLine)));
    expect(tokens.filter(t => !t.startsWith('N'))).toEqual(['Citation','1','1']);
    const output = testInfo.outputPath('note.html');writeFileSync(output,result.html!);
    await page.goto(pathToFileURL(output).href);await page.evaluate(() => document.fonts.ready);
    const fragments = await page.locator('.page-footnotes p[data-source-anchor-id]').evaluateAll(ps => ps.filter(p => p.textContent?.trim().startsWith('N')).map(p => {
      const box = p.getBoundingClientRect(), paper = p.closest<HTMLElement>('.page-box')!, origin = paper.getBoundingClientRect();
      return { anchor:p.getAttribute('data-source-anchor-id'),note:p.closest('[data-footnote-id]')?.getAttribute('data-source-anchor-id'),
        page:Number(paper.dataset.pageNumber),text:p.textContent!.trim(),geometry:{ x:(box.left - origin.left) * .75,
          y:(box.top - origin.top) * .75,width:box.width * .75,height:box.height * .75 } };
    }));
    expect(fragments.map(p => p.text).join(' ').split(/\s+/)).toEqual(labels);
    expect(new Set(fragments.map(p => p.anchor)).size).toBe(1);
    expect(new Set(fragments.map(p => p.note)).size).toBe(1);
    expect(fragments[0].anchor).toMatch(/^p:fn:/);expect(fragments[0].note).toMatch(/^fn:fn:/);
    for (const fragment of fragments) {
      const mapped = result.pageMap.fragments.find(f => f.anchorId === fragment.anchor && f.pageNumber === fragment.page && f.story === 'footnote')!;
      expect(mapped).toBeDefined();
      for (const key of ['x','y','width','height'] as const)
        expect(Math.abs(mapped.geometry[key] - fragment.geometry[key])).toBeLessThan(.1);
    }
  });

for (const percent of [80,150]) {
  test(`a mixed-reference ${percent}% note keeps its HTML policy across continuations`, async ({ browser, page }, testInfo) => {
    const { parts,labels } = source(percent,216);
    const marker = '<w:r><w:footnoteRef/></w:r>';
    parts['word/footnotes.xml'] = parts['word/footnotes.xml'].replace(`<w:p>${marker}</w:p>`,'')
      .replace('</w:pPr>',`</w:pPr>${marker}`);
    const bytes = pack(parts);writeFileSync(testInfo.outputPath('source.docx'),bytes);
    const result = await renderDocxArtifacts(bytes,{ browser,outputs:['pdf','html'],reviewProfile:'final',commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');expect(result.pageCount).toBeGreaterThan(1);
    expect((await inspectNotes(result.pdf!,100)).lines.flat()).toEqual(labels);
    await page.setContent(result.html!);
    await expect(page.locator('.page-footnotes svg')).toHaveCount(0);
    const anchors = await page.locator('.page-footnotes p[data-source-anchor-id]').evaluateAll(ps => ps.map(p => p.getAttribute('data-source-anchor-id')));
    expect(anchors.length).toBeGreaterThan(1);expect(new Set(anchors).size).toBe(1);
  });

  test(`a ${percent}% note refits from a 144pt first page into a 216pt continuation band`, async ({ browser, page }, testInfo) => {
    const { parts,labels } = source(percent,216);
    parts['word/footnotes.xml'] = parts['word/footnotes.xml'].replace('<w:ind w:right="6480"/>','');
    parts['word/document.xml'] = parts['word/document.xml'].replace('<w:body>',`<w:body><w:p><w:pPr>${section(144,'nextPage')}</w:pPr><w:r><w:t>Leading</w:t></w:r></w:p>`)
      .replace(/<w:sectPr><w:pgSz[^]*?<\/w:sectPr>/,
        Array.from({ length:40 },(_,i) => `<w:p><w:r><w:t>Body${i}</w:t></w:r></w:p>`).join('') + section(216,'continuous'));
    const bytes = pack(parts);writeFileSync(testInfo.outputPath('source.docx'),bytes);
    const result = await renderDocxArtifacts(bytes,{ browser,outputs:['pdf','html'],reviewProfile:'final',commentProfile:'hidden' });
    expect(result.renderReport.status).toBe('complete');expect(result.pageCount).toBeGreaterThan(1);
    expect((await inspectNotes(result.pdf!,percent)).lines.flat()).toEqual(labels);
    const html = testInfo.outputPath('note.html');writeFileSync(html,result.html!);await page.goto(pathToFileURL(html).href);
    const paragraphs = await page.locator('.page-footnotes p[data-source-anchor-id]').evaluateAll(ps => ps.filter(p => p.textContent?.trim().startsWith('N')).map(p => {
      const rows = new Map<string,string[]>();
      for (const text of Array.from(p.querySelectorAll('svg text'))) {
        const y = text.getAttribute('y')!;rows.set(y,[...(rows.get(y) ?? []),...text.textContent!.trim().split(/\s+/)]);
      }
      return { width:p.getBoundingClientRect().width * .75,anchor:p.getAttribute('data-source-anchor-id'),lines:Array.from(rows.values()) };
    }));
    expect(paragraphs[0].width).toBeCloseTo(144,2);
    for (const p of paragraphs.slice(1)) expect(p.width).toBeCloseTo(216,2);
    expect(new Set(paragraphs.map(p => p.anchor)).size).toBe(1);
    const perLine = percent === 80 ? [4,6] : [2,3];
    for (const [i,p] of paragraphs.entries()) for (const line of p.lines.slice(0,-1)) expect(line).toHaveLength(perLine[i ? 1 : 0]);
  });
}

for (const [indent,first] of [['firstLine',3],['hanging',5]] as const)
  test(`a scaled ${indent} note applies its first-line indent once across page continuations`, async ({ browser }, testInfo) => {
    const { parts,labels } = source(80,216);
    const ind = indent === 'firstLine' ? '<w:ind w:right="6480" w:firstLine="480"/>'
      : '<w:ind w:left="480" w:right="6000" w:hanging="480"/>';
    parts['word/footnotes.xml'] = parts['word/footnotes.xml'].replace('<w:ind w:right="6480"/>',ind)
      .replace('w:line="240"','w:line="360"');
    const bytes = pack(parts);writeFileSync(testInfo.outputPath('source.docx'),bytes);
    const result = await renderDocxArtifacts(bytes,{ browser,outputs:['pdf'],reviewProfile:'final',commentProfile:'hidden' });
    expect(result.pageCount).toBeGreaterThan(1);
    const { lines,pitches } = await inspectNotes(result.pdf!,80);
    expect(pitches.length).toBeGreaterThan(1);
    for (const pitch of pitches) expect(Math.abs(pitch - 2320 / 2048 * 12 * 1.5)).toBeLessThan(.02);
    expect(lines).toEqual([labels.slice(0,first),...Array.from({ length:Math.ceil((labels.length - first) / 4) },
      (_,i) => labels.slice(first + i * 4,first + (i + 1) * 4))]);
  });

test('packed notes reserve their fitted height and keep both references once', async ({ browser }, testInfo) => {
  const { parts,labels } = source(150,108);
  const second = parts['word/footnotes.xml'].match(/<w:footnote w:id="1">[^]*?<\/w:footnote>/)![0]
    .replace('w:id="1"','w:id="2"').replace(/N(\d{4})/g,(_,n) => `N${String(Number(n) + 108).padStart(4,'0')}`);
  parts['word/footnotes.xml'] = parts['word/footnotes.xml'].replace('</w:footnotes>',second + '</w:footnotes>');
  parts['word/document.xml'] = parts['word/document.xml'].replace('</w:p>',
    '<w:r><w:footnoteReference w:id="2"/></w:r></w:p>');
  const bytes = pack(parts);writeFileSync(testInfo.outputPath('source.docx'),bytes);
  const result = await renderDocxArtifacts(bytes,{ browser,outputs:['pdf'],reviewProfile:'final',commentProfile:'hidden' });
  expect(result.renderReport.status).toBe('complete');expect(result.pageCount).toBeGreaterThan(1);
  const { lines,tokens } = await inspectNotes(result.pdf!,150);
  expect(lines.flat()).toEqual([...labels,...labels.map((_,i) => `N${String(i + 108).padStart(4,'0')}`)]);
  expect(tokens.filter(t => !t.startsWith('N'))).toEqual(['Citation','1','2','1','2']);
});

for (const percent of [100,150]) test(`a ${percent}% note entering a narrower page preserves scale or refuses the export`, async ({ browser }, testInfo) => {
  const { parts } = source(percent,60);
  parts['word/footnotes.xml'] = parts['word/footnotes.xml'].replace('<w:ind w:right="6480"/>','')
    .replace(/N\d{4}/g,'ABCDEFGHIJKLMNOP');
  parts['word/document.xml'] = parts['word/document.xml'].replace('<w:body><w:p>',`<w:body><w:p><w:pPr>${section(216,'nextPage')}</w:pPr>`)
    .replace(/<w:sectPr><w:pgSz[^]*?<\/w:sectPr>/,'<w:p><w:r><w:t>After</w:t></w:r></w:p>' + section(144,'continuous'));
  const bytes = pack(parts);writeFileSync(testInfo.outputPath('source.docx'),bytes);
  const result = renderDocxArtifacts(bytes,{ browser,outputs:['pdf'],reviewProfile:'final',commentProfile:'hidden' });
  if (percent === 150) await expect(result).rejects.toMatchObject({ code:'pagination_failure',phase:'pagination' });
  else {
    const rendered = await result;
    expect(rendered.renderReport.status).toBe('complete');expect(rendered.pageCount).toBeGreaterThan(1);
    const widths = rendered.pageMap.fragments.filter(f => f.story === 'footnote' && f.anchorId.startsWith('p:fn:')).map(f => Math.round(f.geometry.width));
    expect(widths).toContain(216);expect(widths).toContain(144);
    const task = getDocument({ data:new Uint8Array(rendered.pdf!),verbosity:0 });
    try {
      const pdf = await task.promise, tokens:string[] = [];
      for (let i = 1; i <= pdf.numPages; i++) for (const item of (await (await pdf.getPage(i)).getTextContent()).items)
        if ('str' in item && item.str.trim()) tokens.push(...item.str.trim().split(/\s+/));
      expect(tokens.filter(t => t === 'ABCDEFGHIJKLMNOP')).toHaveLength(60);
      expect(tokens.filter(t => t !== 'ABCDEFGHIJKLMNOP')).toEqual(['Citation','1','1','After']);
    } finally { await task.destroy(); }
  }
});
