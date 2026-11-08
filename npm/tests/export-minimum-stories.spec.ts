import { expect, test } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox:true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath:process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

type Owner = 'body' | 'header' | 'footer' | 'footnote' | 'endnote';
function source(owner:Owner, natural:boolean, count:number, repeated:boolean, narrow = false) {
  const running = owner === 'header' || owner === 'footer';
  const part = running ? owner : owner + 's';
  const storyParts = owner === 'body' ? [] : [part];
  const labels = Array.from({ length:count }, (_, i) => `${narrow ? 'Violet' : 'Verbena'}${String(i).padStart(3, '0')}`);
  const paragraph = '<w:p><w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" '
    + (natural ? 'w:line="240" w:lineRule="auto"' : 'w:line="320" w:lineRule="atLeast"')
    + '/>' + (narrow ? '<w:ind w:right="7440"/>' : '')
    + `</w:pPr><w:r><w:t>${labels.join(' ')}</w:t></w:r></w:p>`;
  const story = running ? `<w:${owner === 'header' ? 'hdr' : 'ftr'} xmlns:w="${W_NS}">${paragraph}</w:${owner === 'header' ? 'hdr' : 'ftr'}>`
    : `<w:${part} xmlns:w="${W_NS}"><w:${owner} w:id="1"><w:p><w:r><w:${owner}Ref/></w:r></w:p>${paragraph}</w:${owner}></w:${part}>`;
  const body = owner === 'body' ? paragraph : Array.from({ length:running && repeated ? 3 : 1 }, (_, i) => '<w:p>'
    + (i ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '') + `<w:r><w:t>Body${i}</w:t></w:r>`
    + (!running ? `<w:r><w:${owner}Reference w:id="1"/></w:r>` : '') + '</w:p>').join('');
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
  const rels = (content:string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts:Record<string,string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + [['document', 'document.main'], ['styles', 'styles'], ...storyParts.map(name => [name, name])].map(([name, suffix]) =>
        `<Override PartName="/word/${name}.xml" ContentType="${type}${suffix}+xml"/>`).join('') + '</Types>',
    '_rels/.rels': rels(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': rels(['styles', ...storyParts].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')),
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    ...(owner === 'body' ? {} : { [`word/${part}.xml`]:story }),
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${body}<w:sectPr>`
      + (running ? `<w:${owner}Reference w:type="default" r:id="${part}"/>` : '')
      + '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>',
  };
  return { labels, bytes:storedZip(Object.entries(parts).map(([name, value]) => ({ name, data:xml(value) }))) };
}

async function facts(bytes:Uint8Array) {
  const task = getDocument({ data:new Uint8Array(bytes), verbosity:0 });
  try {
    const pdf = await task.promise, pages = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const items = (await (await pdf.getPage(n)).getTextContent()).items;
      const rows = new Map<number,string[]>();
      for (const item of items) if ('str' in item && item.str.includes('Verbena')) {
        const y = item.transform[5];
        rows.set(y, [...(rows.get(y) ?? []), ...item.str.trim().split(/\s+/)]);
      }
      pages.push(Array.from(rows, ([y, labels]) => ({ y, labels })));
    }
    return pages;
  } finally { await task.destroy(); }
}

for (const owner of ['header', 'footer', 'footnote', 'endnote'] as const) for (const continued of [false, true])
  test(`${owner} natural minimum preserves ${continued ? 'repeated or continued' : 'short'} story text`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const running = owner === 'header' || owner === 'footer';
    const count = continued ? running ? 54 : 240 : 18;
    const controlSource = source(owner, true, count, continued), actualSource = source(owner, false, count, continued);
    writeFileSync(testInfo.outputPath('natural.docx'),controlSource.bytes);
    writeFileSync(testInfo.outputPath('minimum.docx'),actualSource.bytes);
    const options = { browser, outputs:['pdf', 'html'] as const, reviewProfile:'final' as const, commentProfile:'hidden' as const };
    const control = await renderDocxArtifacts(controlSource.bytes, options);
    const actual = await renderDocxArtifacts(actualSource.bytes, options);
    const expected = await facts(control.pdf!), observed = await facts(actual.pdf!);
    expect(actual.pageCount).toBe(control.pageCount);
    if (continued) expect(actual.pageCount).toBeGreaterThan(1);
    expect(observed.map(page => page.map(row => row.labels))).toEqual(expected.map(page => page.map(row => row.labels)));
    for (const [index, page] of observed.entries()) for (const [line, row] of page.entries())
      expect(Math.abs(row.y - expected[index][line].y)).toBeLessThan(.5);
    if (owner === 'endnote') {
      await page.setContent(actual.html!);
      const ids = await page.locator('.page-box p').evaluateAll(paragraphs => paragraphs
        .filter(p => p.textContent?.includes('Verbena')).map(p => ({ paragraph:p.getAttribute('data-source-anchor-id'),
          note:p.querySelector('[data-source-anchor-id^="en:en:"]')?.getAttribute('data-source-anchor-id') })));
      expect(ids).toHaveLength(observed.filter(rows => rows.length).length);
      expect(new Set(ids.map(p => p.paragraph)).size).toBe(1);
      expect(new Set(ids.map(p => p.note)).size).toBe(1);
      expect(ids[0].paragraph).toMatch(/^p:en:/); expect(ids[0].note).toMatch(/^en:en:/);
    }
    expect(observed.flatMap(page => page.flatMap(row => row.labels)))
      .toEqual(Array.from({ length:running ? control.pageCount : 1 }, () => actualSource.labels).flat());
  });

for (const owner of ['body', 'header', 'footer'] as const)
  test(`${owner} below-natural minimum preserves fractional glyph geometry`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const geometry = async (bytes:Uint8Array) => {
      const task = getDocument({ data:new Uint8Array(bytes), verbosity:0 });
      try {
        const pdf = await task.promise, words = [];
        for (let page = 1; page <= pdf.numPages; page++)
          for (const item of (await (await pdf.getPage(page)).getTextContent()).items)
            if ('str' in item && item.str.trim().startsWith('Violet'))
              words.push({ page, label:item.str.trim(), rect:[item.transform[4], item.transform[5], item.width, item.height] });
        return words;
      } finally { await task.destroy(); }
    };
    const options = { browser, outputs:['pdf'] as const, reviewProfile:'final' as const, commentProfile:'hidden' as const };
    const count = owner === 'body' ? 64 : 18;
    const naturalSource = source(owner, true, count, true, true), minimumSource = source(owner, false, count, true, true);
    writeFileSync(testInfo.outputPath('natural.docx'), naturalSource.bytes);
    writeFileSync(testInfo.outputPath('minimum.docx'), minimumSource.bytes);
    const natural = await renderDocxArtifacts(naturalSource.bytes, options);
    const minimum = await renderDocxArtifacts(minimumSource.bytes, options);
    const expected = await geometry(natural.pdf!), actual = await geometry(minimum.pdf!);
    expect(natural.pageCount).toBe(3); expect(minimum.pageCount).toBe(3);
    expect(actual.map(({ page, label }) => ({ page, label }))).toEqual(expected.map(({ page, label }) => ({ page, label })));
    expect(actual).toHaveLength(owner === 'body' ? 64 : 54);
    const fragments = (result:typeof minimum) => result.pageMap.fragments
      .map(({ pageNumber, geometry, story }) => ({ pageNumber, geometry, story }));
    expect(fragments(minimum)).toEqual(fragments(natural));
    for (const [index, word] of actual.entries()) for (const [coordinate, value] of word.rect.entries())
      expect(Math.abs(value - expected[index].rect[coordinate])).toBeLessThan(.02);
  });
