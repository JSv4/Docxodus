import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';
import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const words = 'Juniper records the ferry inventory before the winter crossing. '.repeat(12);
const run = (text: string) => `<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

function fixture(position: number, line: number, empty: boolean): Uint8Array {
  const offset = position === 0 ? 0 : position === 1 ? words.indexOf('before') : words.length;
  const content = run(words.slice(0, offset))
    + (empty ? '<w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t/></w:r>' : '')
    + run(words.slice(offset));
  return pageFittingDocx(`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${line}" w:lineRule="auto"/></w:pPr>
    ${content}</w:p><w:p>${run('Inventory complete.')}</w:p>`);
}

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise, pages = [];
    for (let page = 1; page <= pdf.numPages; page++) {
      const content = await (await pdf.getPage(page)).getTextContent();
      pages.push(content.items.flatMap(item => 'str' in item && item.str.trim()
        ? [{ text: item.str, transform: item.transform, width: item.width }] : []));
    }
    return pages;
  } finally { await task.destroy(); }
}

for (const line of [240, 480]) for (const position of [0, 1, 2]) {
  test(`empty run at position ${position}, line multiple ${line / 240}, preserves PDF text and geometry`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const,
      commentProfile: 'hidden' as const, timeoutMs: 120_000 };
    const control = await renderDocxArtifacts(fixture(position, line, false), options);
    const actual = await renderDocxArtifacts(fixture(position, line, true), options);
    const expected = await facts(control.pdf!), observed = await facts(actual.pdf!);
    expect(expected.flat().map(item => item.text).join(' ').replace(/\s+/g, ' ').trim())
      .toBe((words + 'Inventory complete.').trim());
    expect(observed).toEqual(expected);
    await page.setContent(actual.html!);
    const anchors = await page.locator('.page-box p[data-source-anchor-id]')
      .evaluateAll(elements => elements.map(element => element.getAttribute('data-source-anchor-id')));
    expect(anchors).toHaveLength(2);
    expect(new Set(anchors).size).toBe(2);
    for (const anchor of anchors)
      expect(actual.pageMap.fragments.filter(fragment => fragment.anchorId === anchor)).toHaveLength(1);
  });
}

test('empty runs inside a textbox preserve its PDF geometry and source fragments', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const docx = (empty: boolean) => outOfFlowDocx(`<w:p>${run('Body inventory.')}${textBoxRun('paragraph',
    ['Juniper first record.', 'Juniper second record.']).replace(/<w:p>/g, '<w:p>' + (empty
      ? '<w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t/></w:r>' : ''))}</w:p>`);
  const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const,
    commentProfile: 'hidden' as const, timeoutMs: 120_000 };
  const control = await renderDocxArtifacts(docx(false), options);
  const actual = await renderDocxArtifacts(docx(true), options);
  expect(await facts(actual.pdf!)).toEqual(await facts(control.pdf!));
  await page.setContent(actual.html!);
  const anchors = await page.locator('.page-box [data-source-anchor-id]').evaluateAll(elements => elements
    .filter(element => element.textContent?.startsWith('Juniper'))
    .map(element => element.getAttribute('data-source-anchor-id')));
  expect(new Set(anchors).size).toBe(2);
  for (const anchor of anchors)
    expect(actual.pageMap.fragments.filter(fragment => fragment.anchorId === anchor)).toHaveLength(1);
});
