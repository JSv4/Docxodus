import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { wideBodyTableDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

test('an AutoFit asymmetric table fits its text and images on every continuation', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(wideBodyTableDocx(false), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages: string[] = [];
    let images = 0;
    for (let index = 1; index <= pdf.numPages; index++) {
      const pdfPage = await pdf.getPage(index);
      const content = await pdfPage.getTextContent();
      pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      images += (await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject).length;
    }
    expect(pdf.numPages).toBeGreaterThan(1);
    expect(pdf.numPages).toBe(result.pageMap.pages.length);
    expect(pages.join(' ').match(/Ochre\d+[A-E]/g)).toEqual(Array.from({ length: 39 }, (_, row) =>
      Array.from({ length: 5 }, (_, column) => `Ochre${String(row + 1).padStart(2, '0')}${String.fromCharCode(65 + column)}`)).flat());
    expect(pages.join(' ').match(/Batch complete\./g)).toHaveLength(1);
    expect(images).toBe(5);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const geometry = await page.locator('.page-content').evaluateAll(bodies => bodies.flatMap(body =>
    Array.from(body.querySelectorAll('table'), table => ({
      rightOverflow: table.getBoundingClientRect().right - body.getBoundingClientRect().right,
      bottomOverflow: table.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom,
      fonts: Array.from(table.querySelectorAll('p'), p => getComputedStyle(p).fontSize),
    }))));
  expect(geometry.every(table => table.rightOverflow <= 1 && table.bottomOverflow <= 1)).toBe(true);
  expect(geometry.flatMap(table => table.fonts).every(font => parseFloat(font) >= 14)).toBe(true);
});

// A 780pt fixed grid places the fifth cell wholly beyond Letter paper and the fourth across its
// edge. Word 2024 prints the four cells the paper holds and none of the fifth (issue #1733); the
// export clips the fifth the same way and reports each of its paragraphs instead of refusing.
test('the paper clips fixed table cells that lie wholly beyond it', async ({ browser }) => {
  const options = { browser, outputs: ['pdf'] as 'pdf'[], reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
  const result = await renderDocxArtifacts(wideBodyTableDocx(), options);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    let text = '';
    for (let index = 1; index <= pdf.numPages; index++)
      text += (await (await pdf.getPage(index)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ') + ' ';
    expect(text.match(/Ochre\d+[A-E]/g)).toEqual(Array.from({ length: 39 }, (_, row) =>
      ['A', 'B', 'C', 'D'].map(column => `Ochre${String(row + 1).padStart(2, '0')}${column}`)).flat());
  } finally { await task.destroy(); }
  const clipped = result.renderReport.warnings.filter(w => w.code === 'source_text_clipped').map(w => w.anchorId!);
  expect(clipped).toHaveLength(39);
  expect(clipped.every(anchorId => anchorId.startsWith('p:'))).toBe(true);
  expect(result.pageMap.fragments.filter(fragment => clipped.includes(fragment.anchorId))).toEqual([]);
  await expect(renderDocxArtifacts(wideBodyTableDocx(), { ...options, unsupportedContent: 'strict' }))
    .rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});
