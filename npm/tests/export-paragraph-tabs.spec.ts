import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { autoSpacingParagraphDocx } from './docx-auto-spacing-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

test('inline tab advances and dotted leaders survive paragraph continuations exactly once', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(autoSpacingParagraphDocx(false, true), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise, pages: string[] = [];
    for (let index = 1; index <= pdf.numPages; index++) {
      const content = await (await pdf.getPage(index)).getTextContent();
      pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
    }
    expect(pdf.numPages).toBeGreaterThan(1);
    expect(pdf.numPages).toBe(result.pageMap.pages.length);
    expect(pages.join(' ').match(/Juniper\d+/g)).toEqual(Array.from({ length: 83 }, (_, i) => `Juniper${String(i + 1).padStart(3, '0')}`));
    expect(pages.join(' ').match(/checkpoint recorded\./g)).toHaveLength(83);
    expect(pages.join(' ').match(/Checkpoint inventory complete\./g)).toHaveLength(1);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const tabs = await page.locator('.page-content [data-docx-tab]').evaluateAll(elements => elements.map(element => {
    const style = getComputedStyle(element), rect = element.getBoundingClientRect();
    return { width: rect.width, leader: style.borderBottomStyle,
      overflow: rect.bottom - element.closest('.page-content')!.getBoundingClientRect().bottom };
  }));
  expect(tabs).toHaveLength(1);
  expect(tabs.every(tab => tab.width > 0 && tab.leader === 'dotted' && tab.overflow <= 1)).toBe(true);
});
