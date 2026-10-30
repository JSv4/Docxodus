import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { linkedTextBoxDocx } from './docx-linked-textbox-fixture.js';
test.use({ launchOptions: { chromiumSandbox: true, ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });
for (const mode of ['flow', 'fitting', 'unlinked', 'image', 'cross-page', 'tab'] as const) {
  test(`linked textbox ${mode} preserves all source text and fragment geometry`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(linkedTextBoxDocx(mode), { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(mode === 'cross-page' ? 2 : 1);
      let text = '', images = 0;
      for (let number = 1; number <= pdf.numPages; number++) {
        const pdfPage = await pdf.getPage(number);
        text += (await pdfPage.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ') + ' ';
        images += (await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject).length;
      }
      expect(images).toBe(mode === 'image' ? 1 : 0);
      if (mode === 'tab') expect(text.match(/CedarClosing/g)).toHaveLength(1);
      expect(text.match(/Alder\d+/g)).toEqual(Array.from({ length: 70 }, (_, i) => `Alder${String(i).padStart(3, '0')}`));
    } finally { await task.destroy(); }
    expect(result.pageMap.pages).toHaveLength(mode === 'cross-page' ? 2 : 1);
    await page.setContent(result.html!);
    const boxes = page.locator('.page-box [data-docx-textbox]');
    await expect(boxes).toHaveCount(mode === 'unlinked' ? 1 : 2);
    const values = await boxes.evaluateAll(boxes => boxes.map(box => ({ text: box.textContent, width: box.getBoundingClientRect().width * .75,
      overflow: box.scrollHeight - box.clientHeight, ids: Array.from(box.querySelectorAll<HTMLElement>('[data-source-anchor-id]'), p => p.dataset.sourceAnchorId) })));
    expect(values[0].width).toBeCloseTo(144, 1);
    expect(values.every(box => box.overflow <= 1)).toBe(true);
    if (mode === 'flow' || mode === 'image' || mode === 'cross-page' || mode === 'tab') {
      expect(values[1].width).toBeCloseTo(216, 1);
      expect(values[0].text).toContain('Alder000'); expect(values[1].text).toContain('Alder069');
      expect(values[0].ids).toHaveLength(1); expect(values[1].ids).toEqual(values[0].ids);
      const fragments = result.pageMap.fragments.filter(fragment => fragment.anchorId === values[0].ids[0]);
      expect(fragments).toHaveLength(2);
      if (mode === 'cross-page') expect(new Set(fragments.map(fragment => fragment.pageNumber)).size).toBe(2);
    } else if (mode === 'fitting') expect(values[1].text?.trim()).toBe('');
  });
}
