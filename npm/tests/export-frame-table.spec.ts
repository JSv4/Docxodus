import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { frameTableDocx } from './docx-frame-table-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [framed, cellText, nestedDrawing] of [[true, false, false], [true, true, false], [false, false, false], [true, false, true]]) {
  test(`table-owned labels remain addressable with frames=${framed}, cell text=${cellText}, nested drawing=${nestedDrawing}`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(frameTableDocx(framed, cellText, nestedDrawing), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin', commentColumnPlacement: 'balloons' });
    expect(result.pageCount).toBe(1);
    const table = result.pageMap.fragments.filter(f => f.anchorId.startsWith('tbl:'));
    const cells = result.pageMap.fragments.filter(f => f.anchorId.startsWith('tc:'));
    expect(table).toHaveLength(1); expect(cells).toHaveLength(3);
    await page.setContent(result.html!);
    const labels = ['Fern amber', 'Fern blue', 'Fern copper'];
    const ids = await page.locator('.page-box [data-source-anchor-id^="p:"]').evaluateAll((ps, labels) =>
      labels.map(label => (ps.find(p => p.textContent === label) as HTMLElement)?.dataset.sourceAnchorId), labels);
    for (let index = 0; index < labels.length; index++) {
      const label = result.pageMap.fragments.find(f => f.anchorId === ids[index])!;
      expect(label).toBeDefined(); expect(label.inTableCell).toBe(true);
      if (framed && !cellText) {
        for (const owner of [table[0], cells[index]]) {
          expect(owner.geometry.x).toBeLessThanOrEqual(label.geometry.x + 0.5);
          expect(owner.geometry.y).toBeLessThanOrEqual(label.geometry.y + 0.5);
          expect(owner.geometry.x + owner.geometry.width).toBeGreaterThanOrEqual(label.geometry.x + label.geometry.width - 0.5);
          expect(owner.geometry.y + owner.geometry.height).toBeGreaterThanOrEqual(label.geometry.y + label.geometry.height - 0.5);
        }
      }
    }
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (const label of [...labels, 'Fern inventory complete.', ...(cellText ? ['Slot 1 ready.', 'Slot 2 ready.', 'Slot 3 ready.'] : [])])
        expect(text.split(label)).toHaveLength(2);
    } finally { await task.destroy(); }
  });
}
