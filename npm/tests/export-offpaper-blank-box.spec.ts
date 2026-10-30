import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const placement of ['left', 'above', 'visible', 'lost-text'] as const) {
  test(`source-positioned ${placement} text box retains the visible-content contract`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const lost = placement === 'lost-text';
    let box = textBoxRun('page', [lost ? 'Finch box label.' : '']);
    box = box.replace('<wp:wrapSquare wrapText="bothSides"/>', '<wp:wrapNone/>');
    box = box.replace('<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset>',
      `<wp:positionH relativeFrom="page"><wp:posOffset>${placement === 'left' || lost ? -2540000 : 914400}</wp:posOffset>`);
    box = box.replace('<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset>',
      `<wp:positionV relativeFrom="page"><wp:posOffset>${placement === 'above' ? -1143000 : 914400}</wp:posOffset>`);
    const pending = renderDocxArtifacts(outOfFlowDocx(`<w:p>${box}<w:r><w:t>Finch watch complete.</w:t></w:r></w:p>`),
      { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
        commentProfile: 'margin', commentColumnPlacement: 'balloons' });
    if (lost) { await expect(pending).rejects.toThrow(/has no measurable fragment/); return; }
    const result = await pending;
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    const id = await page.locator('[data-docx-textbox-clip="true"] [data-source-anchor-id]').getAttribute('data-source-anchor-id');
    expect(result.pageMap.fragments.some(f => f.anchorId === id)).toBe(placement === 'visible');
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text.split('Finch watch complete.')).toHaveLength(2);
    } finally { await task.destroy(); }
  });
}
