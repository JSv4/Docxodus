import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { outOfFlowDocx, textBoxRun } from './docx-out-of-flow-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const placement of ['left', 'above', 'visible', 'offpaper-text', 'strict-text'] as const) {
  test(`source-positioned ${placement} text box retains the visible-content contract`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const hasText = placement === 'offpaper-text' || placement === 'strict-text';
    let box = textBoxRun('page', [hasText ? 'Finch box label.' : '']);
    box = box.replace('<wp:wrapSquare wrapText="bothSides"/>', '<wp:wrapNone/>');
    box = box.replace('<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset>',
      `<wp:positionH relativeFrom="page"><wp:posOffset>${placement === 'left' || hasText ? -2540000 : 914400}</wp:posOffset>`);
    box = box.replace('<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset>',
      `<wp:positionV relativeFrom="page"><wp:posOffset>${placement === 'above' ? -1143000 : 914400}</wp:posOffset>`);
    const pending = renderDocxArtifacts(outOfFlowDocx(`<w:p>${box}<w:r><w:t>Finch watch complete.</w:t></w:r></w:p>`),
      { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
        commentProfile: 'margin', commentColumnPlacement: 'balloons',
        ...(placement === 'strict-text' ? { unsupportedContent: 'strict' as const } : {}) });
    if (placement === 'strict-text') {
      await expect(pending).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
      return;
    }
    const result = await pending;
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    const id = await page.locator('[data-docx-textbox-clip="true"] [data-source-anchor-id]').getAttribute('data-source-anchor-id');
    const warnings = result.renderReport.warnings.filter(warning => warning.code === 'source_text_clipped');
    expect(warnings).toHaveLength(hasText ? 1 : 0);
    if (hasText) {
      expect(warnings[0].anchorId).toBe(id);
      expect(result.html).toContain('Finch box label.');
    }
    expect(result.pageMap.fragments.some(f => f.anchorId === id)).toBe(placement === 'visible');
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const content = await (await (await task.promise).getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text.split('Finch watch complete.')).toHaveLength(2);
      expect(text).not.toContain('Finch box label.');
    } finally { await task.destroy(); }
  });
}
