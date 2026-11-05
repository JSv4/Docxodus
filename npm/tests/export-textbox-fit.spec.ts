import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { textboxFitDocx, textboxFitLabels, type TextBoxFit } from './docx-textbox-fit-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true } });

for (const inherited of [false, true]) for (const fit of ['implicit', 'fixed', 'shrink', 'auto'] as TextBoxFit[])
for (const placement of ['edge', 'onpage', 'outside'] as const) {
  test(`${inherited ? 'inherited' : 'local'} ${fit} fit preserves ${placement} source geometry`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120000);
    const bytes = textboxFitDocx(inherited ? 'implicit' : fit, placement, inherited ? fit : undefined);
    writeFileSync(testInfo.outputPath('birch-fit.docx'), bytes);
    const render = renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120000 });
    if (placement !== 'onpage' && (fit === 'shrink' || fit === 'auto')) {
      await expect(render).rejects.toThrow(/has no measurable fragment/);
      return;
    }
    const result = await render;
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    const anchors = await page.locator('[data-source-anchor-id]').evaluateAll(elements =>
      Object.fromEntries(elements.map(e => [(e.textContent ?? '').trim(), e.getAttribute('data-source-anchor-id')])));
    for (const label of [...textboxFitLabels, '¶']) {
      expect(anchors[label]).toBeTruthy();
      expect(result.pageMap.fragments.some(f => f.anchorId === anchors[label])).toBe(placement === 'onpage' || (placement === 'edge' && label !== '¶'));
    }
    const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await pdf.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join(' ');
      expect(text).toContain('Birch host complete');
      for (const label of textboxFitLabels) expect(text.includes(label)).toBe(placement !== 'outside');
    } finally { await pdf.destroy(); }
  });
}

for (const story of ['header', 'footer'] as const) test(`${story} inherits fixed fit from the document theme`, async ({ browser }) => {
  const result = await renderDocxArtifacts(textboxFitDocx('implicit', 'edge', 'fixed', story), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120000,
  });
  expect(result.pageCount).toBe(1);
  expect(result.renderReport.warnings.filter(w => w.code === 'source_text_clipped')).toHaveLength(1);
});
