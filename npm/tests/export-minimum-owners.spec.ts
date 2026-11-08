import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { minimumOwnerDocx, minimumOwnerText } from './docx-minimum-owner-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise, pages = [];
    for (let number = 1; number <= pdf.numPages; number++) {
      const text = await (await pdf.getPage(number)).getTextContent();
      pages.push(text.items.flatMap(item => 'str' in item && item.str.trim()
        ? [{ text: item.str, transform: item.transform, width: item.width }] : []));
    }
    return pages;
  } finally { await task.destroy(); }
}

for (const [owner, suffix] of [['cell', 'tab'], ['list', 'tab'], ['list', 'nothing']] as const) for (const repeats of [4, 48])
  test(`${owner} (${suffix}) natural minimum preserves text geometry and continuation (${repeats} repetitions)`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const,
      commentProfile: 'hidden' as const, timeoutMs: 120_000 };
    const control = await renderDocxArtifacts(minimumOwnerDocx(owner, true, repeats, suffix), options);
    const actual = await renderDocxArtifacts(minimumOwnerDocx(owner, false, repeats, suffix), options);
    const expected = await facts(control.pdf!), observed = await facts(actual.pdf!);
    expect(expected.flat().map(i => i.text).join(' ').replace(/\s+/g, ' ').trim())
      .toBe((owner === 'list' ? '1. ' : '') + minimumOwnerText.repeat(repeats).trim());
    expect(observed).toEqual(expected);
    expect(actual.pageCount).toBe(repeats === 4 ? 1 : control.pageCount);
    if (repeats === 48) expect(actual.pageCount).toBeGreaterThan(1);
    await page.setContent(actual.html!);
    const anchors = await page.locator('.page-box p[data-source-anchor-id]').evaluateAll(elements =>
      elements.filter(p => p.textContent?.includes('Teasel')).map(p => p.getAttribute('data-source-anchor-id')));
    expect(new Set(anchors).size).toBe(1);
    expect(anchors.length).toBe(actual.pageCount);
    expect(actual.pageMap.fragments.filter(f => f.anchorId === anchors[0])).toHaveLength(actual.pageCount);
    if (owner === 'list') {
      const markers = await page.locator('.page-box [data-list-marker]').evaluateAll(elements => elements
        .filter(e => !e.parentElement?.closest('[data-list-marker]')).map(e => e.textContent?.trim()));
      expect(markers).toEqual(['1.']);
    }
  });
