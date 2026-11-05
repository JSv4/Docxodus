import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { continuousMarginsDocx, type MarginCase } from './docx-continuous-margins-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true } });

for (const change of ['same', 'left', 'right', 'top', 'bottom', 'next', 'paper'] as MarginCase[]) {
  test(`${change} section geometry reaches standalone HTML and PDF`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(continuousMarginsDocx(change), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const breaks = change === 'next' || change === 'paper';
    expect(result.pageMap.pages).toHaveLength(breaks ? 2 : 1);
    await page.setContent(result.html!);
    const boxes = await page.locator('.page-box').evaluateAll(pages => pages.map(page => {
      const paper = page.getBoundingClientRect();
      return Array.from(page.querySelectorAll<HTMLElement>('.page-content p')).map(p => ({
        text: p.textContent, x: (p.getBoundingClientRect().left - paper.left) * 0.75,
        width: p.getBoundingClientRect().width * 0.75,
      }));
    }));
    expect(boxes[0][0].text).toBe('SequoiaBefore');
    const after = boxes.flat().find(p => p.text === 'SequoiaAfter')!;
    expect(after.x).toBeCloseTo(change === 'left' ? 144 : 72, 1);
    expect(after.width).toBeCloseTo(change === 'left' || change === 'right' ? 396 : change === 'paper' ? 576 : 468, 1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(breaks ? 2 : 1);
      const p = await pdf.getPage(breaks ? 2 : 1);
      const after = (await p.getTextContent()).items.find(i => 'str' in i && i.str === 'SequoiaAfter');
      expect(after && 'transform' in after ? after.transform[4] : NaN).toBeCloseTo(change === 'left' ? 144 : 72, 0);
    } finally { await task.destroy(); }
  });
}

for (const change of ['left', 'right', 'top', 'bottom'] as const) {
  test(`${change} margin survives paragraph fragmentation and page ownership`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(continuousMarginsDocx(change, true), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageMap.pages.length).toBeGreaterThan(1);
    expect(result.pageMap.pages.map(p => p.sectionIndex)).toEqual([0, ...result.pageMap.pages.slice(1).map(() => 1)]);
    await page.setContent(result.html!);
    expect(await page.locator('.page-content').first().textContent()).toContain('SequoiaAfter');
    const body = await page.locator('.page-content').evaluateAll(elements => elements.map(element => ({
      top: parseFloat((element as HTMLElement).style.top), height: parseFloat((element as HTMLElement).style.height),
    })));
    expect(body[0]).toEqual({ top: 72, height: 648 });
    expect(body[1]).toEqual({ top: change === 'top' ? 144 : 72, height: change === 'top' || change === 'bottom' ? 576 : 648 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const words: string[] = [];
      for (let n = 1; n <= pdf.numPages; n++) {
        const p = await pdf.getPage(n);
        words.push(...(await p.getTextContent()).items.flatMap(i => 'str' in i ? i.str.match(/Sequoia\d{3}/g) ?? [] : []));
      }
      expect(words).toEqual(Array.from({ length: 750 }, (_, i) => `Sequoia${String(i).padStart(3, '0')}`));
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
    } finally { await task.destroy(); }
  });
}
