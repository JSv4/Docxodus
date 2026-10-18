import { test, expect } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { paragraphFrameFixture } from './docx-paragraph-frame-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const story of ['body', 'header', 'footer'] as const) {
  test(`public ${story} PDF preserves frame offsets and source geometry`, async ({ browser }, info) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(paragraphFrameFixture(true, story), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    writeFileSync(info.outputPath(`paragraph-frame-${story}.pdf`), result.pdf!);
    writeFileSync(info.outputPath(`paragraph-frame-${story}.html`), result.html!);
    expect(result.pageCount).toBe(1);
    expect(result.html).toContain('data-docx-text-frame="true"');
    const fragment = result.pageMap.fragments.find(f => f.story === story && Math.abs(f.geometry.x - 42) < 0.5);
    expect(fragment, `${story} frame source survives outside clipping`).toBeDefined();
    expect(fragment!.geometry.width).toBeCloseTo(504, 0);
    expect(fragment!.geometry.height).toBeGreaterThan(0);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      writeFileSync(info.outputPath(`paragraph-frame-${story}-text.json`), JSON.stringify(items, null, 2));
      // PDF.js may expose one item per sentence or per word. Retain each item's actual box.
      const visible = items.filter(item => item.str.trim().length > 0);
      const phrase = (text: string) => {
        const matches: typeof visible[] = [];
        for (let start = 0; start < visible.length; start++) {
          let joined = '';
          for (let end = start; end < visible.length; end++) {
            joined += (joined ? ' ' : '') + visible[end].str.trim();
            if (joined === text) matches.push(visible.slice(start, end + 1));
            if (!text.startsWith(joined)) break;
          }
        }
        expect(matches, `complete, unique PDF phrase: ${text}`).toHaveLength(1);
        const group = matches[0];
        const y = group[0].transform[5];
        for (const item of group) expect(item.transform[5]).toBeCloseTo(y, 1);
        return { left: Math.min(...group.map(item => item.transform[4])),
          right: Math.max(...group.map(item => item.transform[4] + item.width)), y };
      };
      const frame = phrase('CEDAR RECORD 731'), control = phrase('Cedar page control');
      expect(frame.left).toBeCloseTo(42, 0);
      expect(frame.y).toBeCloseTo(control.y - 1, 0);
      expect((control.left + control.right) / 2).toBeCloseTo(306, 0);
    } finally { await task.destroy(); }
  });
}

test('a promoted body frame retains its table-cell source ownership', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(paragraphFrameFixture(true, 'body', false, true), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  const fragment = result.pageMap.fragments.find(f => f.story === 'body' && Math.abs(f.geometry.x - 42) < 0.5);
  expect(fragment).toBeDefined();
  expect(fragment!.inTableCell).toBe(true);
  expect(fragment!.geometry.width).toBeCloseTo(504, 0);
});
