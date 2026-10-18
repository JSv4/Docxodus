import { test, expect } from '@playwright/test';
import { commentDateFixture } from './docx-comment-date-fixture.js';

for (const modern of [false, true]) for (const mode of [0, 1, 2]) {
  test(`comment date follows durable UTC metadata (modern=${modern}, mode=${mode})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const result = await page.evaluate(({ bytes, mode }) => {
      const D = (window as any).Docxodus;
      const input = new Uint8Array(bytes);
      const html = D.DocumentConverter.ConvertDocxToHtmlComplete(input, 'Cedar date', 'docx-', false,
        '', mode, 'comment-', 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = parsed.body.innerHTML;
      const bridge = D.DocxSessionBridge;
      const handle = bridge.OpenSession(input, '');
      try {
        const anchor = JSON.parse(bridge.ListBlocks(handle)).body[0].id;
        const blockDates = [];
        for (let pass = 0; pass < 2; pass++) {
          const fragment = bridge.RenderEditorBlockHtml(handle, anchor, '{"comments":true}');
          if (!fragment.startsWith('<')) throw new Error(fragment);
          blockDates.push(...Array.from(new DOMParser().parseFromString(fragment, 'text/html')
            .querySelectorAll('[data-date]')).map(el => el.getAttribute('data-date')));
        }
        return { dates: Array.from(document.querySelectorAll('[data-date]')).map(el => el.getAttribute('data-date')),
          captions: Array.from(document.querySelectorAll('.comment-date, .comment-margin-date')).map(el => el.textContent),
          listed: JSON.parse(bridge.ListComments(handle))[0].date,
          stateless: JSON.parse(D.DocumentComparer.GetCommentsJson(input))[0].date, blockDates };
      } finally { bridge.CloseSession(handle); }
    }, { bytes: Array.from(commentDateFixture(modern)), mode });
    const expected = modern ? '2026-09-10T04:45:00Z' : '2026-09-09T23:45:00Z';
    expect(result.dates.length).toBeGreaterThan(0);
    expect(new Set(result.dates)).toEqual(new Set([expected]));
    expect(result.listed).toBe(expected);
    expect(result.stateless).toBe(expected);
    expect(result.blockDates.length).toBeGreaterThan(0);
    expect(new Set(result.blockDates)).toEqual(new Set([expected]));
    if (mode !== 1) expect(result.captions).toEqual([modern ? 'Sep 10, 2026' : 'Sep 9, 2026']);
  });
}
