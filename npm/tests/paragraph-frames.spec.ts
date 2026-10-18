import { test, expect } from '@playwright/test';
import { paragraphFrameFixture } from './docx-paragraph-frame-fixture.js';

for (const story of ['body', 'header', 'footer'] as const) for (const classes of [false, true]) {
  test(`${story} frame uses page offsets without adding a story line (classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes, classes, story }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar frame', 'docx-', classes, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, true, false, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      const band = host.querySelector<HTMLElement>(story === 'body' ? '.page-content' : `.page-${story}`)!;
      const pageBox = band.parentElement!;
      const paragraphs = Array.from(pageBox.querySelectorAll<HTMLElement>('p'));
      const frame = paragraphs.find(p => p.textContent === 'CEDAR RECORD 731')!.closest<HTMLElement>('[data-docx-text-frame]')!;
      const control = paragraphs.find(p => p.textContent === 'Cedar page control')!;
      const pageRect = pageBox.getBoundingClientRect(), footerRect = band.getBoundingClientRect();
      const frameRect = frame.getBoundingClientRect(), controlRect = control.getBoundingClientRect();
      return { position: getComputedStyle(frame).position,
        left: frameRect.left - pageRect.left, top: frameRect.top - footerRect.top,
        width: frameRect.width, footerHeight: story === 'body' ? controlRect.height : footerRect.height, controlHeight: controlRect.height,
        controlCenter: (controlRect.left + controlRect.right) / 2 - pageRect.left };
    }, { bytes: Array.from(paragraphFrameFixture(true, story)), classes, story });
    expect(result.position).toBe('absolute');
    expect(result.left).toBeCloseTo(56, 0);
    expect(result.top).toBeCloseTo(4 / 3, 0);
    expect(result.width).toBeCloseTo(672, 0);
    expect(result.footerHeight).toBeCloseTo(result.controlHeight, 0);
    expect(result.controlCenter).toBeCloseTo(408, 0);
  });
}
