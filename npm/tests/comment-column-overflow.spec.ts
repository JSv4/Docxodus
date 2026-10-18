import { test, expect } from '@playwright/test';
import { longCommentFixture, LONG_COMMENT_TOKEN } from './docx-long-comment-fixture.js';

for (const classes of [false, true]) for (const count of [8, 240]) for (const longWord of [false, true]) {
  test(`${count}-word comment is fully readable without changing the page (classes=${classes}, longWord=${longWord})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes, classes, last }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar notes', 'docx-', classes, '', 2, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, true, true, false, '', true);
      if (html.startsWith('{')) throw new Error(html);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      const pagination = (window as any).DocxodusPagination.paginateHtml(html, host, {
        showPageNumbers: false, layoutToken: { documentVersion: 1, rendererFingerprint: 'comment-column-test' },
      });
      const column = host.querySelector<HTMLElement>('.page-comment-margin')!;
      const pageBox = host.querySelector<HTMLElement>('.page-box')!;
      const walker = document.createTreeWalker(column, NodeFilter.SHOW_TEXT);
      let node: Node | null, tail: Text | null = null;
      while ((node = walker.nextNode())) if (node.textContent?.includes(last)) tail = node as Text;
      if (!tail) throw new Error(`missing final comment token ${last}`);
      const range = document.createRange(), start = tail.data.indexOf(last);
      range.setStart(tail, start); range.setEnd(tail, start + last.length);
      const box = column.getBoundingClientRect(), pageRect = pageBox.getBoundingClientRect();
      const visible = () => {
        const rect = range.getBoundingClientRect();
        return rect.top >= box.top + column.clientTop - 1 && rect.bottom <= box.top + column.clientTop + column.clientHeight + 1 &&
          rect.left >= box.left + column.clientLeft - 1 && rect.right <= box.left + column.clientLeft + column.clientWidth + 1;
      };
      const initiallyVisible = visible();
      column.scrollTop = column.scrollHeight;
      return { overflow: getComputedStyle(column).overflowY, initiallyVisible, finallyVisible: visible(),
        scrollable: column.scrollHeight > column.clientHeight,
        pageWidth: pageRect.width, pageHeight: pageRect.height, columnLeft: box.left, pageRight: pageRect.right,
        columnWidth: box.width, stories: pagination.pageMap.fragments.map((fragment: any) => fragment.story) };
    }, { bytes: Array.from(longCommentFixture(count, false, longWord)), classes,
      last: longWord ? LONG_COMMENT_TOKEN : `CedarNote${String(count - 1).padStart(3, '0')}` });
    expect(result.pageWidth).toBeCloseTo(816, 0);
    expect(result.pageHeight).toBeCloseTo(1056, 0);
    expect(result.stories).toContain('body');
    expect(result.stories).not.toContain('comment');
    if (count === 240) {
      expect(result.columnLeft).toBeGreaterThanOrEqual(result.pageRight);
      expect(result.columnWidth).toBeCloseTo(192, 0);
      expect(result.scrollable).toBe(true);
      expect(['auto', 'scroll']).toContain(result.overflow);
    } else expect(result.initiallyVisible).toBe(true);
    expect(result.finallyVisible).toBe(true);
  });
}

for (const scale of [0.75, 1, 1.5]) {
  test(`outside comment width and gap follow configuration at scale ${scale}`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const facts = await page.evaluate(({ bytes, scale }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar notes', 'docx-', false, '', 2, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      (window as any).DocxodusPagination.paginateHtml(html, host, {
        scale, showPageNumbers: false, commentColumnWidthPt: 180, commentColumnGapPt: 18,
      });
      const column = host.querySelector<HTMLElement>('.page-comment-margin')!;
      const box = host.querySelector<HTMLElement>('.page-box')!.getBoundingClientRect();
      const rect = column.getBoundingClientRect();
      return { width: rect.width, gap: rect.left - box.right, pageWidth: box.width, pageHeight: box.height,
        focusable: column.tabIndex === 0, label: column.getAttribute('aria-label') };
    }, { bytes: Array.from(longCommentFixture(8)), scale });
    expect(facts.width).toBeCloseTo(240 * scale, 0);
    expect(facts.gap).toBeCloseTo(24 * scale, 0);
    expect(facts.pageWidth).toBeCloseTo(816 * scale, 0);
    expect(facts.pageHeight).toBeCloseTo(1056 * scale, 0);
    expect(facts.focusable).toBe(true);
    expect(facts.label).toBe('Comments for page 1');
  });
}

test('physical margin placement retains comment PageMap fragments and scrolling', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  const facts = await page.evaluate((bytes) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Cedar notes', 'docx-', false, '', 2, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', false, false, false, true, true, false, '', true);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    const result = (window as any).DocxodusPagination.paginateHtml(html, host, {
      showPageNumbers: false, commentColumnPlacement: 'margin',
      layoutToken: { documentVersion: 1, rendererFingerprint: 'margin-column-test' },
    });
    const column = host.querySelector<HTMLElement>('.page-comment-margin')!;
    return { insidePage: Boolean(column.closest('.page-box')), overflow: getComputedStyle(column).overflowY,
      stories: result.pageMap.fragments.map((fragment: any) => fragment.story) };
  }, Array.from(longCommentFixture(8)));
  expect(facts.insidePage).toBe(true);
  expect(facts.stories).toContain('comment');
  expect(facts.stories).toContain('body');
  expect(['auto', 'scroll']).toContain(facts.overflow);
});

test('invalid comment-column dimensions and placement fail before layout', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  const results = await page.evaluate(() => {
    const staging = document.createElement('div'), host = document.createElement('div');
    document.body.append(staging, host);
    return [{ commentColumnWidthPt: 0 }, { commentColumnWidthPt: -1 },
      { commentColumnWidthPt: NaN }, { commentColumnWidthPt: Infinity },
      { commentColumnGapPt: -1 }, { commentColumnGapPt: Infinity },
      { commentColumnPlacement: 'sideways' }].map(options => {
        try { new (window as any).DocxodusPagination.PaginationEngine(staging, host, options); return false; }
        catch (error) { return error instanceof RangeError; }
      });
  });
  expect(results).toEqual(Array(7).fill(true));
});

for (const hostWidth of [700, 1200]) for (const secondPage of [false, true]) {
  test(`comment panels preserve page alignment at width ${hostWidth} (second page=${secondPage})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const facts = await page.evaluate(({ bytes, hostWidth }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar notes', 'docx-', false, '', 2, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = `<div id="host" style="width:${hostWidth}px;overflow:auto"></div>`;
      const host = document.querySelector<HTMLElement>('#host')!;
      const result = (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      const left = host.getBoundingClientRect().left;
      return { totalPages: result.totalPages,
        pageLeft: Array.from(host.querySelectorAll<HTMLElement>('.page-box')).map(p => p.getBoundingClientRect().left - left),
        panels: host.querySelectorAll('[data-docx-comment-panel]').length,
        comments: host.querySelectorAll('.page-comment-margin').length };
    }, { bytes: Array.from(longCommentFixture(8, secondPage)), hostWidth });
    expect(facts.totalPages).toBe(secondPage ? 2 : 1);
    expect(facts.pageLeft.every((left: number) => left >= 0)).toBe(true);
    expect(Math.max(...facts.pageLeft) - Math.min(...facts.pageLeft)).toBeLessThan(1);
    expect(facts.panels).toBe(facts.totalPages);
    expect(facts.comments).toBe(1);
  });
}
