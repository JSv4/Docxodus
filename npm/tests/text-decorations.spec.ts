import { test, expect } from '@playwright/test';
import { decorationFixture } from './docx-text-decoration-fixture.js';

for (const paginated of [false, true]) for (const classes of [false, true]) {
  test(`WASM decorations retain independent lines and patterns (${paginated ? 'paginated' : 'continuous'}, classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const observed = await page.evaluate(({ bytes, paginated, classes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Rules', 'docx-', classes, '', -1, 'comment-',
        paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const host = document.createElement('div'); document.body.replaceChildren(host);
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      else host.innerHTML = html;
      return Array.from(host.querySelectorAll(paginated ? '.page-content p' : 'p')).map(p => ({
        text: p.textContent,
        decorations: Array.from(p.querySelectorAll('span')).map(span => {
          const s = getComputedStyle(span);
          return { line: s.textDecorationLine, pattern: s.textDecorationStyle, color: s.textDecorationColor };
        }).filter(d => d.line !== 'none'),
      }));
    }, { bytes: Array.from(decorationFixture()), paginated, classes });
    const lines = (text: string) => observed.find(p => p.text === text)!.decorations;
    expect(lines('DoubleUnderline')).toContainEqual({ line: 'underline', pattern: 'double', color: 'rgb(0, 0, 0)' });
    expect(lines('DoubleStrike')).toContainEqual({ line: 'line-through', pattern: 'double', color: 'rgb(0, 0, 0)' });
    expect(lines('CombinedRules').flatMap(d => d.line.split(' ')).sort()).toEqual(['line-through', 'underline']);
    expect(lines('BlueDotted')).toContainEqual({ line: 'underline', pattern: 'dotted', color: 'rgb(0, 0, 255)' });
    expect(lines('SingleUnderline')).toContainEqual({ line: 'underline', pattern: 'solid', color: 'rgb(0, 0, 0)' });
    expect(lines('SingleStrike')).toContainEqual({ line: 'line-through', pattern: 'solid', color: 'rgb(0, 0, 0)' });
    expect(lines('IndependentRules')).toEqual(expect.arrayContaining([
      { line: 'underline', pattern: 'dotted', color: 'rgb(0, 0, 255)' },
      { line: 'line-through', pattern: 'double', color: 'rgb(0, 0, 0)' },
    ]));
    expect(lines('DisabledRules')).toEqual([]);
  });
}

test('editor reports both independent decorations at the caret', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const state = await page.evaluate(bytes => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div'); document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, {});
    const states = ['IndependentRules', 'SingleUnderline', 'DisabledRules'].map(label => {
      const block = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor]'))
        .find(p => p.textContent === label)!;
      block.focus();
      const text = document.createTreeWalker(block, NodeFilter.SHOW_TEXT).nextNode()!;
      const range = document.createRange(); range.setStart(text, 2); range.collapse(true);
      const selection = window.getSelection()!; selection.removeAllRanges(); selection.addRange(range);
      return editor.queryFormatState();
    });
    editor.close(); return states;
  }, Array.from(decorationFixture()));
  expect(state.map(s => [s.underline, s.strike])).toEqual([[true, true], [true, false], [false, false]]);
});
