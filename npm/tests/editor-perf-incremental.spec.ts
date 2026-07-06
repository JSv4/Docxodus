import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Editor performance work: multi-block format ops reconcile INCREMENTALLY (N single-block
// swaps) instead of a full remount (whole-document convert — seconds on a real doc), and
// remount itself renders from the live session (RenderHtml) instead of marshaling saved
// bytes out to JS and back. These tests pin the fidelity contract of both changes.
test.describe('DocxEditor — incremental multi-block formatting (perf)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  /** Build a 4-paragraph document (AAA/BBB/CCC/DDD) in a fresh editor via real typing. */
  async function buildFourParagraphs(page: Page) {
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'perf-inc';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__perf = { editor, container };
      const first = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      first.focus();
      const r = document.createRange();
      r.selectNodeContents(first);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.keyboard.type('AAA');
    await page.keyboard.press('Enter');
    await page.keyboard.type('BBB');
    await page.keyboard.press('Enter');
    await page.keyboard.type('CCC');
    await page.keyboard.press('Enter');
    await page.keyboard.type('DDD');
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());
  }

  /** Select from the first text node of blocks[0] through the last text node of blocks[2]. */
  const selectFirstThree = (page: Page) =>
    page.evaluate(() => {
      const { container } = (window as any).__perf;
      const blocks = Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[];
      const firstText = (el: HTMLElement) =>
        document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
      const lastText = (el: HTMLElement) => {
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let t: Node | null, last: Node | null = null;
        while ((t = w.nextNode())) last = t;
        return last;
      };
      const fn = firstText(blocks[0]) || blocks[0];
      const ln = lastText(blocks[2]) || blocks[2];
      const r = document.createRange();
      r.setStart(fn, 0);
      r.setEnd(ln, ln.nodeType === 3 ? (ln.textContent || '').length : ln.childNodes.length);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });

  test('multi-block bold swaps only the selected blocks (untouched block keeps its DOM node)', async ({ page }) => {
    await buildFourParagraphs(page);

    // Tag every block's DOM node BEFORE the op. A full remount rebuilds the whole
    // container, destroying these tags — node identity is the structural proof that the
    // op reconciled incrementally.
    await page.evaluate(() => {
      const { container } = (window as any).__perf;
      (Array.from(container.querySelectorAll('p[data-anchor]')) as any[]).forEach((p, i) => {
        p.__preOpTag = `tag${i}`;
      });
    });

    await selectFirstThree(page);
    await page.evaluate(() => (window as any).__perf.editor.format('bold'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__perf;
      const blocks = (Array.from(container.querySelectorAll('p[data-anchor]')) as any[]).filter(
        (p) => /AAA|BBB|CCC|DDD/.test(p.textContent || ''),
      );
      const state = blocks.map((p) => {
        const span = p.querySelector('span');
        return {
          text: (p.textContent || '').trim(),
          bold: span ? parseInt(getComputedStyle(span).fontWeight, 10) >= 600 : false,
          keptNode: p.__preOpTag !== undefined,
        };
      });

      // The cross-block selection must survive the swaps (consecutive ribbon actions).
      const sel = window.getSelection();
      const selectionText = sel && sel.rangeCount > 0 ? sel.getRangeAt(0).toString() : '';
      const selectionCollapsed = !sel || sel.rangeCount === 0 || sel.isCollapsed;

      // Round-trip fidelity.
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = (window as any).Docxodus.DocxEditor.open(c2, saved, (window as any).Docxodus, {});
      const reopened = (Array.from(c2.querySelectorAll('p[data-anchor]')) as HTMLElement[])
        .filter((p) => /AAA|BBB|CCC|DDD/.test(p.textContent || ''))
        .map((p) => {
          const span = p.querySelector('span');
          return {
            text: (p.textContent || '').trim(),
            bold: span ? parseInt(getComputedStyle(span).fontWeight, 10) >= 600 : false,
          };
        });
      editor.close(); e2.close(); container.remove(); c2.remove();
      return { state, selectionText, selectionCollapsed, reopened };
    });

    expect(out.state.length).toBe(4);
    const [a, b, c, d] = out.state;
    // Formatting applied to the three selected blocks; the fourth untouched.
    expect(a.bold && b.bold && c.bold).toBe(true);
    expect(d.bold).toBe(false);
    // The selected blocks were SWAPPED (fresh nodes); the untouched block kept its node —
    // i.e. no whole-document remount happened.
    expect(a.keptNode || b.keptNode || c.keptNode).toBe(false);
    expect(d.keptNode).toBe(true);
    // Selection restored across the swapped range.
    expect(out.selectionCollapsed).toBe(false);
    expect(out.selectionText).toContain('AAA');
    expect(out.selectionText).toContain('CCC');
    // Survives save → reopen.
    expect(out.reopened.filter((r: any) => r.bold).map((r: any) => r.text).sort()).toEqual(['AAA', 'BBB', 'CCC']);
    expect(out.reopened.find((r: any) => r.text === 'DDD')?.bold).toBe(false);
  });

  test('multi-block alignment swaps only the selected blocks and persists', async ({ page }) => {
    await buildFourParagraphs(page);
    await page.evaluate(() => {
      const { container } = (window as any).__perf;
      (Array.from(container.querySelectorAll('p[data-anchor]')) as any[]).forEach((p) => {
        p.__preOpTag = true;
      });
    });
    await selectFirstThree(page);
    await page.evaluate(() => (window as any).__perf.editor.setAlignment('center'));

    const out = await page.evaluate(() => {
      const { editor, container } = (window as any).__perf;
      const read = (root: HTMLElement) =>
        (Array.from(root.querySelectorAll('p[data-anchor]')) as any[])
          .filter((p) => /AAA|BBB|CCC|DDD/.test(p.textContent || ''))
          .map((p) => ({
            text: (p.textContent || '').trim(),
            align: getComputedStyle(p).textAlign,
            keptNode: p.__preOpTag !== undefined,
          }));
      const live = read(container);
      const saved: Uint8Array = editor.save();
      const c2 = document.createElement('div');
      document.body.appendChild(c2);
      const e2 = (window as any).Docxodus.DocxEditor.open(c2, saved, (window as any).Docxodus, {});
      const reopened = read(c2);
      editor.close(); e2.close(); container.remove(); c2.remove();
      return { live, reopened };
    });

    expect(out.live.length).toBe(4);
    expect(out.live.slice(0, 3).every((b: any) => b.align === 'center' && !b.keptNode)).toBe(true);
    expect(out.live[3].align).not.toBe('center');
    expect(out.live[3].keptNode).toBe(true); // untouched block survived in place — no remount
    expect(out.reopened.slice(0, 3).every((b: any) => b.align === 'center')).toBe(true);
  });

  test('session-attached RenderHtml matches the Save + ConvertDocxToHtmlComplete remount path', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const out = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const bytes = (window as any).testDocxBytes as Uint8Array;

      // Mirror the editor: persistAnchorIds session + a projection (refreshAnchorMap).
      const handle = D.DocxSessionBridge.OpenSession(bytes, '{"persistAnchorIds":true}');
      D.DocxSessionBridge.Project(handle);

      // Old remount path: marshal saved bytes to JS, convert bytes → HTML.
      let start = performance.now();
      const saved: Uint8Array = D.DocxSessionBridge.Save(handle);
      const viaBytes = D.DocumentConverter.ConvertDocxToHtmlComplete(
        saved, 'Document', 'docx-', false, '', -1, 'comment-',
        0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true, false, null, true,
      );
      const bytesMs = performance.now() - start;

      // New remount path: render straight from the live session.
      start = performance.now();
      const viaSession = D.DocxSessionBridge.RenderHtml(handle, 'docx-', false, false, 1);
      const sessionMs = performance.now() - start;

      D.DocxSessionBridge.CloseSession(handle);
      return {
        equal: viaBytes === viaSession,
        startsWithHtml: viaSession.charCodeAt(0) === 0x3c,
        lenBytes: viaBytes.length,
        lenSession: viaSession.length,
        bytesMs,
        sessionMs,
      };
    });

    console.log(
      `remount render: bytes-path ${out.bytesMs.toFixed(1)}ms vs session-path ${out.sessionMs.toFixed(1)}ms ` +
      `(html ${out.lenSession} chars)`,
    );
    expect(out.startsWithHtml).toBe(true);
    // Fidelity oracle: the session-attached render must be byte-identical to the bytes path.
    expect(out.equal).toBe(true);
  });

  test('multi-block format on a real document is dramatically cheaper than a full convert', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const out = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const bytes = (window as any).testDocxBytes as Uint8Array;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bytes, D, {});

      const blocks = (Array.from(
        container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
      ) as HTMLElement[]).filter((p) => (p.textContent || '').trim().length > 5);
      const first = blocks[0];
      const third = blocks[2];
      first.focus();
      const firstText = document.createTreeWalker(first, NodeFilter.SHOW_TEXT).nextNode()!;
      let lastText: Node | null = null;
      const w = document.createTreeWalker(third, NodeFilter.SHOW_TEXT);
      let t: Node | null;
      while ((t = w.nextNode())) lastText = t;
      const r = document.createRange();
      r.setStart(firstText, 0);
      r.setEnd(lastText!, (lastText!.textContent || '').length);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);

      // Reference cost of the old multi-block path: one full-document convert.
      let start = performance.now();
      const saved: Uint8Array = editor.save();
      D.DocumentConverter.ConvertDocxToHtmlComplete(
        saved, 'Document', 'docx-', false, '', -1, 'comment-',
        0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true, false, null, true,
      );
      const fullConvertMs = performance.now() - start;

      // New path: three-block bold via per-block swaps.
      start = performance.now();
      editor.format('bold');
      const multiBlockMs = performance.now() - start;

      const boldCount = blocks.slice(0, 3).filter((p) => {
        const el = container.querySelector(`[data-anchor]`);
        void el;
        const span = (Array.from(container.querySelectorAll('p[data-anchor]')) as HTMLElement[])
          .find((q) => (q.textContent || '').trim() === (p.textContent || '').trim());
        const sp = span?.querySelector('span');
        return sp ? parseInt(getComputedStyle(sp).fontWeight, 10) >= 600 : false;
      }).length;

      editor.close();
      container.remove();
      return { fullConvertMs, multiBlockMs, boldCount };
    });

    console.log(
      `3-block bold: ${out.multiBlockMs.toFixed(1)}ms (old full-convert path reference: ` +
      `${out.fullConvertMs.toFixed(1)}ms — ${(out.fullConvertMs / out.multiBlockMs).toFixed(1)}x)`,
    );
    expect(out.boldCount).toBe(3);
    // Generous bound: the incremental path must beat one full-document convert.
    expect(out.multiBlockMs).toBeLessThan(out.fullConvertMs);
  });
});
