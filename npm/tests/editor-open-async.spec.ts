import { test, expect, Page } from '@playwright/test';
import { pageBreakDocx } from './docx-page-break-fixture.js';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// 117 body units across four sections, bordered headings, tables, a footnote and an endnote.
const fixture = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC031-Complicated-Document.docx'),
));
// Eight paragraphs: a two-paragraph box drawn by a style's border, one drawn by direct borders,
// a differently bordered neighbour, plain paragraphs between them.
const boxes = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/WM001-Bordered-Boxes.docx'),
));

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Where two serializations part ways, with context on each side, for a failure message. */
const DIFF_AT = `
  (a, b) => {
    const at = [...a].findIndex((ch, i) => ch !== b[i]);
    if (at < 0 && a.length === b.length) return null;
    const from = Math.max(0, (at < 0 ? Math.min(a.length, b.length) : at) - 80);
    return { at, a: a.slice(from, from + 200), b: b.slice(from, from + 200) };
  }`;

/** Mount both ways into fresh containers and hand back what the comparison needs. */
const MOUNT_BOTH = `
  async ({ bytes, options, windowSize }) => {
    const D = window.Docxodus;
    const exports = D.getWasmExports ? D.getWasmExports() : D;
    const make = () => { const c = document.createElement('div'); document.body.appendChild(c); return c; };
    const a = make();
    const sync = D.DocxEditor.open(a, new Uint8Array(bytes), exports, options);
    const syncHtml = a.innerHTML;
    sync.close();
    const b = make();
    const progress = [];
    let ticks = 0;
    const ticker = setInterval(() => ticks++, 0);
    const started = performance.now();
    const editor = await D.DocxEditor.openAsync(b, new Uint8Array(bytes), exports, {
      ...options, windowSize, onProgress: (done, total) => progress.push([done, total]),
    });
    clearInterval(ticker);
    return { editor, container: b, syncHtml, asyncHtml: b.innerHTML, progress, ticks, elapsed: performance.now() - started };
  }`;

interface Diff { at: number; a: string; b: string }
interface FlowOutcome {
  diff: Diff | null; progress: number[][]; ticks: number; elapsed: number; edited: boolean;
  saved: number; units: number; sections: number; notes: number;
}
interface PagedOutcome { diff: Diff | null; pages: number; progress: number; ticks: number }

test.describe('DocxEditor.openAsync: windowed, yielding mount (#776)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('a flow mount lands the same DOM as open(), one window at a time', async ({ page }) => {
    const outcome = (await page.evaluate(`(${MOUNT_BOTH})(${JSON.stringify({ bytes: Array.from(fixture), options: {}, windowSize: 24 })})
      .then(async (r) => {
        // Still a working editor: a block edits, the session records it, and the bytes save.
        const block = r.container.querySelector('p[data-anchor][contenteditable="true"]');
        const before = r.editor.version;
        block.textContent = 'edited after a windowed mount';
        block.dispatchEvent(new Event('input', { bubbles: true }));
        block.blur();
        block.dispatchEvent(new FocusEvent('blur'));
        await new Promise((resolve) => setTimeout(resolve, 50));
        const after = r.editor.version;
        const saved = r.editor.save().length;
        const units = r.container.querySelectorAll('[data-section-index] > [data-anchor], [data-section-index] > * > [data-anchor]').length;
        const sections = r.container.querySelectorAll('[data-section-index]').length;
        const notes = r.container.querySelectorAll('section.footnotes li, section.endnotes li').length;
        r.editor.close();
        return { diff: (${DIFF_AT})(r.syncHtml, r.asyncHtml),
          progress: r.progress, ticks: r.ticks, elapsed: r.elapsed, edited: after !== null && before !== null && after > before, saved, units, sections, notes };
      })`)) as FlowOutcome;
    expect(outcome.diff, `open():      …${outcome.diff?.a}…\nopenAsync(): …${outcome.diff?.b}…`).toBeNull();
    expect(outcome.units).toBe(117);
    expect(outcome.sections).toBe(4);
    expect(outcome.notes).toBe(2);
    // Five windows of 24 for 117 units, each reported, and the event loop ran between them.
    expect(outcome.progress.length).toBeGreaterThanOrEqual(5);
    expect(outcome.progress[outcome.progress.length - 1]).toEqual([117, 117]);
    expect(outcome.progress.map((p: number[]) => p[0])).toEqual([...outcome.progress.map((p: number[]) => p[0])].sort((x, y) => x - y));
    expect(outcome.ticks).toBeGreaterThanOrEqual(outcome.progress.length - 1);
    expect(outcome.edited).toBe(true);
    expect(outcome.saved).toBeGreaterThan(0);
  });

  test('a paginated mount lands the same pages as open()', async ({ page }) => {
    const outcome = (await page.evaluate(`(${MOUNT_BOTH})(${JSON.stringify({ bytes: Array.from(fixture), options: { paginated: true }, windowSize: 40 })})
      .then((r) => {
        const pages = r.container.querySelectorAll('.page-box').length;
        r.editor.close();
        return { diff: (${DIFF_AT})(r.syncHtml, r.asyncHtml), pages, progress: r.progress.length, ticks: r.ticks };
      })`)) as PagedOutcome;
    expect(outcome.diff, `open():      …${outcome.diff?.a}…\nopenAsync(): …${outcome.diff?.b}…`).toBeNull();
    expect(outcome.pages).toBeGreaterThan(1);
    expect(outcome.progress).toBeGreaterThanOrEqual(3);
    expect(outcome.ticks).toBeGreaterThanOrEqual(outcome.progress - 1);
  });

  for (const kind of ['own', 'end', 'mid', 'leading'] as const) {
    test(`page-break ${kind} retains every fragment through windowed mounting`, async ({ page }) => {
      const result = await page.evaluate<{ texts: string[]; diff: Diff | null }>(`(${MOUNT_BOTH})(${JSON.stringify({
        bytes: Array.from(pageBreakDocx(kind)), options: { paginated: true }, windowSize: 1,
      })}).then(r => {
        const pages = Array.from(r.container.querySelectorAll('.page-box'));
        const texts = pages.map(p => p.querySelector('.page-content').textContent.replace(/\\s/g, ''));
        const diff = (${DIFF_AT})(r.syncHtml, r.asyncHtml);
        r.editor.close(); return { texts, diff };
      })`);
      expect(result.diff).toBeNull();
      expect(result.texts).toEqual(kind === 'leading' ? ['', 'Last']
        : ['First', kind === 'mid' ? 'AfterLast' : 'Last']);
    });
  }

  for (const kind of ['mid', 'leading'] as const) {
    test(`editing a split paragraph ${kind} repaints its later pages without stale fragments`, async ({ page }) => {
      const result = await page.evaluate(async bytes => {
        const D = (window as any).Docxodus;
        const exports = D.getWasmExports ? D.getWasmExports() : D;
        const container = document.createElement('div'); document.body.appendChild(container);
        let edits = 0;
        const editor = await D.DocxEditor.openAsync(container, new Uint8Array(bytes), exports,
          { paginated: true, windowSize: 1, onEdit: () => edits++ });
        const texts = [];
        for (const value of ['Changed', 'Again']) {
          const block = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
          block.textContent = value;
          block.dispatchEvent(new FocusEvent('blur'));
          texts.push(Array.from(container.querySelectorAll('.page-box')).map(p => p.querySelector('.page-content')?.textContent?.replace(/\s/g, '')));
        }
        const saved = editor.save(); editor.close();
        const reopened = D.DocxEditor.open(container, saved, exports, { paginated: true });
        const final = Array.from(container.querySelectorAll('.page-box')).map(p => p.querySelector('.page-content')?.textContent?.replace(/\s/g, ''));
        reopened.close(); return { texts, final, edits };
      }, Array.from(pageBreakDocx(kind)));
      const tail = kind === 'mid' ? 'AfterLast' : 'Last';
      expect(result.texts).toEqual([['Changed', tail], ['Again', tail]]);
      expect(result.final).toEqual(['Again', tail]);
      expect(result.edits).toBe(2);
    });
  }

  test('committing a split head preserves focus moving to another block or outside the editor', async ({ page }) => {
    const result = await page.evaluate(async bytes => {
      const D = (window as any).Docxodus, exports = D.getWasmExports ? D.getWasmExports() : D;
      const container = document.createElement('div'), button = document.createElement('button');
      document.body.append(container, button);
      let edits = 0;
      const editor = await D.DocxEditor.openAsync(container, new Uint8Array(bytes), exports,
        { paginated: true, windowSize: 1, onEdit: () => edits++ });
      const head = container.querySelector('p[data-docx-split-paragraph]') as HTMLElement;
      const last = Array.from(container.querySelectorAll<HTMLElement>('p[contenteditable="true"]')).find(p => p.textContent === 'Last')!;
      head.focus(); head.textContent = 'Changed'; last.focus();
      const focused = document.activeElement as HTMLElement;
      const destination = focused.textContent;
      focused.textContent = 'End'; button.focus();
      const outside = document.activeElement === button;
      const saved = editor.save(); editor.close();
      const reopened = D.DocxEditor.open(container, saved, exports, { paginated: true });
      const texts = Array.from(container.querySelectorAll('.page-content')).map(p => p.textContent?.replace(/\s/g, ''));
      reopened.close(); return { destination, outside, edits, texts };
    }, Array.from(pageBreakDocx('mid')));
    expect(result).toEqual({ destination: 'Last', outside: true, edits: 2, texts: ['Changed', 'AfterEnd'] });
  });

  test('a window never cuts through a border box, and a bundle without the renders mounts synchronously', async ({ page }) => {
    const outcome = await page.evaluate(async ({ boxed, complicated, diffAt }: { boxed: number[]; complicated: number[]; diffAt: string }) => {
      const D = (window as any).Docxodus;
      const exports = D.getWasmExports ? D.getWasmExports() : D;
      const bridge = exports.DocxSessionBridge;
      const handle = bridge.OpenSession(new Uint8Array(boxed), '{}');
      const plan = JSON.parse(bridge.ListRenderedBlocks(handle, false));
      bridge.CloseSession(handle);
      const units = plan.body;
      const groups = new Set(units.map((u: any) => u.group));
      const make = () => { const c = document.createElement('div'); document.body.appendChild(c); return c; };
      const sync = make();
      const syncEditor = D.DocxEditor.open(sync, new Uint8Array(boxed), exports, {});
      const syncHtml = sync.innerHTML;
      syncEditor.close();
      // A window of one unit forces a cut wherever a group allows; the mount's rule extends past
      // a box, so both boxes still arrive whole.
      const container = make();
      const editor = await D.DocxEditor.openAsync(container, new Uint8Array(boxed), exports, { windowSize: 1 });
      const mounted = container.querySelectorAll('[data-section-index] [data-anchor]').length;
      const boxSizes = Array.from(container.querySelectorAll('[data-section-index] > div[style*="border"]'))
        .map((box: any) => box.querySelectorAll(':scope > [data-anchor]').length);
      const diff = new Function('a', 'b', 'return (' + diffAt + ')(a, b)')(syncHtml, container.innerHTML);
      editor.close();
      // Strip the windowed renders from a copy of the exports: openAsync must fall back to open().
      const legacy = { ...exports, DocxSessionBridge: { ...bridge, RenderEditorChromeHtml: undefined, RenderEditorRangeHtml: undefined } };
      const fallbackContainer = make();
      let progressCalls = 0;
      const fallback = await D.DocxEditor.openAsync(fallbackContainer, new Uint8Array(complicated), legacy, { onProgress: () => progressCalls++ });
      const fallbackUnits = fallbackContainer.querySelectorAll('[data-section-index] [data-anchor]').length;
      fallback.close();
      return { total: units.length, groups: groups.size, mounted, boxSizes, diff, progressCalls, fallbackUnits };
    }, { boxed: Array.from(boxes), complicated: Array.from(fixture), diffAt: DIFF_AT });
    expect(outcome.total).toBe(8);
    // Two boxes of two paragraphs each: six groups for eight units.
    expect(outcome.groups).toBe(6);
    expect(outcome.mounted).toBe(8);
    expect(outcome.boxSizes.filter((n: number) => n === 2)).toHaveLength(2);
    expect(outcome.diff, `open():      …${outcome.diff?.a}…\nopenAsync(): …${outcome.diff?.b}…`).toBeNull();
    expect(outcome.progressCalls).toBe(0);
    expect(outcome.fallbackUnits).toBeGreaterThanOrEqual(117);
  });
});
