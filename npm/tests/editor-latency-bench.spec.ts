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

/**
 * Editor latency benchmark — a breakdown of where each interactive editor
 * operation spends its time, instrumented at the WASM bridge boundary.
 *
 * This is a MEASUREMENT harness, not a fidelity pin: assertions are minimal
 * (ops must succeed), and the numbers go to stdout so perf work has a stable
 * before/after instrument. Timings on CI vary; nothing here asserts absolute ms.
 */
test.describe('DocxEditor — operation latency benchmark', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('per-op latency breakdown on a real document', async ({ page }) => {
    test.setTimeout(180000);
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const results = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const bytes = (window as any).testDocxBytes as Uint8Array;

      // ── Instrument every bridge call ──────────────────────────────────
      const callLog: Record<string, { count: number; ms: number }> = {};
      const wrap = (obj: any, tag: string) => {
        const out: any = {};
        for (const k of Object.keys(obj)) {
          const v = obj[k];
          if (typeof v !== 'function') { out[k] = v; continue; }
          out[k] = (...args: any[]) => {
            const t0 = performance.now();
            try {
              return v.apply(obj, args);
            } finally {
              const dt = performance.now() - t0;
              const key = `${tag}.${k}`;
              (callLog[key] ??= { count: 0, ms: 0 });
              callLog[key].count++;
              callLog[key].ms += dt;
            }
          };
        }
        return out;
      };
      const instrumented = {
        DocxSessionBridge: wrap(D.DocxSessionBridge, 'S'),
        DocumentConverter: wrap(D.DocumentConverter, 'C'),
      };

      const container = document.createElement('div');
      container.id = 'bench';
      document.body.appendChild(container);

      const results: Record<string, { totalMs: number; calls: Record<string, { count: number; ms: number }>; fallback?: string | null }> = {};
      const measure = async (name: string, fn: () => void | Promise<void>) => {
        for (const k of Object.keys(callLog)) delete callLog[k];
        const t0 = performance.now();
        await fn();
        const total = performance.now() - t0;
        results[name] = {
          totalMs: total,
          calls: JSON.parse(JSON.stringify(callLog)),
          fallback: editor ? (editor as any).lastReconcileFallback : undefined,
        };
      };

      let editor: any;
      await measure('open', () => {
        editor = D.DocxEditor.open(container, bytes, instrumented, {});
      });

      const blocks = () =>
        (Array.from(
          container.querySelectorAll('p[data-anchor][contenteditable="true"]'),
        ) as HTMLElement[]).filter(
          (p) => !p.closest('table') && !p.closest('section.footnotes, section.endnotes') &&
            (p.textContent || '').trim().length > 10,
        );

      const focusBlock = (el: HTMLElement) => {
        el.focus();
        const t = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
        if (t) {
          const r = document.createRange();
          r.setStart(t, 1);
          r.collapse(true);
          const s = window.getSelection()!;
          s.removeAllRanges();
          s.addRange(r);
        }
      };

      // Warm-up: one commit so first-call JIT/lazy-init doesn't pollute op 1.
      {
        const b = blocks()[0];
        focusBlock(b);
        const t = document.createTreeWalker(b, NodeFilter.SHOW_TEXT).nextNode()!;
        t.textContent = t.textContent + 'w';
        b.blur();
      }

      // 1. Text commit (type + blur).
      await measure('textCommit', () => {
        const b = blocks()[1];
        focusBlock(b);
        const t = document.createTreeWalker(b, NodeFilter.SHOW_TEXT).nextNode()!;
        t.textContent = t.textContent + 'X';
        b.blur();
      });

      // 2. Bold on a sub-selection of one block.
      await measure('formatBold', () => {
        const b = blocks()[2];
        focusBlock(b);
        const t = document.createTreeWalker(b, NodeFilter.SHOW_TEXT).nextNode()!;
        const r = document.createRange();
        r.setStart(t, 0);
        r.setEnd(t, Math.min(5, (t.textContent || '').length));
        const s = window.getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        editor.format('bold');
      });

      // 3. Font size on a sub-selection.
      await measure('fontSize', () => {
        const b = blocks()[3];
        focusBlock(b);
        const t = document.createTreeWalker(b, NodeFilter.SHOW_TEXT).nextNode()!;
        const r = document.createRange();
        r.setStart(t, 0);
        r.setEnd(t, Math.min(5, (t.textContent || '').length));
        const s = window.getSelection()!;
        s.removeAllRanges();
        s.addRange(r);
        editor.setFontSize(14);
      });

      // 4. Alignment (paragraph op) on one block.
      await measure('alignCenter', () => {
        const b = blocks()[4];
        focusBlock(b);
        editor.setAlignment('center');
      });

      // 5. Enter (split at caret).
      await measure('enterSplit', () => {
        const b = blocks()[5];
        focusBlock(b);
        (editor as any).splitAtCaret(b);
      });

      // 6. Backspace merge (merge the split back).
      await measure('backspaceMerge', () => {
        const all = blocks();
        const el = all[6];
        const prev = (editor as any).previousEditable(el);
        (editor as any).mergeWithPrevious(prev, el);
      });

      // 7. Structural: insert a table (reconcile path).
      await measure('insertTable', () => {
        const b = blocks()[7];
        focusBlock(b);
        editor.insertTable(2, 2, {});
      });

      // 8. Structural: insert a row into that table.
      await measure('insertTableRow', () => {
        const cellP = container.querySelector(
          'table p[data-anchor][contenteditable="true"]',
        ) as HTMLElement;
        cellP.focus();
        (editor as any).activeBlock = cellP;
        editor.insertTableRow('below');
      });

      // 9. Structural: delete a block (reconcile path).
      await measure('deleteBlock', () => {
        const b = blocks()[8];
        focusBlock(b);
        (editor as any).activeBlock = b;
        editor.deleteBlock();
      });

      // 10. Undo / redo (reconcile path).
      await measure('undo', () => {
        editor.undo();
      });
      await measure('redo', () => {
        editor.redo();
      });

      // 11. Save (lossless serialize).
      let savedLen = 0;
      await measure('save', () => {
        savedLen = editor.save().length;
      });

      const blockCount = container.querySelectorAll('[data-anchor]').length;
      editor.close();
      container.remove();
      return { results, savedLen, blockCount };
    });

    // Emit the breakdown.
    const fmt = (n: number) => n.toFixed(1).padStart(8);
    console.log(`\n=== editor latency benchmark (HC031, ${results.blockCount} anchored nodes) ===`);
    for (const [op, r] of Object.entries(results.results)) {
      const calls = Object.entries(r.calls)
        .sort((a, b) => b[1].ms - a[1].ms)
        .map(([k, v]) => `${k}×${v.count} ${v.ms.toFixed(1)}ms`)
        .join(', ');
      const fb = r.fallback ? `  ⟲remount: ${r.fallback}` : '';
      console.log(`${op.padEnd(16)} ${fmt(r.totalMs)}ms   [${calls}]${fb}`);
    }

    expect(results.savedLen).toBeGreaterThan(0);
    // Every measured op must actually have hit the session bridge.
    for (const key of ['textCommit', 'formatBold', 'enterSplit', 'undo', 'save']) {
      expect(Object.keys(results.results[key].calls).length).toBeGreaterThan(0);
    }
  });
});
