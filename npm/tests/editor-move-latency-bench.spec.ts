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
 * Block-move latency benchmark — the drag/reorder surface measured on a real
 * bookmark-dense charter (NVCA's published Model Certificate of Incorporation:
 * 234 body blocks, 392 bookmarks, 94 footnotes, 3 section breaks). That shape is
 * the point: the move guards are driven by cross-block ranges and section
 * breaks, so a document with six bookmarks (HC031, which
 * `editor-latency-bench.spec.ts` uses) cannot exercise them.
 *
 * Like that bench this is primarily a MEASUREMENT harness — the numbers go to
 * stdout so perf work has a stable before/after instrument. The budgets below
 * are deliberately loose regression guards, not targets: they are set well
 * above the current measurement so ordinary machine variance cannot fail CI,
 * while an order-of-magnitude regression (the state this bench was written to
 * catch — a hover that cost 624 ms and a drag start that cost 4.3 s) does.
 */
const BUDGETS: Record<string, number> = {
  // Pointer-path work. These run continuously while the mouse moves, so they
  // must not touch the engine at document scale.
  'hover (block boundary crossed)': 100,
  'blockUnitOf (DOM resolve)': 20,
  'resolveDropAt (per dragover)': 5,
  'captureDropZones': 100,
  // One-per-gesture engine queries.
  'ValidMoveTargets (bridge)': 250,
  'ValidMoveTargets review (bridge)': 250,
  'openBlockMoveMenu': 250,
  // The moves themselves, incrementally reconciled.
  'moveBlock direct (down)': 600,
  'moveBlock direct (back)': 600,
  'undo (after move)': 600,
  // Review mode additionally renders both halves of the move pair, so it is
  // the most expensive interaction on the surface. Still ~13x faster than the
  // whole-document remount it used to cost.
  'moveBlock tracked': 1500,
};

test.describe('DocxEditor — block-move latency benchmark', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('move-path latency on a bookmark-dense charter', async ({ page }) => {
    test.setTimeout(300000);
    const bytes = readTestFile('NVCA-Model-COI.docx');
    await page.evaluate((bytesArray: number[]) => {
      (window as any).testDocxBytes = new Uint8Array(bytesArray);
    }, Array.from(bytes));

    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const bytes = (window as any).testDocxBytes as Uint8Array;

      const container = document.createElement('div');
      container.id = 'move-bench';
      container.style.cssText = 'width:800px;margin:0 auto;padding:24px;background:white';
      document.body.appendChild(container);

      const results: Record<string, { ms: number; n: number }> = {};
      const measure = (name: string, n: number, fn: () => void) => {
        const t0 = performance.now();
        for (let i = 0; i < n; i++) fn();
        results[name] = { ms: (performance.now() - t0) / n, n };
      };

      let editor: any;
      measure('open', 1, () => {
        editor = D.DocxEditor.open(container, bytes, D, { blockDrag: true });
      });

      const units: HTMLElement[] = editor['bodyUnitNodes']();
      const movable = units.filter((el: HTMLElement) => editor['isMovableBlockUnit'](el));
      const anchorOf = (el: HTMLElement) => editor['anchorIdOf'](el) as string;

      // ── 1. The raw engine query one drag start / menu open costs. ──────
      const src = movable[Math.floor(movable.length / 2)];
      const srcId = anchorOf(src);
      measure('ValidMoveTargets (bridge)', 5, () => {
        D.DocxSessionBridge.ValidMoveTargets(editor.sessionHandle, srcId);
      });

      // ── 2. Hover: what the pointer crossing one block boundary costs. ──
      // This is the interaction the user pays for continuously while simply
      // moving the mouse across the document, so it is the one that decides
      // whether the surface feels alive or stuck.
      const hoverTargets = movable.slice(0, 12);
      measure('hover (block boundary crossed)', hoverTargets.length, (() => {
        let i = 0;
        return () => {
          const el = hoverTargets[i++ % hoverTargets.length];
          el.dispatchEvent(new PointerEvent('pointermove', { bubbles: true }));
        };
      })());

      // ── 3. blockUnitOf alone (the DOM resolve inside every hover). ─────
      measure('blockUnitOf (DOM resolve)', 20, (() => {
        let i = 0;
        return () => { editor['blockUnitOf'](hoverTargets[i++ % hoverTargets.length]); };
      })());

      // ── 4. Menu open (ValidMoveTargets + 4 destination scans). ─────────
      measure('openBlockMoveMenu', 3, () => {
        editor['blockDragSource'] = src;
        editor['openBlockMoveMenu']();
        editor['closeBlockMoveMenu']();
      });

      // ── 5. Drag-start block measurement, and the per-dragover resolve it
      //       buys: the latter runs on every pointer move during a drag, so it
      //       is the one that decides whether the drop line tracks smoothly.
      measure('captureDropZones', 3, () => { editor['captureDropZones'](); });
      editor['blockDragSource'] = src;
      editor['refreshBlockMoveTargets'](src);
      measure('resolveDropAt (per dragover)', 60, (() => {
        let i = 0;
        return () => { editor['resolveDropAt'](100 + ((i++ * 37) % 600)); };
      })());

      // ── 6. The move itself, direct mode (reconcile repaint). ───────────
      const a = movable[10];
      const b = movable[14];
      let ok = true;
      measure('moveBlock direct (down)', 1, () => {
        ok = editor.moveBlock(anchorOf(a), anchorOf(b), 'after') && ok;
      });
      measure('moveBlock direct (back)', 1, () => {
        const now: HTMLElement[] = editor['bodyUnitNodes']().filter(
          (el: HTMLElement) => editor['isMovableBlockUnit'](el),
        );
        ok = editor.moveBlock(anchorOf(a), anchorOf(now[10]), 'before') && ok;
      });

      // ── 7. Undo of a move. ────────────────────────────────────────────
      measure('undo (after move)', 1, () => { editor.undo(); });
      const directFallbacks = editor.lastReconcileFallback ?? null;

      const blockCount = container.querySelectorAll('[data-anchor]').length;
      const unitCount = units.length;
      editor.close();
      container.remove();

      // ── 8. Tracked move (review mode: named w:moveFrom/w:moveTo pair). ─
      // A separate editor because the tracking mode is fixed at open. Review
      // mode renders BOTH halves of the pair, so it is the costliest move.
      const trackedHost = document.createElement('div');
      trackedHost.style.cssText = 'width:800px;margin:0 auto;padding:24px;background:white';
      document.body.appendChild(trackedHost);
      let trackedOk = true;
      let tracked: any;
      measure('open (review mode)', 1, () => {
        tracked = D.DocxEditor.open(trackedHost, bytes, D, {
          blockDrag: true,
          trackedChanges: 1 /* TrackedChangeMode.RenderInline */,
        });
      });
      const trackedAnchor = (el: HTMLElement) => tracked['anchorIdOf'](el) as string;
      const trackedUnits: HTMLElement[] = tracked['bodyUnitNodes']().filter(
        (el: HTMLElement) => tracked['isMovableBlockUnit'](el),
      );
      // Pick the pair from the engine's own answer: section breaks partition this charter, so a
      // fixed index pair is not necessarily a legal move.
      const tSrcId = trackedAnchor(trackedUnits[Math.floor(trackedUnits.length / 2)]);
      measure('ValidMoveTargets review (bridge)', 3, () => {
        D.DocxSessionBridge.ValidMoveTargets(tracked.sessionHandle, tSrcId);
      });
      const legal = (JSON.parse(
        D.DocxSessionBridge.ValidMoveTargets(tracked.sessionHandle, tSrcId),
      ) as Array<{ anchorId: string; after: boolean }>).filter((t) => t.after);
      measure('moveBlock tracked', 1, () => {
        trackedOk = tracked.moveBlock(tSrcId, legal[legal.length - 1].anchorId, 'after');
      });
      const trackedError = tracked.lastMoveError ?? null;
      // A tracked move must RECONCILE. Falling back to the whole-document remount is the
      // regression this whole bench exists to catch, and it is invisible in a timing average.
      const trackedFallback = tracked.lastReconcileFallback ?? null;
      tracked.close();
      trackedHost.remove();

      return {
        results, blockCount, unitCount, movable: movable.length, ok, trackedOk, trackedError,
        trackedFallback, directFallback: directFallbacks,
      };
    });

    const pad = (n: number) => n.toFixed(1).padStart(9);
    console.log(
      `\n=== block-move latency (NVCA Model COI: ${out.unitCount} body units, ` +
        `${out.movable} movable, ${out.blockCount} anchored nodes) ===`,
    );
    for (const [op, r] of Object.entries(out.results)) {
      const budget = BUDGETS[op];
      console.log(
        `${op.padEnd(32)} ${pad(r.ms)}ms  (avg of ${r.n})` +
          (budget ? `   budget ${budget}ms` : ''),
      );
    }

    expect(out.ok).toBe(true);
    expect(out.trackedOk, `tracked move failed: ${out.trackedError}`).toBe(true);
    // Both moves must reconcile incrementally; a remount fallback is the cliff.
    expect(out.directFallback).toBeNull();
    expect(out.trackedFallback).toBeNull();

    const over = Object.entries(out.results)
      .filter(([op]) => BUDGETS[op] !== undefined && out.results[op].ms > BUDGETS[op])
      .map(([op, r]) => `${op}=${r.ms.toFixed(0)}ms (budget ${BUDGETS[op]}ms)`);
    expect(over, 'over the interaction regression budget').toEqual([]);
  });
});
