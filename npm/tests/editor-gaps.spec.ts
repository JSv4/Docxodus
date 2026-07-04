import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  return Array.from(new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath))));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Regressions for the four gaps found in the S-1 web-editor smoke test.
test.describe('DocxEditor — smoke-test gap regressions', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // GAP 1: paginated mode rendered BLANK for documents with hard page breaks.
  // The converter emitted an EMPTY <div class="page-break"> which serialized self-closing;
  // the browser HTML parser then nested the visible #pagination-container inside the
  // display:none #pagination-staging, so every page box rendered 0x0. HC031 has one hard
  // page break. After the fix the container must be a SIBLING of staging with visible pages.
  test('GAP1: paginated render is visible for a doc with hard page breaks', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.style.width = '960px';
      document.body.appendChild(container);

      D.DocxEditor.open(container, bin, D, { paginated: true });

      const pc = container.querySelector('#pagination-container') as HTMLElement | null;
      const staging = container.querySelector('#pagination-staging') as HTMLElement | null;
      const pageBoxes = pc ? Array.from(pc.querySelectorAll('.page-box')) as HTMLElement[] : [];
      const firstBox = pageBoxes[0];
      const firstBoxRect = firstBox ? firstBox.getBoundingClientRect() : null;
      const visibleText = pc ? (pc.innerText || '').replace(/\s+/g, ' ').trim() : '';

      return {
        hasContainer: !!pc,
        containerInsideStaging: !!(pc && staging && staging.contains(pc)),
        pageBoxCount: pageBoxes.length,
        firstBoxWidth: firstBoxRect ? Math.round(firstBoxRect.width) : 0,
        firstBoxHeight: firstBoxRect ? Math.round(firstBoxRect.height) : 0,
        visibleTextLen: visibleText.length,
      };
    }, bytes);

    // The visible page container must NOT be swallowed by the hidden staging subtree.
    expect(result.hasContainer).toBe(true);
    expect(result.containerInsideStaging).toBe(false);
    // At least one real, non-zero-size page box with the document's content in it.
    expect(result.pageBoxCount).toBeGreaterThan(0);
    expect(result.firstBoxWidth).toBeGreaterThan(100);
    expect(result.firstBoxHeight).toBeGreaterThan(100);
    expect(result.visibleTextLen).toBeGreaterThan(50);
  });

  // GAP 2: rapid focus/blur churn threw uncaught NotFoundError from
  // commitBlock/splitAtCaret's replaceWith ("node ... no longer a child ... moved in a blur
  // event handler"). After the fix, no edit-path DOM swap may throw uncaught, and edits persist.
  test('GAP2: rapid cross-block edits never throw uncaught replaceWith errors', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(String(e)));

    const bytes = readTestFile('HC031-Complicated-Document.docx');
    await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'gap2-editor';
      document.body.appendChild(container);
      (window as any).__ed = D.DocxEditor.open(container, bin, D, {});
    }, bytes);

    // Drive a realistic churn: type into a block then immediately move focus to the next,
    // commit-on-blur racing the next block's focus — repeated across several blocks. Use real
    // keyboard so the synchronous blur->commit->replaceWith interleaving is faithful.
    const blocks = page.locator('#gap2-editor p[data-anchor][contenteditable="true"]');
    const n = Math.min(6, await blocks.count());
    for (let i = 0; i < n; i++) {
      const b = blocks.nth(i);
      await b.click();
      await page.keyboard.type(` x${i}`);
      // Press End+Enter on some to exercise focused-split while dirty, then move on fast.
      if (i % 2 === 0) await page.keyboard.press('Enter');
    }
    // Force a final commit by moving focus away.
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());

    const intactText = await page.evaluate(
      () => ((document.getElementById('gap2-editor') as HTMLElement).innerText || '').includes('x0'),
    );

    expect(errors, `uncaught errors:\n${errors.join('\n')}`).toEqual([]);
    expect(intactText).toBe(true);
  });

  // GAP 3: table cells are editable for TEXT and Enter now SPLITS the cell paragraph into two
  // paragraphs within the SAME cell (the engine keeps the new w:p in the w:tc — the grid is
  // unchanged), so a cell can hold stacked lines. Grid-changing ops (cross-cell Backspace-merge,
  // Tab focus-jump / list-nest) still need whole-table context the single-block model lacks, so
  // they stay inert.
  test('GAP3: Enter splits within a cell (grid unchanged); Tab stays inert; text editing works', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(String(e)));

    const bytes = readTestFile('HC031-Complicated-Document.docx');
    await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'gap3-editor';
      document.body.appendChild(container);
      D.DocxEditor.open(container, bin, D, {});
      // Tag the first editable cell paragraph with enough text to edit.
      const cellP = Array.from(
        container.querySelectorAll('td p[data-anchor][contenteditable="true"]'),
      ).find((p) => (p as HTMLElement).innerText.trim().length > 3) as HTMLElement | undefined;
      if (cellP) cellP.setAttribute('data-testid', 'gap3-cell');
    }, bytes);

    const cell = page.locator('[data-testid="gap3-cell"]');
    await expect(cell).toHaveCount(1);

    // Robust signal: total <p> and table shape across the WHOLE table the cell belongs to.
    const tableStats = () =>
      page.evaluate(() => {
        const c =
          (document.querySelector('[data-testid="gap3-cell"]') as HTMLElement) ||
          (document.querySelector('#gap3-editor td p[contenteditable="true"]') as HTMLElement);
        const table = c.closest('table') as HTMLElement;
        return {
          tableParagraphs: table.querySelectorAll('td p, th p').length,
          tableCells: table.querySelectorAll('td,th').length,
          listMarkers: table.querySelectorAll('[data-list-marker]').length,
        };
      });
    // Resilient: a structural op may re-render (and drop the testid) the cell paragraph,
    // so re-find the first editable cell paragraph and re-tag it before placing the caret.
    const caretToEndOfCell = () =>
      page.evaluate(() => {
        let c = document.querySelector('[data-testid="gap3-cell"]') as HTMLElement | null;
        if (!c) {
          c = document.querySelector('#gap3-editor td p[contenteditable="true"]') as HTMLElement | null;
          c?.setAttribute('data-testid', 'gap3-cell');
        }
        if (!c) return;
        c.focus();
        const r = document.createRange(); r.selectNodeContents(c); r.collapse(false);
        const s = getSelection(); s!.removeAllRanges(); s!.addRange(r);
      });

    const before = await tableStats();

    // Enter splits the cell paragraph IN PLACE: one more paragraph, same cell GRID.
    await caretToEndOfCell();
    await page.keyboard.press('Enter');
    const afterEnter = await tableStats();

    // Tab must be inert (no list nesting / marker, no focus jump that re-shapes the table).
    await caretToEndOfCell();
    await page.keyboard.press('Tab');
    const afterTab = await tableStats();

    // Text editing inside the cell still works and commits.
    await caretToEndOfCell();
    await page.keyboard.type('CELLEDIT');
    await page.evaluate(() => (document.activeElement as HTMLElement)?.blur());
    const textCommitted = await page.evaluate(() =>
      Array.from(document.querySelectorAll('#gap3-editor td p'))
        .some((p) => (p as HTMLElement).innerText.includes('CELLEDIT')),
    );

    expect(errors, `uncaught errors:\n${errors.join('\n')}`).toEqual([]);
    // Enter split the cell paragraph in place: +1 paragraph, SAME cell grid, no list marker.
    expect(afterEnter.tableParagraphs).toBe(before.tableParagraphs + 1);
    expect(afterEnter.tableCells).toBe(before.tableCells);
    expect(afterEnter.listMarkers).toBe(before.listMarkers);
    // Tab stays inert: nothing changes from the post-Enter state.
    expect(afterTab.tableParagraphs).toBe(afterEnter.tableParagraphs);
    expect(afterTab.tableCells).toBe(before.tableCells);
    expect(afterTab.listMarkers).toBe(before.listMarkers);
    // Text editing still works.
    expect(textCommitted).toBe(true);
  });

  // GAP 4: switching pagination mode must PRESERVE session edits (the demo previously re-opened
  // the original bytes on toggle, silently discarding edits). DocxEditor.setPaginated re-renders
  // from the LIVE session, so an edit survives a paginated round-trip.
  test('GAP4: toggling pagination preserves session edits', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const result = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'gap4-editor';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const firstEditable = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]'))
          .find((e) => norm((e as HTMLElement).textContent || '').length > 5) as HTMLElement;

      // Edit a block and commit.
      const MARKER = 'GAP4EDIT marker content';
      const target = firstEditable();
      target.focus();
      target.textContent = MARKER;
      target.dispatchEvent(new Event('blur'));

      const hasMarker = () => norm(container.innerText || '').includes('GAP4EDIT');
      const editedAfterCommit = hasMarker();
      const hasSetPaginated = typeof (editor as any).setPaginated === 'function';

      // Toggle to paginated and back — edit must survive both transitions.
      (editor as any).setPaginated(true);
      const survivedToPaginated = norm(container.textContent || '').includes('GAP4EDIT');
      (editor as any).setPaginated(false);
      const survivedBackToContinuous = hasMarker();

      return { hasSetPaginated, editedAfterCommit, survivedToPaginated, survivedBackToContinuous };
    }, bytes);

    expect(result.hasSetPaginated).toBe(true);
    expect(result.editedAfterCommit).toBe(true);
    expect(result.survivedToPaginated).toBe(true);
    expect(result.survivedBackToContinuous).toBe(true);
  });

  // GAP5 (2nd-round): PROGRAMMATIC drivers (Playwright/automation that move focus + set selection
  // via script, and the paginated duplicate-anchor case) threw uncaught NotFoundError from the
  // edit-path replaceWith — the el.isConnected guard didn't cover a node detached by a synchronous
  // blur during focus transfer. After the fix (replaceNode: parent recheck + try/catch) no edit
  // path throws uncaught, in either render mode.
  test('GAP5: scripted focus/selection churn never throws uncaught replaceWith errors', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(String(e)));

    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const result = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      // Drive a doc in BOTH modes the way an automated harness does: programmatically focus a
      // block, dirty it via execCommand, then immediately focus another — racing commit-on-blur
      // against the next focus, repeatedly. (No awaits between, mirroring the smoke-test driver.)
      const churn = (root: HTMLElement) => {
        const blocks = Array.from(
          root.querySelectorAll('p[data-anchor][contenteditable="true"]'),
        ) as HTMLElement[];
        const n = Math.min(8, blocks.length);
        for (let i = 0; i < n; i++) {
          const el = blocks[i];
          el.focus();
          const range = document.createRange();
          range.selectNodeContents(el);
          range.collapse(false);
          const sel = window.getSelection()!;
          sel.removeAllRanges();
          sel.addRange(range);
          document.execCommand('insertText', false, ` scripted${i}`);
          // Focus the NEXT block synchronously → blur(el) → commitBlock(el) → replaceWith race.
          (blocks[(i + 1) % n]).focus();
        }
        (document.activeElement as HTMLElement)?.blur();
      };

      const cont = document.createElement('div');
      cont.id = 'gap5-cont';
      document.body.appendChild(cont);
      const e1 = D.DocxEditor.open(cont, bin, D, {});
      churn(cont);

      const pag = document.createElement('div');
      pag.id = 'gap5-pag';
      document.body.appendChild(pag);
      const e2 = D.DocxEditor.open(pag, bin, D, { paginated: true });
      const pageRoot = (pag.querySelector('#pagination-container') as HTMLElement) || pag;
      churn(pageRoot);

      return {
        contHasEdit: (cont.innerText || '').includes('scripted0'),
        pagHasEdit: (pageRoot.innerText || '').includes('scripted0'),
      };
    }, bytes);

    expect(errors, `uncaught errors:\n${errors.join('\n')}`).toEqual([]);
    // Edits still land (the fix tolerates the race without losing the committed content).
    expect(result.contHasEdit).toBe(true);
  });

  // GAP6 (2nd-round): paginated mode left the hidden #pagination-staging subtree in the DOM, so
  // every data-anchor existed twice (staging + page-box copy) — querySelector('[data-anchor]') was
  // ambiguous and the staging copy went stale. After the fix staging is removed post-measure, so
  // each data-anchor is unique and the page-box copies are the single source of truth.
  test('GAP6: paginated render has unique data-anchors and no leftover staging', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const result = await page.evaluate((bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'gap6-editor';
      container.style.width = '960px';
      document.body.appendChild(container);
      D.DocxEditor.open(container, bin, D, { paginated: true });

      const anchored = Array.from(container.querySelectorAll('[data-anchor]')) as HTMLElement[];
      const counts = new Map<string, number>();
      for (const el of anchored) {
        const a = el.getAttribute('data-anchor')!;
        counts.set(a, (counts.get(a) || 0) + 1);
      }
      const duplicated = Array.from(counts.entries()).filter(([, c]) => c > 1).length;

      // Pick a real editable anchor and confirm a single, page-box-resident match.
      const sample = anchored.find((e) => e.getAttribute('contenteditable') === 'true');
      const sampleAnchor = sample?.getAttribute('data-anchor') || '';
      const matchesForSample = sampleAnchor
        ? container.querySelectorAll(`[data-anchor="${sampleAnchor}"]`).length
        : 0;

      return {
        stagingPresent: !!container.querySelector('#pagination-staging, .page-staging'),
        anchorCount: anchored.length,
        duplicatedAnchors: duplicated,
        matchesForSample,
        sampleInPageBox: !!sample?.closest('.page-box'),
      };
    }, bytes);

    expect(result.stagingPresent).toBe(false);
    expect(result.anchorCount).toBeGreaterThan(0);
    expect(result.duplicatedAnchors).toBe(0);
    expect(result.matchesForSample).toBe(1);
    expect(result.sampleInPageBox).toBe(true);
  });
});
