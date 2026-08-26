import { test, expect, Page } from '@playwright/test';

// docs/demo/golf.html — DOCX GOLF (copied into the webroot as demo-golf.html
// by pretest). In production it imports the pinned CDN engine; the game and
// its caddie panel (`docx-golf.js`) are demo content living beside the page.
// `?engine=./embed.bundle.js` retargets only the library, which is how this
// spec drives the page locally.
//
// The pure game logic (course validation, stroke counting, caddie phrasing)
// is exercised headlessly by docs/demo/tools/docx-golf.test.mjs. This spec
// guards what only a real browser proves — and, above all, it keeps the
// COURSE honest the way the old puzzle-eval CI did its levels:
//   - no hole starts already solved (a mis-built target would otherwise let
//     an empty round score 100%),
//   - every hole's reference solution reaches zero DocxDiff revisions,
//   - and it does so within par.
// It also proves the human loop end-to-end: a real typed edit, committed by
// blur, is noticed by the byte-fingerprint poll, scored by the diff engine,
// and declared HOLE CLEAR — plus the target/redline views and hole reset.

const OVERRIDE = 'engine=./embed.bundle.js';

async function bootCourse(page: Page) {
  await page.goto(`/demo-golf.html?${OVERRIDE}`);
  await page.waitForFunction(
    () =>
      (window as any).__golf !== undefined ||
      (window as any).__golfError !== undefined,
    { timeout: 120000 },
  );
  const err = await page.evaluate(() => (window as any).__golfError);
  expect(err, `golf boot failed: ${err}`).toBeUndefined();
  // Hole 1 has finished loading once its first score (the honesty check)
  // has run.
  await page.waitForFunction(() => (window as any).__golf.revisionsLeft() >= 0, {
    timeout: 60000,
  });
}

test.describe('DOCX GOLF', () => {
  test('boots the shipped surface with hole 1 on the tee, unsolved', async ({ page }) => {
    await bootCourse(page);

    expect(await page.evaluate(() => (window as any).__golf.holeIndex())).toBe(0);
    expect(await page.evaluate(() => (window as any).__golf.strokes())).toBe(0);
    expect(await page.evaluate(() => (window as any).__golf.revisionsLeft())).toBeGreaterThan(0);
    expect(await page.evaluate(() => (window as any).__golf.cleared())).toBe(false);

    // The caddie panel narrates the hole.
    await expect(page.locator('[data-dxg="title"]')).toContainText('First tee');
    await expect(page.locator('[data-dxg="parchip"]')).toHaveText('par 1');
    // Player → target phrasing: the caddie asks for the typo's removal (and
    // the correction's insertion) — assert on the list, not its ordering.
    await expect(page.locator('[data-dxg="hints"]')).toContainText('Purchasr');
    // The ball is a real document in the real editor.
    await expect(page.locator('#app [contenteditable]').first()).toBeVisible();
  });

  test('every hole is honest: unsolved at the tee, cleared by its reference within par', async ({ page }) => {
    test.setTimeout(300000);
    await bootCourse(page);

    const holes = await page.evaluate(() => (window as any).__golf.course.length);
    expect(holes).toBeGreaterThanOrEqual(4);

    for (let i = 0; i < holes; i++) {
      await page.evaluate(async (idx) => {
        await (window as any).__golf.loadHole(idx);
      }, i);
      await page.waitForFunction(
        (idx) => (window as any).__golf.holeIndex() === idx && (window as any).__golf.revisionsLeft() >= 0,
        i,
        { timeout: 60000 },
      );

      const tee = await page.evaluate(() => (window as any).__golf.check());
      expect(tee, `hole ${i + 1} must not start solved`).toBeGreaterThan(0);

      const ref = await page.evaluate(async () => {
        const g = (window as any).__golf;
        const r = await g.playReference();
        return { ...r, par: g.hole().par, results: undefined };
      });
      expect(ref.allOk, `hole ${i + 1}: every reference op must succeed`).toBe(true);
      expect(ref.ops, `hole ${i + 1}: reference must play within par`).toBeLessThanOrEqual(ref.par);
      expect(ref.revisionsLeft, `hole ${i + 1}: reference must diff to zero`).toBe(0);

      expect(await page.evaluate(() => (window as any).__golf.cleared())).toBe(true);
      await expect(page.locator('[data-dxg="banner"]')).toContainText('HOLE CLEAR');
    }

    // A full round earns a complete scorecard.
    const scorecard = await page.evaluate(() => (window as any).__golf.scorecard());
    expect(scorecard.every((s: unknown) => s !== null)).toBe(true);
  });

  test('hole 1 is playable by hand: type the fix, click away, the caddie calls it', async ({ page }) => {
    test.setTimeout(180000);
    await bootCourse(page);

    // Select exactly the misspelled word (a dblclick at the element center
    // lands on whatever word happens to be mid-paragraph), then retype it —
    // the ordinary human gesture, committed when focus leaves the block.
    await page.evaluate(() => {
      for (const block of Array.from(document.querySelectorAll<HTMLElement>('#app [contenteditable]'))) {
        const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          const text = node as Text;
          const at = text.data.indexOf('Purchasr');
          if (at < 0) continue;
          block.focus();
          const range = document.createRange();
          range.setStart(text, at);
          range.setEnd(text, at + 'Purchasr'.length);
          const selection = window.getSelection()!;
          selection.removeAllRanges();
          selection.addRange(range);
          return;
        }
      }
      throw new Error('typo not found in the editor');
    });
    await page.keyboard.type('Purchaser');
    await page.locator('#app [contenteditable]', { hasText: 'Master Services Agreement' }).first().click();

    // Fingerprint poll notices the commit, the diff engine scores it, and the
    // hole clears in one stroke.
    await page.waitForFunction(() => (window as any).__golf.cleared(), { timeout: 30000 });
    expect(await page.evaluate(() => (window as any).__golf.revisionsLeft())).toBe(0);
    expect(await page.evaluate(() => (window as any).__golf.strokes())).toBeGreaterThanOrEqual(1);
    await expect(page.locator('[data-dxg="banner"]')).toContainText('hole in one');
  });

  test('the target and redline views render the other side of the hole', async ({ page }) => {
    await bootCourse(page);

    await page.locator('[data-dxg="tabs"] button[data-view="target"]').click();
    const frame = page.frameLocator('#caddie iframe');
    await expect(frame.locator('body')).toContainText('the Purchaser identified');

    await page.locator('[data-dxg="tabs"] button[data-view="redline"]').click();
    // The redline of an untouched hole 1 shows the correction as a revision:
    // the target's spelling must appear in the compared output.
    await expect(page.frameLocator('#caddie iframe').locator('body')).toContainText('Purchaser');
  });

  test('show me: the caddie plays the line and the scorecard says so', async ({ page }) => {
    test.setTimeout(180000);
    await bootCourse(page);

    await page.locator('[data-dxg="showme"]').click();
    await page.waitForFunction(() => (window as any).__golf.cleared(), { timeout: 60000 });
    expect(await page.evaluate(() => (window as any).__golf.revisionsLeft())).toBe(0);
    const entry = await page.evaluate(() => (window as any).__golf.scorecard()[0]);
    expect(entry.assisted, 'a shown hole must be marked caddie-assisted').toBe(true);
    await expect(page.locator('[data-dxg="banner"]')).toContainText('caddie');

    // Reset clears the assist and re-tees.
    await page.locator('[data-dxg="reset"]').click();
    await page.waitForFunction(() => (window as any).__golf.revisionsLeft() > 0, { timeout: 60000 });
    expect(await page.evaluate(() => (window as any).__golf.scorecard()[0])).toBeNull();
  });

  test('reset re-tees the hole: strokes and diffs return to the start', async ({ page }) => {
    test.setTimeout(180000);
    await bootCourse(page);

    const before = await page.evaluate(() => (window as any).__golf.revisionsLeft());
    await page.evaluate(async () => { await (window as any).__golf.playReference(); });
    expect(await page.evaluate(() => (window as any).__golf.cleared())).toBe(true);

    await page.locator('[data-dxg="reset"]').click();
    await page.waitForFunction(() => (window as any).__golf.revisionsLeft() > 0, { timeout: 60000 });
    expect(await page.evaluate(() => (window as any).__golf.strokes())).toBe(0);
    expect(await page.evaluate(() => (window as any).__golf.revisionsLeft())).toBe(before);
    expect(await page.evaluate(() => (window as any).__golf.cleared())).toBe(false);
    await expect(page.locator('[data-dxg="banner"]')).not.toBeVisible();
  });
});

test.describe('DOCX GOLF on a phone', () => {
  test.use({ viewport: { width: 390, height: 844 }, hasTouch: true });

  test('the caddie collapses to a strip and expands on demand', async ({ page }) => {
    test.setTimeout(180000);
    await bootCourse(page);

    // Compact: the panel body starts collapsed behind a toggle, with a mini
    // score strip keeping strokes/par/diffs visible, and the document keeps
    // the bulk of the screen.
    const panel = page.locator('#caddie.dxg');
    await expect(panel).toHaveAttribute('data-compact', 'true');
    await expect(page.locator('[data-dxg="brief"]')).not.toBeVisible();
    await expect(page.locator('[data-dxg="mini"]')).toBeVisible();
    await expect(page.locator('[data-dxg="mini"]')).toContainText('2'); // hole 1 tees with 2 diffs

    await page.locator('[data-dxg="toggle"]').click();
    await expect(page.locator('[data-dxg="brief"]')).toBeVisible();
    await expect(page.locator('[data-dxg="hints"]')).toBeVisible();

    await page.locator('[data-dxg="toggle"]').click();
    await expect(page.locator('[data-dxg="brief"]')).not.toBeVisible();
  });
});
