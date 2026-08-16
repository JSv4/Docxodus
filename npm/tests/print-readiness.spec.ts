import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';

const here = dirname(fileURLToPath(import.meta.url));
const fontBytes = readFileSync(join(
  here,
  '..',
  '..',
  'docs',
  'demo',
  'fonts',
  'docxodus-canvas-mono.woff2',
));
const imageBytes = Buffer.from(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" fill="#0ea5e9"/></svg>',
);

async function ready(page: Page): Promise<void> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
}

test.describe('deterministic print readiness barrier', () => {
  test.beforeEach(async ({ page }) => ready(page));

  test('waits for delayed fonts, images, charts, and a stable final page tree', async ({ page }, testInfo) => {
    let releaseFont!: () => void;
    let releaseImage!: () => void;
    let fontRequested = false;
    let imageRequested = false;
    const fontGate = new Promise<void>((resolve) => { releaseFont = resolve; });
    const imageGate = new Promise<void>((resolve) => { releaseImage = resolve; });

    await page.route('**/fonts/readiness-delayed.woff2', async (route) => {
      fontRequested = true;
      await fontGate;
      await route.fulfill({
        status: 200,
        contentType: 'font/woff2',
        body: fontBytes,
      });
    });
    await page.route('**/readiness-delayed.svg', async (route) => {
      imageRequested = true;
      await imageGate;
      await route.fulfill({
        status: 200,
        contentType: 'image/svg+xml',
        body: imageBytes,
      });
    });

    await page.evaluate(() => (window as any).DocxodusStandalone.startReadinessProbe({
      timeoutMs: 5_000,
      fontUrl: '/fonts/readiness-delayed.woff2',
      imageUrl: '/readiness-delayed.svg',
    }));
    await expect.poll(() => fontRequested).toBe(true);
    await expect.poll(() => imageRequested).toBe(true);
    expect(await page.evaluate(() =>
      (window as any).DocxodusStandalone.readinessProbeSettled())).toBe(false);

    releaseFont();
    await page.waitForTimeout(25);
    expect(await page.evaluate(() =>
      (window as any).DocxodusStandalone.readinessProbeSettled())).toBe(false);

    releaseImage();
    await page.waitForTimeout(25);
    expect(await page.evaluate(() =>
      (window as any).DocxodusStandalone.readinessProbeSettled())).toBe(false);

    await page.evaluate(() => (window as any).DocxodusStandalone.releaseReadinessGraphic());
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());

    expect(outcome.ok).toBe(true);
    expect(outcome.result.fonts).toContainEqual(expect.objectContaining({
      requestedFamily: expect.stringContaining('Docxodus Delayed Readiness'),
      available: true,
    }));
    expect(outcome.result.images).toEqual([
      expect.objectContaining({
        kind: 'image',
        resource: 'delayed-readiness-image',
        status: 'complete',
      }),
    ]);
    expect(outcome.result.graphics).toEqual([
      expect.objectContaining({
        kind: 'chart',
        resource: 'delayed-chart',
        status: 'complete',
      }),
    ]);
    expect(outcome.result.pageTree).toEqual(expect.objectContaining({
      pageCount: 1,
      quietIntervalMs: 100,
      animationFrames: 4,
      mutations: 0,
      resizes: 0,
    }));

    const screenshot = await page.screenshot({ fullPage: true });
    await testInfo.attach('print-readiness-evidence.json', {
      body: Buffer.from(`${JSON.stringify(outcome, null, 2)}\n`),
      contentType: 'application/json',
    });
    await testInfo.attach('print-readiness-final-tree.png', {
      body: screenshot,
      contentType: 'image/png',
    });
  });

  test('times out with the exact incomplete phase and pending resource', async ({ page }, testInfo) => {
    await page.evaluate(() => (window as any).DocxodusStandalone.startReadinessProbe({
      timeoutMs: 100,
    }));
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());

    expect(outcome.ok).toBe(false);
    expect(outcome.elapsedMs).toBeLessThan(1_000);
    expect(outcome.error).toEqual(expect.objectContaining({
      phase: 'chart_svg_materialization',
      pending: ['materialization:delayed-chart'],
    }));
    expect(outcome.error.message).toContain('timed out during chart_svg_materialization');
    await testInfo.attach('print-readiness-timeout.json', {
      body: Buffer.from(`${JSON.stringify(outcome, null, 2)}\n`),
      contentType: 'application/json',
    });
  });

  test('times out deliberately delayed pagination with the page-layout resource pending', async ({ page }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(
      here,
      '..',
      '..',
      'TestFiles',
      'CA',
      'CA001-Plain.docx',
    )));
    const outcome = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertWithDelayedPagination(
        bytes,
        {
          reviewProfile: 'final',
          commentProfile: 'hidden',
          timeoutMs: 5_000,
        },
        5_500,
      ), Array.from(source));

    await testInfo.attach('print-readiness-delayed-pagination.json', {
      body: Buffer.from(`${JSON.stringify(outcome, null, 2)}\n`),
      contentType: 'application/json',
    });
    expect(outcome.paginationEntered, JSON.stringify(outcome, null, 2)).toBe(true);
    expect(outcome.unexpectedSuccess).toBeUndefined();
    expect(outcome.error).toEqual(expect.objectContaining({
      code: 'readiness_timeout',
      phase: 'pagination',
    }));
    expect(outcome.error.report.readiness).toContainEqual(expect.objectContaining({
      phase: 'pagination',
      status: 'failed',
      pending: ['page layout'],
    }));
  });

  test('rejects a delayed page-tree mutation without racing or hanging', async ({ page }, testInfo) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 1_000,
        quietIntervalMs: 150,
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphicWithDelayedTreeMutation(75);
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());

    expect(outcome.ok).toBe(false);
    expect(outcome.elapsedMs).toBeLessThan(1_000);
    expect(outcome.error).toEqual(expect.objectContaining({
      phase: 'page_tree_stability',
      pending: ['page-tree:1-pages'],
    }));
    expect(outcome.error.message).toContain('changed during the quiet interval');
    await testInfo.attach('print-readiness-delayed-page-tree.json', {
      body: Buffer.from(`${JSON.stringify(outcome, null, 2)}\n`),
      contentType: 'application/json',
    });
  });

  test('repeated final-tree checks produce identical page count and geometry signature', async ({ page }) => {
    const run = async () => {
      await page.evaluate(() => {
        (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 2_000 });
        (window as any).DocxodusStandalone.releaseReadinessGraphic();
      });
      return page.evaluate(() => (window as any).DocxodusStandalone.finishReadinessProbe());
    };

    const first = await run();
    const second = await run();
    expect(first.ok).toBe(true);
    expect(second.ok).toBe(true);
    expect(second.result.pageTree.pageCount).toBe(first.result.pageTree.pageCount);
    expect(second.result.pageTree.signature).toBe(first.result.pageTree.signature);
  });
});
