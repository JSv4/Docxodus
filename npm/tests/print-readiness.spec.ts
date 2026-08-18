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
const tinyPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

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
        resource: 'html-image:delayed-readiness-image',
        status: 'complete',
      }),
    ]);
    expect(outcome.result.graphics).toEqual([
      expect.objectContaining({
        kind: 'chart',
        resource: 'graphic:chart:delayed-chart',
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
      pending: ['materialization:graphic:chart:delayed-chart'],
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

  test('uses a monotonic clock even when the wall clock is perturbed', async ({ page }) => {
    const outcome = await page.evaluate(async () => {
      const originalDateNow = Date.now;
      try {
        Date.now = () => Number.MAX_SAFE_INTEGER;
        (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 2_000 });
        (window as any).DocxodusStandalone.releaseReadinessGraphic();
        return await (window as any).DocxodusStandalone.finishReadinessProbe();
      } finally {
        Date.now = originalDateNow;
      }
    });
    expect(outcome.ok).toBe(true);
  });

  test('aborts a pending phase promptly with no readiness timeout', async ({ page }) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 5_000 });
      (window as any).DocxodusStandalone.abortReadinessProbe();
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(false);
    expect(outcome.elapsedMs).toBeLessThan(1_000);
    expect(outcome.error).toEqual(expect.objectContaining({ name: 'AbortError' }));
  });

  test('tracks a replacement graphic rather than a stale element snapshot', async ({ page }) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 2_000 });
      (window as any).DocxodusStandalone.replaceReadinessGraphic();
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(true);
    expect(outcome.result.graphics).toEqual([
      expect.objectContaining({ resource: 'graphic:chart:delayed-chart', status: 'complete' }),
    ]);
  });

  test('rejects a visual-resource inventory one over the configured bound', async ({ page }) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 2_000,
        additionalImages: 1,
        limits: { visualResources: 1 },
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphic();
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(false);
    expect(outcome.error).toEqual(expect.objectContaining({
      phase: 'chart_svg_materialization',
      reason: 'resource_limit',
      pending: ['visual-resource-limit:1'],
    }));
  });

  test('does not miss a broken image inserted by a later graphic producer', async ({ page }) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 2_000 });
      (window as any).DocxodusStandalone.releaseReadinessGraphicWithLateBrokenImage();
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(false);
    expect(outcome.error).toEqual(expect.objectContaining({
      phase: 'image_decoding',
      pending: ['image:html-image:late-broken-image'],
    }));
  });

  test('bounds a resource chain that never reaches a global fixed point', async ({ page }) => {
    await page.evaluate(() => {
      (window as any).DocxodusStandalone.startReadinessProbe({ timeoutMs: 5_000 });
      (window as any).DocxodusStandalone.releaseReadinessGraphicWithUnsettledResourceChain();
    });
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(false);
    expect(outcome.elapsedMs).toBeLessThan(5_000);
    expect(outcome.error).toEqual(expect.objectContaining({
      phase: 'page_tree_stability',
      pending: expect.arrayContaining(['resource-fixed-point:8/8']),
    }));
    expect(outcome.error.message).toContain('did not settle within 8 passes');
  });

  test('probes every CSS image-set candidate and a late background inserted by a graphic', async ({ page }) => {
    await page.evaluate((png) => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 3_000,
        backgroundCss: `image-set(url(${png}) 1x, url(${png}) 2x)`,
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphicWithLateBackground(`url(${png})`);
    }, tinyPng);
    const outcome = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(outcome.ok).toBe(true);
    const backgrounds = outcome.result.images.filter((probe: any) =>
      probe.source === 'css-background');
    expect(backgrounds.length).toBeGreaterThanOrEqual(3);
    expect(backgrounds.every((probe: any) =>
      probe.status === 'complete' && /^[0-9a-f]{64}$/.test(probe.contentKey))).toBe(true);
  });

  test('rejects broken CSS pixels and bounds computed-style observation work', async ({ page }) => {
    await page.evaluate((png) => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 2_000,
        // DPR 1 selects the valid candidate; readiness must still decode the
        // corrupt, non-selected image-set candidate.
        backgroundCss: `image-set('${png}' 1x, 'data:image/png;base64,AAAA' 2x)`,
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphic();
    }, tinyPng);
    const broken = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(broken.ok).toBe(false);
    expect(broken.error).toEqual(expect.objectContaining({
      phase: 'image_decoding',
      pending: ['image:css-background:p:main:background'],
    }));

    await page.evaluate((png) => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 2_000,
        backgroundCss: `url(${png})`,
        limits: { automaticResourceBytes: 16 },
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphic();
    }, tinyPng);
    const bounded = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(bounded.ok).toBe(false);
    expect(bounded.error).toEqual(expect.objectContaining({
      phase: 'image_decoding',
      reason: 'resource_limit',
      pending: ['css-background-code-unit-limit:16'],
    }));

    await page.evaluate((png) => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 2_000,
        backgroundCss: `url(${png})`,
        backgroundCopies: 2,
        // One canonical computed value fits, but charging every painted
        // occurrence makes two uses exceed this bound even though tokenization
        // and digesting are cached by unique value.
        limits: { automaticResourceBytes: 150 },
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphic();
    }, tinyPng);
    const repeated = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(repeated.ok).toBe(false);
    expect(repeated.error).toEqual(expect.objectContaining({
      phase: 'image_decoding',
      reason: 'resource_limit',
      pending: ['css-background-code-unit-limit:150'],
    }));
  });

  test('decodes SVG image pixels and rejects missing or external SVG use targets', async ({ page }) => {
    await page.evaluate((png) => {
      (window as any).DocxodusStandalone.startReadinessProbe({
        timeoutMs: 2_000,
        svgImageUrl: png,
        svgUseHref: '#readiness-use-target',
      });
      (window as any).DocxodusStandalone.releaseReadinessGraphic();
    }, tinyPng);
    const valid = await page.evaluate(() =>
      (window as any).DocxodusStandalone.finishReadinessProbe());
    expect(valid.ok).toBe(true);
    expect(valid.result.images).toContainEqual(expect.objectContaining({
      source: 'svg-image', status: 'complete',
    }));
    expect(valid.result.graphics).toContainEqual(expect.objectContaining({
      source: 'svg-use', status: 'complete',
    }));

    for (const href of ['#missing-readiness-target', 'https://example.invalid/icon.svg#shape']) {
      await page.evaluate((target) => {
        (window as any).DocxodusStandalone.startReadinessProbe({
          timeoutMs: 2_000,
          svgUseHref: target,
        });
        (window as any).DocxodusStandalone.releaseReadinessGraphic();
      }, href);
      const invalid = await page.evaluate(() =>
        (window as any).DocxodusStandalone.finishReadinessProbe());
      expect(invalid.ok).toBe(false);
      expect(invalid.error).toEqual(expect.objectContaining({
        phase: 'chart_svg_materialization',
        pending: ['materialization:svg-use:p:main:svg-use'],
      }));
    }
  });

  test('keys same-family font requests by face attributes and exact accumulated sample', async ({ page }) => {
    const run = async (tail: string) => {
      await page.evaluate((fontTail) => {
        (window as any).DocxodusStandalone.startReadinessProbe({
          timeoutMs: 2_000,
          fontEvidence: true,
          fontTail,
        });
        (window as any).DocxodusStandalone.releaseReadinessGraphic();
      }, tail);
      return page.evaluate(() => (window as any).DocxodusStandalone.finishReadinessProbe());
    };
    const first = await run('Beta');
    const second = await run('Delta');
    expect(first.ok).toBe(true);
    expect(second.ok).toBe(true);
    const firstEvidence = first.result.fonts.filter((probe: any) =>
      probe.requestedFamily.includes('Docxodus Evidence'));
    const secondEvidence = second.result.fonts.filter((probe: any) =>
      probe.requestedFamily.includes('Docxodus Evidence'));
    expect(firstEvidence).toHaveLength(2);
    expect(new Set(firstEvidence.map((probe: any) => probe.requestKey)).size).toBe(2);
    expect(firstEvidence.map((probe: any) => probe.requestKey).sort())
      .not.toEqual(secondEvidence.map((probe: any) => probe.requestKey).sort());
  });
});
