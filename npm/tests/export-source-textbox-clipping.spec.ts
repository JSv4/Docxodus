import { expect, test, type Browser, type BrowserContext, type Route } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { sourceClipLines, sourceTextboxClipDocx } from './docx-source-textbox-clip-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['drawing', 'vml'] as const) for (const story of ['body', 'header', 'footer'] as const)
for (const mode of (story === 'body' ? ['fixed', 'auto', 'off-paper'] : ['fixed']) as ('fixed' | 'auto' | 'off-paper')[]) {
  test(`${kind} ${story} ${mode} box reports only authored clipping`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = sourceTextboxClipDocx(kind, mode === 'auto' ? 'auto' : 'fixed', mode === 'off-paper', story);
    writeFileSync(testInfo.outputPath('quail-window.docx'), bytes);
    const result = await renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    const warnings = result.renderReport.warnings.filter(w => w.code === 'source_text_clipped');
    expect(warnings).toHaveLength(mode === 'auto' ? 0 : mode === 'fixed' ? 1 : 3);
    await page.setContent(result.html!);
    const anchors = await page.locator('[data-source-anchor-id]').evaluateAll(elements => {
      return Object.fromEntries(elements.map(e => [e.textContent, e.getAttribute('data-source-anchor-id')]));
    });
    const visibleLines = mode === 'auto' ? sourceClipLines : mode === 'fixed' ? sourceClipLines.slice(0, 2) : [];
    for (const line of sourceClipLines) {
      expect(anchors[line]).toBeTruthy(); // Source text survives in the clipped HTML story.
      expect(result.pageMap.fragments.some(f => f.anchorId === anchors[line])).toBe(visibleLines.includes(line));
      expect(warnings.some(w => w.anchorId === anchors[line])).toBe(!visibleLines.includes(line));
    }
    const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await pdf.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join(' ');
      expect(text).toContain('Quail ledger complete');
      for (const line of visibleLines) expect(text).toContain(line);
      for (const line of sourceClipLines.filter(line => !visibleLines.includes(line))) expect(text).not.toContain(line);
    } finally { await pdf.destroy(); }
  });
}

test('strict export rejects nonblank authored clipping with a source diagnostic', async ({ browser }) => {
  test.setTimeout(120_000);
  await expect(renderDocxArtifacts(sourceTextboxClipDocx('drawing'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});

test('an implicit fixed box does not claim explicit no-autofit provenance', async ({ browser }) => {
  test.setTimeout(120_000);
  await expect(renderDocxArtifacts(sourceTextboxClipDocx('drawing', 'implicit'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden',
  })).rejects.toThrow(/has no measurable fragment/);
});

for (const paragraphRelative of [false, true]) test(`an implicit ${paragraphRelative ? 'paragraph-relative' : 'page-relative'} box wholly outside paper reports source placement without internal clipping eligibility`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(sourceTextboxClipDocx('drawing', 'implicit', true, 'body', paragraphRelative), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden',
  });
  expect(result.renderReport.warnings.filter(w => w.code === 'source_text_clipped')).toHaveLength(3);
  expect(result.html).toContain('data-docx-textbox-source-position="true"');
  expect(result.html).not.toContain('data-docx-textbox-source-clip="true"');
  const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const items = (await (await (await pdf.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const text = items.map(item => item.str).join(' ');
    expect(text).toContain('Quail ledger complete');
    for (const line of sourceClipLines) expect(text).not.toContain(line);
  } finally { await pdf.destroy(); }
});

for (const fit of ['overflow', 'ellipsis'] as const) test(`explicit ${fit} does not certify unsupported internal clipping`, async ({ browser }) => {
  test.setTimeout(120_000);
  await expect(renderDocxArtifacts(sourceTextboxClipDocx('drawing', fit), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden',
  })).rejects.toThrow(/has no measurable fragment/);
});

/** Mutate only the script-disabled print snapshot, after the source map has been finalized. */
function shrinkPrintedBox(browser: Browser, evidence: { changed: number }): Browser {
  const bind = (target: any, key: string | symbol) => {
    const value = Reflect.get(target, key);
    return typeof value === 'function' ? value.bind(target) : value;
  };
  return new Proxy(browser, { get(target, key) {
    if (key !== 'newContext') return bind(target, key);
    return async (options: Parameters<Browser['newContext']>[0]) => {
      const context = await target.newContext(options);
      if (options?.javaScriptEnabled !== false) return context;
      return new Proxy(context, { get(ctx, contextKey) {
        if (contextKey !== 'route') return bind(ctx, contextKey);
        return (url: Parameters<BrowserContext['route']>[0], handler: Parameters<BrowserContext['route']>[1],
          routeOptions: Parameters<BrowserContext['route']>[2]) => ctx.route(url, (route, request) =>
          handler(new Proxy(route, { get(targetRoute, routeKey) {
            if (routeKey !== 'fulfill') return bind(targetRoute, routeKey);
            return (response: Parameters<Route['fulfill']>[0] = {}) => {
              if (typeof response.body !== 'string') return targetRoute.fulfill(response);
              const body = response.body.replace('</head>',
                '<style>[data-docx-textbox-source-clip="true"] { height:14pt !important; }</style></head>');
              expect(body).not.toBe(response.body);
              evidence.changed++;
              return targetRoute.fulfill({ ...response, body });
            };
          } }), request), routeOptions);
      } });
    };
  } });
}

test('print verification rejects newly clipped text in an authored fixed box', async ({ browser }) => {
  test.setTimeout(120_000);
  const evidence = { changed: 0 };
  await expect(renderDocxArtifacts(sourceTextboxClipDocx('drawing'), {
    browser: shrinkPrintedBox(browser, evidence), outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden',
  })).rejects.toMatchObject({ code: 'output_verification_failure', phase: 'pdf_print' });
  expect(evidence.changed).toBe(1);
});
