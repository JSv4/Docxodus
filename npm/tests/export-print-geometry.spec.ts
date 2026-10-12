import { expect, test, type Browser, type BrowserContext, type Route } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generatePrintGeometryDocx, type PrintGeometryCase } from './docx-print-geometry-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function pdfText(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const content = await (await pdf.getPage(1)).getTextContent();
    return content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
  } finally { await task.destroy(); }
}

/** Intercept only the owned print snapshot; all Playwright methods keep their original receiver. */
function observePrint(browser: Browser, evidence: { contexts: number; mutations: number },
  mutate?: (html: string) => string): Browser {
  const bound = (target: any, key: string | symbol) => {
    const value = Reflect.get(target, key);
    return typeof value === 'function' ? value.bind(target) : value;
  };
  return new Proxy(browser, { get(target, key) {
    if (key !== 'newContext') return bound(target, key);
    return async (options: Parameters<Browser['newContext']>[0]) => {
      const context = await target.newContext(options);
      if (options?.javaScriptEnabled !== false) return context;
      evidence.contexts++;
      return new Proxy(context, { get(targetContext, contextKey) {
        if (contextKey !== 'route' || !mutate) return bound(targetContext, contextKey);
        return (url: Parameters<BrowserContext['route']>[0], handler: Parameters<BrowserContext['route']>[1],
          routeOptions: Parameters<BrowserContext['route']>[2]) => targetContext.route(url, (route, request) =>
            handler(new Proxy(route, { get(targetRoute, routeKey) {
              if (routeKey !== 'fulfill') return bound(targetRoute, routeKey);
              return (response: Parameters<Route['fulfill']>[0] = {}) => {
                if (typeof response.body === 'string') {
                  const changed = mutate(response.body);
                  expect(changed).not.toBe(response.body);
                  evidence.mutations++;
                  return targetRoute.fulfill({ ...response, body: changed });
                }
                return targetRoute.fulfill(response);
              };
            } }), request), routeOptions);
      } });
    };
  } });
}

for (const kind of ['zero-width', 'ordinary', 'drawing-only', 'drawing-with-text'] as const) {
  test(`${kind} prints the same measurable source fragments in an inert snapshot`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const evidence = { contexts: 0, mutations: 0 };
    const result = await renderDocxArtifacts(generatePrintGeometryDocx(kind), {
      browser: observePrint(browser, evidence), outputs: ['html', 'pdf'],
      reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
    });
    expect(evidence).toEqual({ contexts: 1, mutations: 0 });
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.readiness.some(p => p.phase === 'pdf_print' && p.status === 'complete')).toBe(true);
    expect(result.pageMap.fragments.length).toBe(kind.startsWith('drawing') ? 3 : 1);
    expect(result.pageMap.fragments.every(f => f.geometry.width > 0 && f.geometry.height > 0)).toBe(true);
    const text = await pdfText(result.pdf!);
    if (kind.startsWith('drawing')) {
      expect(text).toContain('BoxMarker Amber sample tray.');
      expect(text).toContain('FollowingMarker The tray is ready.');
      if (kind === 'drawing-only') expect(result.pageMap.fragments[0].geometry.height).toBeGreaterThanOrEqual(144);
      else expect(result.pageMap.fragments[0].geometry.height).toBeLessThan(20);
      if (kind === 'drawing-with-text') expect(text).toContain('HostMarker The archive label.');
    } else expect(text.replace(/\s+/g, '')).toBe('ABC');
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const ownWidth = await page.locator('.page-box p[data-source-anchor-id]').first().evaluate(p => p.getBoundingClientRect().width);
    if (kind === 'zero-width') expect(ownWidth).toBe(0);
    else expect(ownWidth).toBeGreaterThan(0);
    expect(await page.locator('script, iframe, object, embed').count()).toBe(0);
    writeFileSync(testInfo.outputPath(`${kind}.pdf`), result.pdf!);
  });
}

const mutations: Array<{ name: string; kind: PrintGeometryCase; mutate: (html: string) => string }> = [
  { name: 'changed fragment identity', kind: 'ordinary',
    mutate: html => html.replace(/data-page-fragment-id="[^"]+"/, 'data-page-fragment-id="changed-fragment"') },
  { name: 'shifted source geometry', kind: 'ordinary',
    mutate: html => html.replace('</head>', '<style>[data-source-anchor-id] { transform: translateX(10px) !important; }</style></head>') },
  { name: 'active-content element', kind: 'ordinary',
    mutate: html => html.replace('</body>', '<script type="application/json">{"marker":"Inert"}</script></body>') },
  { name: 'lost promoted-drawing association', kind: 'drawing-only',
    mutate: html => html.replace(/ data-docx-anchor-host="[^"]*"/g, '') },
];

for (const { name, kind, mutate } of mutations) {
  test(`print verification still rejects ${name}`, async ({ browser }) => {
    test.setTimeout(120_000);
    const evidence = { contexts: 0, mutations: 0 };
    await expect(renderDocxArtifacts(generatePrintGeometryDocx(kind), {
      browser: observePrint(browser, evidence, mutate), outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    })).rejects.toMatchObject({ code: 'output_verification_failure', phase: 'pdf_print' });
    expect(evidence).toEqual({ contexts: 1, mutations: 1 });
  });
}
