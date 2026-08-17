import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { expect, test, type Page, type TestInfo } from '@playwright/test';
import { generateCorruptImageDocx } from './docx-corrupt-image-fixture.js';
import { generateFootnoteDocx } from './docx-footnote-fixture.js';
import { generateTableCommentDocx } from './docx-page-map-fixture.js';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

const here = dirname(fileURLToPath(import.meta.url));
const testFiles = join(here, '..', '..', 'TestFiles');

interface BrowserExportResult {
  html: string;
  pageCount: number;
  pageMap: {
    rendererFingerprint: string;
    pages: Array<{ pageNumber: number; width: number; height: number; sectionIndex?: number }>;
    fragments: unknown[];
  };
  renderReport: {
    status: 'complete';
    source: { rawPackageBytesDigest: string };
    options: { layoutDigest: string };
    environment: { rendererFingerprint: string; verification: string };
    pages: Array<{ pageNumber: number; width: number; height: number; sectionIndex?: number }>;
    bindings: { pageMapDigest: string; htmlDigest: string };
    readiness: Array<{
      phase: string;
      status: string;
      pending: string[];
      diagnostics?: Array<{ code: string; count: number }>;
    }>;
    fonts: Array<{ requestedFamily: string; status: string; source: string }>;
    resources: Array<{
      kind: string;
      status: string;
      readiness?: string;
      resource?: string;
    }>;
    warnings: Array<{ code: string; severity: string; phase: string }>;
  };
  warnings: unknown[];
  rendererFingerprint: string;
}

function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`;
  if (value !== null && typeof value === 'object') {
    return `{${Object.keys(value as Record<string, unknown>).sort().map((key) =>
      `${JSON.stringify(key)}:${canonical((value as Record<string, unknown>)[key])}`).join(',')}}`;
  }
  return JSON.stringify(value);
}

function digest(value: Uint8Array | string): string {
  return createHash('sha256').update(value).digest('hex');
}

function generateTrackedRevisionDocx(): Uint8Array {
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>
    <w:del w:id="1" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del>
    <w:ins w:id="2" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>
    <w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

async function ready(page: Page): Promise<void> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
}

async function convert(
  page: Page,
  source: Uint8Array,
  mutateCaller = false,
  overrides: Record<string, unknown> = {},
): Promise<BrowserExportResult> {
  return page.evaluate(async ({ bytes, mutate, optionsOverride }) => {
    const api = (window as any).DocxodusStandalone;
    const options = {
      reviewProfile: 'final',
      commentProfile: 'endnotes',
      documentVersion: 17,
      ...optionsOverride,
    };
    return mutate
      ? api.convertAfterCallerMutation(bytes, options)
      : api.convert(bytes, options);
  }, { bytes: Array.from(source), mutate: mutateCaller, optionsOverride: overrides });
}

async function attachSuccessArtifacts(
  testInfo: TestInfo,
  result: BrowserExportResult,
  screenshot: Buffer,
  requests: string[],
): Promise<void> {
  const gallery = testInfo.outputPath('artifact-gallery');
  mkdirSync(gallery, { recursive: true });
  const files = {
    html: Buffer.from(result.html),
    map: Buffer.from(JSON.stringify(result.pageMap, null, 2)),
    report: Buffer.from(JSON.stringify(result.renderReport, null, 2)),
    requests: Buffer.from(JSON.stringify(requests, null, 2)),
    screenshot,
  };
  writeFileSync(join(gallery, 'standalone-final.html'), files.html);
  writeFileSync(join(gallery, 'page-map.json'), files.map);
  writeFileSync(join(gallery, 'render-report.json'), files.report);
  writeFileSync(join(gallery, 'request-log.json'), files.requests);
  writeFileSync(join(gallery, 'offline-reopen.png'), files.screenshot);
  const index = `<!doctype html><meta charset="utf-8"><title>Docxodus #438 artifacts</title>
<h1>Standalone export proof</h1><ul>
<li><a href="standalone-final.html">Final offline HTML</a></li>
<li><a href="page-map.json">PageMap</a></li>
<li><a href="render-report.json">Render report</a></li>
<li><a href="request-log.json">Intercepted request log</a></li>
<li><a href="offline-reopen.png">Offline reopen screenshot</a></li>
</ul>`;
  writeFileSync(join(gallery, 'index.html'), index);

  const dataLink = (type: string, body: Buffer) =>
    `data:${type};base64,${body.toString('base64')}`;
  const viewer = `<!doctype html><meta charset="utf-8"><title>Docxodus #438 evidence</title>
<style>body{font:14px system-ui;margin:24px}iframe{width:100%;height:70vh;border:1px solid #bbb}
img{max-width:100%;border:1px solid #bbb}li{margin:.5em}</style>
<h1>Standalone export proof</h1><ul>
<li><a download="standalone-final.html" href="${dataLink('text/html', files.html)}">Download final offline HTML</a></li>
<li><a download="page-map.json" href="${dataLink('application/json', files.map)}">Download PageMap</a></li>
<li><a download="render-report.json" href="${dataLink('application/json', files.report)}">Download render report</a></li>
<li><a download="request-log.json" href="${dataLink('application/json', files.requests)}">Download request log</a></li>
</ul><h2>Offline reopen</h2><img alt="Offline reopen" src="${dataLink('image/png', screenshot)}">
<h2>Final HTML preview</h2><iframe sandbox="allow-same-origin" src="${dataLink('text/html', files.html)}"></iframe>`;
  writeFileSync(join(gallery, 'view-artifacts.html'), viewer);

  for (const [name, contentType] of [
    ['standalone-final.html', 'text/html'],
    ['page-map.json', 'application/json'],
    ['render-report.json', 'application/json'],
    ['request-log.json', 'application/json'],
    ['offline-reopen.png', 'image/png'],
    ['view-artifacts.html', 'text/html'],
  ] as const) {
    await testInfo.attach(name, { path: join(gallery, name), contentType });
  }
}

test.describe('standalone paginated HTML', () => {
  test.beforeEach(async ({ page }) => ready(page));

  test('materializes one offline tree and binds its report, PageMap, and immutable source', async ({ page, context }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const result = await convert(page, source, true);

    expect(result.html.toLowerCase().startsWith('<!doctype html>')).toBe(true);
    expect(result.html).toContain('data-docxodus-standalone="v1"');
    expect(result.html).not.toContain('id="pagination-staging"');
    expect(result.html).not.toMatch(/<script\b/i);
    expect(result.pageCount).toBeGreaterThan(0);
    expect(result.pageMap.pages).toHaveLength(result.pageCount);
    expect(result.renderReport.pages).toEqual(result.pageMap.pages.map((entry) => ({
      pageNumber: entry.pageNumber,
      width: entry.width,
      height: entry.height,
      sectionIndex: entry.sectionIndex,
    })));
    expect(result.rendererFingerprint).toBe(result.pageMap.rendererFingerprint);
    expect(result.rendererFingerprint).toBe(result.renderReport.environment.rendererFingerprint);
    expect(result.renderReport.environment.verification).toBe('browserObserved');
    expect(result.renderReport.fonts.every((font) =>
      font.requestedFamily.length > 0
      && font.status === 'unverified'
      && font.source === 'browser')).toBe(true);
    expect(result.renderReport.warnings.some((warning) =>
      warning.code === 'font_environment_unverified'
      && warning.severity === 'warning'
      && warning.phase === 'font_loading')).toBe(result.renderReport.fonts.length > 0);
    for (const phase of [
      'wasm_initialization',
      'docx_conversion',
      'font_loading',
      'image_decoding',
      'chart_svg_materialization',
      'pagination',
      'running_story_placement',
      'page_tree_stability',
    ]) {
      expect(result.renderReport.readiness).toContainEqual(expect.objectContaining({
        phase,
        status: 'complete',
        pending: [],
      }));
    }
    expect(result.renderReport.readiness).toContainEqual(expect.objectContaining({
      phase: 'pagination',
      diagnostics: expect.arrayContaining([
        expect.objectContaining({ code: 'sections_processed', count: 1 }),
        expect.objectContaining({ code: 'page_runs_processed', count: 1 }),
      ]),
    }));
    expect(result.renderReport.source.rawPackageBytesDigest).toBe(digest(source));
    expect(result.renderReport.bindings.htmlDigest).toBe(digest(result.html));
    expect(result.renderReport.bindings.pageMapDigest).toBe(digest(canonical(result.pageMap)));

    await page.addStyleTag({ content: `
      .page-box { width: 1px !important; transform: scale(.01) !important; }
      #pagination-container { gap: 999px !important; margin: 777px !important; }
    ` });
    const repeated = await convert(page, source);
    expect(repeated.rendererFingerprint).toBe(result.rendererFingerprint);
    expect(repeated.pageMap).toEqual(result.pageMap);
    expect(repeated.html).toBe(result.html);

    const offline = await context.newPage();
    const requests: string[] = [];
    const offlinePath = testInfo.outputPath('standalone-file-reopen.html');
    writeFileSync(offlinePath, result.html);
    const offlineUrl = pathToFileURL(offlinePath).href;
    offline.on('request', (request) => {
      if (request.url() !== offlineUrl) requests.push(request.url());
    });
    await offline.goto(offlineUrl, { waitUntil: 'load' });
    const audit = await offline.evaluate(() => {
      const pages = Array.from(document.querySelectorAll<HTMLElement>('.page-box'));
      const ids = Array.from(document.querySelectorAll<HTMLElement>('[id]'), (node) => node.id);
      const fragmentLinks = Array.from(document.querySelectorAll<HTMLAnchorElement>('a[href^="#"]'));
      const selection = document.createRange();
      const textNode = document.querySelector('.page-box')?.firstChild;
      if (textNode) selection.selectNodeContents(document.querySelector('.page-box')!);
      return {
        pages: pages.length,
        geometries: pages.map((node) => ({
          width: node.getBoundingClientRect().width * 72 / 96,
          height: node.getBoundingClientRect().height * 72 / 96,
          sectionIndex: Number(node.dataset.sectionIndex ?? 0),
        })),
        idsUnique: ids.length === new Set(ids).size,
        fragmentsResolve: fragmentLinks.every((link) => {
          const target = link.getAttribute('href')!.slice(1);
          return ids.filter((id) => id === target).length === 1;
        }),
        hasSelectableText: selection.toString().trim().length > 0,
        activeElements: document.querySelectorAll('script, iframe, object, embed, link[rel="stylesheet"]').length,
      };
    });
    expect(audit.pages).toBe(result.pageCount);
    expect(audit.idsUnique).toBe(true);
    expect(audit.fragmentsResolve).toBe(true);
    expect(audit.hasSelectableText).toBe(true);
    expect(audit.activeElements).toBe(0);
    expect(requests).toEqual([]);
    for (let index = 0; index < audit.geometries.length; index++) {
      expect(audit.geometries[index].width).toBeCloseTo(result.pageMap.pages[index].width, 1);
      expect(audit.geometries[index].height).toBeCloseTo(result.pageMap.pages[index].height, 1);
      expect(audit.geometries[index].sectionIndex).toBe(result.pageMap.pages[index].sectionIndex ?? 0);
    }

    const screenshot = await offline.screenshot({ fullPage: true });
    await attachSuccessArtifacts(testInfo, result, screenshot, requests);
    await offline.close();
  });

  test('keeps a header-owned embedded image in the offline page tree', async ({ page, context }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'DB005-Headers-With-Images.docx')));
    const result = await convert(page, source);
    const offline = await context.newPage();
    const requests: string[] = [];
    offline.on('request', (request) => requests.push(request.url()));
    await offline.setContent(result.html, { waitUntil: 'load' });
    const headerImages = await offline.locator('.page-header img[src^="data:image/png;base64,"]').count();
    expect(headerImages).toBeGreaterThan(0);
    expect(requests).toEqual([]);
    const screenshot = await offline.screenshot({ fullPage: true });
    await attachSuccessArtifacts(testInfo, result, screenshot, requests);
    await offline.close();
  });

  test('keeps supported charts, notes, margin comments, revisions, and fragment targets', async ({ page }) => {
    const audit = (html: string) => page.evaluate((source) => {
      const parsed = new DOMParser().parseFromString(source, 'text/html');
      const ids = Array.from(parsed.querySelectorAll<HTMLElement>('[id]'), (element) => element.id);
      const fragmentLinks = Array.from(parsed.querySelectorAll<HTMLAnchorElement>('a[href^="#"]'));
      return {
        charts: parsed.querySelectorAll('svg').length,
        footnotes: parsed.querySelectorAll('.page-footnotes [data-footnote-id]').length,
        // Pagination deliberately flattens the converter's endnote section into
        // ordinary page-flow paragraphs. Canonical en-scoped provenance is the
        // durable final-tree identity; the staging class is not.
        endnotes: parsed.querySelectorAll(
          '[data-source-anchor-id^="p:en:"], [data-source-anchor-id^="en:en:"]',
        ).length,
        marginComments: parsed.querySelectorAll('.page-comment-margin [data-comment-id]').length,
        revisions: parsed.querySelectorAll('ins, del, .rev-format-change').length,
        fragmentLinks: fragmentLinks.length,
        fragmentsResolve: fragmentLinks.every((link) => {
          const target = link.getAttribute('href')!.slice(1);
          return ids.filter((id) => id === target).length === 1;
        }),
      };
    }, html);

    const chart = await convert(
      page,
      new Uint8Array(readFileSync(join(testFiles, 'HC043-Chart.docx'))),
    );
    expect((await audit(chart.html)).charts).toBeGreaterThan(0);

    const footnotes = await convert(page, generateFootnoteDocx(2));
    const footnoteAudit = await audit(footnotes.html);
    expect(footnoteAudit.footnotes).toBeGreaterThanOrEqual(2);
    expect(footnoteAudit.fragmentsResolve).toBe(true);

    const endnotes = await convert(
      page,
      new Uint8Array(readFileSync(join(testFiles, 'RC', 'RC007-Endnotes-After.docx'))),
    );
    expect((await audit(endnotes.html)).endnotes).toBeGreaterThan(0);

    const comments = await convert(
      page,
      generateTableCommentDocx(),
      false,
      { commentProfile: 'margin' },
    );
    expect((await audit(comments.html)).marginComments).toBeGreaterThan(0);

    const markup = await convert(page, generateTrackedRevisionDocx(), false, {
      reviewProfile: 'markup',
      commentProfile: 'hidden',
    });
    expect((await audit(markup.html)).revisions).toBeGreaterThan(0);

    const denseLinks = await convert(
      page,
      new Uint8Array(readFileSync(join(testFiles, 'DD', 'DD001-DenseBookmarkXrefFootnote.docx'))),
    );
    const linkAudit = await audit(denseLinks.html);
    expect(linkAudit.fragmentLinks).toBeGreaterThan(0);
    expect(linkAudit.fragmentsResolve).toBe(true);
  });

  test('reports a failed supported-image decode according to warn or strict policy', async ({ page }, testInfo) => {
    const source = generateCorruptImageDocx();
    const warned = await convert(page, source, false, { unsupportedContent: 'warn' });
    expect(warned.html).toContain('docxodus-export-resource-placeholder');
    expect(warned.renderReport.resources).toContainEqual(expect.objectContaining({
      kind: 'image',
      status: 'omitted',
      readiness: 'failed',
    }));
    expect(warned.renderReport.warnings).toContainEqual(expect.objectContaining({
      code: 'image_decode_failed',
      severity: 'warning',
      phase: 'image_decoding',
    }));

    const strictFailure = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertFailure(bytes, {
        reviewProfile: 'final',
        commentProfile: 'hidden',
        unsupportedContent: 'strict',
      }), Array.from(source));
    expect(strictFailure.code).toBe('resource_policy_failure');
    expect(strictFailure.phase).toBe('image_decoding');
    expect(strictFailure.report.readiness).toContainEqual(expect.objectContaining({
      phase: 'image_decoding',
      status: 'failed',
    }));
    expect(strictFailure.report.resources).toContainEqual(expect.objectContaining({
      kind: 'image',
      status: 'omitted',
      readiness: 'failed',
    }));
    await testInfo.attach('image-readiness-policy.json', {
      body: Buffer.from(`${JSON.stringify({
        warning: warned.renderReport,
        strictFailure,
      }, null, 2)}\n`),
      contentType: 'application/json',
    });
  });

  test('preserves a structured failed report for strict unsupported content', async ({ page }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'WC', 'WC012-Math-After.docx')));
    const failure = await page.evaluate(async (bytes) => (window as any).DocxodusStandalone.convertFailure(
      bytes,
      { reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict' },
    ), Array.from(source));

    expect(failure.unexpectedSuccess).toBeUndefined();
    expect(failure.code).toBe('resource_policy_failure');
    expect(failure.phase).toBe('docx_conversion');
    expect(failure.report.status).toBe('failed');
    expect(failure.report.failure.code).toBe('resource_policy_failure');
    await testInfo.attach('failed-render-report.json', {
      body: Buffer.from(JSON.stringify(failure.report, null, 2)),
      contentType: 'application/json',
    });
  });

  test('fails closed with a report when an indivisible body block would clip', async ({ page }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'HC006-Test-01.docx')));
    const failure = await page.evaluate(async (bytes) => (window as any).DocxodusStandalone.convertFailure(
      bytes,
      { reviewProfile: 'final', commentProfile: 'hidden' },
    ), Array.from(source));

    expect(failure.unexpectedSuccess).toBeUndefined();
    expect(failure.code).toBe('pagination_failure');
    expect(failure.phase).toBe('running_story_placement');
    expect(failure.report.status).toBe('failed');
    expect(failure.report.failure.message).toContain('body content is clipped');
    await testInfo.attach('clipped-content-render-report.json', {
      body: Buffer.from(JSON.stringify(failure.report, null, 2)),
      contentType: 'application/json',
    });
  });

  test('publishes complete PageMaps when long footnote paragraphs continue', async ({
    page,
  }, testInfo) => {
    const cases = [
      {
        id: 'single-oversized-paragraph',
        // Issue #489 case C: one paragraph is taller than the maximum note band and must
        // continue across pages without clipping.
        source: generateFootnoteDocx(1, 2, 1, [700]),
      },
      {
        id: 'oversized-leading-paragraph-with-tail',
        // Issue #489 case C2: the long leader and its tail must all survive continuation.
        source: generateFootnoteDocx(1, 2, 3, [700, 8, 8]),
      },
    ];
    const evidence: Array<{
      id: string;
      sourceSha256: string;
      pageMap: BrowserExportResult['pageMap'];
      renderReport: BrowserExportResult['renderReport'];
    }> = [];

    for (const entry of cases) {
      const result = await convert(page, entry.source, false, { commentProfile: 'hidden' });
      const footnotePages = new Set(result.pageMap.fragments
        .filter((fragment: any) => fragment.story === 'footnote')
        .map((fragment: any) => fragment.pageNumber));

      expect(result.pageCount, entry.id).toBeGreaterThan(2);
      expect(footnotePages.size, `${entry.id} must continue its note across pages`)
        .toBeGreaterThan(1);
      expect(result.renderReport.status, entry.id).toBe('complete');
      expect(result.renderReport.source.rawPackageBytesDigest, entry.id).toBe(digest(entry.source));
      expect(result.renderReport.readiness, entry.id).toContainEqual(expect.objectContaining({
        phase: 'running_story_placement',
        status: 'complete',
        pending: [],
      }));
      expect(result.renderReport.bindings.htmlDigest, entry.id).toBe(digest(result.html));
      expect(result.renderReport.bindings.pageMapDigest, entry.id)
        .toBe(digest(canonical(result.pageMap)));
      evidence.push({
        id: entry.id,
        sourceSha256: digest(entry.source),
        pageMap: result.pageMap,
        renderReport: result.renderReport,
      });
    }

    await testInfo.attach('long-footnote-continuation-exports.json', {
      body: Buffer.from(`${JSON.stringify(evidence, null, 2)}\n`),
      contentType: 'application/json',
    });
  });
});

test('PaginationEngine uses the element realm and applies scale exactly once', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  const result = await page.evaluate(async () => {
    const frame = document.createElement('iframe');
    frame.style.position = 'fixed';
    frame.style.left = '-10000px';
    frame.style.width = '1200px';
    frame.style.height = '900px';
    const loaded = new Promise<void>((resolve) => frame.addEventListener('load', () => resolve(), { once: true }));
    frame.srcdoc = `<!doctype html><style>
      body{margin:0}.page-box{box-sizing:border-box;background:white}.page-container{display:flex}
      table{border-collapse:collapse}td{height:20pt}
    </style><div id="staging"><section data-section-index="0" data-page-width="612"
      data-page-height="792" data-content-width="468" data-content-height="648"
      data-margin-top="72" data-margin-right="72" data-margin-bottom="72" data-margin-left="72">
      <div><table><tbody>${'<tr><td>row</td></tr>'.repeat(45)}</tbody></table></div>
    </section></div><div id="pages" class="page-container"></div>`;
    document.body.appendChild(frame);
    await loaded;
    try {
      const foreign = frame.contentDocument!;
      const engine = new (window as any).DocxodusPagination.PaginationEngine(
        foreign.getElementById('staging'),
        foreign.getElementById('pages'),
        { scale: 0.8, showPageNumbers: false, pageGap: 0 },
      );
      const pagination = engine.paginate();
      let secondCall = '';
      try { engine.paginate(); } catch (error) { secondCall = String(error); }
      const first = foreign.querySelector<HTMLElement>('.page-box')!;
      return {
        pages: pagination.totalPages,
        readiness: pagination.readiness,
        width: first.getBoundingClientRect().width,
        authoredWidth: first.style.width,
        zoom: first.style.zoom,
        transform: first.style.transform,
        secondCall,
      };
    } finally {
      frame.remove();
    }
  });

  expect(result.pages).toBeGreaterThan(1);
  expect(result.readiness.status).toBe('ready');
  expect(result.readiness.pageCount).toBe(result.pages);
  expect(result.readiness.diagnostics).toEqual(expect.arrayContaining([
    expect.objectContaining({ code: 'sections_processed', count: 1 }),
    expect.objectContaining({ code: 'page_runs_processed', count: 1 }),
  ]));
  expect(result.authoredWidth).toBe('612pt');
  expect(result.width).toBeGreaterThan(0);
  expect(Number(result.zoom)).toBeCloseTo(0.8, 5);
  expect(result.transform).toBe('');
  expect(result.secondCall).toContain('one-shot');
});
