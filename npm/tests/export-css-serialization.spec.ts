import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { cssCascadeFixture } from './docx-css-cascade-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const css = 'p:first-child > span { color: rgb(255, 0, 0) !important; background-color: rgb(0, 255, 0) !important; }';

test('delivered stylesheet preserves child selectors in HTML and PDF', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(cssCascadeFixture(), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', additionalCss: css,
  });
  await writeFile(testInfo.outputPath('css-cascade.pdf'), result.pdf!);
  await writeFile(testInfo.outputPath('css-cascade.html'), result.html!);
  expect(result.pageCount).toBe(1);
  expect(result.renderReport.options.additionalCss).toBe(css);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise, page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(pdf.numPages).toBe(1);
    expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
    expect(text).toContain('Cedar cascade sample.'); expect(text).toContain('Plain control.');
  } finally { await task.destroy(); }
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    const target = page.locator('.page-content p').first().locator('span').first();
    await expect(target).toHaveCSS('color', 'rgb(255, 0, 0)');
    await expect(target).toHaveCSS('background-color', 'rgb(0, 255, 0)');
    await expect(page.locator('.page-content p').nth(1).locator('span').first()).not.toHaveCSS('color', 'rgb(255, 0, 0)');
  } finally { await page.close(); }
});

test('styles participate in layout identity while empty CSS uses the default', async ({ browser }) => {
  test.setTimeout(120_000);
  const options = { browser, outputs: ['html'] as ['html'], reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
  const plain = await renderDocxArtifacts(cssCascadeFixture(), options);
  const empty = await renderDocxArtifacts(cssCascadeFixture(), { ...options, additionalCss: '' });
  const styled = await renderDocxArtifacts(cssCascadeFixture(), { ...options, additionalCss: css });
  expect(empty.renderReport.options.layoutDigest).toBe(plain.renderReport.options.layoutDigest);
  expect(styled.renderReport.options.layoutDigest).not.toBe(plain.renderReport.options.layoutDigest);
  expect(plain.renderReport.options).not.toHaveProperty('additionalCss');
  expect(empty.renderReport.options).not.toHaveProperty('additionalCss');
});

test('delivered CSS literals retain ampersands and mixed-case closing-style text', async ({ browser }) => {
  test.setTimeout(120_000);
  const literal = '</StYlE><p id=css-escape>A&B ]]> C';
  const additionalCss = `p:first-child::after { content: ${JSON.stringify(literal)}; }`;
  const result = await renderDocxArtifacts(cssCascadeFixture(), {
    browser, outputs: ['html'], reviewProfile: 'final', commentProfile: 'hidden', additionalCss,
  });
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    expect(await page.locator('.page-content p').first().evaluate(element => getComputedStyle(element, '::after').content))
      .toBe(JSON.stringify(literal));
    await expect(page.locator('#css-escape')).toHaveCount(0);
  } finally { await page.close(); }
});

// Inspect the artifact across the host's separate exact-key boundary.
test('framed export host applies its additionalCss batch option', async ({ browser }) => {
  test.setTimeout(120_000);
  const source = Buffer.from(cssCascadeFixture());
  const frame = (bytes: Buffer) => { const size = Buffer.alloc(4); size.writeUInt32BE(bytes.length); return Buffer.concat([size, bytes]); };
  const request = { schemaVersion: 1, sources: [{ id: 'css-source', byteLength: source.length,
    sha256: createHash('sha256').update(source).digest('hex'),
    mediaType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }],
    batches: [{ id: 'css-batch', sourceId: 'css-source', artifactRequestIds: ['css-html'],
      options: { outputs: ['html'], reviewProfile: 'final', commentProfile: 'hidden', additionalCss: css } }] };
  const result = spawnSync(process.execPath, [fileURLToPath(new URL('../../npm-export/dist/host.js', import.meta.url))], {
    input: Buffer.concat([frame(Buffer.from(JSON.stringify(request))), frame(source)]),
    env: process.env, maxBuffer: 16 * 1024 * 1024, timeout: 110_000,
  });
  expect(result.status, result.stderr.toString()).toBe(0);
  let offset = 0;
  const next = () => { const size = result.stdout.readUInt32BE(offset); offset += 4;
    const bytes = result.stdout.subarray(offset, offset + size); offset += size; return bytes; };
  const response = JSON.parse(next().toString());
  expect(response.fatal).toBeUndefined();
  const blobs = new Map<string, Buffer>();
  for (const artifact of response.artifacts) blobs.set(artifact.kind, next());
  expect(offset).toBe(result.stdout.length);
  expect(JSON.parse(blobs.get('renderReport')!.toString()).options.additionalCss).toBe(css);
  const page = await browser.newPage();
  try {
    await page.setContent(blobs.get('html')!.toString());
    await expect(page.locator('.page-content p').first().locator('span').first()).toHaveCSS('color', 'rgb(255, 0, 0)');
  } finally { await page.close(); }
});

test('browser export rejects malformed CSS options before package parsing', async ({ page }) => {
  await page.goto('/standalone-export-harness.html');
  const results = await page.evaluate(async () => {
    // @ts-expect-error Browser module is served from the staged export bundle.
    const { convertDocxToPaginatedHtml } = await import('/export-browser.js');
    const errors = [];
    for (const additionalCss of [null, 42, {}, '\ud800']) {
      try { await convertDocxToPaginatedHtml(new Uint8Array([1]), {
        reviewProfile: 'final', commentProfile: 'hidden', additionalCss,
      }); errors.push(null); }
      catch (error: any) { errors.push({ code: error.code, phase: error.phase, message: error.message }); }
    }
    return errors;
  });
  expect(results).toHaveLength(4);
  for (const error of results) expect(error).toMatchObject({ code: 'invalid_argument', phase: 'input_validation', message: expect.stringContaining('additionalCss') });
});
