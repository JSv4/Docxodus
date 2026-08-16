import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

const here = dirname(fileURLToPath(import.meta.url));
const validFont = readFileSync(join(
  here,
  '..',
  '..',
  'docs',
  'demo',
  'fonts',
  'docxodus-canvas-mono.woff2',
));
const corruptFont = Buffer.concat([Buffer.from('wOF2'), Buffer.alloc(60)]);

function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function fontPlan(bytes: Uint8Array, overrides: Record<string, unknown> = {}) {
  return {
    mode: 'exact',
    format: 'woff2',
    mediaType: 'font/woff2',
    byteLength: bytes.byteLength,
    sha256: digest(bytes),
    bytesBase64: Buffer.from(bytes).toString('base64'),
    licenseIdentity: 'a'.repeat(64),
    ...overrides,
  };
}

function generateInlineFontDocx(): Uint8Array {
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
    <w:rFonts w:ascii="Docxodus Requested A" w:hAnsi="Docxodus Requested A"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p>
    <w:r><w:rPr><w:rFonts w:ascii="Docxodus Requested A" w:hAnsi="Docxodus Requested A"/></w:rPr><w:t>Alpha text </w:t></w:r>
    <w:r><w:rPr><w:rFonts w:ascii="Docxodus Requested B" w:hAnsi="Docxodus Requested B"/><w:b/><w:i/></w:rPr><w:t>Bold italic beta text.</w:t></w:r>
  </w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

async function ready(page: Page): Promise<void> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
}

async function convertWithResolver(
  page: Page,
  plan: Record<string, unknown>,
  strictFonts = false,
): Promise<any> {
  const source = generateInlineFontDocx();
  return page.evaluate(async ({ bytes, resolverPlan, strict }) => {
    const api = (window as any).DocxodusStandalone;
    return api.convertWithFontResolver(bytes, {
      reviewProfile: 'final',
      commentProfile: 'hidden',
      strictFonts: strict,
      documentVersion: 442,
    }, resolverPlan);
  }, { bytes: Array.from(source), resolverPlan: plan, strict: strictFonts });
}

async function convertFailure(
  page: Page,
  plan: Record<string, unknown>,
  strictFonts = false,
): Promise<any> {
  const source = generateInlineFontDocx();
  return page.evaluate(async ({ bytes, resolverPlan, strict }) => {
    const api = (window as any).DocxodusStandalone;
    return api.convertWithFontResolverFailure(bytes, {
      reviewProfile: 'final',
      commentProfile: 'hidden',
      strictFonts: strict,
      documentVersion: 442,
    }, resolverPlan);
  }, { bytes: Array.from(source), resolverPlan: plan, strict: strictFonts });
}

test.describe('browser configured font runtime', () => {
  test.beforeEach(async ({ page }) => ready(page));

  test('parses CSS stacks and inventories every rendered text-node face canonically', async ({ page }) => {
    const parsed = await page.evaluate(() => (window as any).DocxodusStandalone.parseFontFamilies(
      '"Comma, Family", Escaped\\20 Name, serif',
    ));
    expect(parsed).toEqual(['Comma, Family', 'Escaped Name', 'serif']);

    const requests = await page.evaluate(() => (window as any).DocxodusStandalone.inventoryFontFixture(
      '<div style="font-family:&quot;Outer Face&quot;,serif">A'
      + '<span style="font-family:&quot;Inner Face&quot;,monospace;font-weight:700;font-style:italic">B</span>'
      + '<span style="display:none;font-family:&quot;Hidden Display&quot;">C</span>'
      + '<span style="visibility:hidden;font-family:&quot;Hidden Visibility&quot;">D</span>'
      + '<span style="content-visibility:hidden;font-family:&quot;Hidden Content&quot;">E</span>'
      + '</div>',
    ));
    expect(requests).toHaveLength(2);
    expect(requests.map((request: any) => request.id)).toEqual(['font-0001', 'font-0002']);
    expect(requests).toEqual(expect.arrayContaining([
      expect.objectContaining({ familyStack: ['Inner Face', 'monospace'], style: 'italic', weight: 700 }),
      expect.objectContaining({ familyStack: ['Outer Face', 'serif'], style: 'normal', weight: 400 }),
    ]));
    expect(requests.flatMap((request: any) => request.sampleCodePoints)).toEqual(
      expect.arrayContaining(['A'.codePointAt(0), 'B'.codePointAt(0)]),
    );
  });

  test('fails closed when the canonical inventory bounds are exceeded', async ({ page }) => {
    await expect(page.evaluate(() => (window as any).DocxodusStandalone.inventoryFontFixture(
      '<span style="font-family:A">A</span><span style="font-family:B">B</span>',
      { fontRequests: 1 },
    ))).rejects.toThrow(/fontRequests limit exceeded/);
  });

  test('loads exact configured faces, serializes them, and passes strict reopened readiness', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont), true);
    expect(result.renderReport.fontIdentity.resolutionDigest).toMatch(/^[0-9a-f]{64}$/);
    expect(result.renderReport.fonts.length).toBeGreaterThanOrEqual(2);
    expect(result.renderReport.fonts.every((font: any) =>
      font.status === 'resolved'
      && font.source === 'configured'
      && font.faceMatch === 'exact'
      && font.glyphCoverage === 'complete')).toBe(true);
    expect(result.html).toContain('id="docxodus-configured-fonts"');
    expect(result.html).toContain('data:font/woff2;base64,');
    expect(result.html).toMatch(/__DocxodusConfigured_[0-9a-f_]+/);
    const syntheticFamilies = Array.from(result.html.matchAll(
      /@font-face\{font-family:"(__DocxodusConfigured_[0-9a-f_]+)"/g,
    ), (match: RegExpMatchArray) => match[1]);
    expect(new Set(syntheticFamilies).size).toBeGreaterThanOrEqual(2);
  });

  test('warns for metric-changing substitution and retains deterministic configured evidence', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'substituted' }));
    expect(result.renderReport.fonts.every((font: any) => font.status === 'substituted')).toBe(true);
    expect(new Set(result.renderReport.fonts.map((font: any) => font.requestedFamily)).size)
      .toBeGreaterThanOrEqual(2);
    const syntheticFamilies = Array.from(result.html.matchAll(
      /@font-face\{font-family:"(__DocxodusConfigured_[0-9a-f_]+)"/g,
    ), (match: RegExpMatchArray) => match[1]);
    expect(syntheticFamilies).toHaveLength(result.renderReport.fonts.length);
    expect(new Set(syntheticFamilies).size).toBe(1);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_substituted', phase: 'font_loading' }),
      expect.objectContaining({ code: 'font_metric_mismatch', phase: 'font_loading' }),
    ]));
  });

  test('keeps an explicit resolver miss authoritative while recording browser fallback', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'missing' }));
    expect(result.renderReport.fonts.length).toBeGreaterThan(0);
    expect(result.renderReport.fonts.every((font: any) =>
      font.status === 'missing'
      && font.source === 'browser'
      && font.browserFallbackAvailable === true)).toBe(true);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_unavailable', phase: 'font_loading' }),
    ]));
  });

  test('reports corrupt selected faces as load_failed by default and rejects them in strict mode', async ({ page }) => {
    const permissive = await convertWithResolver(page, fontPlan(corruptFont));
    expect(permissive.renderReport.fonts.every((font: any) => font.status === 'load_failed')).toBe(true);
    expect(permissive.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_load_failed', severity: 'warning' }),
    ]));
    expect(permissive.html).not.toContain('id="docxodus-configured-fonts"');

    const strict = await convertFailure(page, fontPlan(corruptFont), true);
    expect(strict).toEqual(expect.objectContaining({
      code: 'resource_policy_failure',
      phase: 'font_loading',
    }));
    expect(strict.report.fonts.every((font: any) => font.status === 'load_failed')).toBe(true);
  });

  test('rejects digest mismatches and resolver drift across pristine retries', async ({ page }) => {
    const invalidDigest = await convertFailure(page, fontPlan(validFont, { sha256: '0'.repeat(64) }));
    expect(invalidDigest).toEqual(expect.objectContaining({
      code: 'resource_policy_failure',
      phase: 'font_loading',
    }));
    expect(invalidDigest.message).toContain('do not match sha256');

    const drift = await convertFailure(page, fontPlan(validFont, { drift: true }));
    expect(drift).toEqual(expect.objectContaining({
      code: 'resource_policy_failure',
      phase: 'font_loading',
    }));
    expect(drift.message).toContain('different configuration');
  });
});
