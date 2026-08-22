import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import {
  FONT_RESOLVER_CONTRACT_ID,
  FONT_RESOLVER_SCHEMA_VERSION,
  FONT_SUBSTITUTION_CONTRACT_MATERIAL,
  FONT_SUBSTITUTION_CONTRACT_VERSION,
} from '../src/font-contract.js';
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
const corruptFont = Buffer.alloc(64);
corruptFont.write('wOF2', 0, 'ascii');
corruptFont.writeUInt32BE(corruptFont.byteLength, 8);
corruptFont.writeUInt32BE(corruptFont.byteLength, 16);

function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === 'string' || typeof value === 'boolean'
    || typeof value === 'number') return value;
  if (Array.isArray(value)) return value.map(canonicalValue);
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(value as Record<string, unknown>).sort()) {
    const member = (value as Record<string, unknown>)[key];
    if (member !== undefined) result[key] = canonicalValue(member);
  }
  return result;
}

const substitutionContractDigest = digest(Buffer.from(
  JSON.stringify(canonicalValue(FONT_SUBSTITUTION_CONTRACT_MATERIAL)),
));

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

type AdversarialResolverMode =
  | 'stable-snapshot'
  | 'complete-with-missing'
  | 'selected-unverified-coverage'
  | 'exact-descriptor-mismatch'
  | 'resolved-family-mismatch'
  | 'missing-selection-metadata';

async function convertWithAdversarialResolver(
  page: Page,
  mode: AdversarialResolverMode,
): Promise<any> {
  const source = generateInlineFontDocx();
  const alternate = Buffer.from(validFont);
  alternate[alternate.byteLength - 1] ^= 1;
  return page.evaluate(async ({
    bytes,
    font,
    alternateBytesBase64,
    resolverMode,
    contract,
  }) => {
    const api = (window as any).DocxodusStandalone;
    const reads: Array<{ count: number }> = [];
    let requestFrozen = true;
    const resolver = async (request: any) => {
      requestFrozen = requestFrozen
        && Object.isFrozen(request)
        && Object.isFrozen(request.requests)
        && request.requests.every((item: any) => Object.isFrozen(item)
          && Object.isFrozen(item.familyStack)
          && Object.isFrozen(item.familyKinds)
          && Object.isFrozen(item.sampleCodePoints));
      const faces = request.requests.map((item: any) => ({
        id: `fixture-${item.id}`,
        resolvedFamily: resolverMode === 'resolved-family-mismatch'
          ? 'Docxodus Deliberately Different'
          : item.familyStack[0],
        postscriptName: `DocxodusFixture-${item.id}`,
        version: 'fixture-v1',
        style: item.style,
        weight: resolverMode === 'exact-descriptor-mismatch'
          ? (item.weight === 400 ? 700 : 400)
          : item.weight,
        stretch: item.stretch,
        format: 'woff2',
        mediaType: 'font/woff2',
        byteLength: font.byteLength,
        sha256: font.sha256,
        bytesBase64: font.bytesBase64,
        licenseEvidence: {
          kind: 'installable',
          identity: 'a'.repeat(64),
          noSubsetting: false,
        },
      }));
      if (resolverMode === 'stable-snapshot') {
        for (const face of faces) {
          const counter = { count: 0 };
          reads.push(counter);
          Object.defineProperty(face, 'bytesBase64', {
            configurable: true,
            enumerable: true,
            get() {
              counter.count++;
              return counter.count === 1 ? font.bytesBase64 : alternateBytesBase64;
            },
          });
        }
      }
      const outcomes = request.requests.map((item: any, index: number) => {
        if (resolverMode === 'missing-selection-metadata') {
          return {
            requestId: item.id,
            requestedFamily: item.familyStack[0],
            resolvedFamily: 'Docxodus Forbidden Selection',
            status: 'missing',
            glyphCoverage: 'unverified',
          };
        }
        const firstPoint = item.sampleCodePoints[0];
        return {
          requestId: item.id,
          requestedFamily: item.familyStack[0],
          resolvedFamily: faces[index].resolvedFamily,
          status: 'resolved',
          faceId: faces[index].id,
          metricCompatible: true,
          faceMatch: 'exact',
          glyphCoverage: resolverMode === 'selected-unverified-coverage'
            ? 'unverified'
            : 'complete',
          ...(resolverMode === 'complete-with-missing'
            ? { missingCodePoints: [firstPoint] }
            : {}),
        };
      });
      return {
        schemaVersion: contract.schemaVersion,
        resolverContract: contract.id,
        substitutionContractVersion: contract.substitutionVersion,
        substitutionContractDigest: contract.digest,
        outcomes,
        faces: resolverMode === 'missing-selection-metadata' ? [] : faces,
      };
    };
    try {
      const result = await api.convert(bytes, {
        reviewProfile: 'final',
        commentProfile: 'hidden',
        strictFonts: true,
        documentVersion: 442,
        fontResolver: resolver,
      });
      return {
        ok: true,
        result,
        requestFrozen,
        maximumBytesBase64Reads: Math.max(0, ...reads.map(entry => entry.count)),
      };
    } catch (error) {
      const candidate = error as { message?: string; toJSON?: () => unknown };
      return {
        ok: false,
        error: candidate.toJSON?.() ?? { message: candidate.message ?? String(error) },
        requestFrozen,
        maximumBytesBase64Reads: Math.max(0, ...reads.map(entry => entry.count)),
      };
    }
  }, {
    bytes: Array.from(source),
    font: {
      byteLength: validFont.byteLength,
      sha256: digest(validFont),
      bytesBase64: validFont.toString('base64'),
    },
    alternateBytesBase64: alternate.toString('base64'),
    resolverMode: mode,
    contract: {
      schemaVersion: FONT_RESOLVER_SCHEMA_VERSION,
      id: FONT_RESOLVER_CONTRACT_ID,
      substitutionVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
      digest: substitutionContractDigest,
    },
  });
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
      expect.objectContaining({
        familyStack: ['Inner Face', 'monospace'],
        familyKinds: ['named', 'generic'],
        style: 'italic',
        weight: 700,
      }),
      expect.objectContaining({
        familyStack: ['Outer Face', 'serif'],
        familyKinds: ['named', 'generic'],
        style: 'normal',
        weight: 400,
      }),
    ]));
    expect(requests.flatMap((request: any) => request.sampleCodePoints)).toEqual(
      expect.arrayContaining(['A'.codePointAt(0), 'B'.codePointAt(0)]),
    );
  });

  test('keeps quoted generic words distinct from CSS generic families', async ({ page }) => {
    const requests = await page.evaluate(() => (window as any).DocxodusStandalone.inventoryFontFixture(
      '<span style="font-family:serif">A</span>'
      + '<span style="font-family:&quot;serif&quot;">B</span>',
    ));
    expect(requests).toHaveLength(2);
    expect(requests).toEqual(expect.arrayContaining([
      expect.objectContaining({ familyStack: ['serif'], familyKinds: ['generic'] }),
      expect.objectContaining({ familyStack: ['serif'], familyKinds: ['named'] }),
    ]));
  });

  test('fails closed when the canonical inventory bounds are exceeded', async ({ page }) => {
    await expect(page.evaluate(() => (window as any).DocxodusStandalone.inventoryFontFixture(
      '<span style="font-family:A">A</span><span style="font-family:B">B</span>',
      { fontRequests: 1 },
    ))).rejects.toThrow(/fontRequests limit exceeded/);

    const tooManyFamilies = Array.from({ length: 65 }, (_, index) => `"Family ${index}"`).join(',');
    await expect(page.evaluate((families) =>
      (window as any).DocxodusStandalone.inventoryFontFixture(
        `<span style='font-family:${families}'>A</span>`,
      ), tooManyFamilies)).rejects.toThrow(/at most 64 families/);

    await expect(page.evaluate((family) =>
      (window as any).DocxodusStandalone.inventoryFontFixture(
        `<span style='font-family:"${family}"'>A</span>`,
      ), 'A'.repeat(257))).rejects.toThrow(/at most 256/);

    const oversizedStack = Array.from({ length: 17 }, (_, index) =>
      `"${'A'.repeat(238)}${String(index).padStart(3, '0')}"`).join(',');
    await expect(page.evaluate((families) =>
      (window as any).DocxodusStandalone.inventoryFontFixture(
        `<span style='font-family:${families}'>A</span>`,
      ), oversizedStack)).rejects.toThrow(/at most 4096 characters/);

    await expect(page.evaluate(() =>
      (window as any).DocxodusStandalone.parseFontFamilies('\ud800')))
      .rejects.toThrow(/unpaired UTF-16 surrogate/);
  });

  test('samples large repeated text without materializing one scalar array per node', async ({ page }) => {
    const requests = await page.evaluate((text) =>
      (window as any).DocxodusStandalone.inventoryFontFixture(
        `<span style="font-family:Bounded">${text}</span>`,
      ), 'A'.repeat(250_000));
    expect(requests).toHaveLength(1);
    expect(requests[0].sampleCodePoints).toEqual(['A'.codePointAt(0)]);
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
    expect(new Set(syntheticFamilies).size).toBe(result.renderReport.fonts.length);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_substituted', phase: 'font_loading' }),
      expect.objectContaining({ code: 'font_metric_mismatch', phase: 'font_loading' }),
    ]));
  });

  test('warns when a resolved face required style synthesis', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'synthesized' }));
    expect(result.renderReport.fonts.length).toBeGreaterThan(0);
    expect(result.renderReport.fonts.every((font: any) =>
      font.status === 'resolved' && font.faceMatch === 'synthesized' && font.verified === false)).toBe(true);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_face_synthesized', phase: 'font_loading' }),
    ]));
  });

  test('warns when a resolved face covers only part of the requested text', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'partial-coverage' }));
    expect(result.renderReport.fonts.length).toBeGreaterThan(0);
    expect(result.renderReport.fonts.every((font: any) =>
      font.status === 'resolved' && font.glyphCoverage === 'partial'
      && font.missingCodePointCount === 1 && font.verified === false)).toBe(true);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_glyph_coverage_partial', phase: 'font_loading' }),
    ]));
  });

  test('keeps an explicit resolver miss authoritative while measuring browser fallback', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'missing' }));
    expect(result.renderReport.fonts.length).toBeGreaterThan(0);
    // The document requests fictional families ("Docxodus Requested A"/"B") that Chromium has
    // never heard of and that carry no substitution rule, so a genuine advance-width measurement
    // reports no browser fallback — document.fonts.check() would wrongly report true here for
    // every family regardless of availability, which is exactly the bug this measures against.
    expect(result.renderReport.fonts.every((font: any) =>
      font.status === 'missing'
      && font.source === 'browser'
      && font.glyphCoverage === 'unverified'
      && font.browserFallbackAvailable === false)).toBe(true);
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_unavailable', phase: 'font_loading' }),
    ]));
  });

  test('does not claim the whole environment is unattested when a configured resolver reports one family missing', async ({ page }) => {
    const result = await convertWithResolver(page, fontPlan(validFont, { mode: 'partial-missing' }));
    expect(result.renderReport.fonts.length).toBe(2);
    const [resolved, missing] = result.renderReport.fonts;
    expect(resolved.status).toBe('resolved');
    expect(resolved.source).toBe('attested');
    expect(missing.status).toBe('missing');
    expect(missing.source).toBe('browser');
    // font_unavailable is the accurate, family-scoped warning for the missing font. It must
    // not also trigger font_environment_unverified — that would misreport a fully configured,
    // attested resolver as an unattested browser-observed environment, and misdirect the
    // caller ("use an explicit verified font resolver") when they already have one.
    expect(result.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_unavailable', phase: 'font_loading' }),
    ]));
    expect(result.renderReport.warnings).not.toEqual(expect.arrayContaining([
      expect.objectContaining({ code: 'font_environment_unverified' }),
    ]));
  });

  test('reports corrupt selected faces as load_failed by default and rejects them in strict mode', async ({ page }) => {
    const permissive = await convertWithResolver(page, fontPlan(corruptFont));
    expect(permissive.renderReport.fonts.every((font: any) => font.status === 'load_failed')).toBe(true);
    // The Font Loading API's own rejection reason must reach the report, not just the fact
    // that loading failed — otherwise a corrupt font and an internal engine break both read
    // as the same opaque "could not be decoded or loaded" with nothing to tell them apart.
    expect(permissive.renderReport.fonts.every((font: any) =>
      typeof font.loadFailureDetail === 'string' && font.loadFailureDetail.length > 0)).toBe(true);
    expect(permissive.renderReport.warnings).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'font_load_failed',
        severity: 'warning',
        detail: expect.any(String),
      }),
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

  test('snapshots resolver primitives once and freezes the request contract', async ({ page }) => {
    const outcome = await convertWithAdversarialResolver(page, 'stable-snapshot');
    expect(outcome.ok).toBe(true);
    expect(outcome.requestFrozen).toBe(true);
    expect(outcome.maximumBytesBase64Reads).toBe(1);
    expect(outcome.result.renderReport.fonts.every((font: any) =>
      font.status === 'resolved' && font.faceMatch === 'exact')).toBe(true);
  });

  test('rejects internally contradictory resolver outcomes', async ({ page }) => {
    for (const [mode, message] of [
      ['complete-with-missing', /complete coverage cannot name missing code points/],
      ['selected-unverified-coverage', /must declare complete coverage when no code points are missing/],
      ['exact-descriptor-mismatch', /claims an exact face with different/],
      ['resolved-family-mismatch', /resolved status must match the request's primary family/],
      ['missing-selection-metadata', /cannot attach selection metadata to missing status/],
    ] as const) {
      const outcome = await convertWithAdversarialResolver(page, mode);
      expect(outcome.ok).toBe(false);
      expect(outcome.error).toEqual(expect.objectContaining({
        code: 'resource_policy_failure',
        phase: 'font_loading',
      }));
      expect(outcome.error.message).toMatch(message);
    }
  });

  test('rejects webfonts whose declared expansion exceeds browser font limits', async ({ page }) => {
    const lengthMismatch = Buffer.alloc(64);
    lengthMismatch.write('wOF2', 0, 'ascii');
    lengthMismatch.writeUInt32BE(lengthMismatch.byteLength - 1, 8);
    lengthMismatch.writeUInt32BE(lengthMismatch.byteLength, 16);
    const mismatched = await convertFailure(page, fontPlan(lengthMismatch));
    expect(mismatched).toEqual(expect.objectContaining({
      code: 'resource_policy_failure',
      phase: 'font_loading',
    }));
    expect(mismatched.message).toContain('declared length does not match');

    const expandedBomb = Buffer.alloc(64);
    expandedBomb.write('wOF2', 0, 'ascii');
    expandedBomb.writeUInt32BE(expandedBomb.byteLength, 8);
    expandedBomb.writeUInt32BE(0xffff_ffff, 16);
    const outcome = await convertFailure(page, fontPlan(expandedBomb));
    expect(outcome).toEqual(expect.objectContaining({
      code: 'resource_limit',
      phase: 'font_loading',
    }));
    expect(outcome.message).toContain('expanded bytes');
  });
});
