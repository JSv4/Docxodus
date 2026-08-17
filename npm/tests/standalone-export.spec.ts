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
const reportSchemaV1Bytes = readFileSync(
  join(here, '..', '..', 'docs', 'schemas', 'render-report-v1.schema.json'),
);
const reportSchemaV1 = JSON.parse(reportSchemaV1Bytes.toString('utf8'));
const reportSchemaV2Bytes = readFileSync(
  join(here, '..', '..', 'docs', 'schemas', 'render-report-v2.schema.json'),
);
const reportSchemaV2 = JSON.parse(reportSchemaV2Bytes.toString('utf8'));
const reportSchemaV3 = JSON.parse(readFileSync(
  join(here, '..', '..', 'docs', 'schemas', 'render-report-v3.schema.json'),
  'utf8',
));

interface BrowserExportResult {
  html: string;
  pageCount: number;
  pageMap: {
    rendererFingerprint: string;
    pages: Array<{
      pageNumber: number;
      pageInSection: number;
      pageName: string;
      width: number;
      height: number;
      sectionIndex?: number;
    }>;
    fragments: unknown[];
  };
  renderReport: {
    schema: 'https://docxodus.dev/schemas/render/render-report/v3';
    schemaVersion: 3;
    status: 'complete';
    source: { rawPackageBytesDigest: string };
    derivedProfileSource?: { rawPackageBytesDigest: string; byteLength: number };
    options: {
      reviewProfile: 'final' | 'original' | 'markup';
      reviewProfileAlreadyApplied: boolean;
      title: string;
      outputs: Array<'html' | 'pdf'>;
      layoutDigest: string;
      policy: { limits: Record<string, number> };
    };
    environment: { rendererFingerprint: string; verification: string };
    pages: BrowserExportResult['pageMap']['pages'];
    bindings: { pageMapDigest: string; htmlDigest: string };
    readiness: Array<{
      phase: string;
      status: string;
      pending: string[];
      diagnostics?: Array<{ code: string; count: number }>;
    }>;
    fontIdentity: { resolutionDigest: string };
    fonts: Array<{
      requestId: string;
      requestedFamily: string;
      requestedFamilies: string[];
      requestedFamilyKinds: Array<'named' | 'generic'>;
      status: string;
      source: string;
    }>;
    fontReadiness: Array<{ requestKey: string; requestedFamily: string; available: boolean }>;
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

function schemaErrors(
  root: any,
  schema: any,
  value: any,
  path = '$',
): string[] {
  if (typeof schema === 'boolean') return schema ? [] : [`${path}: rejected`];
  if (!schema || typeof schema !== 'object') return [];
  const errors: string[] = [];
  if (schema.$ref) {
    const target = schema.$ref.split('/').slice(1).reduce(
      (current: any, token: string) => current[token.replace(/~1/g, '/').replace(/~0/g, '~')],
      root,
    );
    errors.push(...schemaErrors(root, target, value, path));
  }
  const equal = (left: unknown, right: unknown) => canonical(left) === canonical(right);
  if ('const' in schema && !equal(value, schema.const)) errors.push(`${path}: const`);
  if (schema.enum && !schema.enum.some((entry: unknown) => equal(value, entry))) {
    errors.push(`${path}: enum`);
  }
  const typeMatches = schema.type === undefined
    || (schema.type === 'object' && value !== null && typeof value === 'object' && !Array.isArray(value))
    || (schema.type === 'array' && Array.isArray(value))
    || (schema.type === 'string' && typeof value === 'string')
    || (schema.type === 'number' && typeof value === 'number' && Number.isFinite(value))
    || (schema.type === 'integer' && Number.isSafeInteger(value))
    || (schema.type === 'boolean' && typeof value === 'boolean')
    || (schema.type === 'null' && value === null);
  if (!typeMatches) return [...errors, `${path}: type ${schema.type}`];
  if (schema.allOf) schema.allOf.forEach((entry: any) =>
    errors.push(...schemaErrors(root, entry, value, path)));
  if (schema.anyOf && !schema.anyOf.some((entry: any) =>
    schemaErrors(root, entry, value, path).length === 0)) errors.push(`${path}: anyOf`);
  if (schema.oneOf) {
    const matches = schema.oneOf.filter((entry: any) =>
      schemaErrors(root, entry, value, path).length === 0).length;
    if (matches !== 1) errors.push(`${path}: oneOf(${matches})`);
  }
  if (schema.not && schemaErrors(root, schema.not, value, path).length === 0) {
    errors.push(`${path}: not`);
  }
  if (schema.if) {
    const branch = schemaErrors(root, schema.if, value, path).length === 0
      ? schema.then
      : schema.else;
    if (branch) errors.push(...schemaErrors(root, branch, value, path));
  }
  if (typeof value === 'string') {
    if (schema.minLength !== undefined && value.length < schema.minLength) {
      errors.push(`${path}: minLength`);
    }
    if (schema.pattern && !(new RegExp(schema.pattern).test(value))) errors.push(`${path}: pattern`);
  }
  if (typeof value === 'number') {
    if (schema.minimum !== undefined && value < schema.minimum) errors.push(`${path}: minimum`);
    if (schema.maximum !== undefined && value > schema.maximum) errors.push(`${path}: maximum`);
    if (schema.exclusiveMinimum !== undefined && value <= schema.exclusiveMinimum) {
      errors.push(`${path}: exclusiveMinimum`);
    }
  }
  if (Array.isArray(value)) {
    if (schema.minItems !== undefined && value.length < schema.minItems) errors.push(`${path}: minItems`);
    if (schema.maxItems !== undefined && value.length > schema.maxItems) errors.push(`${path}: maxItems`);
    schema.prefixItems?.forEach((entry: any, index: number) => {
      if (index < value.length) errors.push(...schemaErrors(root, entry, value[index], `${path}[${index}]`));
    });
    if (schema.items && typeof schema.items === 'object') value.forEach((entry, index) =>
      errors.push(...schemaErrors(root, schema.items, entry, `${path}[${index}]`)));
    if (schema.contains) {
      const matches = value.filter((entry, index) =>
        schemaErrors(root, schema.contains, entry, `${path}[${index}]`).length === 0).length;
      if (matches < (schema.minContains ?? 1) || matches > (schema.maxContains ?? Infinity)) {
        errors.push(`${path}: contains(${matches})`);
      }
    }
  }
  if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
    for (const key of schema.required ?? []) {
      if (!Object.prototype.hasOwnProperty.call(value, key)) errors.push(`${path}.${key}: required`);
    }
    for (const [key, entry] of Object.entries(schema.properties ?? {})) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        errors.push(...schemaErrors(root, entry, value[key], `${path}.${key}`));
      }
    }
    if (schema.additionalProperties === false) {
      const allowed = new Set(Object.keys(schema.properties ?? {}));
      for (const key of Object.keys(value)) if (!allowed.has(key)) errors.push(`${path}.${key}: additional`);
    }
  }
  return errors;
}

test('keeps frozen render-report v1/v2 disjoint from the closed v3 font schema', () => {
  const v1Bytes = reportSchemaV1Bytes;
  const v1 = reportSchemaV1;
  const schema = reportSchemaV2;
  const limits = JSON.parse(readFileSync(join(here, '..', 'src', 'export-resource-limits-v1.json'), 'utf8'));
  const definitions = schema.$defs;
  expect(digest(v1Bytes)).toBe('50476223d2707ebd178a08239273af14b5dd9fb47504a7f152e17230f13accad');
  expect(digest(reportSchemaV2Bytes)).toBe('a9f80b7cc6134034b376b57b43aa80ead51301227b0d38b39f5120aa9ad1e78c');
  expect(v1.$id).toBe('https://docxodus.dev/schemas/render/render-report/v1');
  expect(v1.$defs.complete.properties.schemaVersion.const).toBe(1);
  expect(v1.$defs.baseProperties.readiness.items.properties.diagnostics).toBeUndefined();
  expect(v1.$defs.baseProperties.resources.items.properties.readiness).toBeUndefined();
  expect(schema.$id).toBe('https://docxodus.dev/schemas/render/render-report/v2');
  expect(reportSchemaV3.$id).toBe('https://docxodus.dev/schemas/render/render-report/v3');
  expect(reportSchemaV3.$defs.complete.properties.schemaVersion.const).toBe(3);
  expect(reportSchemaV3.$defs.complete.required).toContain('fontReadiness');
  expect(reportSchemaV3.$defs.failed.required).toContain('fontReadiness');
  expect(reportSchemaV3.$defs.baseProperties.fonts.items).toEqual({
    $ref: '#/$defs/fontResolution',
  });
  expect(reportSchemaV3.$defs.baseProperties.fontReadiness.items).toEqual({
    $ref: '#/$defs/fontReadinessProbe',
  });
  expect(definitions.complete.properties.schemaVersion.const).toBe(2);
  expect(definitions.complete.properties.schema.const).not.toBe(
    v1.$defs.complete.properties.schema.const,
  );
  expect(schema.$schema).toBe('https://json-schema.org/draft/2020-12/schema');
  expect(schema.oneOf).toEqual([
    { $ref: '#/$defs/complete' },
    { $ref: '#/$defs/failed' },
  ]);
  expect(definitions.complete.additionalProperties).toBe(false);
  expect(definitions.failed.additionalProperties).toBe(false);
  expect(definitions.complete.required).toEqual(expect.arrayContaining([
    'fontIdentity', 'environment', 'pages', 'bindings',
  ]));
  expect(definitions.failed.required).toEqual(expect.arrayContaining(['failure', 'unavailable']));
  expect(definitions.baseProperties.options.properties.outputs.oneOf.map(
    (entry: { const: string[] }) => entry.const,
  )).toEqual([[], ['html'], ['pdf'], ['html', 'pdf']]);
  expect(new Set(definitions.limits.required)).toEqual(new Set(Object.keys(limits.defaults)));
  expect(Object.keys(limits.defaults)).toEqual(Object.keys(limits.hardCeilings));
  expect(definitions.runtimeAttestation.required).toEqual(expect.arrayContaining([
    'chromiumProduct', 'chromiumBuild', 'executableSha256', 'launchFlags',
    'hostFontsDigest', 'basis',
  ]));
  expect(reportSchemaV3.$defs.runtimeAttestation.required).not.toContain('executableSha256');
  expect(definitions.errorCode.enum).toEqual(expect.arrayContaining([
    'invalid_argument', 'source_digest_mismatch', 'document_version_unrepresentable',
    'operation_cancelled', 'resource_limit',
  ]));
  const readinessItem = definitions.baseProperties.readiness.items;
  const paginationRule = readinessItem.allOf.find((entry: any) =>
    entry.if?.properties?.phase?.const === 'pagination');
  expect(paginationRule.then.required).toContain('diagnostics');
  expect(paginationRule.else).toEqual({ not: { required: ['diagnostics'] } });
  expect(readinessItem.properties.diagnostics).toEqual(expect.objectContaining({
    minItems: 4,
    maxItems: 4,
  }));
  expect(paginationRule.then.properties.diagnostics.allOf.map(
    (entry: any) => entry.contains.properties.code.const,
  )).toEqual([
    'sections_processed',
    'page_runs_processed',
    'source_anchors_inventoried',
    'note_references_inventoried',
  ]);
  const resourceRules = definitions.baseProperties.resources.items.allOf;
  const externalLinkRule = resourceRules.find((entry: any) =>
    entry.if?.properties?.kind?.const === 'external_link');
  expect(externalLinkRule.then.not.anyOf).toEqual([
    { required: ['readiness'] },
    { required: ['contentKey'] },
  ]);
  expect(resourceRules).toEqual(expect.arrayContaining([
    expect.objectContaining({
      if: { properties: { readiness: { const: 'failed' } }, required: ['readiness'] },
      then: expect.objectContaining({
        properties: expect.objectContaining({ status: { const: 'omitted' } }),
      }),
    }),
  ]));
});

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

  test('uses strict RFC 8785-compatible canonical JSON at the browser boundary', async ({ page }) => {
    const canonicalized = await page.evaluate(() => (window as any).DocxodusStandalone.canonicalJson({
      z: -0,
      nested: { b: true, a: '\u20ac' },
      array: [3, null, 'é'],
      omitted: undefined,
    }));
    expect(canonicalized).toBe('{"array":[3,null,"é"],"nested":{"a":"€","b":true},"z":0}');
    const rejection = await page.evaluate(() => {
      try {
        (window as any).DocxodusStandalone.canonicalJson({ invalid: '\ud800' });
        return '';
      } catch (error) {
        return String(error);
      }
    });
    expect(rejection).toContain('unpaired UTF-16 surrogates');
  });

  test('materializes one offline tree and binds its report, PageMap, and immutable source', async ({ page, context }, testInfo) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const result = await convert(page, source, true);

    expect(result.html.toLowerCase().startsWith('<!doctype html>')).toBe(true);
    expect(result.html).toContain('data-docxodus-standalone="v1"');
    expect(result.html).not.toContain('id="pagination-staging"');
    expect(result.html).not.toMatch(/<script\b/i);
    expect(result.pageCount).toBeGreaterThan(0);
    expect(result.pageMap.pages).toHaveLength(result.pageCount);
    expect(result.renderReport.pages).toEqual(result.pageMap.pages);
    expect(result.renderReport.options.outputs).toEqual(['html']);
    expect(result.renderReport.options.title).toBe('');
    expect(result.renderReport.options.reviewProfileAlreadyApplied).toBe(false);
    expect(result.renderReport.derivedProfileSource).toBeDefined();
    expect(result.rendererFingerprint).toBe(result.pageMap.rendererFingerprint);
    expect(result.rendererFingerprint).toBe(result.renderReport.environment.rendererFingerprint);
    expect(result.renderReport.environment.verification).toBe('browserObserved');
    expect(result.renderReport).toEqual(expect.objectContaining({
      schema: 'https://docxodus.dev/schemas/render/render-report/v3',
      schemaVersion: 3,
    }));
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, result.renderReport)).toEqual([]);
    expect(schemaErrors(reportSchemaV1, reportSchemaV1, result.renderReport)).not.toEqual([]);
    expect(schemaErrors(reportSchemaV2, reportSchemaV2, result.renderReport)).not.toEqual([]);
    const legacyV2Report = structuredClone(result.renderReport) as any;
    legacyV2Report.schema = 'https://docxodus.dev/schemas/render/render-report/v2';
    legacyV2Report.schemaVersion = 2;
    legacyV2Report.fonts = legacyV2Report.fontReadiness.map((entry: any) => ({
      requestKey: entry.requestKey,
      requestedFamily: entry.requestedFamily,
      status: entry.available ? 'unverified' : 'missing',
      source: 'browser',
    }));
    delete legacyV2Report.fontReadiness;
    legacyV2Report.fontIdentity = {
      schemaVersion: 1,
      digest: result.renderReport.fontIdentity.resolutionDigest,
      verification: 'browserObserved',
    };
    expect(schemaErrors(reportSchemaV2, reportSchemaV2, legacyV2Report)).toEqual([]);
    expect(schemaErrors(reportSchemaV1, reportSchemaV1, legacyV2Report)).not.toEqual([]);
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, legacyV2Report)).not.toEqual([]);
    const legacyV1Report = structuredClone(legacyV2Report) as any;
    legacyV1Report.schema = 'https://docxodus.dev/schemas/render/render-report/v1';
    legacyV1Report.schemaVersion = 1;
    legacyV1Report.readiness.forEach((entry: any) => { delete entry.diagnostics; });
    legacyV1Report.resources.forEach((entry: any) => {
      delete entry.readiness;
      delete entry.contentKey;
      delete entry.anchorId;
      delete entry.message;
    });
    legacyV1Report.fonts.forEach((entry: any) => { delete entry.requestKey; });
    expect(schemaErrors(reportSchemaV1, reportSchemaV1, legacyV1Report)).toEqual([]);
    expect(schemaErrors(reportSchemaV2, reportSchemaV2, legacyV1Report)).not.toEqual([]);
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, {
      ...result.renderReport,
      schemaVersion: 2,
    })).not.toEqual([]);
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
        expect.objectContaining({ code: 'source_anchors_inventoried' }),
        expect.objectContaining({ code: 'note_references_inventoried' }),
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

  test('exports a converter-produced multi-page footnote losslessly and deterministically', async ({
    page,
    context,
  }) => {
    const source = generateFootnoteDocx(1, 1, 1, 600, {
      normalText: 'NORMAL-SEPARATOR-STORY',
      continuationText: 'CONTINUATION-SEPARATOR-STORY',
    });
    const first = await convert(page, source);
    const second = await convert(page, source);

    expect(first.pageCount).toBeGreaterThan(2);
    expect(second.html).toBe(first.html);
    expect(canonical(second.pageMap)).toBe(canonical(first.pageMap));
    expect(second.renderReport.bindings).toEqual(first.renderReport.bindings);
    expect(first.renderReport.pages).toEqual(first.pageMap.pages);

    const audit = await page.evaluate((html) => {
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      const ids = Array.from(parsed.querySelectorAll<HTMLElement>('[id]'), (element) => element.id);
      const noteBands = Array.from(parsed.querySelectorAll<HTMLElement>('.page-footnotes'));
      return {
        notePages: noteBands.length,
        noteText: noteBands.map((band) => band.textContent ?? '').join(' ')
          .replace(/\s+/g, ' ').trim(),
        separators: parsed.querySelectorAll(
          '.page-footnotes > [data-footnote-separator]',
        ).length,
        separatorKinds: noteBands.map((band) =>
          band.firstElementChild?.getAttribute('data-footnote-separator') ?? ''),
        separatorText: noteBands.map((band) =>
          band.firstElementChild?.textContent?.replace(/\s+/g, ' ').trim() ?? ''),
        numbers: parsed.querySelectorAll('.page-footnotes .footnote-number').length,
        uniqueIds: new Set(ids).size === ids.length,
        continuationIds: Array.from(
          parsed.querySelectorAll<HTMLElement>('.footnote-continuation'),
          (element) => element.dataset.footnoteId,
        ),
      };
    }, first.html);
    expect(audit.notePages).toBe(first.pageCount);
    expect(audit.separators).toBe(audit.notePages);
    expect(audit.separatorKinds).toEqual([
      'normal',
      ...Array.from({ length: audit.notePages - 1 }, () => 'continuation'),
    ]);
    expect(audit.separatorText[0]).toContain('NORMAL-SEPARATOR-STORY');
    expect(audit.separatorText.slice(1).every((text) =>
      text.includes('CONTINUATION-SEPARATOR-STORY')
      && !text.includes('NORMAL-SEPARATOR-STORY'))).toBe(true);
    expect(audit.numbers).toBe(1);
    expect(audit.uniqueIds).toBe(true);
    expect(audit.continuationIds.length).toBeGreaterThan(1);
    expect(audit.continuationIds.every((id) => id === '1')).toBe(true);
    for (const index of [1, 2, 150, 300, 450, 599, 600]) {
      const token = `footnote-1-1-${index}`;
      expect(audit.noteText.match(new RegExp(`\\b${token}\\b`, 'g'))).toHaveLength(1);
    }

    const fragments = first.pageMap.fragments as Array<{
      story?: string;
      anchorId?: string;
      pageNumber?: number;
      fragmentIndex?: number;
    }>;
    const paragraphFragments = fragments.filter((fragment) =>
      fragment.story === 'footnote' && fragment.anchorId?.startsWith('p:fn:'));
    expect(paragraphFragments.length).toBeGreaterThan(2);
    expect(new Set(paragraphFragments.map((fragment) => fragment.pageNumber)).size)
      .toBeGreaterThan(2);
    expect(paragraphFragments.map((fragment) => fragment.fragmentIndex))
      .toEqual(paragraphFragments.map((_, index) => index));

    const offline = await context.newPage();
    const requests: string[] = [];
    offline.on('request', (request) => requests.push(request.url()));
    await offline.setContent(first.html, { waitUntil: 'load' });
    expect(await offline.locator('.page-box').count()).toBe(first.pageCount);
    expect(requests).toEqual([]);
    await offline.close();
  });

  test('applies each review profile exactly once and proves already-applied input', async ({ page }) => {
    const source = generateTrackedRevisionDocx();
    const sourceDigest = digest(source);
    const visibleText = (html: string) => page.evaluate((value) => {
      const parsed = new DOMParser().parseFromString(value, 'text/html');
      return parsed.body.textContent?.replace(/\s+/g, ' ').trim() ?? '';
    }, html);

    const finalResult = await convert(page, source, false, {
      reviewProfile: 'final',
      commentProfile: 'hidden',
    });
    const originalResult = await convert(page, source, false, {
      reviewProfile: 'original',
      commentProfile: 'hidden',
    });
    const markupResult = await convert(page, source, false, {
      reviewProfile: 'markup',
      commentProfile: 'hidden',
    });

    expect(await visibleText(finalResult.html)).toContain('Before added after.');
    expect(await visibleText(finalResult.html)).not.toContain('removed');
    expect(await visibleText(originalResult.html)).toContain('Before removed after.');
    expect(await visibleText(originalResult.html)).not.toContain('added');
    expect(await visibleText(markupResult.html)).toContain('Before removedadded after.');
    expect(finalResult.renderReport.source.rawPackageBytesDigest).toBe(sourceDigest);
    expect(originalResult.renderReport.source.rawPackageBytesDigest).toBe(sourceDigest);
    expect(markupResult.renderReport.source.rawPackageBytesDigest).toBe(sourceDigest);
    expect(finalResult.renderReport.derivedProfileSource?.rawPackageBytesDigest).not.toBe(sourceDigest);
    expect(originalResult.renderReport.derivedProfileSource?.rawPackageBytesDigest).not.toBe(sourceDigest);
    expect(markupResult.renderReport.derivedProfileSource).toBeUndefined();

    const noRevisions = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const alreadyApplied = await convert(page, noRevisions, false, {
      reviewProfile: 'final',
      reviewProfileAlreadyApplied: true,
    });
    expect(alreadyApplied.renderReport.options.reviewProfileAlreadyApplied).toBe(true);
    expect(alreadyApplied.renderReport.derivedProfileSource).toBeUndefined();

    const falseClaim = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertFailure(bytes, {
        reviewProfile: 'final',
        reviewProfileAlreadyApplied: true,
        commentProfile: 'hidden',
      }), Array.from(source));
    expect(falseClaim.code).toBe('invalid_argument');
    expect(falseClaim.phase).toBe('package_preflight');
    expect(falseClaim.report.status).toBe('failed');
    expect(falseClaim.report.derivedProfileSource).toBeUndefined();
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, falseClaim.report)).toEqual([]);
  });

  test('fails exact source identity and caller-lowered package ceilings closed', async ({ page }) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const mismatch = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertFailure(bytes, {
        reviewProfile: 'markup',
        commentProfile: 'hidden',
        expectedSourceDigest: '0'.repeat(64),
      }), Array.from(source));
    expect(mismatch.code).toBe('source_digest_mismatch');
    expect(mismatch.phase).toBe('package_preflight');
    expect(mismatch.report.source.rawPackageBytesDigest).toBe(digest(source));
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, mismatch.report)).toEqual([]);

    const limited = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertFailure(bytes, {
        reviewProfile: 'markup',
        commentProfile: 'hidden',
        limits: { opcEntries: 1 },
      }), Array.from(source));
    expect(limited.code).toBe('resource_limit');
    expect(limited.phase).toBe('package_preflight');
    expect(limited.report.options.policy.limits.opcEntries).toBe(1);
    expect(limited.report.readiness.at(-1).status).toBe('failed');
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, limited.report)).toEqual([]);
  });

  test('honors AbortSignal and always removes its isolated render realm', async ({ page }) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const outcome = await page.evaluate(async (bytes) => {
      const api = (window as any).DocxodusStandalone;
      const controller = new AbortController();
      controller.abort();
      const before = document.querySelectorAll('iframe[data-docxodus-export-realm]').length;
      const failure = await api.convertFailure(bytes, {
        reviewProfile: 'markup',
        commentProfile: 'hidden',
        signal: controller.signal,
      });
      await new Promise((resolve) => setTimeout(resolve, 0));
      return {
        failure,
        before,
        after: document.querySelectorAll('iframe[data-docxodus-export-realm]').length,
      };
    }, Array.from(source));
    expect(outcome.failure.code).toBe('operation_cancelled');
    expect(outcome.failure.phase).toBe('input_validation');
    expect(outcome.after).toBe(outcome.before);
  });

  test('requires pristine attempts to agree on font outcomes as well as page markup', async ({ page }) => {
    const source = new Uint8Array(readFileSync(join(testFiles, 'CA', 'CA001-Plain.docx')));
    const outcome = await page.evaluate(async (bytes) =>
      (window as any).DocxodusStandalone.convertWithAlternatingFontOutcomes(bytes, {
        reviewProfile: 'final',
        commentProfile: 'hidden',
        timeoutMs: 15_000,
      }), Array.from(source));
    expect(outcome.contextCount).toBeGreaterThanOrEqual(4);
    expect(outcome.result.renderReport.warnings).toContainEqual(expect.objectContaining({
      code: 'page_tree_retry',
      phase: 'page_tree_stability',
    }));
    expect(outcome.result.renderReport.fonts.length).toBeGreaterThan(0);
    expect(outcome.result.renderReport.fonts.every((font: any) =>
      font.status === 'missing'
      && /^font-[0-9]{4,}$/.test(font.requestId)
      && /^[0-9a-f]{64}$/.test(font.sampleDigest))).toBe(true);
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
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, strictFailure.report)).toEqual([]);
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
    expect(failure.report.readiness.slice(0, -1).every(
      (entry: { status: string }) => entry.status === 'complete',
    )).toBe(true);
    expect(failure.report.readiness.at(-1).status).toBe('failed');
    expect(new Set(failure.report.unavailable.map(
      (entry: { field: string }) => entry.field,
    )).size).toBe(failure.report.unavailable.length);
    expect(failure.report.unavailable).toContainEqual(expect.objectContaining({
      field: 'bindings.pdfDigest',
      reasonCode: 'notRequested',
    }));
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, failure.report)).toEqual([]);
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
    expect(failure.report.readiness.at(-1).status).toBe('failed');
    expect(schemaErrors(reportSchemaV3, reportSchemaV3, failure.report)).toEqual([]);
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
      table{border-collapse:collapse}td{height:20pt}.body{height:20pt;margin:0}
      .page-footnotes{font:10pt/10pt Arial}.page-footnotes hr{height:1px;margin:0 0 3pt;border:0}
      .footnote-item,.footnote-content p,.footnote-continuation p{margin:0}
    </style><div id="staging">
      <div id="pagination-footnote-registry"><div class="footnote-item" data-footnote-id="realm-note">
        <span class="footnote-number">1</span><span class="footnote-content"><p>${
          Array.from({ length: 240 }, (_, index) => `realm-${index}`).join(' ')
        }</p></span></div></div>
      <section data-section-index="0" data-page-width="612"
      data-page-height="792" data-content-width="468" data-content-height="648"
      data-margin-top="72" data-margin-right="72" data-margin-bottom="72" data-margin-left="72">
      <p class="body">foreign realm note <sup data-footnote-id="realm-note">1</sup></p>
      <div><table><tbody>${'<tr><td>row</td></tr>'.repeat(45)}</tbody></table></div>
    </section></div><div id="pages" class="page-container"></div>`;
    document.body.appendChild(frame);
    await loaded;
    try {
      const foreign = frame.contentDocument!;
      const hostCreateElement = document.createElement;
      document.createElement = (() => {
        throw new Error('host document.createElement must not be used');
      }) as typeof document.createElement;
      try {
        const engine = new (window as any).DocxodusPagination.PaginationEngine(
          foreign.getElementById('staging'),
          foreign.getElementById('pages'),
          { scale: 0.8, showPageNumbers: false, pageGap: 0 },
        );
        const pagination = engine.paginate();
        let secondCall = '';
        try { engine.paginate(); } catch (error) { secondCall = String(error); }
        const first = foreign.querySelector<HTMLElement>('.page-box')!;
        const noteText = Array.from(foreign.querySelectorAll<HTMLElement>('.page-footnotes'))
          .map((band) => band.textContent ?? '').join(' ').replace(/\s+/g, ' ').trim();
        return {
          pages: pagination.totalPages,
          readiness: pagination.readiness,
          width: first.getBoundingClientRect().width,
          authoredWidth: first.style.width,
          zoom: first.style.zoom,
          transform: first.style.transform,
          secondCall,
          notePages: foreign.querySelectorAll('.page-footnotes').length,
          noteText,
        };
      } finally {
        document.createElement = hostCreateElement;
      }
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
    expect.objectContaining({ code: 'source_anchors_inventoried' }),
    expect.objectContaining({ code: 'note_references_inventoried' }),
  ]));
  expect(result.authoredWidth).toBe('612pt');
  expect(result.width).toBeGreaterThan(0);
  expect(Number(result.zoom)).toBeCloseTo(0.8, 5);
  expect(result.transform).toBe('');
  expect(result.secondCall).toContain('one-shot');
  expect(result.notePages).toBeGreaterThan(1);
  expect(result.noteText).toContain('realm-0');
  expect(result.noteText).toContain('realm-239');
});
