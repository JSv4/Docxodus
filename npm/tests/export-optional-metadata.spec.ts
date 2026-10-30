import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { optionalMetadataDocx } from './docx-optional-metadata-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function text(bytes: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent();
      pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
    }
    return pages.join(' ').replace(/\s+/g, ' ');
  } finally { await task.destroy(); }
}

for (const kind of ['broken-properties', 'missing-thumbnail', 'orphan-xml'] as const) {
  for (const reviewProfile of ['markup', 'final', 'original'] as const) {
    test(`${kind} preserves ${reviewProfile} content and reports the optional defect`, async ({ browser }) => {
      test.setTimeout(120_000);
      const source = optionalMetadataDocx(kind), before = new Uint8Array(source);
      const options = { browser, outputs: ['pdf'] as const, reviewProfile, commentProfile: 'hidden' as const };
      const result = await renderDocxArtifacts(source, options);
      const control = await renderDocxArtifacts(optionalMetadataDocx('clean'), options);
      const actual = await text(result.pdf!);
      expect(actual).toBe(await text(control.pdf!));
      expect(actual).toContain('Quartz receipt: twelve gauges arrived.');
      expect(actual).toContain('QuartzCache 12.');
      expect(actual.includes('Old rack.')).toBe(reviewProfile !== 'final');
      expect(actual.includes('New rack.')).toBe(reviewProfile !== 'original');
      expect(source).toEqual(before);
      expect(result.renderReport.warnings.some(warning => warning.code ===
        (kind === 'missing-thumbnail' ? 'missing_target' : 'malformed_xml'))).toBe(true);
    });
  }

  test(`${kind} remains an invalid package under strict policy`, async ({ browser }) => {
    await expect(renderDocxArtifacts(optionalMetadataDocx(kind), {
      browser, outputs: ['pdf'], reviewProfile: 'markup', commentProfile: 'hidden', unsupportedContent: 'strict',
    })).rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight' });
  });
}

for (const kind of ['broken-styles', 'missing-image', 'properties-used-as-styles',
  'document-thumbnail', 'thumbnail-used-as-image', 'broken-relationships'] as const) {
  test(`${kind} still fails preflight`, async ({ browser }) => {
    await expect(renderDocxArtifacts(optionalMetadataDocx(kind), {
      browser, outputs: ['pdf'], reviewProfile: 'markup', commentProfile: 'hidden',
    })).rejects.toMatchObject({ code: 'invalid_document', phase: 'package_preflight' });
  });
}
