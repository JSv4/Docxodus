import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// DB012 contains real numbered lists AND plain paragraphs / headings, so we can
// exercise both the list-membership path and the "non-list paragraph -> null"
// path against the same fixture without needing a bespoke test doc.
const FIXTURE = 'DB012-Lists-With-Different-Numberings.docx';

test.describe('block-metadata (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('getBlockMetadata returns kind+scope for body p/h/li anchors', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        const firstBlock = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .find(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));
        if (!firstBlock) throw new Error('no body block anchor found in fixture');

        const meta = JSON.parse(bridge.GetBlockMetadata(h, firstBlock.id));
        return { firstBlock, meta };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    expect(result.meta).not.toBeNull();
    expect(result.meta.anchorId).toBe(result.firstBlock.id);
    expect(result.meta.kind).toBe(result.firstBlock.kind);
    expect(result.meta.scope).toBe('body');
    expect(typeof result.meta.hasInlineFormatting).toBe('boolean');
  });

  test('getBlockMetadata returns null for unknown anchor', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const unknown = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(bin, '');
      try {
        return JSON.parse(bridge.GetBlockMetadata(h, 'p:body:does-not-exist'));
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    // Unknown anchors must surface as JSON null, not throw.
    expect(unknown).toBeNull();
  });

  test('getBlockMetadatas dedupes ids and maps unknown ids to null', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        const blocks = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));
        if (blocks.length === 0) throw new Error('no body block anchors found');

        const knownId = blocks[0].id;
        // Pass duplicates AND an unknown id; expect dedupe + explicit null.
        const map = JSON.parse(bridge.GetBlockMetadatas(
          h,
          JSON.stringify([knownId, knownId, 'p:body:unknown']),
        ));

        return {
          knownId,
          keyCount: Object.keys(map).length,
          known: map[knownId],
          unknown: map['p:body:unknown'],
          keys: Object.keys(map).sort(),
        };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    // Dedupe: two copies of knownId + one unknown -> exactly 2 keys.
    expect(result.keyCount).toBe(2);
    expect(result.keys).toEqual([result.knownId, 'p:body:unknown'].sort());
    expect(result.known).not.toBeNull();
    expect(result.known.anchorId).toBe(result.knownId);
    expect(result.unknown).toBeNull();
  });

  test('getListMembership: null for plain paragraphs, populated for list items', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        const blocks = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));

        // Bulk-read metadata for every body block so we can find a list item and
        // (independently) a non-list paragraph without relying on document order.
        const metaMap = JSON.parse(bridge.GetBlockMetadatas(
          h, JSON.stringify(blocks.map(b => b.id)),
        ));

        const listBlock = blocks.find(b => metaMap[b.id]?.list != null);
        const nonListBlock = blocks.find(b => metaMap[b.id] && metaMap[b.id].list == null);

        const listMembership = listBlock
          ? JSON.parse(bridge.GetListMembership(h, listBlock.id))
          : null;
        const nonListMembership = nonListBlock
          ? JSON.parse(bridge.GetListMembership(h, nonListBlock.id))
          : null;

        return {
          listBlockId: listBlock?.id ?? null,
          nonListBlockId: nonListBlock?.id ?? null,
          listMembership,
          nonListMembership,
        };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    // DB012 is the lists fixture — it MUST have at least one list-item anchor.
    expect(result.listBlockId, 'expected at least one list-item anchor in DB012').not.toBeNull();
    expect(result.listMembership).not.toBeNull();
    expect(result.listMembership.anchorId).toBe(result.listBlockId);
    expect(typeof result.listMembership.numId).toBe('number');
    expect(typeof result.listMembership.abstractNumId).toBe('number');
    expect(typeof result.listMembership.level).toBe('number');
    expect(typeof result.listMembership.format).toBe('string');
    expect(typeof result.listMembership.start).toBe('number');
    expect(result.listMembership.isAutoNumbered).toBe(true);

    // If a non-list paragraph exists, its membership MUST be null. If the fixture
    // is 100% list items the assertion is skipped — but we still verified the
    // list-side contract above.
    if (result.nonListBlockId !== null) {
      expect(result.nonListMembership).toBeNull();
    }
  });

  test('getSectionInfo returns page geometry for body anchors', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        const firstBlock = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .find(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));
        if (!firstBlock) throw new Error('no body block anchor found in fixture');

        const info = JSON.parse(bridge.GetSectionInfo(h, firstBlock.id));
        return { info };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    expect(result.info).not.toBeNull();
    expect(typeof result.info.anchorId).toBe('string');
    expect(typeof result.info.sectionUnid).toBe('string');
    expect(result.info.pageWidthTwips).toBeGreaterThan(0);
    expect(result.info.pageHeightTwips).toBeGreaterThan(0);
    expect(result.info.columns).toBeGreaterThanOrEqual(1);
    expect(typeof result.info.landscape).toBe('boolean');
    expect(Array.isArray(result.info.headerPartUris)).toBe(true);
    expect(Array.isArray(result.info.footerPartUris)).toBe(true);
  });

  test('style catalog and direct/effective spans feed returned ids into mutations', async ({ page }) => {
    const bytes = readTestFile(FIXTURE);

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const h = bridge.OpenSession(new Uint8Array(bytesArray), '');
      try {
        const projection = JSON.parse(bridge.Project(h));
        const anchor = (Object.entries(projection.anchorIndex) as [string, any][])
          .map(([id, value]) => ({ id, ...value }))
          .find(value => value.scope === 'body' && ['p', 'h', 'li'].includes(value.kind)
            && value.textPreview.length > 0)?.id;
        if (!anchor) throw new Error('fixture has no text-bearing body paragraph');

        const styles = JSON.parse(bridge.ListStyles(h));
        const paragraphStyle = styles.find((style: any) => style.type === 'paragraph' && style.id === 'Normal')
          ?? styles.find((style: any) => style.type === 'paragraph' && !/^Heading[1-9]$/.test(style.id));
        if (!paragraphStyle) throw new Error('fixture has no paragraph style');
        const styleMutation = JSON.parse(bridge.SetParagraphStyle(h, anchor, paragraphStyle.id));
        const mutationAnchor = styleMutation.modified?.[0]?.id ?? anchor;

        const formatting = JSON.parse(bridge.GetFormatting(h, mutationAnchor));
        const spans = JSON.parse(bridge.ListInlineSpans(h, mutationAnchor));
        const span = spans[0];
        const spanMutation = JSON.parse(bridge.ApplyFormat(
          h, span.anchorId, JSON.stringify(span.span), JSON.stringify({ bold: true }),
        ));
        const after = JSON.parse(bridge.GetFormatting(h, mutationAnchor));

        return { anchor: mutationAnchor, styles, paragraphStyle, styleMutation, formatting, span, spanMutation, after };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    expect(result.styles.length).toBeGreaterThan(0);
    expect(result.paragraphStyle.id).toBeTruthy();
    expect(result.styleMutation.success).toBe(true);
    expect(result.formatting.anchorId).toBe(result.anchor);
    expect(result.formatting.directParagraph).toBeDefined();
    expect(result.formatting.effectiveParagraph).toBeDefined();
    expect(result.span.anchorId).toBe(result.anchor);
    expect(result.span.span.length).toBeGreaterThan(0);
    expect(result.span.direct).toBeDefined();
    expect(result.span.effective).toBeDefined();
    expect(result.spanMutation.success).toBe(true);
    expect(result.after.runs[0].direct.bold).toBe(true);
  });
});
