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

test.describe('anchor-introspection (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('session projection + getAnchorInfo(s) surface textPreview, fn boilerplate filtered', async ({ page }) => {
    // HC031 is the "Complicated Document" fixture: real body content + one user
    // footnote ("This is a footnote.") + the two Word-reserved separator
    // footnotes (type="separator" and type="continuationSeparator"). Lets us
    // exercise the body-anchor textPreview path AND the fn boilerplate filter
    // introduced by Unit B in a single end-to-end run.
    //
    // Anchor ids derive from PtOpenXml.Unid attributes — those are assigned
    // lazily and the unid values are random Guids, so two separate sessions
    // (or a session vs. a standalone ConvertWmlToMarkdown call) over the same
    // raw bytes mint *different* anchor ids. To keep ids stable across the
    // assertions below, every lookup goes through the *same* session handle.
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const entries = Object.entries(proj.anchorIndex) as [string, any][];

        // Sort body p/h/li anchors in markdown document order so we can pick
        // stable id1/id2 picks for the bulk-lookup assertion.
        const bodyBlocks = entries
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));
        const orderedBody = bodyBlocks
          .map(t => ({ t, idx: proj.markdown.indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a, b) => a.idx - b.idx)
          .map(x => x.t);

        // (1) Bulk lookup textPreview for every body block, then assert that the
        //     *majority* carry non-empty text. (Empty paragraphs — spacing,
        //     section breaks — legitimately produce empty previews per the
        //     `ComputeTextPreview` contract documented on `AnchorTarget`.)
        const bodyIds = bodyBlocks.map(t => t.id);
        const bodyInfoMap = JSON.parse(bridge.GetAnchorInfos(handle, JSON.stringify(bodyIds)));
        const bodyPreviews = bodyIds
          .map(id => bodyInfoMap[id]?.textPreview)
          .filter((p): p is string => typeof p === 'string');
        const nonEmptyBodyPreviews = bodyPreviews.filter(p => p.length > 0);

        const id1 = orderedBody[0]?.id;
        const id2 = orderedBody[1]?.id;
        const preview1FromBulk = bodyInfoMap[id1]?.textPreview;

        // (2) Single-anchor lookup returns the same shape and same preview as
        //     the bulk lookup. No drift between the two surfaces.
        const singleInfo = JSON.parse(bridge.GetAnchorInfo(handle, id1));

        // (3) Bulk shape: explicit null for unknown ids, real shape for known ids.
        const bulk = JSON.parse(bridge.GetAnchorInfos(
          handle,
          JSON.stringify([id1, id2, 'unknown-anchor-id-xyz']),
        ));

        // (5) Footnote-scope anchors: any fn-scope p/h/li entries must be REAL
        //     footnotes — Unit B filters out the boilerplate separator pair
        //     (w:type="separator" and "continuationSeparator") that Word inserts
        //     in every doc with a FootnotesPart. HC031 ships exactly one real
        //     footnote (" This is a footnote."), so we expect at least one
        //     fn-scope paragraph whose textPreview is non-empty and recognizable.
        const fnEntries = entries
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'fn');
        const fnBlockIds = fnEntries
          .filter(t => ['p', 'h', 'li'].includes(t.kind))
          .map(t => t.id);
        const fnInfoMap = JSON.parse(bridge.GetAnchorInfos(handle, JSON.stringify(fnBlockIds)));
        const fnPreviews = fnBlockIds
          .map(id => fnInfoMap[id]?.textPreview)
          .filter((p): p is string => typeof p === 'string' && p.length > 0);

        return {
          bodyBlockCount: bodyBlocks.length,
          bodyPreviewSamples: bodyPreviews.slice(0, 5),
          nonEmptyBodyCount: nonEmptyBodyPreviews.length,
          id1,
          id2,
          preview1FromBulk,
          singleInfo,
          bulkKeys: Object.keys(bulk).sort(),
          bulkPreview1: bulk[id1]?.textPreview,
          bulkPreview2: bulk[id2]?.textPreview,
          bulkUnknown: bulk['unknown-anchor-id-xyz'],
          fnEntryCount: fnEntries.length,
          fnPreviews,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    // (1) Body p/h/li anchors carry textPreview when the underlying paragraph
    //     has text. Most body anchors should be non-empty — and at least one of
    //     the sampled previews must contain real content, proving textPreview
    //     wired up end-to-end and isn't returning empty strings everywhere.
    expect(result.bodyBlockCount).toBeGreaterThan(0);
    expect(
      result.nonEmptyBodyCount,
      `bodyPreviewSamples=${JSON.stringify(result.bodyPreviewSamples)}`,
    ).toBeGreaterThan(result.bodyBlockCount / 2);
    expect(
      result.bodyPreviewSamples.some(p => typeof p === 'string' && p.length > 0),
      `bodyPreviewSamples=${JSON.stringify(result.bodyPreviewSamples)}`,
    ).toBe(true);

    // (2) Single-anchor lookup returns the same shape (id/kind/scope/textPreview)
    //     and the same preview as the bulk lookup for the same id.
    expect(result.singleInfo).not.toBeNull();
    expect(result.singleInfo.id).toBe(result.id1);
    expect(['body', 'hdr', 'ftr', 'fn', 'en']).toContain(result.singleInfo.scope);
    expect(['p', 'h', 'li', 'tc', 'tbl', 'tr', 'fn', 'en']).toContain(result.singleInfo.kind);
    expect(result.singleInfo.textPreview).toBe(result.preview1FromBulk);
    expect(typeof result.singleInfo.textPreview).toBe('string');

    // (3) Bulk lookup returns a record keyed by every requested id, with the
    //     same previews as single lookup, and explicit null for unknown ids.
    expect(result.bulkKeys).toEqual([result.id1, result.id2, 'unknown-anchor-id-xyz'].sort());
    expect(result.bulkPreview1).toBe(result.preview1FromBulk);
    expect(typeof result.bulkPreview2).toBe('string');
    expect(result.bulkUnknown).toBeNull();

    // (5) The separator pair (w:type="separator" and "continuationSeparator")
    //     must not leak into the AnchorIndex. HC031 has one real footnote, so
    //     we expect at least one fn-scope paragraph anchor with a non-empty
    //     preview containing the word "footnote" — the boilerplate separators
    //     contain no visible text, so this assertion uniquely fails if they
    //     leak through.
    expect(result.fnPreviews.length).toBeGreaterThan(0);
    expect(
      result.fnPreviews.some(p => p.toLowerCase().includes('footnote')),
      `fnPreviews=${JSON.stringify(result.fnPreviews)}`,
    ).toBe(true);
  });

  test('session Project() serializer emits textPreview on anchorIndex entries', async ({ page }) => {
    // Follow-up regression: DocxSessionBridge.SerializeProjection (the JSExport
    // backing session.project() / bridge.Project(handle)) was originally a third
    // serialization path that the Unit D textPreview rollout missed. This test
    // pins the fix — every body p/h/li entry in proj.anchorIndex must carry a
    // string `textPreview`, and at least one must be non-empty for HC031 which
    // has real body content.
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const entries = Object.entries(proj.anchorIndex) as [string, any][];
        const bodyBlocks = entries
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind));
        const allHaveStringPreview = bodyBlocks.every(
          t => typeof t.textPreview === 'string',
        );
        const nonEmptyPreviews = bodyBlocks
          .map(t => t.textPreview)
          .filter((p: string) => typeof p === 'string' && p.length > 0);
        return {
          bodyBlockCount: bodyBlocks.length,
          allHaveStringPreview,
          nonEmptyPreviewCount: nonEmptyPreviews.length,
          sampleEntry: bodyBlocks[0] ?? null,
          previewSamples: nonEmptyPreviews.slice(0, 3),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.bodyBlockCount).toBeGreaterThan(0);
    expect(
      result.allHaveStringPreview,
      `sampleEntry=${JSON.stringify(result.sampleEntry)}`,
    ).toBe(true);
    expect(
      result.nonEmptyPreviewCount,
      `previewSamples=${JSON.stringify(result.previewSamples)}`,
    ).toBeGreaterThan(0);
    // The sample entry must conform to the documented anchorIndex value shape:
    // partUri, unid, kind, scope, textPreview — all strings.
    expect(result.sampleEntry).not.toBeNull();
    expect(typeof result.sampleEntry!.partUri).toBe('string');
    expect(typeof result.sampleEntry!.unid).toBe('string');
    expect(typeof result.sampleEntry!.kind).toBe('string');
    expect(typeof result.sampleEntry!.scope).toBe('string');
    expect(typeof result.sampleEntry!.textPreview).toBe('string');
  });
});
