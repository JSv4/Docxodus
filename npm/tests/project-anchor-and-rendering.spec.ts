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

// AnchorIdRendering enum mirror (kept inline for in-page eval — types.ts can't be
// imported into page.evaluate code, only into the spec scope).
const AnchorIdRendering = {
  FullUnid: 0,
  Abbreviated: 1,
  Sequential: 2,
} as const;

// ProjectionDepth enum mirror.
const ProjectionDepth = {
  SelfOnly: 0,
  Subtree: 1,
  SubtreeAndFollowingSiblings: 2,
} as const;

// HC031 has 10 body headings (Heading1 + Heading2 styles) and dozens of body
// paragraphs — gives us multiple heading anchors (so "section ends at the next
// heading" is observable) and far more than two body paragraphs (so Sequential
// :1 and :2 both appear in the bucket counter).
const TEST_DOC = 'HC031-Complicated-Document.docx';

test.describe('ProjectAnchor + AnchorIdRendering (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Abbreviated rendering shrinks anchor tokens', async ({ page }) => {
    const bytes = readTestFile(TEST_DOC);

    const result = await page.evaluate(async (args: { bytesArray: number[]; abbr: number; full: number }) => {
      const bin = new Uint8Array(args.bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;

      // Two sessions — one Full, one Abbreviated — so we can compare the rendered markdown.
      const fullHandle = bridge.OpenSession(bin, JSON.stringify({
        projectionSettings: { anchorIdRendering: args.full },
      }));
      const abbrHandle = bridge.OpenSession(bin, JSON.stringify({
        projectionSettings: { anchorIdRendering: args.abbr },
      }));
      try {
        const fullProj = JSON.parse(bridge.Project(fullHandle));
        const abbrProj = JSON.parse(bridge.Project(abbrHandle));

        // Full Unid tokens look like {#p:body:<32 hex chars>}; Abbreviated uses a
        // shorter prefix per (kind, scope) bucket with a 4-char floor.
        const fullTokenMatch = /\{#[a-z]+:body:([0-9a-f]{32})\}/.exec(fullProj.markdown);
        const abbrTokenMatch = /\{#[a-z]+:body:([0-9a-f]{4,32})\}/.exec(abbrProj.markdown);
        const fullUnidLen = fullTokenMatch?.[1].length ?? -1;
        const abbrUnidLen = abbrTokenMatch?.[1].length ?? -1;

        return {
          fullMdLen: fullProj.markdown.length,
          abbrMdLen: abbrProj.markdown.length,
          fullUnidLen,
          abbrUnidLen,
          // Sanity — anchor indices should be non-empty and the abbreviated index
          // is dual-keyed so it has at least as many entries as the full one.
          fullEntries: Object.keys(fullProj.anchorIndex).length,
          abbrEntries: Object.keys(abbrProj.anchorIndex).length,
        };
      } finally {
        bridge.CloseSession(fullHandle);
        bridge.CloseSession(abbrHandle);
      }
    }, { bytesArray: Array.from(bytes), abbr: AnchorIdRendering.Abbreviated, full: AnchorIdRendering.FullUnid });

    // Full Unid is always 32 hex chars; abbreviated must be strictly shorter for
    // any non-trivial document (HC001 has plenty of anchors per bucket).
    expect(result.fullUnidLen).toBe(32);
    expect(result.abbrUnidLen).toBeGreaterThanOrEqual(4);
    expect(result.abbrUnidLen).toBeLessThan(32);
    // The whole markdown should therefore be shorter.
    expect(result.abbrMdLen).toBeLessThan(result.fullMdLen);
    expect(result.fullEntries).toBeGreaterThan(0);
    // Dual-keyed index in Abbreviated mode → at least as many entries as full.
    expect(result.abbrEntries).toBeGreaterThanOrEqual(result.fullEntries);
  });

  test('Sequential rendering produces 1-based numeric ids per bucket', async ({ page }) => {
    const bytes = readTestFile(TEST_DOC);

    const result = await page.evaluate(async (args: { bytesArray: number[]; seq: number }) => {
      const bin = new Uint8Array(args.bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, JSON.stringify({
        projectionSettings: { anchorIdRendering: args.seq },
      }));
      try {
        const proj = JSON.parse(bridge.Project(handle));
        return {
          markdownExcerpt: proj.markdown.substring(0, 800),
          markdown: proj.markdown,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: Array.from(bytes), seq: AnchorIdRendering.Sequential });

    // Sequential tokens: {#<kind>:body:<digits>} with no hex Unid. We expect at
    // least body:1 and body:2 (heading or paragraph bucket — HC001 has ≥14 of each).
    expect(result.markdown).toMatch(/\{#[a-z]+:body:1\}/);
    expect(result.markdown).toMatch(/\{#[a-z]+:body:2\}/);
    // And critically — NO 32-char hex Unid tokens should appear (everything is numeric).
    expect(result.markdown).not.toMatch(/\{#[a-z]+:body:[0-9a-f]{32}\}/);
  });

  test('projectAnchor on a heading returns the section (heading + content up to next same-or-higher heading)', async ({ page }) => {
    const bytes = readTestFile(TEST_DOC);

    const result = await page.evaluate(async (args: { bytesArray: number[]; depth: number }) => {
      const bin = new Uint8Array(args.bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const full = JSON.parse(bridge.Project(handle));
        // Walk the markdown to extract every heading token paired with its
        // markdown level (`#` count immediately after the token). The projection
        // emits one heading per `{#h:body:UNID} ##... <text>` line, so we can
        // reconstruct the H1/H2/... hierarchy without re-parsing the OOXML.
        const re = /\{#(h:body:[0-9a-f]{32})\}\s+(#+)\s/g;
        const headings: { id: string; level: number; pos: number }[] = [];
        let m: RegExpExecArray | null;
        while ((m = re.exec(full.markdown)) !== null) {
          headings.push({ id: m[1], level: m[2].length, pos: m.index });
        }
        if (headings.length < 2) {
          return { error: `fixture has only ${headings.length} body headings — need ≥2 with hierarchy` };
        }
        // Find the first heading that has a strictly-later heading at the same
        // OR higher (smaller `level` number) level — its section is bounded.
        let firstIdx = -1;
        let boundaryIdx = -1;
        for (let i = 0; i < headings.length; i++) {
          for (let j = i + 1; j < headings.length; j++) {
            if (headings[j].level <= headings[i].level) {
              firstIdx = i; boundaryIdx = j; break;
            }
          }
          if (firstIdx >= 0) break;
        }
        if (firstIdx < 0) {
          return { error: 'no heading in fixture has a same-or-higher-level successor' };
        }
        const first = headings[firstIdx];
        const boundary = headings[boundaryIdx];

        const slice = JSON.parse(bridge.ProjectAnchor(handle, first.id, args.depth));
        return {
          firstHeadingId: first.id,
          firstLevel: first.level,
          boundaryHeadingId: boundary.id,
          boundaryLevel: boundary.level,
          fullMdLen: full.markdown.length,
          sliceMdLen: slice.markdown.length,
          // Heading token for the target must appear; the boundary heading token
          // must NOT appear (section ends before the next same-or-higher heading).
          sliceContainsFirst: slice.markdown.includes('{#' + first.id + '}'),
          sliceContainsBoundary: slice.markdown.includes('{#' + boundary.id + '}'),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: Array.from(bytes), depth: ProjectionDepth.SubtreeAndFollowingSiblings });

    expect(result.error).toBeUndefined();
    expect(result.sliceMdLen).toBeGreaterThan(0);
    expect(result.sliceMdLen).toBeLessThan(result.fullMdLen!);
    expect(result.sliceContainsFirst).toBe(true);
    expect(result.sliceContainsBoundary).toBe(false);
  });

  test('projectAnchor with SelfOnly returns just one block', async ({ page }) => {
    const bytes = readTestFile(TEST_DOC);

    const result = await page.evaluate(async (args: { bytesArray: number[]; selfOnly: number }) => {
      const bin = new Uint8Array(args.bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const full = JSON.parse(bridge.Project(handle));
        // First body paragraph by document order — same lookup pattern as above.
        const paraEntries = Object.entries(full.anchorIndex)
          .filter(([_id, t]: [string, any]) => t.kind === 'p' && t.scope === 'body')
          .map(([id, t]: [string, any]) => ({ id, t, idx: full.markdown.indexOf('{#' + id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a, b) => a.idx - b.idx);
        if (paraEntries.length === 0) {
          return { error: 'fixture has no body paragraphs' };
        }
        const target = paraEntries[0];
        const slice = JSON.parse(bridge.ProjectAnchor(handle, target.id, args.selfOnly));
        return {
          targetId: target.id,
          fullMdLen: full.markdown.length,
          sliceMdLen: slice.markdown.length,
          markdown: slice.markdown,
          // Number of distinct anchor-bearing block tokens in the slice. SelfOnly
          // on a paragraph anchor must include exactly that paragraph — no
          // siblings, no following content.
          anchorTokenCount: (slice.markdown.match(/\{#[a-z]+:body:[0-9a-f]{32}\}/g) ?? []).length,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, { bytesArray: Array.from(bytes), selfOnly: ProjectionDepth.SelfOnly });

    expect(result.error).toBeUndefined();
    expect(result.sliceMdLen).toBeGreaterThan(0);
    expect(result.sliceMdLen).toBeLessThan(result.fullMdLen!);
    // SelfOnly on a paragraph → exactly one anchor token in the slice.
    expect(result.anchorTokenCount).toBe(1);
    // And that single token must address the requested paragraph.
    expect(result.markdown).toContain('{#' + result.targetId + '}');
  });
});
