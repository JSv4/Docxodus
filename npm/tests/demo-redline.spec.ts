import { test, expect, Page } from '@playwright/test';

// docs/demo/redline.html — REDLINE THEATER (copied into the webroot as
// demo-redline.html by pretest). In production it imports the pinned CDN engine;
// the show (`redline-theater.js`) and the browser MCP endpoint (`mcp-wire.js`)
// are demo content living beside the page. `?engine=./embed.bundle.js` retargets
// only the library, which is how this spec drives the page locally.
//
// The pure logic (frame shapes, binding substitution, script validation,
// telemetry, and the contract check against the real ToolCatalog.cs) is exercised
// headlessly by docs/demo/tools/redline-theater.test.mjs. This spec guards what
// only a real browser proves, and the claims the page makes on screen:
//
//   - the baseline is authored CLEAN (a show that starts with revisions in the
//     document would attribute someone else's marks to the counsel),
//   - every scripted MCP call succeeds against the live WASM session,
//   - the edits land as NATIVE tracked changes attributed to all three counsel,
//   - the redline is provably reversible — reject-all reproduces the baseline,
//   - and the saved bytes are a real DOCX carrying that markup.
//
// It also measures the thing the demo asserts about itself: per-call latency, so
// a regression that makes the show stutter fails here rather than on stage.

const OVERRIDE = 'engine=./embed.bundle.js';

async function bootTheater(page: Page, extra = '') {
  await page.goto(`/demo-redline.html?${OVERRIDE}${extra}`);
  await page.waitForFunction(
    () =>
      (window as any).__theater !== undefined ||
      (window as any).__theaterError !== undefined,
    { timeout: 120000 },
  );
  const err = await page.evaluate(() => (window as any).__theaterError);
  expect(err, `theater boot failed: ${err}`).toBeUndefined();
  // The stage is set once the baseline has been authored and the wire carries
  // the two set_mode frames that bracket it.
  await page.waitForFunction(() => (window as any).__theater.ready, { timeout: 60000 });
  await page.evaluate(() => (window as any).__theater.ready);
}

/** The markdown projection escapes markdown-significant characters, so a clause
 *  reads back as `one and one\\-half times \\(1.5x\\)`. Comparisons here are about
 *  the words, not the projection's escaping. */
const unescapeMarkdown = (markdown: string) => markdown.replace(/\\(.)/g, '$1');

/** Run the whole negotiation at MAX (no pacing) and wait for the finale. */
async function runToFinale(page: Page) {
  await page.evaluate(() => (window as any).__theater.setSpeed('max'));
  await page.evaluate(() => (window as any).__theater.run());
  await page.waitForFunction(() => (window as any).__theater.redline() != null, {
    timeout: 120000,
  });
}

test.describe('REDLINE THEATER', () => {
  test('sets a clean stage: the baseline carries no tracked changes', async ({ page }) => {
    await bootTheater(page);

    // The show measures every mark against this document, so it must start with
    // none of its own. A baseline authored with recording ON would silently
    // attribute the setup to the first counsel.
    const revisions = await page.evaluate(() => (window as any).__theater.stats().revisions);
    expect(revisions).toBe(0);

    // The baseline is a real document with real structure to negotiate over.
    const baselineSize = await page.evaluate(
      () => (window as any).__theater.baseline().length);
    expect(baselineSize).toBeGreaterThan(3000);

    // Both recording-mode switches went over the wire as real tool calls.
    await expect(page.locator('[data-dxt="wire"]')).toContainText(
      'docxodus_track_changes/set_mode');
    await expect(page.locator('[data-dxt="revisions"]')).toHaveText('0');
    await expect(page.locator('[data-dxt="save"]')).toBeDisabled();
  });

  test('runs the whole negotiation: every MCP call succeeds', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    // Not one frame came back isError — the driver throws on the first that does,
    // so an error would have surfaced in the panel instead of the proof.
    await expect(page.locator('[data-dxt="error"]')).not.toHaveAttribute('data-on', 'true');

    const stats = await page.evaluate(() => (window as any).__theater.stats());
    // Three acts of scripted calls plus the two staging set_mode frames.
    expect(stats.calls).toBeGreaterThanOrEqual(30);
    expect(stats.revisions).toBeGreaterThan(10);
    // Exactly one call is expected to be refused — the list-format step that
    // demonstrates the engine declining an edit it cannot encode reversibly.
    expect(stats.refusals).toBe(1);

    // Every tool the script exercises actually got called — a script that
    // silently degraded to nothing but replace_text would still "pass" above.
    const tools = stats.perTool.map((row: { tool: string }) => row.tool);
    expect(tools).toEqual(expect.arrayContaining([
      'docxodus_search', 'docxodus_edit', 'docxodus_create',
      'docxodus_format', 'docxodus_comment', 'docxodus_list',
    ]));
  });

  test('attributes the marks to all three counsel, as Word groups them', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    const authors = await page.evaluate(() => {
      const session = (window as any).__theater.session;
      const counts: Record<string, number> = {};
      for (const revision of session.listRevisions()) {
        counts[revision.author] = (counts[revision.author] ?? 0) + 1;
      }
      return counts;
    });

    // The three counsel are three values of the session's revision author, so
    // the resulting markup is genuinely per-reviewer.
    expect(Object.keys(authors).sort()).toEqual(
      ['Dana Whitfield', 'Marcus Oyelaran', 'Priya Raghunathan']);
    for (const [author, count] of Object.entries(authors)) {
      expect(count, `${author} wrote nothing`).toBeGreaterThan(0);
    }

    // And the panel shows that breakdown rather than a single blended total.
    const rows = page.locator('[data-dxt="authors"] .dxt-author');
    await expect(rows).toHaveCount(3);
    await expect(rows.first()).toContainText('Dana Whitfield');
  });

  test('proves the redline is reversible, on screen', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    // The claim the whole demo is built to make: rejecting every generated
    // revision reproduces the baseline the show started from.
    await expect(page.locator('[data-dxt="proof"]')).toHaveAttribute('data-on', 'true');
    await expect(page.locator('[data-dxt="proof"]')).toHaveAttribute('data-ok', 'true');
    await expect(page.locator('[data-dxt="proof"]')).toContainText('REVERSIBLE');
    await expect(page.locator('[data-dxt="proof"]')).toContainText('zero content differences');

    // Independently of the panel: reject-all on the saved redline must land back
    // on a document with no revisions left and the baseline's defect restored.
    const rejected = await page.evaluate(() => {
      const engine = (window as any).__theaterEngine;
      const bytes = (window as any).__theater.redline();
      const shadow = engine.openDocxSession(bytes);
      try {
        shadow.rejectAllRevisions();
        return {
          revisionsLeft: shadow.listRevisions().length,
          markdown: shadow.project().markdown as string,
        };
      } finally {
        shadow.close();
      }
    });
    expect(rejected.revisionsLeft).toBe(0);
    // "The Vendor" is the defect Act III conformed; rejecting must bring it back.
    const restored = unescapeMarkdown(rejected.markdown);
    expect(restored).toContain('The Vendor');
    expect(restored).toContain('sixty (60) days');
    expect(restored).toContain('Section 8 — Governing Law');
    // And the negotiated positions are gone from the restored text.
    expect(restored).not.toContain('one and one-half times (1.5x)');
    expect(restored).not.toContain('Data Protection');
    // The footnote is the beat that used to break this (issue #614): its citation
    // rejected away but the definition stayed in /word/footnotes.xml, so the note
    // text survived a "full" rejection. #625 made the citation the reversible unit
    // and prunes the uncited definition, so the side letter must be gone too.
    expect(restored).not.toContain('side letter of even date');
  });

  test('the footnote the proof covers is really a tracked citation', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    // Guard against the reversibility proof passing vacuously: it would come back
    // clean if the note were never written at all. So read the XML out of the
    // redline and check BOTH halves of the note record, which is what makes the
    // round trip work — the citation run inside a w:ins (#625), and the
    // definition's own content insertion-marked (#638). The second half is not
    // decoration: #625 left the definition unmarked and compensated with an
    // unconditional prune on reject, which ate baseline-owned note husks. With
    // both marked, the prune is guarded on the definition being emptied too.
    const note = await page.evaluate(() => {
      const engine = (window as any).__theaterEngine;
      const shadow = engine.openDocxSession((window as any).__theater.redline());
      try {
        const match = shadow.grep('months preceding the claim')[0];
        const defs = shadow.findByKind('p', 'fn');
        return {
          anchor: match?.enclosingAnchor?.id ?? null,
          citingXml: match ? shadow.raw.getXml(match.enclosingAnchor.id) : '',
          defXml: defs.map((a: { id: string }) => shadow.raw.getXml(a.id)),
        };
      } finally {
        shadow.close();
      }
    });
    expect(note.anchor, 'the cap clause is still findable').not.toBeNull();
    expect(note.citingXml).toContain('footnoteReference');
    const insBlocks = note.citingXml.match(/<w:ins\b[\s\S]*?<\/w:ins>/g) ?? [];
    expect(
      insBlocks.some((block: string) => block.includes('footnoteReference')),
      'the footnote reference must be wrapped in w:ins, not written as plain content',
    ).toBe(true);

    // The definition Act II authored — identified by its text, since the part also
    // holds Word's two reserved separator notes, which carry no markup and must not.
    const authored = note.defXml.filter((xml: string) => xml.includes('side letter of even date'));
    expect(authored.length, 'the side-letter note definition is in the redline').toBe(1);
    expect(
      authored[0],
      'the definition content must record too, or a stateless reject cannot tell '
      + 'this note from a baseline-owned husk the counterpart merely cites',
    ).toContain('<w:ins');
  });

  test('accepting the redline lands the negotiated terms', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    const acceptedRaw = await page.evaluate(() => {
      const engine = (window as any).__theaterEngine;
      const shadow = engine.openDocxSession((window as any).__theater.redline());
      try {
        shadow.acceptAllRevisions();
        return shadow.project().markdown as string;
      } finally {
        shadow.close();
      }
    });
    const accepted = unescapeMarkdown(acceptedRaw);

    // The end state of the negotiation, clause by clause: Supplier's counter on
    // the cap, the compromise payment terms, the new section, and the
    // defined-term sweep.
    expect(accepted).toContain('one and one-half times (1.5x)');
    expect(accepted).toContain('forty-five (45) days');
    expect(accepted).toContain('Section 8 — Data Protection');
    expect(accepted).toContain('Section 9 — Governing Law');
    expect(accepted).not.toContain('The Vendor');
    // Neither side's intermediate position survives as text.
    expect(accepted).not.toContain('two times (2x)');
    expect(accepted).not.toContain('sixty (60) days');
  });

  test('the saved file is a real DOCX carrying native tracked-change markup', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    const saved = await page.evaluate(() => {
      const bytes = (window as any).__theater.redline() as Uint8Array;
      return { length: bytes.length, zip: bytes[0] === 0x50 && bytes[1] === 0x4b };
    });
    expect(saved.zip, 'a .docx is a zip').toBe(true);
    expect(saved.length).toBeGreaterThan(3000);

    // The markup is `w:ins`/`w:del`, not a rendering of a diff — the demo's
    // central claim about what it produced.
    const families = await page.evaluate(() =>
      (window as any).__theater.session.listRevisions()
        .map((r: { family: string }) => r.family));
    expect(families.some((f: string) => f.includes('insert'))).toBe(true);
    expect(families.some((f: string) => f.includes('delete'))).toBe(true);

    await expect(page.locator('[data-dxt="save"]')).toBeEnabled();
  });

  test('the wire shows real request/response frames with measured latency', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    const wire = page.locator('[data-dxt="wire"]');
    // Requests and responses are paired and labelled the way MCP names them.
    await expect(wire.locator('.dxt-frame.req').first()).toContainText('docxodus_');
    expect(await wire.locator('.dxt-frame.req').count()).toBeGreaterThan(20);
    expect(await wire.locator('.dxt-frame.res').count()).toBeGreaterThan(20);
    // No response frame rendered as a fault. The one refusal gets its own
    // colour, because the engine declining an irreversible edit is not an error.
    expect(await wire.locator('.dxt-frame.res.err').count()).toBe(0);
    expect(await wire.locator('.dxt-frame.res.refused').count()).toBe(1);
    await expect(wire.locator('.dxt-frame.res').last()).toContainText('ms');

    // The HUD reports what was actually measured, not a placeholder.
    await expect(page.locator('[data-dxt="calls"]')).not.toHaveText('0');
    await expect(page.locator('[data-dxt="p50"]')).not.toHaveText('–');
  });

  test('is fast enough to animate: median tool call stays in single digits', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);

    const stats = await page.evaluate(() => (window as any).__theater.stats());
    // The demo's performance claim, guarded. These are generous against a cold
    // CI runner — the point is to catch an order-of-magnitude regression (an
    // accidental full-document reconvert per edit), not to benchmark.
    expect(stats.p50, `p50 was ${stats.p50}ms`).toBeLessThan(120);
    expect(stats.p95, `p95 was ${stats.p95}ms`).toBeLessThan(600);
    // Repaints are driven by mutations, not by calls: the seven searches and the
    // two set_mode frames change nothing and must not repaint. Plus the repaint
    // is frame-dropped, so several mutations landing inside one animation frame
    // coalesce into a single refresh. Both together put this comfortably below
    // the call count rather than one under it.
    expect(stats.renderCount).toBeGreaterThan(0);
    expect(stats.renderCount).toBeLessThanOrEqual(stats.calls - 9);
  });

  test('reset returns to a clean stage and can run again', async ({ page }) => {
    await bootTheater(page);
    await runToFinale(page);
    expect(await page.evaluate(() => (window as any).__theater.stats().revisions))
      .toBeGreaterThan(10);

    await page.locator('[data-dxt="reset"]').click();
    await page.waitForFunction(
      () => (window as any).__theater.stats().revisions === 0, { timeout: 60000 });

    // A second run must be as clean as the first: leftover notes, comments or
    // ghost headings from the previous take would show up as extra revisions.
    await runToFinale(page);
    await expect(page.locator('[data-dxt="proof"]')).toHaveAttribute('data-ok', 'true');
    await expect(page.locator('[data-dxt="error"]')).not.toHaveAttribute('data-on', 'true');
  });
});

// ─── DIFF STRESS ──────────────────────────────────────────────────────
//
// The theater RECORDS its redline; this mode COMPUTES one from scratch after
// every edit and times it. The point of the mode is the gap between those two
// costs, so these assertions are about the measurement being real: a diff
// genuinely runs each frame, the deeper pipelines genuinely cost more, and the
// ratio the panel reports is derived from both halves rather than hardcoded.

async function runStress(page: Page, depth: string, frames: number) {
  await page.evaluate(() => (window as any).__theater.setMode('stress'));
  await page.evaluate((d) => (window as any).__theater.stress.setDepth(d), depth);
  // Fire and forget: run() resolves only when the whole loop is done.
  await page.evaluate(() => { void (window as any).__theater.stress.run(); });
  await page.waitForFunction(
    (n) => (window as any).__theater.stress.stats().frames >= n,
    frames, { timeout: 180000 });
  await page.evaluate(() => (window as any).__theater.stress.stop());
  await page.waitForFunction(
    () => !(window as any).__theater.stress.isRunning(), { timeout: 60000 });
  return page.evaluate(() => (window as any).__theater.stress.stats());
}

test.describe('DIFF STRESS', () => {
  test('computes a real redline every frame, against a growing document', async ({ page }) => {
    await bootTheater(page);
    const stats = await runStress(page, 'revisions', 6);

    expect(stats.frames).toBeGreaterThanOrEqual(6);
    // One clause appended per frame, so the revision count tracks the frames —
    // this is what proves a diff actually ran rather than a cached number being
    // re-displayed.
    expect(stats.revisions).toBeGreaterThanOrEqual(6);
    // The document grew under the engine; a fixed input would measure nothing.
    expect(stats.documentBytes).toBeGreaterThan(stats.startBytes);
    expect(stats.fps).toBeGreaterThan(0);
    expect(stats.p50).toBeGreaterThan(0);

    // The panel shows the measurement, not a placeholder.
    await expect(page.locator('[data-dxs="fps"]')).not.toHaveText('–');
    await expect(page.locator('[data-dxs="frames"]')).not.toHaveText('0');
    const path = await page.locator('[data-dxs="path"]').getAttribute('d');
    expect(path, 'the frame-time trace is drawn').toBeTruthy();
    expect(path!.split('L').length).toBeGreaterThan(3);
  });

  test('a deeper pipeline costs more, and the panel says so', async ({ page }) => {
    await bootTheater(page);
    const cheap = await runStress(page, 'revisions', 5);
    const full = await runStress(page, 'full', 5);

    // Rendering the redline to HTML on top of computing it cannot be free. This
    // is the assertion that would catch the depth selector silently not applying
    // — the bug the first draft of this mode actually had.
    expect(full.p50, `revisions ${cheap.p50}ms vs full ${full.p50}ms`)
      .toBeGreaterThan(cheap.p50);
    expect(full.fps).toBeLessThan(cheap.fps);
    expect(full.depth).toBe('full');

    // Each run is its own experiment — the meter resets, so the second run's
    // frame count is not the first run's plus more.
    expect(full.frames).toBeLessThan(cheap.frames + full.frames);
  });

  test('reports how much dearer computing a redline is than recording one', async ({ page }) => {
    await bootTheater(page);
    const stats = await runStress(page, 'redline', 5);

    // Both halves are measured through the same MCP endpoint: the mutation is a
    // real docxodus_create frame, the diff is a real comparison pass.
    expect(stats.avgMutateMs).toBeGreaterThan(0);
    expect(stats.avgMutateMs).toBeLessThan(60);
    // Computing is expected to be at least an order of magnitude dearer. The
    // bound is deliberately loose — the claim is the order of magnitude, not a
    // number a slower runner would fail.
    expect(stats.ratio, `ratio was ${stats.ratio}`).toBeGreaterThan(10);
    await expect(page.locator('[data-dxs="ratiox"]')).toContainText('×');
  });

  test('switching away from the stress pane stops the loop', async ({ page }) => {
    await bootTheater(page);
    await page.evaluate(() => (window as any).__theater.setMode('stress'));
    await page.evaluate(() => { void (window as any).__theater.stress.run(); });
    await page.waitForFunction(
      () => (window as any).__theater.stress.stats().frames >= 2, { timeout: 120000 });

    // A loop left running behind a hidden pane would keep burning the engine.
    await page.locator('[data-dxt="modes"] button[data-mode="wire"]').click();
    await page.waitForFunction(
      () => !(window as any).__theater.stress.isRunning(), { timeout: 60000 });
    expect(await page.evaluate(() => (window as any).__theater.stress.isRunning())).toBe(false);
  });
});
