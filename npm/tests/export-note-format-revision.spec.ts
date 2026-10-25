import { expect, test, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { NOTE_TEXT, noteFormatRevisionDocx } from './docx-note-format-revision-fixture.js';

/**
 * A fitting footnote whose text carries a formatting revision exports (issue #1271).
 *
 * The default revision CSS draws a 2px dotted border under `span.rev-format-change`. CSS paints an
 * inline box's bottom border below its text without making the line taller, so on a one-line note
 * it hung past the note band, and the clipping check, which measured every element's border box,
 * failed the export with "A footnote continuation is clipped" although every glyph showed.
 */

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const markup = { reviewProfile: 'markup', commentProfile: 'hidden' };

/** The page-tree export in the harness: its outcome, and the text of each page's note band. */
async function exportNotes(page: Page, source: Uint8Array, options: Record<string, unknown> = {}) {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  return page.evaluate(async ({ bytes, options }) => {
    let result;
    try {
      result = await (window as any).DocxodusStandalone.convert(bytes, options);
    } catch (error: any) {
      return { failure: `${error.code}: ${error.message}` };
    }
    const html = new DOMParser().parseFromString(result.html, 'text/html');
    return { status: result.renderReport.status,
      notes: Array.from(html.querySelectorAll('.page-footnotes'), notes => notes.textContent!.replace(/\s+/g, ' ').trim()) };
  }, { bytes: Array.from(source), options: { ...markup, ...options } });
}

test.describe('a footnote with a formatting revision (#1271)', () => {
  test('exports, with the whole note on its page', async ({ page }) => {
    const result = await exportNotes(page, noteFormatRevisionDocx());
    expect(result.failure).toBeUndefined();
    expect(result.status).toBe('complete');
    expect(result.notes).toHaveLength(1);
    expect(result.notes?.[0]).toContain(NOTE_TEXT);
  });

  test('controls: the same note without the revision, and under the Word presentation', async ({ page }) => {
    for (const [source, options] of [
      [noteFormatRevisionDocx({ formatRevision: false }), {}],
      [noteFormatRevisionDocx(), { markupPresentation: 'word' }],
    ] as const) {
      const result = await exportNotes(page, source, options);
      expect(result.failure).toBeUndefined();
      expect(result.notes?.[0]).toContain(NOTE_TEXT);
    }
  });

  test('a note that really is cut off still fails, revision or not', async ({ page }) => {
    // A 750pt exact-height row cannot split and is taller than the page, so the text at its foot is lost.
    for (const formatRevision of [false, true]) {
      const result = await exportNotes(page, noteFormatRevisionDocx({ formatRevision, exactRowTwips: 15000 }));
      expect(result.failure).toBe('pagination_failure: A footnote continuation is clipped in the final page tree.');
    }
  });

  test('the delivered PDF carries the note text', async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(noteFormatRevisionDocx(), { browser, outputs: ['pdf'],
      reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const text = (await (await pdf.getPage(1)).getTextContent()).items
        .filter(item => 'str' in item).map(item => item.str).join('');
      expect(text.replace(/\s+/g, ' ')).toContain(NOTE_TEXT);
    } finally { await task.destroy(); }
  });
});

const lineMetrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');

test('laidOutBottom leaves out only an inline box\'s own bottom padding and border', async ({ page }) => {
  for (const zoom of [1, 2]) {
    await page.setContent(`<div id="band" style="zoom:${zoom};width:400px;font:16px/20px 'Liberation Sans'">
      <p style="margin:0"><span id="decorated" style="border-bottom:6px dotted orange;padding-bottom:3px">ruled</span>
        <span id="lowered" style="vertical-align:-14px">low</span>
        <img id="image" alt="" style="width:10px;height:10px;border-bottom:6px solid red">
        <span id="block" style="display:inline-block;border-bottom:6px solid blue">box</span></p></div>`);
    await page.addScriptTag({ content: lineMetrics });
    const result = await page.evaluate(() => {
      const band = document.getElementById('band')!;
      const measure = (id: string) => {
        const element = document.getElementById(id)!;
        const range = document.createRange();
        range.selectNodeContents(element);
        return { box: element.getBoundingClientRect().bottom, text: range.getBoundingClientRect().bottom,
          laidOut: (window as any).laidOutBottom(element, band) };
      };
      return { scale: band.getBoundingClientRect().width / 400,
        decorated: measure('decorated'), lowered: measure('lowered'), image: measure('image'), block: measure('block') };
    });
    expect(result.scale).toBeCloseTo(zoom, 5);
    // The decorated run's border and padding are left out, down to its text (9 CSS px at this zoom).
    expect(result.decorated.box - result.decorated.laidOut).toBeCloseTo(9 * zoom, 3);
    expect(result.decorated.laidOut).toBeCloseTo(result.decorated.text, 0);
    // A lowered run's text, an image's border and an inline-block's border all stay measured.
    expect(result.lowered.laidOut).toBe(result.lowered.box);
    expect(result.image.laidOut).toBe(result.image.box);
    expect(result.block.laidOut).toBe(result.block.box);
  }
});
