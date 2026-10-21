import { test, expect, Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Whether the editor shows the revisions a document already holds is a display choice; how it
// records new edits is a separate one. They used to be the same switch: the editor showed
// revisions only while recording them, and the ribbon's "No markup" turned recording off
// (issue #1279).

const REVISION = 'w:author="Counsel A" w:date="2026-05-01T09:00:00Z"';
const ACCEPTED = 'The tenant shall repaint the fence each spring.';
const WITH_MARKUP = 'The tenant shall promptly repaint the fence each spring.';

function revisionFixture(): number[] {
  const body = '<w:p><w:r><w:t xml:space="preserve">The tenant shall </w:t></w:r>'
    + `<w:del w:id="1" ${REVISION}><w:r><w:delText xml:space="preserve">promptly </w:delText></w:r></w:del>`
    + '<w:r><w:t xml:space="preserve">repaint the fence</w:t></w:r>'
    + `<w:ins w:id="2" ${REVISION}><w:r><w:t xml:space="preserve"> each spring</w:t></w:r></w:ins>`
    + '<w:r><w:t>.</w:t></w:r></w:p>';
  return Array.from(storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) }))));
}

/** What the editor shows: rendered revision marks, the paragraph text, and the recording mode. */
type Shown = { ins: number; del: number; text: string; trackedChanges: number };

async function openEditor(page: Page, options: Record<string, unknown>) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ bytes, options }) => {
    const w = window as any;
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [] as unknown[];
    w.editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {
      ...options,
      onEditFailed: (info: unknown) => w.failures.push(info),
    });
    w.shown = (): Shown => {
      const p = container.querySelector<HTMLElement>('p[data-anchor]')!;
      return {
        ins: container.querySelectorAll('ins').length,
        del: container.querySelectorAll('del').length,
        text: (p.textContent ?? '').replace(/ /g, ' '),
        trackedChanges: w.editor.trackedChanges,
      };
    };
    w.revisions = () => w.editor.listRevisions().map((r: any) => `${r.type}:${r.text}`);
  }, { bytes: revisionFixture(), options });
}

const shown = (page: Page): Promise<Shown> => page.evaluate(() => (window as any).shown());

/** Type `text` at the end of the paragraph and commit it. */
async function typeAtEnd(page: Page, text: string) {
  await page.evaluate(() => (document.querySelector('p[data-anchor]') as HTMLElement).focus());
  await page.keyboard.press('End');
  await page.keyboard.type(text);
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
}

test.describe('DocxEditor — showing revisions apart from recording them (#1279)', () => {
  test('control: with neither option set, the document opens as if every change were accepted', async ({ page }) => {
    await openEditor(page, {});
    expect(await shown(page)).toEqual({ ins: 0, del: 0, text: ACCEPTED, trackedChanges: 0 });
  });

  test('showRevisions: true renders existing revisions while edits stay untracked', async ({ page }) => {
    await openEditor(page, { showRevisions: true });
    expect(await shown(page)).toEqual({ ins: 1, del: 1, text: WITH_MARKUP, trackedChanges: 0 });
    expect(await page.evaluate(() => (window as any).editor.revisionsVisible)).toBe(true);
    const before = await page.evaluate(() => (window as any).revisions());

    await typeAtEnd(page, ' Z');

    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
    expect(await page.evaluate(() => (window as any).revisions())).toEqual(before);
    expect((await shown(page)).text).toBe(`${WITH_MARKUP} Z`);
  });

  test('with edits untracked, a Backspace that would reach shown deleted text is refused', async ({ page }) => {
    await openEditor(page, { showRevisions: true });
    const before = await page.evaluate(() => (window as any).revisions());
    await page.evaluate(() => {
      const p = document.querySelector<HTMLElement>('p[data-anchor]')!;
      p.focus();
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        if (!(node.textContent ?? '').startsWith('repaint')) continue;
        const range = document.createRange();
        range.setStart(node, 0);
        range.collapse(true);
        window.getSelection()!.removeAllRanges();
        window.getSelection()!.addRange(range);
        return;
      }
      throw new Error('no text node starts with "repaint"');
    });
    await page.keyboard.press('Backspace');
    await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
    expect(await page.evaluate(() => (window as any).revisions())).toEqual(before);
    expect(await shown(page)).toEqual({ ins: 1, del: 1, text: WITH_MARKUP, trackedChanges: 0 });
  });

  test('showRevisions(false) hides markup while edits keep recording as revisions', async ({ page }) => {
    await openEditor(page, { trackedChanges: 1 /* RenderInline */ });
    expect(await shown(page)).toEqual({ ins: 1, del: 1, text: WITH_MARKUP, trackedChanges: 1 });

    await page.evaluate(() => (window as any).editor.showRevisions(false));
    expect(await shown(page)).toEqual({ ins: 0, del: 0, text: ACCEPTED, trackedChanges: 1 });
    expect(await page.evaluate(() => (window as any).editor.revisionsVisible)).toBe(false);

    await typeAtEnd(page, ' Z');
    expect(await page.evaluate(() => (window as any).failures)).toEqual([]);
    expect(await page.evaluate(() => (window as any).revisions())).toContain('insert: Z');
    expect(await shown(page)).toEqual({ ins: 0, del: 0, text: `${ACCEPTED} Z`, trackedChanges: 1 });
  });

  test('a display choice survives a change of recording mode', async ({ page }) => {
    await openEditor(page, { showRevisions: false });
    await page.evaluate(() => (window as any).editor.setTrackedChanges(1));
    expect(await shown(page)).toEqual({ ins: 0, del: 0, text: ACCEPTED, trackedChanges: 1 });
    await page.evaluate(() => (window as any).editor.showRevisions(true));
    await page.evaluate(() => (window as any).editor.setTrackedChanges(0));
    expect(await shown(page)).toEqual({ ins: 1, del: 1, text: WITH_MARKUP, trackedChanges: 0 });
  });

  test('control: left unset, the display follows the recording mode as before', async ({ page }) => {
    await openEditor(page, {});
    await page.evaluate(() => (window as any).editor.setTrackedChanges(1));
    expect(await shown(page)).toEqual({ ins: 1, del: 1, text: WITH_MARKUP, trackedChanges: 1 });
    await page.evaluate(() => (window as any).editor.setTrackedChanges(0));
    expect(await shown(page)).toEqual({ ins: 0, del: 0, text: ACCEPTED, trackedChanges: 0 });
  });
});

async function openRibbon(page: Page) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo?.getEditor(), { timeout: 60000 });
  await page.evaluate((bytes) => {
    const w = window as any;
    w.__ribbon.open(new Uint8Array(bytes), 'revisions.docx');
    w.__selectTab('review');
  }, revisionFixture());
}

/** The ribbon's review controls and what the live editor shows. */
async function review(page: Page) {
  return page.evaluate(() => {
    const w = window as any;
    const surface: HTMLElement = w.__ribbon.surface;
    return {
      trackedChanges: w.__ribbon.editor.trackedChanges,
      checkbox: (document.querySelector('[data-dxr="trackchanges"]') as HTMLInputElement).checked,
      markup: (document.querySelector('[data-dxr="markup"]') as HTMLSelectElement).value,
      marks: surface.querySelectorAll('ins, del').length,
    };
  });
}

test.describe('ribbon Markup control — display only (#1279)', () => {
  test('"No markup" hides revisions and leaves Track Changes on', async ({ page }) => {
    await openRibbon(page);
    await page.locator('[data-dxr="trackchanges"]').check();
    expect(await review(page)).toEqual({ trackedChanges: 1, checkbox: true, markup: 'all', marks: 2 });

    await page.locator('[data-dxr="markup"]').selectOption('none');
    expect(await review(page)).toEqual({ trackedChanges: 1, checkbox: true, markup: 'none', marks: 0 });
  });

  test('"All markup" shows revisions and leaves Track Changes off', async ({ page }) => {
    await openRibbon(page);
    expect(await review(page)).toEqual({ trackedChanges: 0, checkbox: false, markup: 'none', marks: 0 });

    await page.locator('[data-dxr="markup"]').selectOption('all');
    expect(await review(page)).toEqual({ trackedChanges: 0, checkbox: false, markup: 'all', marks: 2 });

    // The choice outlives the document: a newly opened file keeps it.
    await page.evaluate((bytes) => (window as any).__ribbon.open(new Uint8Array(bytes), 'again.docx'), revisionFixture());
    expect(await review(page)).toEqual({ trackedChanges: 0, checkbox: false, markup: 'all', marks: 2 });
  });

  test('control: before any Markup choice, Track Changes still turns markup on with it', async ({ page }) => {
    await openRibbon(page);
    await page.locator('[data-dxr="trackchanges"]').check();
    expect(await review(page)).toEqual({ trackedChanges: 1, checkbox: true, markup: 'all', marks: 2 });
    await page.locator('[data-dxr="trackchanges"]').uncheck();
    expect(await review(page)).toEqual({ trackedChanges: 0, checkbox: false, markup: 'none', marks: 0 });
  });
});
