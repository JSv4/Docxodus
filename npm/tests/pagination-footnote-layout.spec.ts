import { test, expect, type Page, type TestInfo } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Layout contract for the paginated footnote area. The engine bottom-anchors the note block and
 * lets it grow upward into body space, so two things have to hold that nothing pinned before:
 * the body must never be drawn underneath the notes, and reserving note space must not throw the
 * rest of the page away. Both were broken on a real 94-footnote document — one page rendered body
 * and note glyphs on top of each other, and the median page ran 57% full (93% with notes off).
 */

/** A page of `bodyBlocks` paragraphs, the first `citing` of which each cite a distinct footnote. */
function stagingHtml(bodyBlocks: number, citing: number, noteLines: number): string {
  const notes = Array.from({ length: citing }, (_, i) => `
    <div class="footnote-item" data-footnote-id="f${i}">
      <span class="footnote-number">${i + 1}. </span>
      <span class="footnote-content">${
        Array.from({ length: noteLines }, () => `<p>note ${i} line</p>`).join('')
      }</span>
    </div>`).join('');

  const body = Array.from({ length: bodyBlocks }, (_, i) =>
    i < citing
      ? `<p class="body">body ${i} <sup data-footnote-id="f${i}">${i + 1}</sup></p>`
      : `<p class="body">body ${i}</p>`).join('');

  return `
    <style>
      #staging { font: 16px/16px Arial; }
      .body { height: 20pt; margin: 0; }
      .footnote-item { margin: 0; }
      .footnote-content > p { height: 10pt; margin: 0; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">${notes}</div>
      <div data-section-index="0"
           data-page-width="312" data-page-height="312"
           data-content-width="300" data-content-height="300"
           data-margin-top="6" data-margin-right="6"
           data-margin-bottom="6" data-margin-left="6">
        ${body}
      </div>
    </div>
    <div id="container"></div>`;
}

/** The document's final body block cites several whole notes, forcing at least one to defer. */
function finalBlockDeferralHtml(): string {
  const notes = Array.from({ length: 4 }, (_, i) => `
    <div class="footnote-item" data-footnote-id="last${i}">
      <span class="footnote-number">${i + 1}. </span>
      <span class="footnote-content">${
        Array.from({ length: 10 }, () => `<p>last note ${i} line</p>`).join('')
      }</span>
    </div>`).join('');
  const citations = Array.from({ length: 4 }, (_, i) =>
    `<sup data-footnote-id="last${i}">${i + 1}</sup>`).join('');
  return `
    <style>
      #staging { font: 16px/16px Arial; }
      .body { height: 100pt; margin: 0; }
      .footnote-item { margin: 0; }
      .footnote-content > p { height: 10pt; margin: 0; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">${notes}</div>
      <div data-section-index="0"
           data-page-width="312" data-page-height="312"
           data-content-width="300" data-content-height="300"
           data-margin-top="6" data-margin-right="6"
           data-margin-bottom="6" data-margin-left="6">
        <p class="body">only and final body block ${citations}</p>
      </div>
    </div>
    <div id="container"></div>`;
}

const LONG_NOTE_WORDS = Array.from({ length: 120 }, (_, index) => `word-${index}`).join(' ');

/**
 * One ordinary long paragraph is deliberately wider than several 60%-height note bands. The
 * trailing paragraphs make the former C2 failure visible: an oversized leader must not evacuate
 * the citing page or strand otherwise splittable siblings inside one clipped continuation.
 */
function longParagraphFootnoteHtml(
  trailingParagraphs = '',
  leadingContent = LONG_NOTE_WORDS,
  extraCss = '',
  itemAttributes = '',
  contentAttributes = '',
  separatorStories = '',
): string {
  return `
    <style>
      #staging { font: 12pt/12pt Arial; }
      .body { height: 20pt; margin: 0; }
      .page-footnotes { font-size: 10pt; line-height: 10pt; }
      .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
      .footnote-item { margin: 0; }
      .footnote-number { display: inline; vertical-align: super; }
      .footnote-content { display: inline; }
      .footnote-content p:first-of-type { display: inline; }
      .footnote-content p:not(:first-of-type) { display: block; margin: 0; }
      .footnote-continuation p { margin: 0; }
      ${extraCss}
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">
        ${separatorStories}
        <div class="footnote-item" data-footnote-id="f-long"
             data-source-anchor-id="fn:fn:f-long" ${itemAttributes}>
          <span class="footnote-number">1</span>
          <span class="footnote-content" ${contentAttributes}>
            <p id="source-long" data-anchor="long-unid"
               data-source-anchor-id="p:fn:long">${leadingContent}</p>
            ${trailingParagraphs}
          </span>
        </div>
      </div>
      <div id="pagination-comment-margin-registry">
        <aside data-comment-id="7" data-source-anchor-id="cmt:fn:seven">
          <p>margin-comment-seven</p>
        </aside>
      </div>
      <div data-section-index="0"
           data-page-width="140" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="19"
           data-margin-bottom="1" data-margin-left="1">
        <p class="body" data-source-anchor-id="p:body:citation">
          citing body <sup data-footnote-id="f-long">1</sup>
        </p>
      </div>
    </div>
    <div id="container"></div>`;
}

function continuationThenNewNoteHtml(): string {
  const longWords = Array.from({ length: 140 }, (_, index) => `first-${index}`).join(' ');
  return `
    <style>
      #staging { font: 12pt/12pt Arial; }
      .body { height: 20pt; margin: 0; }
      .page-footnotes { font: 10pt/10pt Arial; }
      .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
      .footnote-item, .footnote-content p, .footnote-continuation p { margin: 0; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f1">
          <span class="footnote-number">1</span>
          <span class="footnote-content"><p>${longWords}</p></span>
        </div>
        <div class="footnote-item" data-footnote-id="f2">
          <span class="footnote-number">2</span>
          <span class="footnote-content"><p>second-note-alpha second-note-omega</p></span>
        </div>
      </div>
      <div data-section-index="0"
           data-page-width="122" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p class="body">cite first <sup data-footnote-id="f1">1</sup></p>
        <p class="body">body after first A</p>
        <p class="body">cite second <sup data-footnote-id="f2">2</sup></p>
        <p class="body">body after second B</p>
      </div>
    </div>
    <div id="container"></div>`;
}

function noPrefixFitsCitationBandHtml(): string {
  return `
    <style>
      #staging { font: 12pt/12pt Arial; }
      .body { height: 95pt; margin: 0; }
      .page-footnotes { font: 10pt/10pt Arial; }
      .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
      .footnote-item { margin: 0; }
      .footnote-content p, .footnote-continuation p {
        display: block; height: 40pt; line-height: 40pt; margin: 0;
      }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="fresh">
          <span class="footnote-number">1</span>
          <span class="footnote-content"><p>fresh-band-safe tail-safe</p></span>
        </div>
      </div>
      <div data-section-index="0"
           data-page-width="122" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p class="body">citation <sup data-footnote-id="fresh">1</sup></p>
      </div>
    </div>
    <div id="container"></div>`;
}

function indivisibleFitsFreshBandHtml(): string {
  return `
    <style>
      #staging { font: 12pt/12pt Arial; }
      .body { height: 95pt; margin: 0; }
      .page-footnotes { font: 10pt/10pt Arial; }
      .page-footnotes hr { height: 1px; margin: 0 0 3pt; border: 0; }
      .footnote-item, .footnote-content p, .footnote-continuation p { margin: 0; }
      .unsafe-fresh-band { display: inline-block; height: 45pt; width: 20pt; }
    </style>
    <div id="staging">
      <div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="unsafe-fresh">
          <span class="footnote-number">1</span>
          <span class="footnote-content">
            <p><span class="unsafe-fresh-band">indivisible-safe</span></p>
            <p>later-safe-sibling</p>
          </span>
        </div>
      </div>
      <div data-section-index="0"
           data-page-width="122" data-page-height="122"
           data-content-width="120" data-content-height="120"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        <p class="body">citation <sup data-footnote-id="unsafe-fresh">1</sup></p>
      </div>
    </div>
    <div id="container"></div>`;
}

async function paginateScenario(page: Page, html: string) {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const result = new PaginationEngine(staging, container, {
      showPageNumbers: false,
      checkCancellation: (() => {
        let checkpoints = 0;
        return () => {
          if (++checkpoints > 50_000) throw new Error('pagination did not make forward progress');
        };
      })(),
    }).paginate();
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const normalized = (value: string | null | undefined) =>
      (value ?? '').replace(/\s+/g, ' ').trim();
    const clipped: Array<{ page: number; text: string }> = [];
    const overlaps: number[] = [];
    boxes.forEach((box, index) => {
      const band = box.querySelector<HTMLElement>('.page-footnotes');
      const content = box.querySelector<HTMLElement>('.page-content');
      if (!band) return;
      const bandRect = band.getBoundingClientRect();
      for (const element of Array.from(band.querySelectorAll<HTMLElement>('p, .footnote-item'))) {
        if (element.getBoundingClientRect().bottom > bandRect.bottom + 1) {
          clipped.push({ page: index + 1, text: normalized(element.textContent) });
        }
      }
      const body = content ? Array.from(content.children) as HTMLElement[] : [];
      const bodyBottom = body.at(-1)?.getBoundingClientRect().bottom;
      if (bodyBottom !== undefined && bodyBottom > bandRect.top + 1) overlaps.push(index + 1);
    });
    return {
      totalPages: result.totalPages,
      bodyByPage: boxes.map((box) => Array.from(
        box.querySelector<HTMLElement>('.page-content')?.children ?? [],
        (element) => normalized(element.textContent),
      )),
      noteTextByPage: boxes.map((box) => {
        const clone = box.querySelector<HTMLElement>('.page-footnotes')?.cloneNode(true) as
          HTMLElement | undefined;
        clone?.querySelectorAll('.footnote-number').forEach((number) => number.remove());
        return normalized(clone?.textContent);
      }),
      notePages: boxes.map((box, index) =>
        box.querySelector('.page-footnotes') ? index + 1 : 0).filter(Boolean),
      f1Parts: container.querySelectorAll('.page-footnotes [data-footnote-id="f1"]').length,
      f2Parts: container.querySelectorAll('.page-footnotes [data-footnote-id="f2"]').length,
      freshPages: Array.from(container.querySelectorAll<HTMLElement>(
        '.page-footnotes [data-footnote-id="fresh"]',
      ))
        .map((element) => Number(element.closest<HTMLElement>('.page-box')?.dataset.pageNumber)),
      unsafeFreshPages: Array.from(container.querySelectorAll<HTMLElement>(
        '.page-footnotes [data-footnote-id="unsafe-fresh"]',
      )).map((element) => Number(element.closest<HTMLElement>('.page-box')?.dataset.pageNumber)),
      numberText: Array.from(container.querySelectorAll<HTMLElement>('.footnote-number'))
        .map((number) => normalized(number.textContent)),
      clipped,
      overlaps,
    };
  });
}

async function paginateLongFootnote(
  page: Page,
  trailingParagraphs = '',
  leadingContent = LONG_NOTE_WORDS,
  extraCss = '',
  itemAttributes = '',
  contentAttributes = '',
  separatorStories = '',
) {
  const expectedText = `${leadingContent.replace(/<[^>]+>/g, ' ')} ${
    trailingParagraphs.replace(/<[^>]+>/g, ' ')
  }`.replace(/\s+/g, ' ').trim();
  await page.setContent(longParagraphFootnoteHtml(
    trailingParagraphs,
    leadingContent,
    extraCss,
    itemAttributes,
    contentAttributes,
    separatorStories,
  ));
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ expectedText }: { expectedText: string }) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    const pagination = new PaginationEngine(staging, container, {
      showPageNumbers: false,
      layoutToken: { documentVersion: 1, rendererFingerprint: 'issue-489-evidence' },
    }).paginate();
    const boxes = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    const noteBands = boxes.map((box) => box.querySelector<HTMLElement>('.page-footnotes'));
    const paragraphFragments = Array.from(
      container.querySelectorAll<HTMLElement>('[data-source-anchor-id="p:fn:long"]'),
    );
    const visibleNoteText = boxes.map((box) => {
      const clone = box.querySelector<HTMLElement>('.page-footnotes')?.cloneNode(true) as
        HTMLElement | undefined;
      clone?.querySelectorAll('.footnote-number').forEach((number) => number.remove());
      clone?.querySelectorAll('[data-footnote-separator]').forEach((separator) => separator.remove());
      clone?.querySelectorAll('p').forEach((paragraph) => paragraph.after(' '));
      return clone?.textContent ?? '';
    }).join(' ').replace(/\s+/g, ' ').trim();
    const clipped = noteBands.flatMap((band, pageIndex) => {
      if (!band) return [];
      const bandRect = band.getBoundingClientRect();
      return Array.from(band.querySelectorAll<HTMLElement>('p'))
        .filter((paragraph) => paragraph.getBoundingClientRect().bottom > bandRect.bottom + 1)
        .map((paragraph) => ({
          page: pageIndex + 1,
          anchor: paragraph.dataset.sourceAnchorId,
          paragraphBottom: paragraph.getBoundingClientRect().bottom,
          bandBottom: bandRect.bottom,
        }));
    });
    const pageMapFragments = pagination.pageMap.fragments
      .filter((fragment: any) => fragment.anchorId === 'p:fn:long');
    const pageNumberFor = (element: Element) => Number(
      element.closest<HTMLElement>('.page-box')?.dataset.pageNumber ?? '0',
    );
    const links = Array.from(container.querySelectorAll<HTMLAnchorElement>('[data-case="link"]'));
    const clusterParts = Array.from(
      container.querySelectorAll<HTMLElement>('[data-cluster-part]'),
    );
    const continuationShells = Array.from(
      container.querySelectorAll<HTMLElement>('.footnote-continuation'),
    );
    return {
      totalPages: pagination.totalPages,
      notePages: noteBands.filter(Boolean).length,
      paragraphFragments: paragraphFragments.length,
      pageMapFragments: pageMapFragments.map((fragment: any) => ({
        fragmentId: fragment.fragmentId,
        fragmentIndex: fragment.fragmentIndex,
        pageNumber: fragment.pageNumber,
        story: fragment.story,
        geometry: fragment.geometry,
      })),
      bareIds: paragraphFragments.filter((paragraph) => paragraph.id === 'source-long').length,
      bareAnchors: paragraphFragments.filter((paragraph) =>
        paragraph.dataset.anchor === 'long-unid').length,
      visibleNoteText,
      expectedText,
      clipped,
      firstPageText: noteBands[0]?.textContent?.replace(/\s+/g, ' ').trim() ?? '',
      pagesText: noteBands.map((band) => band?.textContent?.replace(/\s+/g, ' ').trim() ?? ''),
      separatorCount: container.querySelectorAll(
        '.page-footnotes > [data-footnote-separator]',
      ).length,
      separatorKinds: noteBands.map((band) =>
        band?.firstElementChild?.getAttribute('data-footnote-separator') ?? ''),
      separatorText: noteBands.map((band) =>
        band?.firstElementChild?.textContent?.replace(/\s+/g, ' ').trim() ?? ''),
      separatorFirst: noteBands.every((band) => !band ||
        Boolean(band.firstElementChild?.hasAttribute('data-footnote-separator'))),
      separatorIdentityCount: container.querySelectorAll(
        '.page-footnotes #separator-source, .page-footnotes [data-anchor="separator-anchor"], ' +
        '.page-footnotes [data-source-anchor-id="p:separator-source"]',
      ).length,
      separatorLocalLinks: container.querySelectorAll(
        '.page-footnotes [data-footnote-separator] a[href^="#"]',
      ).length,
      separatorExternalLinks: Array.from(container.querySelectorAll<HTMLAnchorElement>(
        '.page-footnotes [data-footnote-separator] a[href^="https://"]',
      )).map((link) => link.href),
      numberCount: container.querySelectorAll('.page-footnotes .footnote-number').length,
      continuationFootnoteIds: Array.from(
        container.querySelectorAll<HTMLElement>('.footnote-continuation'),
      ).map((wrapper) => wrapper.dataset.footnoteId ?? ''),
      continuationShells: continuationShells.map((wrapper) => {
        const content = wrapper.querySelector<HTMLElement>(':scope > .footnote-content');
        const firstParagraph = content?.querySelector<HTMLElement>(
          'p:not(.footnote-continuation-position-sentinel)',
        );
        const style = firstParagraph ? getComputedStyle(firstParagraph) : null;
        return {
          isItem: wrapper.classList.contains('footnote-item'),
          hasDirectContent: Boolean(content),
          shellToken: wrapper.dataset.shellToken ?? '',
          contentToken: content?.dataset.contentToken ?? '',
          direction: getComputedStyle(wrapper).direction,
          lang: wrapper.lang,
          color: style?.color ?? '',
          display: style?.display ?? '',
          marginLeft: style?.marginLeft ?? '',
          sentinelCount: content?.querySelectorAll(
            ':scope > .footnote-continuation-position-sentinel',
          ).length ?? 0,
          duplicateIds: wrapper.querySelectorAll('#source-note-content').length
            + (wrapper.id === 'source-note-item' ? 1 : 0),
          duplicateAnchors: wrapper.querySelectorAll('[data-anchor="source-content-anchor"]').length
            + (wrapper.dataset.anchor === 'source-item-anchor' ? 1 : 0),
        };
      }),
      listMarkerPages: Array.from(
        container.querySelectorAll<HTMLElement>('[data-list-marker]'),
      ).map(pageNumberFor),
      bookmarkTargetCount: container.querySelectorAll('#note-bookmark').length,
      inlineAnchorCount: container.querySelectorAll('[data-anchor="inline-note-anchor"]').length,
      fieldCount: container.querySelectorAll('[data-field="PAGE"]').length,
      linkHrefs: links.map((link) => link.getAttribute('href')),
      linkText: links.map((link) => link.textContent ?? '').join(''),
      commentText: Array.from(
        container.querySelectorAll<HTMLElement>('[data-comment-id="7"]'),
      ).map((comment) => comment.textContent ?? '').join(''),
      commentPageMapFragments: pagination.pageMap.fragments.filter(
        (fragment: any) => fragment.anchorId === 'cmt:fn:seven',
      ).length,
      commentMarkerCount: container.querySelectorAll('a[data-comment-id="7"]').length,
      marginCommentCount: container.querySelectorAll(
        '.page-comment-margin > [data-comment-id="7"]',
      ).length,
      marginCommentText: Array.from(container.querySelectorAll<HTMLElement>(
        '.page-comment-margin > [data-comment-id="7"]',
      )).map((comment) => comment.textContent?.replace(/\s+/g, ' ').trim() ?? '').join(' '),
      richPayloadText: Array.from(
        container.querySelectorAll<HTMLElement>('[data-case="rich-payload"]'),
      ).map((fragment) => fragment.textContent ?? '').join(''),
      richPayloadFragments: container.querySelectorAll('[data-case="rich-payload"]').length,
      crossRunFragments: Array.from(container.querySelectorAll<HTMLElement>(
        '[data-case="cross-run"]',
      )).map((fragment) => fragment.textContent ?? ''),
      cjkFragments: Array.from(container.querySelectorAll<HTMLElement>(
        '[data-case="cjk"]',
      )).map((fragment) => fragment.textContent ?? ''),
      clusterPartCounts: ['left', 'right'].map((part) =>
        clusterParts.filter((element) => element.dataset.clusterPart === part).length),
      clusterPages: clusterParts.map(pageNumberFor),
      tableText: container.querySelector<HTMLElement>('[data-case="note-table"]')?.textContent
        ?.replace(/\s+/g, ' ').trim() ?? '',
      tablePageMapFragments: pagination.pageMap.fragments.filter(
        (fragment: any) => fragment.anchorId === 'tbl:fn:tail-table',
      ).length,
    };
  }, { expectedText });
}

type LongFootnoteResult = Awaited<ReturnType<typeof paginateLongFootnote>>;

function describeFailure(error: unknown): { name?: string; message: string; stack?: string } {
  if (error instanceof Error) {
    return { name: error.name, message: error.message, stack: error.stack };
  }
  return { message: String(error) };
}

async function attachLongFootnoteEvidence(
  page: Page,
  testInfo: TestInfo,
  scenario: string,
  result: LongFootnoteResult | undefined,
  failure: ReturnType<typeof describeFailure> | undefined,
): Promise<void> {
  const captureWarnings: string[] = [];
  let html: Buffer | undefined;
  let screenshot: Buffer | undefined;
  try {
    html = Buffer.from(await page.content());
  } catch (error) {
    captureWarnings.push(`HTML capture failed: ${describeFailure(error).message}`);
  }
  try {
    const container = page.locator('#container');
    if (await container.count()) screenshot = await container.screenshot();
  } catch (error) {
    captureWarnings.push(`screenshot capture failed: ${describeFailure(error).message}`);
  }

  await testInfo.attach(`${scenario}.json`, {
    body: Buffer.from(JSON.stringify({ result, failure, captureWarnings }, null, 2)),
    contentType: 'application/json',
  });
  if (html) {
    await testInfo.attach(`${scenario}.html`, { body: html, contentType: 'text/html' });
  }
  if (screenshot) {
    await testInfo.attach(`${scenario}.png`, { body: screenshot, contentType: 'image/png' });
  }
}

async function withLongFootnoteEvidence(
  page: Page,
  testInfo: TestInfo,
  scenario: string,
  paginate: () => Promise<LongFootnoteResult>,
  verify: (result: LongFootnoteResult) => void | Promise<void>,
): Promise<void> {
  let result: LongFootnoteResult | undefined;
  let failure: ReturnType<typeof describeFailure> | undefined;
  try {
    result = await paginate();
    await verify(result);
  } catch (error) {
    failure = describeFailure(error);
    throw error;
  } finally {
    await attachLongFootnoteEvidence(page, testInfo, scenario, result, failure);
  }
}

/**
 * Geometry of every rendered page. `usedPct` is BODY + NOTES against the content box: notes
 * legitimately consume page height, so body extent alone understates how full a page is — the
 * measure that matters is how much of the page is doing any work at all.
 */
const MEASURE = () => {
  const container = document.getElementById('container') as HTMLElement;
  return Array.from(container.querySelectorAll('.page-box')).map((box, i) => {
    const content = box.querySelector('.page-content') as HTMLElement | null;
    const notes = box.querySelector('.page-footnotes') as HTMLElement | null;
    const kids = content ? Array.from(content.children) : [];
    const contentTop = content ? content.getBoundingClientRect().top : 0;
    const bodyBottom = kids.length
      ? kids[kids.length - 1].getBoundingClientRect().bottom
      : contentTop;
    const avail = content ? content.getBoundingClientRect().height : 0;
    const bodyPx = bodyBottom - contentTop;
    const notesPx = notes ? notes.getBoundingClientRect().height : 0;
    return {
      page: i + 1,
      overlapPx: notes ? Math.round(bodyBottom - notes.getBoundingClientRect().top) : 0,
      usedPct: avail ? Math.round(((bodyPx + notesPx) / avail) * 100) : 0,
      bodyPx: Math.round(bodyPx),
      notesPx: Math.round(notesPx),
      hasNotes: !!notes,
    };
  });
};

/** Every note id cited on a page must be rendered somewhere in the document. */
const MEASURE_NOTES = () => {
  const container = document.getElementById('container') as HTMLElement;
  const renderedItems = Array.from(container.querySelectorAll<HTMLElement>('.footnote-item'));
  const rendered = new Set(renderedItems.map((i) => i.getAttribute('data-footnote-id')));
  const cited = Array.from(
    new Set(Array.from(container.querySelectorAll('[data-footnote-id]'))
      .filter((e) => e.tagName === 'SUP' || e.tagName === 'A')
      .map((e) => e.getAttribute('data-footnote-id'))),
  );
  return {
    cited: cited.length,
    rendered: rendered.size,
    lost: cited.filter((id) => !rendered.has(id)),
    nested: container.querySelectorAll('.footnote-item .footnote-item').length,
    clipped: renderedItems.filter((item) => {
      const band = item.closest('.page-footnotes');
      if (!band) return true;
      const itemRect = item.getBoundingClientRect();
      const bandRect = band.getBoundingClientRect();
      return itemRect.top < bandRect.top - 1 || itemRect.bottom > bandRect.bottom + 1;
    }).map((item) => {
      const band = item.closest('.page-footnotes')!;
      const itemRect = item.getBoundingClientRect();
      const bandRect = band.getBoundingClientRect();
      return {
        id: item.dataset.footnoteId,
        itemTop: Math.round(itemRect.top),
        itemBottom: Math.round(itemRect.bottom),
        bandTop: Math.round(bandRect.top),
        bandBottom: Math.round(bandRect.bottom),
      };
    }),
  };
};

async function paginate(page: Page, html: string) {
  await page.setContent(html);
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(
    ({ measure }: { measure: string }) => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
      // eslint-disable-next-line no-eval
      return (0, eval)(`(${measure})`)();
    },
    { measure: MEASURE.toString() },
  );
}

test.describe('Paginated footnote layout', () => {
  test('body text is never drawn underneath the footnote area', async ({ page }) => {
    // Many notes on one page is the shape that produced overlapping glyphs on the real document.
    const pages = await paginate(page, stagingHtml(14, 8, 3));

    const overlapping = pages.filter((p: any) => p.overlapPx > 1);
    expect(
      overlapping,
      `pages where body overlaps notes: ${JSON.stringify(overlapping)}`,
    ).toEqual([]);
  });

  test('reserving footnote space does not leave the rest of the page empty', async ({ page }) => {
    // One early citation reserves note space; the remaining body must still fill the page.
    const pages = await paginate(page, stagingHtml(40, 3, 4));

    // Every page but the last should be reasonably full, counting body AND notes.
    const starved = pages.slice(0, -1).filter((p: any) => p.usedPct < 80);
    expect(
      starved,
      `under-filled pages: ${JSON.stringify(starved)} of ${pages.length}`,
    ).toEqual([]);
  });

  test('no cited note is dropped when several will not fit on one page', async ({ page }) => {
    // The engine carried an unfitted note in a SINGLE continuation slot, so when two notes on the
    // same page both failed to fit, the second overwrote the first and that note vanished from the
    // document — four notes disappeared from a real 94-footnote file. Deferred notes now queue.
    await page.setContent(stagingHtml(12, 10, 10));
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const notes = await page.evaluate(
      ({ measure }: { measure: string }) => {
        const staging = document.getElementById('staging') as HTMLElement;
        const container = document.getElementById('container') as HTMLElement;
        const { PaginationEngine } = (window as any).DocxodusPagination;
        new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
        // eslint-disable-next-line no-eval
        return (0, eval)(`(${measure})`)();
      },
      { measure: MEASURE_NOTES.toString() },
    );

    expect(notes.cited).toBeGreaterThan(0);
    expect(notes.lost, `notes cited but never rendered: ${JSON.stringify(notes.lost)}`).toEqual([]);
    // A note must never be wrapped inside another note (breaks the number/text line).
    expect(notes.nested).toBe(0);
  });

  test('a page whose notes are capped still fills its remaining body space', async ({ page }) => {
    // Notes long enough to hit the 60% cap: the other 40% must still take body text.
    const pages = await paginate(page, stagingHtml(30, 6, 8));

    const withNotes = pages.filter((p: any) => p.hasNotes);
    expect(withNotes.length).toBeGreaterThan(0);
    // No page may sit mostly idle while body blocks are still pending.
    const wasteful = pages.slice(0, -1).filter((p: any) => p.usedPct < 80);
    expect(
      wasteful,
      `pages left nearly empty: ${JSON.stringify(wasteful)} of ${pages.length}`,
    ).toEqual([]);
  });

  test('whole notes deferred from the final body block drain onto note-only pages', async ({ page }) => {
    await page.setContent(finalBlockDeferralHtml());
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const result = await page.evaluate(
      ({ measure }: { measure: string }) => {
        const staging = document.getElementById('staging') as HTMLElement;
        const container = document.getElementById('container') as HTMLElement;
        const { PaginationEngine } = (window as any).DocxodusPagination;
        new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
        // eslint-disable-next-line no-eval
        const notes = (0, eval)(`(${measure})`)();
        return {
          ...notes,
          noteOnlyPages: Array.from(container.querySelectorAll('.page-box')).filter((box) =>
            box.querySelector('.page-footnotes') && !box.querySelector('.page-content')?.children.length,
          ).length,
        };
      },
      { measure: MEASURE_NOTES.toString() },
    );
    expect(result.cited).toBe(4);
    expect(result.lost, `notes cited but never rendered: ${JSON.stringify(result.lost)}`).toEqual([]);
    expect(result.clipped, `notes rendered outside their visible band: ${JSON.stringify(result.clipped)}`)
      .toEqual([]);
    expect(result.noteOnlyPages).toBeGreaterThan(0);
  });

  test('splits one long footnote paragraph across as many note bands as needed', async ({
    page,
  }, testInfo) => {
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-long-footnote',
      () => paginateLongFootnote(page),
      (result) => {
        expect(result.totalPages).toBeGreaterThan(2);
        expect(result.notePages).toBe(result.totalPages);
        expect(result.paragraphFragments).toBeGreaterThan(2);
        expect(result.pageMapFragments.length).toBe(result.paragraphFragments);
        expect(new Set(result.pageMapFragments.map((fragment: any) => fragment.pageNumber)).size)
          .toBeGreaterThan(2);
        expect(result.pageMapFragments.map((fragment: any) => fragment.fragmentIndex))
          .toEqual(result.pageMapFragments.map((_: any, index: number) => index));
        expect(new Set(result.pageMapFragments.map((fragment: any) => fragment.fragmentId)).size)
          .toBe(result.pageMapFragments.length);
        expect(result.pageMapFragments.every((fragment: any) =>
          fragment.story === 'footnote'
          && fragment.geometry.x >= 0
          && fragment.geometry.y >= 0
          && fragment.geometry.width > 0
          && fragment.geometry.height > 0)).toBe(true);
        expect(result.bareIds).toBe(1);
        expect(result.bareAnchors).toBe(1);
        expect(result.visibleNoteText).toBe(LONG_NOTE_WORDS);
        expect(result.clipped).toEqual([]);
        expect(result.firstPageText).toContain('word-0');
      },
    );
  });

  test('keeps a long leading paragraph with its citation while later note paragraphs continue', async ({
    page,
  }, testInfo) => {
    const trailing = `
      <p data-source-anchor-id="p:fn:after-one">after-one</p>
      <p data-source-anchor-id="p:fn:after-two">after-two</p>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-leading-paragraph-with-siblings',
      () => paginateLongFootnote(page, trailing),
      (result) => {
        expect(result.firstPageText).toContain('word-0');
        expect(result.pagesText.some((text: string) => text.includes('after-one'))).toBe(true);
        expect(result.pagesText.some((text: string) => text.includes('after-two'))).toBe(true);
        expect(result.visibleNoteText).toBe(`${LONG_NOTE_WORDS} after-one after-two`);
        expect(result.clipped).toEqual([]);
      },
    );
  });

  test('an indivisible oversized note paragraph does not swallow later siblings', async ({
    page,
  }, testInfo) => {
    const trailing = `
      <p data-source-anchor-id="p:fn:after-one">after-one</p>
      <p data-source-anchor-id="p:fn:after-two">after-two</p>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-indivisible-leading-paragraph',
      () => paginateLongFootnote(
        page,
        trailing,
        '<span class="unsafe-note-box">oversized-unsafe</span>',
        '.unsafe-note-box { display: inline-block; height: 150pt; width: 10pt; }',
      ),
      (result) => {
        expect(result.totalPages).toBeGreaterThan(1);
        expect(result.firstPageText).toContain('oversized-unsafe');
        expect(result.firstPageText).not.toContain('after-one');
        expect(result.pagesText.slice(1).some((text: string) =>
          text.includes('after-one'))).toBe(true);
        expect(result.pagesText.slice(1).some((text: string) =>
          text.includes('after-two'))).toBe(true);
        expect(result.pageMapFragments).toHaveLength(1);
      },
    );
  });

  test('drains an older continuation before admitting later body and note payloads', async ({
    page,
  }) => {
    const result = await paginateScenario(page, continuationThenNewNoteHtml());
    expect(result.totalPages).toBeGreaterThan(3);
    expect(result.bodyByPage.flat()).toEqual([
      'cite first 1',
      'body after first A',
      'cite second 2',
      'body after second B',
    ]);
    const noteText = result.noteTextByPage.join(' ');
    for (const token of ['first-0', 'first-70', 'first-139', 'second-note-alpha', 'second-note-omega']) {
      expect(noteText.match(new RegExp(`\\b${token}\\b`, 'g'))).toHaveLength(1);
    }
    expect(result.f1Parts).toBeGreaterThan(2);
    expect(result.f2Parts).toBe(1);
    expect(result.numberText).toEqual(['1', '2']);
    expect(result.clipped).toEqual([]);
    expect(result.overlaps).toEqual([]);
  });

  test('defers a splittable paragraph when no safe prefix fits the citation band', async ({
    page,
  }) => {
    const result = await paginateScenario(page, noPrefixFitsCitationBandHtml());
    expect(result.totalPages).toBe(2);
    expect(result.bodyByPage).toEqual([['citation 1'], []]);
    expect(result.noteTextByPage[0]).toBe('');
    expect(result.noteTextByPage[1]).toContain('fresh-band-safe tail-safe');
    expect(result.freshPages).toEqual([2]);
    expect(result.numberText).toEqual(['1']);
    expect(result.clipped).toEqual([]);
    expect(result.overlaps).toEqual([]);
  });

  test('defers indivisible content that fits a fresh band instead of clipping the residual band', async ({
    page,
  }) => {
    const result = await paginateScenario(page, indivisibleFitsFreshBandHtml());
    expect(result.totalPages).toBe(2);
    expect(result.noteTextByPage[0]).toBe('');
    expect(result.noteTextByPage[1]).toContain('indivisible-safe later-safe-sibling');
    expect(result.unsafeFreshPages).toEqual([2]);
    expect(result.numberText).toEqual(['1']);
    expect(result.clipped).toEqual([]);
    expect(result.overlaps).toEqual([]);
  });

  test('preserves Unicode clusters and semantic inline/table content across a long note', async ({
    page,
  }, testInfo) => {
    const linkText = Array.from({ length: 45 }, (_, index) => `linked-${index} `).join('');
    const cjk = '漢字仮名交じり文。'.repeat(24);
    const payload = `• A\u00a0B C\u2060D e\u0301 👨‍👩‍👧‍👦 שלום 42 [7] ` +
      `LatinCrossRunBoundary ${linkText}${cjk}`;
    const leading = `<span data-case="rich-payload" data-anchor="inline-note-anchor">` +
      `<span data-list-marker="true">• </span>` +
      `<a id="note-bookmark"></a>` +
      `<span>A\u00a0B C\u2060D e\u0301 </span>` +
      `<span data-cluster-part="left">👨</span><span data-cluster-part="right">‍👩‍👧‍👦</span>` +
      `<b dir="rtl"> שלום </b>` +
      `<span data-field="PAGE">42</span> ` +
      `<a id="comment-ref-7" href="#comment-7" data-comment-id="7" ` +
      `>[7]</a> ` +
      `<span data-case="cross-run"><span>LatinCross</span><span>RunBoundary</span></span> ` +
      `<a data-case="link" href="https://example.test/note">${linkText}</a>` +
      `<span data-case="cjk">${cjk}</span></span>`;
    const trailing = `<table data-case="note-table" data-source-anchor-id="tbl:fn:tail-table">` +
      `<tbody><tr><td>tail-table-cell</td></tr></tbody></table>`;

    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-rich-lossless-continuation',
      () => paginateLongFootnote(page, trailing, leading),
      (result) => {
        expect(result.totalPages).toBeGreaterThan(3);
        expect(result.richPayloadFragments).toBeGreaterThan(2);
        expect(result.richPayloadText).toBe(payload);
        expect(result.crossRunFragments).toEqual(['LatinCrossRunBoundary']);
        expect(result.cjkFragments.join('')).toBe(cjk);
        expect(result.cjkFragments.every((fragment: string) =>
          !fragment.startsWith('。') && !/[「『（［｛]$/.test(fragment))).toBe(true);
        expect(result.clusterPartCounts).toEqual([1, 1]);
        expect(result.listMarkerPages).toHaveLength(1);
        expect(result.bookmarkTargetCount).toBe(1);
        expect(result.inlineAnchorCount).toBe(1);
        expect(result.fieldCount).toBe(1);
        expect(result.commentMarkerCount).toBe(1);
        expect(result.marginCommentCount).toBe(1);
        expect(result.marginCommentText).toBe('margin-comment-seven');
        expect(result.commentText).toContain('[7]');
        expect(result.commentText).toContain('margin-comment-seven');
        expect(result.commentPageMapFragments).toBe(1);
        expect(result.linkHrefs.length).toBeGreaterThan(1);
        expect(result.linkHrefs.every((href: string | null) =>
          href === 'https://example.test/note')).toBe(true);
        expect(result.linkText).toBe(linkText);
        expect(result.tableText).toBe('tail-table-cell');
        expect(result.tablePageMapFragments).toBe(1);
        expect(result.numberCount).toBe(1);
        expect(result.separatorCount).toBe(result.notePages);
        expect(result.continuationFootnoteIds.every((id: string) => id === 'f-long')).toBe(true);
        expect(result.clipped).toEqual([]);
      },
    );
  });

  test('continues inside the source note/content shells without duplicating their identities', async ({
    page,
  }, testInfo) => {
    const trailing = `<p data-source-anchor-id="p:fn:styled-tail">${LONG_NOTE_WORDS}</p>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-continuation-shell-inheritance',
      () => paginateLongFootnote(
        page,
        trailing,
        'short first paragraph',
        `[data-shell-token="kept"] > [data-content-token="kept"] { color: rgb(1, 2, 3); }
         [data-content-token="kept"] p:not(:first-of-type) { margin-left: 19px; }`,
        `id="source-note-item" data-anchor="source-item-anchor"
         data-shell-token="kept" dir="rtl" lang="ar"`,
        `id="source-note-content" data-anchor="source-content-anchor"
         data-content-token="kept"`,
      ),
      (result) => {
        expect(result.continuationShells.length).toBeGreaterThan(2);
        expect(result.continuationShells.every((shell: any) =>
          shell.isItem
          && shell.hasDirectContent
          && shell.shellToken === 'kept'
          && shell.contentToken === 'kept'
          && shell.direction === 'rtl'
          && shell.lang === 'ar'
          && shell.color === 'rgb(1, 2, 3)'
          && shell.display === 'block'
          && shell.marginLeft === '19px'
          && shell.sentinelCount === 1
          && shell.duplicateIds === 0
          && shell.duplicateAnchors === 0)).toBe(true);
        expect(result.visibleNoteText).toBe(`short first paragraph ${LONG_NOTE_WORDS}`);
        expect(result.clipped).toEqual([]);
      },
    );
  });

  test('selects and sanitizes authored normal and continuation separator stories', async ({
    page,
  }, testInfo) => {
    const separatorStories = `
      <div data-footnote-separator="normal" id="separator-source"
           data-anchor="separator-anchor" data-source-anchor-id="p:separator-source">
        <span>NORMAL-STORY</span><a href="#separator-source">local</a>
        <a href="https://example.test/separator">external</a>
        <span data-footnote-id="not-a-definition">nested-semantic</span>
      </div>
      <div data-footnote-separator="continuation" id="separator-source"
           data-anchor="separator-anchor" data-source-anchor-id="p:separator-source">
        <span>CONTINUATION-STORY</span><a href="#separator-source">local</a>
        <a href="https://example.test/separator">external</a>
      </div>`;
    await withLongFootnoteEvidence(
      page,
      testInfo,
      'issue-489-custom-footnote-separators',
      () => paginateLongFootnote(page, '', LONG_NOTE_WORDS, '', '', '', separatorStories),
      (result) => {
        expect(result.notePages).toBeGreaterThan(2);
        expect(result.separatorCount).toBe(result.notePages);
        expect(result.separatorKinds).toEqual([
          'normal',
          ...Array.from({ length: result.notePages - 1 }, () => 'continuation'),
        ]);
        expect(result.separatorText[0]).toContain('NORMAL-STORY');
        expect(result.separatorText.slice(1).every((text: string) =>
          text.includes('CONTINUATION-STORY') && !text.includes('NORMAL-STORY'))).toBe(true);
        expect(result.separatorFirst).toBe(true);
        expect(result.separatorIdentityCount).toBe(0);
        expect(result.separatorLocalLinks).toBe(0);
        expect(result.separatorExternalLinks).toHaveLength(result.notePages);
        expect(result.visibleNoteText).toBe(LONG_NOTE_WORDS);
        expect(result.clipped).toEqual([]);
      },
    );
  });

  test('checks the prospective page cap before appending a continuation page', async ({ page }) => {
    await page.setContent(longParagraphFootnoteHtml());
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const result = await page.evaluate(() => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const checks: number[] = [];
      let message = '';
      try {
        new PaginationEngine(staging, container, {
          showPageNumbers: false,
          checkPageCount: (count: number) => {
            checks.push(count);
            if (count > 1) throw new Error('finalPages limit 1');
          },
        }).paginate();
      } catch (error) {
        message = String(error);
      }
      return {
        checks,
        message,
        pages: container.querySelectorAll('.page-box').length,
      };
    });

    expect(result.message).toContain('finalPages limit 1');
    expect(result.checks).toEqual([1, 2]);
    expect(result.pages).toBe(1);
  });

  test('charges the next page before cloning an adversarial continuation tail', async ({ page }) => {
    const tail = Array.from({ length: 1_200 }, (_, index) =>
      `<p style="height:0;line-height:0;margin:0">tail-${index}</p>`).join('');
    await page.setContent(longParagraphFootnoteHtml(tail));
    await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
    const result = await page.evaluate(() => {
      const staging = document.getElementById('staging') as HTMLElement;
      const container = document.getElementById('container') as HTMLElement;
      const { PaginationEngine } = (window as any).DocxodusPagination;
      const originalCloneNode = Node.prototype.cloneNode;
      let cloneCalls = 0;
      Node.prototype.cloneNode = function patchedCloneNode(deep?: boolean) {
        cloneCalls++;
        return originalCloneNode.call(this, deep);
      };
      const admissions: Array<{ page: number; cloneCalls: number }> = [];
      let message = '';
      try {
        new PaginationEngine(staging, container, {
          showPageNumbers: false,
          checkPageCount: (prospectivePage: number) => {
            admissions.push({ page: prospectivePage, cloneCalls });
            if (prospectivePage === 2) throw new Error('finalPages adversarial cap');
          },
          checkCancellation: (() => {
            let checks = 0;
            return () => {
              if (++checks > 50_000) throw new Error('unbounded continuation work');
            };
          })(),
        }).paginate();
      } catch (error) {
        message = String(error);
      } finally {
        Node.prototype.cloneNode = originalCloneNode;
      }
      return {
        admissions,
        message,
        pages: container.querySelectorAll('.page-box').length,
      };
    });

    expect(result.message).toContain('finalPages adversarial cap');
    expect(result.admissions.map((entry) => entry.page)).toEqual([1, 2]);
    // Page-one materialization may clone the visible head, but the 1,200-node
    // tail must not be partitioned/recloned before page two is admitted.
    expect(result.admissions[1].cloneCalls - result.admissions[0].cloneCalls).toBeLessThan(250);
    expect(result.pages).toBe(1);
  });
});
