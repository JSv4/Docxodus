import { test, expect, Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function section(blocks: string, footnoteRegistry = ''): string {
  return `
    <style>
      #staging { font: 12px/12px Arial; }
      #staging p { box-sizing: border-box; margin: 0; padding: 0; width: 100%; }
    </style>
    <div id="staging">
      <div data-section-index="0"
           data-page-width="102" data-page-height="102"
           data-content-width="100" data-content-height="100"
           data-margin-top="1" data-margin-right="1"
           data-margin-bottom="1" data-margin-left="1">
        ${blocks}
      </div>
      ${footnoteRegistry}
    </div>
    <div id="container"></div>`;
}

/** A footnote whose paragraphs are `heights` points tall. */
function note(id: string, number: string, heights: number[]): string {
  return `
    <div class="footnote-item" data-footnote-id="${id}">
      <span class="footnote-number">${number}</span>
      <div class="footnote-content">${heights.map((height, index) =>
        `<p style="height: ${height}pt; margin: 0">${id}-${index}</p>`).join('')}</div>
    </div>`;
}

/** Three 30pt paragraphs: one measures 40.5pt with the separator, the last two 73.5pt. */
const threeParagraphNote = note('f1', '1', [30, 30, 30]);

function registry(...notes: string[]): string {
  return `<div id="pagination-footnote-registry">${notes.join('')}</div>`;
}

async function paginate(page: Page, fragmentParagraphs = false): Promise<string[][]> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate((fragmentParagraphs) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false, fragmentParagraphs }).paginate();

    return Array.from(container.querySelectorAll<HTMLElement>('.page-content')).map(content =>
      Array.from(content.children).map(block => (block.textContent || '').trim())
    );
  }, fragmentParagraphs);
}

async function paginateWithFootnoteInfo(page: Page, fragmentParagraphs = false): Promise<{
  content: string[][];
  footnotePages: boolean[];
}> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate((fragmentParagraphs) => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false, fragmentParagraphs }).paginate();

    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    return {
      content: pages.map(page =>
        Array.from(page.querySelector<HTMLElement>('.page-content')!.children).map(
          block => (block.textContent || '').trim()
        )
      ),
      footnotePages: pages.map(page => page.querySelector('.page-footnotes') !== null),
    };
  }, fragmentParagraphs);
}

test.describe('Pagination keep-with-next chains', () => {
  for (const mode of ['split', 'keep-lines', 'widow-control', 'page-break', 'fragmentation-disabled']) {
    test(`a heading keeps a valid first paragraph fragment: ${mode}`, async ({ page }) => {
      const lines = Array.from({ length: 6 }, (_, i) => `Cedar${i + 1}`);
      await page.setContent(section(`
        <p style="height: 40pt">lead</p>
        <p data-keep-with-next="true" style="height: 20pt">heading</p>
        <p data-keep-lines="${mode === 'keep-lines'}" data-page-break-before="${mode === 'page-break'}"
           data-widow-control="${mode === 'widow-control'}"
           style="font: 10pt/10pt Arial">${lines.join('<br>')}</p>`));

      const pages = await paginate(page, mode !== 'fragmentation-disabled');
      expect(pages).toEqual(mode === 'split'
        ? [['lead', 'heading', lines.slice(0, 4).join('')], [lines.slice(4).join('')]]
        : mode === 'page-break'
          ? [['lead', 'heading'], [lines.join('')]]
          : [['lead'], ['heading', lines.join('')]]);
    });
  }

  for (const [citationLine, leadHeight] of [[0, 15], [5, 15], [0, 30]]) {
    test(`a kept paragraph fragment carries line ${citationLine + 1}'s note after ${leadHeight}pt`, async ({ page }) => {
      const lines = Array.from({ length: 6 }, (_, i) => `Cedar${i + 1}`);
      const markup = lines.map((line, i) => line + (i === citationLine ? '<sup data-footnote-id="f1">1</sup>' : ''));
      await page.setContent(section(`<p style="height: ${leadHeight}pt">lead</p>
        <p data-keep-with-next="true" style="height: 20pt">heading</p>
        <p style="font: 10pt/10pt Arial">${markup.join('<br>')}</p>`, registry(note('f1', '1', [10]))));
      const result = await paginateWithFootnoteInfo(page, true);
      const text = lines.map((line, i) => line + (i === citationLine ? '1' : ''));
      const firstLines = citationLine === 0 ? 2 : 5;
      expect(result.content).toEqual(leadHeight === 30
        ? [['lead'], ['heading', text.slice(0, 3).join('')], [text.slice(3).join('')]]
        : [['lead', 'heading', text.slice(0, firstLines).join('')], [text.slice(firstLines).join('')]]);
      expect(result.footnotePages).toEqual(leadHeight === 30 ? [false, true, false]
        : citationLine === 0 ? [true, false] : [false, true]);
    });
  }

  test('moves a feasible multi-block keep chain to the next page', async ({ page }) => {
    await page.setContent(section(`
      <p style="height: 50pt">lead</p>
      <p data-keep-with-next="true" style="height: 20pt">heading</p>
      <p data-keep-with-next="true" style="height: 20pt">subheading</p>
      <p style="height: 20pt">body</p>`));

    const pages = await paginate(page);

    expect(pages).toEqual([
      ['lead'],
      ['heading', 'subheading', 'body'],
    ]);
  });

  test('keeps greedy placement for a chain too tall for a fresh page', async ({ page }) => {
    await page.setContent(section(`
      <p style="height: 60pt">lead</p>
      <p data-keep-with-next="true" style="height: 20pt">heading</p>
      <p data-keep-with-next="true" style="height: 50pt">subheading</p>
      <p style="height: 50pt">body</p>`));

    const pages = await paginate(page);

    expect(pages).toEqual([
      ['lead', 'heading'],
      ['subheading', 'body'],
    ]);
  });

  test('starts a new keep chain after page-break-before with a footnote continuation', async ({ page }) => {
    await page.setContent(section(
      `
        <p data-keep-with-next="true" style="height: 20pt">lead<sup data-footnote-id="f1">1</sup></p>
        <p data-page-break-before="true" data-keep-with-next="true" style="height: 25pt">heading</p>
        <p style="height: 25pt">body</p>`,
      `<div id="pagination-footnote-registry">
        <div class="footnote-item" data-footnote-id="f1">
          <span class="footnote-number">1</span>
          <div class="footnote-content">
            <p style="height: 30pt; margin: 0">first</p>
            <p style="height: 30pt; margin: 0">second</p>
            <p style="height: 30pt; margin: 0">third</p>
          </div>
        </div>
      </div>`
    ));

    const result = await paginateWithFootnoteInfo(page);

    expect(result.content).toEqual([
      ['lead1'],
      [],
      ['heading', 'body'],
    ]);
    // The note is 3 x 30pt and the band carries the separator: one paragraph measures 40.5pt and
    // two 70.5pt. Page 1 shows the first beside `lead`. Page 2 opens with the 70.5pt tail, which
    // ends there but leaves 29.5pt, too little for the 50pt chain the page break starts; as in
    // Word reference PDFs, the tail keeps the page and the chain opens page 3 whole rather than
    // leaving `heading` alone at page 2's foot (issue #1178).
    expect(result.footnotePages).toEqual([true, true, false]);
  });

  test('a chain that a carried note leaves no room for waits for the page where it fits', async ({ page }) => {
    await page.setContent(section(
      `
        <p style="height: 20pt">lead<sup data-footnote-id="f1">1</sup></p>
        <p data-keep-with-next="true" style="height: 20pt">heading</p>
        <p style="height: 20pt">body</p>`,
      registry(threeParagraphNote)
    ));

    const result = await paginateWithFootnoteInfo(page);

    // `heading` would fit beside the first note paragraph on page 1, `body` would not. The 73.5pt
    // tail opens page 2 and ends there, leaving 26.5pt, too little for the 40pt chain. As in Word
    // reference PDFs, the tail keeps page 2 and the chain opens page 3 whole (issue #1180); the
    // chain used to be refused the move because the whole tail was counted against page 2, and
    // `heading` was left alone at page 1's foot.
    expect(result.content).toEqual([['lead1'], [], ['heading', 'body']]);
    expect(result.footnotePages).toEqual([true, true, false]);
  });

  for (const [name, blocks, notes, content, footnotePages] of [
    // The issue's shape: the chain's paragraph cites a one-line note, which opens page 3 with it.
    ['whose paragraph cites a note', `<h4 data-keep-with-next="true" style="height: 20pt; margin: 0">heading</h4>
      <p style="height: 20pt">body<sup data-footnote-id="f2">2</sup></p>`, note('f2', '2', [10]),
      [['lead1'], [], ['heading', 'body2']], [true, true, true]],
    // A 45pt chain, not on any band boundary, and a block after it.
    ['of 45pt', `<h4 data-keep-with-next="true" style="height: 20pt; margin: 0">heading</h4>
      <p style="height: 25pt">body</p><p style="height: 20pt">after</p>`, '',
      [['lead1'], [], ['heading', 'body', 'after']], [true, true, false]],
    // The chain's paragraph cites a 60pt note: only the note's start must sit beside it, the rest
    // runs on.
    ['whose paragraph cites a long note', `<h4 data-keep-with-next="true" style="height: 20pt; margin: 0">heading</h4>
      <p style="height: 20pt">body<sup data-footnote-id="f2">2</sup></p><p style="height: 20pt">after</p>`,
      note('f2', '2', [20, 20, 20]),
      [['lead1'], [], ['heading', 'body2'], ['after']], [true, true, true, true]],
  ] as const) {
    test(`a chain ${name} is not stranded beside a carried note (issue #1180)`, async ({ page }) => {
      await page.setContent(section(
        `<p style="height: 20pt">lead<sup data-footnote-id="f1">1</sup></p>${blocks}`,
        registry(threeParagraphNote, notes)
      ));

      const result = await paginateWithFootnoteInfo(page);

      expect(result.content).toEqual(content);
      expect(result.footnotePages).toEqual(footnotePages);
    });
  }

  test('a chain waits behind a carried note that runs on for pages', async ({ page }) => {
    await page.setContent(section(
      `
        <p style="height: 20pt">lead<sup data-footnote-id="f1">1</sup></p>
        <h4 data-keep-with-next="true" style="height: 20pt; margin: 0">heading</h4>
        <p style="height: 20pt">body</p>`,
      registry(note('f1', '1', [30, 30, 30, 30, 30, 30]))
    ));

    const result = await paginateWithFootnoteInfo(page);

    // Pages 2 and 3 hold the note alone; the chain opens the page where it ends.
    expect(result.content).toEqual([['lead1'], [], [], ['heading', 'body']]);
    expect(result.footnotePages).toEqual([true, true, true, true]);
  });
});
