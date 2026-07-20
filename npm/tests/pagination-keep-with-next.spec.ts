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

async function paginate(page: Page): Promise<string[][]> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();

    return Array.from(container.querySelectorAll<HTMLElement>('.page-content')).map(content =>
      Array.from(content.children).map(block => (block.textContent || '').trim())
    );
  });
}

async function paginateWithFootnoteInfo(page: Page): Promise<{
  content: string[][];
  footnotePages: boolean[];
}> {
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

  return page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { showPageNumbers: false }).paginate();

    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-box'));
    return {
      content: pages.map(page =>
        Array.from(page.querySelector<HTMLElement>('.page-content')!.children).map(
          block => (block.textContent || '').trim()
        )
      ),
      footnotePages: pages.map(page => page.querySelector('.page-footnotes') !== null),
    };
  });
}

test.describe('Pagination keep-with-next chains', () => {
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
    expect(result.footnotePages).toEqual([true, true, false]);
  });

  test('does not force a chain onto a fresh page without room for its pending continuation', async ({ page }) => {
    await page.setContent(section(
      `
        <p style="height: 20pt">lead<sup data-footnote-id="f1">1</sup></p>
        <p data-keep-with-next="true" style="height: 20pt">heading</p>
        <p style="height: 20pt">body</p>`,
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
      ['lead1', 'heading'],
      ['body'],
    ]);
    expect(result.footnotePages).toEqual([true, true]);
  });
});
