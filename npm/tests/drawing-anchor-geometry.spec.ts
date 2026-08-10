import { expect, test, type Page } from '@playwright/test';
import {
  generateDrawingAnchorDocx,
  type DrawingAnchorFixtureOptions,
} from './docx-drawing-anchor-fixture.js';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined, {
    timeout: 30_000,
  });
}

async function renderAnchor(page: Page, options: DrawingAnchorFixtureOptions) {
  const bytes = generateDrawingAnchorDocx(options);
  return page.evaluate((input) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input),
      'Document', 'docx-', false, '', -1, 'comment-',
      1, 1, 'page-',
      false, 0, 'annot-',
      false, false,
      false, false, false,
    );
    document.body.innerHTML = '<main id="drawing-anchor-root"></main>';
    const root = document.getElementById('drawing-anchor-root')!;
    (window as any).DocxodusPagination.paginateHtml(html, root, {
      scale: 1,
      showPageNumbers: false,
      fragmentParagraphs: false,
    });

    const pageBox = root.querySelector<HTMLElement>('.page-box')!;
    const content = pageBox.querySelector<HTMLElement>('.page-content')!;
    const anchor = pageBox.querySelector<HTMLElement>('[data-docx-drawing-anchor="true"]')!;
    const textBlock = anchor.firstElementChild as HTMLElement;
    const paragraph = content.querySelector<HTMLElement>('p')!;
    const prefix = Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
      .find(element => element.textContent === 'PREFIX');
    const pageRect = pageBox.getBoundingClientRect();
    const anchorRect = anchor.getBoundingClientRect();
    const textRect = textBlock.getBoundingClientRect();
    const paragraphRect = paragraph.getBoundingClientRect();
    const prefixRect = prefix?.getBoundingClientRect();

    return {
      anchor: {
        left: anchorRect.left - pageRect.left,
        top: anchorRect.top - pageRect.top,
        width: anchorRect.width,
        height: anchorRect.height,
      },
      text: {
        left: textRect.left - pageRect.left,
        top: textRect.top - pageRect.top,
      },
      paragraph: {
        left: paragraphRect.left - pageRect.left,
        top: paragraphRect.top - pageRect.top,
      },
      prefixRight: prefixRect ? prefixRect.right - pageRect.left : undefined,
      wrapLeftPt: Number(anchor.dataset.docxAnchorWrapLeft),
      parentClass: anchor.parentElement?.className,
    };
  }, Array.from(bytes));
}

test.describe('generated DrawingML anchor geometry', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
  });

  test('page offsets pin the stored extent and textbox text origin', async ({ page }) => {
    const geometry = await renderAnchor(page, {
      horizontal: { relativeFrom: 'page', offsetEmu: 914400 },
      vertical: { relativeFrom: 'page', offsetEmu: 1143000 },
    });

    expect(geometry.parentClass).toBe('page-box');
    expect(geometry.anchor.left).toBeCloseTo(96, 0);
    expect(geometry.anchor.top).toBeCloseTo(120, 0);
    expect(geometry.anchor.width).toBeCloseTo(192, 0);
    expect(geometry.anchor.height).toBeCloseTo(96, 0);
    expect(geometry.text.left).toBeCloseTo(108, 0); // 9pt bodyPr left inset
    expect(geometry.text.top).toBeCloseTo(128, 0); // 6pt bodyPr top inset
    expect(geometry.wrapLeftPt).toBe(36); // clearance is preserved, not added to left
  });

  test('margin offsets start at the declared page margins', async ({ page }) => {
    const geometry = await renderAnchor(page, {
      horizontal: { relativeFrom: 'margin', offsetEmu: 228600 },
      vertical: { relativeFrom: 'margin', offsetEmu: 304800 },
    });

    expect(geometry.anchor.left).toBeCloseTo(120, 0); // 72pt margin + 18pt offset
    expect(geometry.anchor.top).toBeCloseTo(128, 0); // 72pt margin + 24pt offset
  });

  test('column alignment and paragraph-relative vertical offsets use different bases', async ({ page }) => {
    const geometry = await renderAnchor(page, {
      horizontal: { relativeFrom: 'column', align: 'center' },
      vertical: { relativeFrom: 'paragraph', offsetEmu: 228600 },
    });

    expect(geometry.anchor.left).toBeCloseTo(312, 0);
    expect(geometry.anchor.top - geometry.paragraph.top).toBeCloseTo(24, 0);
  });

  test('character and line origins follow the laid-out anchor location', async ({ page }) => {
    const geometry = await renderAnchor(page, {
      horizontal: { relativeFrom: 'character', offsetEmu: 114300 },
      vertical: { relativeFrom: 'line', offsetEmu: 0 },
      prefix: 'PREFIX',
    });

    expect(geometry.prefixRight).toBeDefined();
    expect(geometry.anchor.left - geometry.prefixRight!).toBeCloseTo(12, 0);
    expect(geometry.anchor.top).toBeCloseTo(geometry.paragraph.top, 0);
  });
});
