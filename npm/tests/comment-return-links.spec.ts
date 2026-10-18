import { test, expect } from '@playwright/test';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { commentReturnFixture } from './docx-comment-return-fixture.js';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function synthetic(page: import('@playwright/test').Page,
  placement: 'margin' | 'outside', target: 'body' | 'missing' | 'hidden' | 'ambiguous' | 'hidden-body' | 'transparent-body' | 'clipped-body' | 'visible-excluded', later = false, linkAttributes = '') {
  const anchor = `<a id="comment-ref-6"${target === 'hidden-body' ? ' style="display:none"' : ''}${target === 'visible-excluded' ? ' data-page-map-exclude="true"' : ''}>MapleTarget</a>`;
  const marker = target === 'transparent-body' ? `<span style="opacity:0">${anchor}</span>` :
    target === 'clipped-body' ? `<span style="display:inline-block;width:0;height:0;overflow:hidden">${anchor}</span>` : anchor;
  await page.setContent(`<style>body{margin:0}p{margin:0 0 12px}</style>
    ${target === 'hidden' || target === 'ambiguous' ? `<div hidden>${marker}</div>` : ''}
    <div id="pagination-staging"><aside id="pagination-comment-margin-registry" style="display:none">
      <div data-comment-id="6"><aside class="review-margin-note" id="comment-6">
        <a class="review-margin-backref" href="#comment-ref-6" ${linkAttributes}>ReturnMaple</a>
        <a class="authored-link" href="#MapleDestination">Other fragment</a><p>Maple note.</p>
      </aside></div></aside>
      <div data-section-index="0" data-page-width="612" data-page-height="792"
        data-content-width="468" data-content-height="648" data-margin-top="72"
        data-margin-right="72" data-margin-bottom="72" data-margin-left="72">
        <p><span data-comment-id="6">Maple comment range.</span>${!later && ['body', 'ambiguous', 'hidden-body', 'transparent-body', 'clipped-body', 'visible-excluded'].includes(target) ? marker : ''}</p>
        ${later ? `<p data-page-break-before="true">${target === 'body' ? marker : ''}Second page.</p>` : ''}
        <p id="MapleDestination">Maple bookmark.</p><p><a href="#MapleDestination" class="body-return">Body return</a></p>
      </div></div><div id="host"></div>`);
  await page.addScriptTag({ path: resolve(bundle) });
  await page.evaluate(placement => {
    const staging = document.querySelector('#pagination-staging')!;
    const html = staging.outerHTML + '<div id="pagination-container" class="page-container"></div>';
    staging.remove();
    (window as any).DocxodusPagination.paginateHtml(html,
      document.querySelector('#host'), { showPageNumbers: false, commentColumnPlacement: placement });
  }, placement);
}

for (const placement of ['margin', 'outside'] as const) test(`unique rendered reference retains its return link (${placement})`, async ({ page }) => {
  await synthetic(page, placement, 'body');
  const link = page.locator('#host .page-comment-margin .review-margin-backref').first();
  await expect(page.locator('#host .page-content #comment-ref-6')).toHaveCount(1);
  await expect(link).toHaveAttribute('href', '#comment-ref-6');
  await expect(link).not.toHaveAttribute('aria-disabled', 'true');
  expect(await link.evaluate(element => (element as HTMLAnchorElement).tabIndex)).toBe(0);
  await expect(page.locator('#host .page-comment-margin [id]')).toHaveCount(0);
  await expect(page.locator('#host .page-comment-margin .authored-link')).not.toHaveAttribute('href');
  await expect(page.locator('#host .page-content .body-return')).toHaveAttribute('href', '#MapleDestination');
});

for (const target of ['missing', 'hidden', 'ambiguous', 'hidden-body', 'transparent-body', 'clipped-body'] as const) test(`unavailable or ambiguous reference stays inert (${target})`, async ({ page }) => {
  await synthetic(page, 'outside', target);
  const link = page.locator('#host .page-comment-margin .review-margin-backref').first();
  await expect(link).not.toHaveAttribute('href');
  await expect(link).toHaveAttribute('aria-disabled', 'true');
  expect(await link.evaluate(element => (element as HTMLAnchorElement).tabIndex)).toBe(-1);
});

test('return navigation resolves after the target on a later page is rendered', async ({ page }) => {
  await synthetic(page, 'outside', 'body', true);
  await expect(page.locator('#host .page-box')).toHaveCount(2);
  const link = page.locator('#host .page-comment-margin .review-margin-backref').first();
  await expect(link).toHaveAttribute('href', '#comment-ref-6');
  await link.focus(); await page.keyboard.press('Enter');
  expect(await page.evaluate(() => location.hash)).toBe('#comment-ref-6');
  expect(await page.evaluate(() => scrollY)).toBeGreaterThan(500);
});

test('a visible UI reference outside the PageMap source inventory remains a return target', async ({ page }) => {
  await synthetic(page, 'outside', 'visible-excluded');
  await expect(page.locator('#host .page-comment-margin .review-margin-backref')).toHaveAttribute('href', '#comment-ref-6');
});

test('return links preserve explicit focus and accessibility attributes', async ({ page }) => {
  await synthetic(page, 'outside', 'body', false, 'tabindex="2" aria-disabled="false" contenteditable="true"');
  const link = page.locator('#host .page-comment-margin .review-margin-backref');
  await expect(link).toHaveAttribute('href', '#comment-ref-6');
  await expect(link).toHaveAttribute('tabindex', '2');
  await expect(link).toHaveAttribute('aria-disabled', 'false');
  await expect(link).toHaveAttribute('contenteditable', 'false');
});

test('a deliberately disabled source return link stays inert', async ({ page }) => {
  await synthetic(page, 'outside', 'body', false, 'aria-disabled="true"');
  const link = page.locator('#host .page-comment-margin .review-margin-backref');
  await expect(link).not.toHaveAttribute('href');
  await expect(link).toHaveAttribute('aria-disabled', 'true');
});

for (const { later, frame } of [{ later: false, frame: false }, { later: true, frame: false }, { later: false, frame: true }]) test(`native margin return link remains active (later target=${later}, frame=${frame})`, async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ path: bundle });
  await page.evaluate(bytes => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Maple return', 'docx-', false, '', 2, 'comment-', 1, 1,
      'page-', false, 0, 'annot-', false, false, false, true, true);
    if (!html.startsWith('<')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = '<div id="host"></div>';
    (window as any).DocxodusPagination.paginateHtml(html, document.querySelector('#host'), {
      showPageNumbers: false, commentColumnPlacement: 'margin',
    });
  }, Array.from(commentReturnFixture(true, later, frame)));
  await expect(page.locator('#host .page-box #comment-ref-6')).toHaveCount(1);
  if (frame) await expect(page.locator('#host [data-docx-text-frame] #comment-ref-6')).toHaveCount(1);
  const link = page.locator('#host .page-comment-margin .comment-margin-backref').first();
  await expect(link).toHaveAttribute('href', '#comment-ref-6');
  await expect(page.locator('#host .page-content a', { hasText: 'MapleBodyReturn' })).toHaveAttribute('href', '#MapleDestination');
  await link.focus(); await page.keyboard.press('Enter');
  expect(await page.evaluate(() => location.hash)).toBe('#comment-ref-6');
  if (later) expect(await page.evaluate(() => scrollY)).toBeGreaterThan(500);
});
