import { expect, test, type Browser } from '@playwright/test';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { imageIdentityDocx } from './docx-image-identity-fixture.js';
import { backgroundPictureDocx } from './docx-background-picture-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

/** Corrupt the healthy sibling only as the serialized document is reopened. */
function corruptReopenedImage(browser: Browser): Browser {
  return new Proxy(browser, { get(target, key) {
    if (key === 'newContext') return async (options: Parameters<Browser['newContext']>[0]) => {
      const context = await target.newContext(options);
      await context.addInitScript(() => {
        const descriptor = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'srcdoc')!;
        Object.defineProperty(HTMLIFrameElement.prototype, 'srcdoc', { ...descriptor, set(html: string) {
          if (html.includes('data-page-number=')) {
            html = html.replace(/(<img\b[^>]*\bsrc=")data:image\/png;base64,iVBOR[^"]+/g,
              '$1data:image/png;base64,SnVuaXBlciBjb3JydXB0ZWQ=');
          }
          descriptor.set!.call(this, html);
        } });
      });
      return context;
    };
    const value = Reflect.get(target, key);
    return typeof value === 'function' ? value.bind(target) : value;
  } });
}

for (const promoted of [false, true]) {
  test(`warned ${promoted ? 'promoted' : 'inline'} image retains its identity through serialization`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(imageIdentityDocx(promoted), { browser, reviewProfile: 'markup', commentProfile: 'hidden', outputs: ['html', 'pdf'] });
    expect(result.renderReport.warnings.filter(warning => warning.code === 'image_decode_failed')).toHaveLength(1);
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    expect(await page.locator('body').textContent()).toContain('Juniper sample register.');
    if (promoted) expect(await page.locator('.page-box > [data-docx-drawing-anchor] img').count()).toBe(1);
  });
}

test('a healthy promoted image survives and a strict broken image is rejected', async ({ browser }) => {
  const result = await renderDocxArtifacts(imageIdentityDocx(true, false), { browser, reviewProfile: 'markup', commentProfile: 'hidden', outputs: ['html', 'pdf'] });
  expect(result.renderReport.warnings.filter(warning => warning.code === 'image_decode_failed')).toHaveLength(0);
  await expect(renderDocxArtifacts(imageIdentityDocx(true), { browser, reviewProfile: 'markup', commentProfile: 'hidden', outputs: ['pdf'], unsupportedContent: 'strict' }))
    .rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'image_decoding' });
});

test('a warned image cannot conceal new damage to another image in the same paragraph', async ({ browser }) => {
  await expect(renderDocxArtifacts(imageIdentityDocx(false), { browser: corruptReopenedImage(browser), reviewProfile: 'markup', commentProfile: 'hidden', outputs: ['pdf'] }))
    .rejects.toMatchObject({ code: 'output_verification_failure', phase: 'output_verification' });
});

test('a repeated undecodable header image retains one warning and one source identity', async ({ browser, page }) => {
  const result = await renderDocxArtifacts(backgroundPictureDocx('page', true, false, true), {
    browser, reviewProfile: 'markup', commentProfile: 'hidden', outputs: ['html', 'pdf'],
  });
  expect(result.pageCount).toBe(2);
  expect(result.renderReport.warnings.filter(warning => warning.code === 'image_decode_failed')).toHaveLength(1);
  await page.setContent(result.html!);
  const identities = await page.locator('.page-box img').evaluateAll(images =>
    images.map(image => image.getAttribute('data-docx-visual-resource-id')));
  expect(identities).toHaveLength(2);
  expect(identities[0]).toBeTruthy();
  expect(new Set(identities).size).toBe(1);
});
