import { expect, test } from '@playwright/test';
import { bookmarkedTableDocx } from './docx-bookmarked-table-fixture.js';

test('standalone and session renders select each requested table row', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const rows = await page.evaluate(bytes => {
    const bin = new Uint8Array(bytes);
    const { DocumentConverter: dc, DocxSessionBridge: sb } = (window as any).Docxodus;
    const handle = sb.OpenSession(bin, '');
    try {
      const anchors = Object.entries(JSON.parse(sb.Project(handle)).anchorIndex)
        .filter(([, target]: [string, any]) => target.kind === 'tr').slice(0, 2).map(([id]) => id);
      const batch = JSON.parse(sb.RenderBlocksHtml(handle, JSON.stringify(anchors), 'docx-', false));
      return anchors.map(id => [dc.RenderBlockHtml(bin, id, 'docx-', false),
        sb.RenderBlockHtml(handle, id, 'docx-', false), batch[id]].map(html => {
        const row = new DOMParser().parseFromString(html, 'application/xml').documentElement;
        return { tag: row.localName, text: row.textContent };
      }));
    } finally { sb.CloseSession(handle); }
  }, Array.from(bookmarkedTableDocx()));
  expect(rows).toHaveLength(2);
  rows.forEach((paths, index) => paths.forEach(row => {
    expect(row.tag).toBe('tr');
    expect(row.text).toContain(`Birch0${index + 1}A`);
    expect(row.text).toContain(`Birch0${index + 1}B`);
    expect(row.text).not.toContain(`Birch0${2 - index}`);
  }));
});
