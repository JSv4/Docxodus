import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

/** See editor-headerfooter.spec.ts — read one entry out of a DOCX without a zip dependency. */
function readZipEntry(zip: Buffer, name: string): string {
  const eocd = zip.lastIndexOf(Buffer.from('PK\x05\x06', 'latin1'));
  if (eocd < 0) throw new Error('not a zip');
  const count = zip.readUInt16LE(eocd + 10);
  let p = zip.readUInt32LE(eocd + 16);
  for (let i = 0; i < count; i++) {
    const nameLen = zip.readUInt16LE(p + 28);
    const extraLen = zip.readUInt16LE(p + 30);
    const commentLen = zip.readUInt16LE(p + 32);
    const entry = zip.toString('latin1', p + 46, p + 46 + nameLen);
    if (entry === name) {
      const method = zip.readUInt16LE(p + 10);
      const compSize = zip.readUInt32LE(p + 20);
      const localOff = zip.readUInt32LE(p + 42);
      const lNameLen = zip.readUInt16LE(localOff + 26);
      const lExtraLen = zip.readUInt16LE(localOff + 28);
      const start = localOff + 30 + lNameLen + lExtraLen;
      const data = zip.subarray(start, start + compSize);
      return (method === 0 ? data : zlib.inflateRawSync(data)).toString('utf8');
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error(`entry not found: ${name}`);
}

function readTestFile(relativePath: string): number[] {
  return Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/**
 * `DocxEditor.save()` must not leak the projector's `PtOpenXml:Unid` bookkeeping into the file the
 * user downloads.
 *
 * The editor used to open its session with `persistAnchorIds: true`, which suppresses the Unid
 * strip on EVERY save — so a document saved without a single edit came back roughly 6x its original
 * size. It went unnoticed because the attributes live in a custom namespace that Word and
 * LibreOffice both ignore, so the output renders identically; only the byte count betrays it.
 *
 * The setting existed for the remount's re-render, which needs anchor ids to survive a
 * save/re-render hop. That is now a per-CALL request (`SaveWithAnchorIds`), so the two consumers
 * can't contaminate each other — hence the second test, which covers the invariant a size
 * assertion cannot see.
 */
test.describe('DocxEditor.save() — no anchor-id bookkeeping in user output', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('a save with zero edits carries no Unid attributes', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const res = await page.evaluate((arr: number[]) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, new Uint8Array(arr), D, {});
      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();
      return { saved: Array.from(saved) };
    }, bytes);

    const documentXml = readZipEntry(Buffer.from(res.saved), 'word/document.xml');
    expect(documentXml).not.toContain('Unid=');
    // The bloat is the symptom a byte-count guard catches even if the attribute is ever renamed.
    expect(res.saved.length).toBeLessThan(bytes.length * 2);
  });

  test('a save after edits is still clean, and the remount keeps blocks wired', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const res = await page.evaluate((arr: number[]) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, new Uint8Array(arr), D, {});

      const blockAt = (i: number) =>
        container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]')[i];

      // toggleList produces a list item, which forces the FULL remount path — the one the
      // anchor-id round trip exists to protect.
      const before = blockAt(1);
      before.focus();
      editor.toggleList('bullet');

      const after = blockAt(1);
      const wired = !!after && after.getAttribute('contenteditable') === 'true' &&
        !!after.getAttribute('data-anchor');

      // A follow-up edit only lands if the remounted block is genuinely wired to the session.
      let followUpApplied = false;
      if (after) {
        after.focus();
        const sel = window.getSelection()!;
        const r = document.createRange();
        r.selectNodeContents(after);
        sel.removeAllRanges();
        sel.addRange(r);
        document.execCommand('insertText', false, 'REMOUNT PROBE');
        after.dispatchEvent(new Event('blur'));
        followUpApplied = (container.textContent || '').includes('REMOUNT PROBE');
      }

      const saved: Uint8Array = editor.save();
      editor.close();
      container.remove();
      return { wired, followUpApplied, saved: Array.from(saved) };
    }, bytes);

    expect(res.wired).toBe(true);
    expect(res.followUpApplied).toBe(true);

    const documentXml = readZipEntry(Buffer.from(res.saved), 'word/document.xml');
    expect(documentXml).not.toContain('Unid=');
    expect(documentXml).toContain('REMOUNT PROBE');
  });
});
