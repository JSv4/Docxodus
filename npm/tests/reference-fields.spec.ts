import { test, expect, Page } from '@playwright/test';
import * as zlib from 'zlib';
import type { EditErrorCode } from '../src/types.js';

/**
 * Issue #607 — reference-field authoring across the WASM bridge.
 *
 * The three tables that narrowing the library to the DOCX toolchain took away with
 * `ReferenceAdder`: contents, figures, authorities. The point of the ops is that a caller never
 * writes a switch string, so this spec asserts the switch string — a malformed one renders as
 * *nothing* in Word, silently, and no schema check catches it.
 *
 * The annotation states that the .NET enum member reached the TypeScript union, which asserting
 * the literal at runtime cannot show. Same reasoning as page-numbering.spec.ts.
 */
const INVALID_REFERENCE_FIELD: EditErrorCode = 'invalid_reference_field';

/** See page-numbering.spec.ts — read one entry out of a DOCX without a zip dependency. */
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

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Reference fields (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('typed options become the field switches, and the field is written dirty', async ({ page }) => {
    const res = await page.evaluate(() => {
      const B = (window as any).Docxodus.DocxSessionBridge;
      const h = B.OpenSession(B.CreateBlankDocx(), '{}');
      try {
        const proj = JSON.parse(B.Project(h));
        const body = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const toc = JSON.parse(B.InsertTableOfContents(h, body, 'before', ''));
        const tocSaved: Uint8Array = B.Save(h);

        const tof = JSON.parse(
          B.InsertTableOfFigures(h, body, 'after', JSON.stringify({ captionLabel: 'Exhibit' })),
        );
        const toa = JSON.parse(
          B.InsertTableOfAuthorities(h, body, 'after', JSON.stringify({ category: 'statutes' })),
        );
        const allSaved: Uint8Array = B.Save(h);

        const bad = JSON.parse(
          B.InsertTableOfContents(h, body, 'before', JSON.stringify({ levels: '0-3' })),
        );

        return {
          tocSuccess: toc.success,
          tofSuccess: tof.success,
          toaSuccess: toa.success,
          badSuccess: bad.success,
          badCode: bad.error?.code ?? null,
          tocSaved: Array.from(tocSaved),
          allSaved: Array.from(allSaved),
        };
      } finally {
        B.CloseSession(h);
      }
    });

    expect(res.tocSuccess).toBe(true);
    expect(res.tofSuccess).toBe(true);
    expect(res.toaSuccess).toBe(true);

    const tocXml = readZipEntry(Buffer.from(res.tocSaved), 'word/document.xml');
    expect(tocXml).toContain('TOC \\o "1-3" \\h \\z \\u');
    // Dirty, so Word repaginates and fills the table instead of trusting a cached result.
    expect(tocXml).toContain('w:fldCharType="begin" w:dirty="true"');
    // Word's own wrapper, which is what puts an "Update Table" control on it.
    expect(tocXml).toContain('Table of Contents');
    // …and the document asks Word to update fields on open, or the reader sees an empty table.
    expect(readZipEntry(Buffer.from(res.tocSaved), 'word/settings.xml')).toContain('updateFields');

    const allXml = readZipEntry(Buffer.from(res.allSaved), 'word/document.xml');
    expect(allXml).toContain('TOC \\c "Exhibit" \\h');
    // The wire name hides Word's numbering; the field carries the number.
    expect(allXml).toContain('TOA \\c "2" \\h');

    // A malformed level range is refused rather than written as a field that renders nothing.
    expect(res.badSuccess).toBe(false);
    expect(res.badCode).toBe(INVALID_REFERENCE_FIELD);
  });
});
