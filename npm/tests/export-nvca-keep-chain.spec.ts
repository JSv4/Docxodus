import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';

/**
 * Issue #1180: in the NVCA model charter the keep-with-next heading "2.3 Deemed Liquidation Events"
 * shares its page with "2.3.1 Definition", whose paragraph cites three notes, as in Word's PDF.
 * Reserving a flat note start for the chain underestimated what those notes need, and the heading
 * was left alone at the foot of the page before.
 */

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

const here = dirname(fileURLToPath(import.meta.url));

test('a heading kept with a paragraph that cites several notes stays on that paragraph\'s page', async ({ browser }) => {
  test.setTimeout(300_000);
  const bytes = new Uint8Array(readFileSync(join(here, '../../TestFiles/NVCA-Model-COI.docx')));
  const result = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], reviewProfile: 'markup',
    commentProfile: 'margin', timeoutMs: 240_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  const pages: string[] = [];
  try {
    const pdf = await task.promise;
    for (let index = 1; index <= pdf.numPages; index++) {
      const content = await (await pdf.getPage(index)).getTextContent();
      pages.push(content.items.map((item) => ('str' in item ? item.str : '')).join(' ').replace(/\s+/g, ' '));
    }
  } finally { await task.destroy(); }

  // The heading carries note 21; the text layer lists section numbers apart from their headings.
  const heading = pages.findIndex((text) => /Deemed Liquidation Events\s*\.\s*21\b/.test(text));
  const definition = pages.findIndex((text) => /Definition\s*\.\s*Each of the following/.test(text));
  expect(heading).toBeGreaterThan(0);
  expect(definition).toBe(heading);
});
