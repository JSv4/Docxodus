import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { renderDocxArtifacts, type FontLicenseAttestation } from '../../npm-export/dist/index.js';
import { drawingGroupDocx, drawingGroupLabels, type DrawingGroupOrientation } from './docx-drawing-group-fixture.js';

const fontDirectories = [fileURLToPath(new URL('../../docs/demo/fonts', import.meta.url))];
const fontLicenseAttestations: FontLicenseAttestation[] = [{
  schemaVersion: 1, usage: 'standalone-document-font-embedding',
  fileSha256: createHash('sha256').update(readFileSync(
    new URL('../../docs/demo/fonts/docxodus-canvas-mono.woff2', import.meta.url))).digest('hex'),
  embeddingPermitted: true, permittedOutputs: ['html', 'pdf'], subsettingPermitted: true,
  basis: 'Bundled renamed DejaVu subset; redistribution and modification permitted by docs/demo/fonts/LICENSE.txt.',
}];

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const header of [false, true]) for (const kind of ['group', 'nested', 'separate'] as const) {
  test(`${header ? 'header' : 'body'} DrawingML ${kind} keeps all stories and their bounds`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(drawingGroupDocx(header, kind), {
      browser, fontDirectories, fontLicenseAttestations, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.warnings).toEqual([]);
    await page.setContent(result.html!);
    const boxes = await page.locator('.page-box [data-docx-textbox]').evaluateAll(elements => elements.map(box => {
      const b = box.getBoundingClientRect(), p = box.closest('.page-box')!.getBoundingClientRect();
      return { text: box.textContent?.trim(), x: (b.x-p.x)*.75, y: (b.y-p.y)*.75,
        width: b.width*.75, height: b.height*.75,
        anchors: Array.from(box.querySelectorAll('[data-source-anchor-id]'), e => e.getAttribute('data-source-anchor-id')) };
    }));
    expect(boxes.map(box => box.text)).toEqual(drawingGroupLabels);
    for (const [index, box] of boxes.entries()) {
      expect(box.x).toBeCloseTo(72 + index*120, 1);
      expect(box.y).toBeCloseTo(200, 1);
      expect(box.width).toBeCloseTo(100, 1);
      expect(box.height).toBeCloseTo(30, 1);
      expect(box.anchors).toHaveLength(1);
      expect(result.pageMap.fragments.some(f => f.anchorId === box.anchors[0])).toBe(true);
    }
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await task.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      for (const [index, label] of drawingGroupLabels.entries()) {
        const matches = items.filter(item => item.str === label);
        expect(matches, JSON.stringify(items.map(item => item.str))).toHaveLength(1);
        expect(matches[0].transform[4]).toBeCloseTo(72 + index*120, 1);
      }
    } finally { await task.destroy(); }
  });
}

for (const [name, rotation, orientation, middleAngle] of [
  ['leaf horizontal flip', 0, { leafFlip: 'H' }, 0],
  ['leaf vertical flip', 0, { leafFlip: 'V' }, 0],
  ['group horizontal flip', 0, { groupFlip: 'H' }, 0],
  ['group vertical flip', 0, { groupFlip: 'V' }, 0],
  ['nested opposing flips', 0, { groupFlip: 'H', nestedFlip: 'H', leafRotation: 30 }, 30],
  ['rotation between nested flips', 0, { groupFlip: 'H', nestedFlip: 'H', nestedRotation: 30 }, -30],
  ['reflected child rotation', 0, { groupFlip: 'H', leafRotation: 30 }, -30],
  ['upright rotated group', 90, { upright: true }, 0],
  ['upright reflected child', 90, { groupFlip: 'H', leafRotation: 30, upright: true }, 0],
] as [string, number, DrawingGroupOrientation, number][]) {
  test(`${name} preserves readable PDF text`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(drawingGroupDocx(true,
      orientation.nestedFlip ? 'nested' : 'group', false, rotation*60000, orientation), {
      browser, fontDirectories, fontLicenseAttestations, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden',
    });
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.warnings).toEqual([]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await task.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      for (const [index, label] of drawingGroupLabels.entries()) {
        const matches = items.filter(item => item.str === label);
        expect(matches, JSON.stringify(items.map(item => item.str))).toHaveLength(1);
        const [a,b,c,d] = matches[0].transform;
        expect(a*d-b*c).toBeGreaterThan(0);
        expect(Math.atan2(-b,a)*180/Math.PI).toBeCloseTo(index === 1 ? middleAngle : orientation.upright ? 0 : rotation, 2);
      }
    } finally { await task.destroy(); }
  });
}

for (const degrees of [0, 12, 90]) {
  test(`inserted header group retains all PDF stories at ${degrees} degrees`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(drawingGroupDocx(true, 'group', true, degrees*60000), {
      browser, fontDirectories, fontLicenseAttestations, outputs: ['pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'hidden',
    });
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.warnings).toEqual([]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await task.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      for (const label of drawingGroupLabels) {
        const matches = items.filter(item => item.str === label);
        expect(matches, JSON.stringify(items.map(item => item.str))).toHaveLength(1);
        const [a, b] = matches[0].transform;
        expect(Math.atan2(-b, a)*180/Math.PI).toBeCloseTo(degrees, 2);
      }
    } finally { await task.destroy(); }
  });
}
