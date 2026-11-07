import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { vmlLineDocx } from './docx-vml-line-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

type Matrix = [number, number, number, number, number, number];
function multiply(a: Matrix, b: Matrix): Matrix {
  return [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
}

type LineCase = { name: string; to: string | null; mixed: boolean; width: number; height: number;
  options?: Parameters<typeof vmlLineDocx>[2]; reverse?: boolean };
const cases: LineCase[] = ['100pt,0', '0,60pt', '96,48'].flatMap(to => [false, true].map(mixed => ({
  name: `${to}, ${mixed ? 'mixed' : 'alone'}`, to, mixed,
  width: to === '100pt,0' ? 100 : to === '96,48' ? 72 : 0,
  height: to === '0,60pt' ? 60 : to === '96,48' ? 36 : 0,
})));
cases.push(
  { name: 'object host alone', to: '0,60pt', mixed: false, width: 0, height: 60, options: { host: 'object' } },
  { name: 'omitted endpoints and stroke', to: null, mixed: true, width: 50, height: 50,
    options: { from: null, strokeWeight: null } },
  { name: 'reversed endpoints', to: '0,0', mixed: true, width: 70, height: 25,
    options: { from: '70pt,25pt' }, reverse: true },
  { name: 'hidden stroke retains layout', to: '70pt,25pt', mixed: true, width: 70, height: 25,
    options: { stroked: false } },
);
for (const { name, to, mixed, width: expectedWidth, height: expectedHeight, options, reverse } of cases) {
  test(`VML line ${name}: strict export retains geometry and source text`, async ({ browser }) => {
    const result = await renderDocxArtifacts(vmlLineDocx(to, mixed, options), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
    });
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.unsupportedContent).toEqual([]);
    expect(result.html).toContain('<line');
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const page = await (await task.promise).getPage(1);
      const items = (await page.getTextContent()).items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join('').replace(/\s/g, '');
      expect(text).toBe(mixed ? 'LineBeforeLineAnchorLineTailLineAfter' : 'LineBeforeLineAfter');
      const ops = await page.getOperatorList();
      let matrix: Matrix = [1, 0, 0, 1, 0, 0], color = '', width = 1;
      const stack: { matrix: Matrix; color: string; width: number }[] = [];
      const strokes: { points: number[][]; width: number }[] = [];
      for (let i = 0; i < ops.fnArray.length; i++) {
        const op = ops.fnArray[i], args = ops.argsArray[i];
        if (op === OPS.save) stack.push({ matrix: [...matrix], color, width });
        else if (op === OPS.restore) ({ matrix, color, width } = stack.pop()!);
        else if (op === OPS.transform) matrix = multiply(matrix, args as Matrix);
        else if (op === OPS.setStrokeRGBColor) color = args[0];
        else if (op === OPS.setLineWidth) width = args[0];
        else if (op === OPS.setGState) {
          for (const [key, value] of args[0]) if (key === 'LW') width = value;
        }
        else if (op === OPS.constructPath && color === '#16777a' && args[0] === OPS.stroke) {
          const points: number[][] = [];
          for (const path of args[1] as Float32Array[]) for (let k = 0; k < path.length;) {
            expect([0, 1]).toContain(path[k++]); // moveTo / lineTo: a curve or filled substitute is not this primitive.
            const x = path[k++], y = path[k++];
            points.push([matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]]);
          }
          strokes.push({ points, width: width * Math.hypot(matrix[0], matrix[1]) });
        }
      }
      const before = items.find(item => item.str === 'LineBefore')!, after = items.find(item => item.str === 'LineAfter')!;
      const anchor = items.find(item => item.str === 'LineAnchor'), tail = items.find(item => item.str === 'LineTail');
      const left = before.transform[4] + (mixed ? anchor!.width : 0) + 9;
      const painted = options?.stroked !== false;
      expect(strokes).toHaveLength(painted ? 1 : 0);
      if (painted) {
        const stroke = strokes[0], xs = stroke.points.map(p => p[0]), ys = stroke.points.map(p => p[1]);
        expect(Math.abs(Math.min(...xs) - left)).toBeLessThan(.8);
        expect(Math.abs(Math.max(...xs) - Math.min(...xs) - expectedWidth)).toBeLessThan(.2);
        expect(Math.abs(Math.max(...ys) - Math.min(...ys) - expectedHeight)).toBeLessThan(.2);
        expect(Math.abs(stroke.width - (options?.strokeWeight === null ? .75 : 2))).toBeLessThan(.02);
        expect(Math.abs(stroke.points[0][0] - Math.min(...xs) - (reverse ? expectedWidth : 0))).toBeLessThan(.2);
        expect(Math.abs(stroke.points[0][1] - (reverse ? Math.min(...ys) : Math.max(...ys)))).toBeLessThan(.2);
        // Existing text-baseline correction plus CSS-pixel snapping can separate ink by up to 1.5pt.
        if (mixed) expect(Math.abs(anchor!.transform[5] - Math.min(...ys))).toBeLessThan(1.5);
        else expect(Math.abs(before.transform[5] - Math.max(...ys) - 2.65)).toBeLessThan(.8);
      }
      // Exact native copies distinguish a sole object's top alignment from a mixed line's baseline.
      // Chromium quantizes glyph placement to CSS pixels; text-to-paint comparisons allow one pixel.
      expect(Math.abs(before.transform[5] - after.transform[5] - 13.8 -
        Math.max(13.8, expectedHeight + (mixed ? 2.65 : 0)))).toBeLessThan(1);
      if (mixed) {
        expect(Math.abs(anchor!.transform[5] - after.transform[5] - 13.8)).toBeLessThan(.8);
        expect(Math.abs(tail!.transform[4] - left - expectedWidth - 9)).toBeLessThan(.8);
        expect(Math.abs(tail!.transform[5] - anchor!.transform[5])).toBeLessThan(.2);
      }
    } finally { await task.destroy(); }
  });
}

for (const [name, to, options] of [
  ['effects', '100pt,0', { extra: '<v:stroke dashstyle="dash"/>' }],
  ['zero-length', '0,0', {}],
  ['hairline', '100pt,0', { strokeWeight: '0' }],
] as const) test(`VML ${name} retains the placeholder and strict-policy refusal`, async ({ browser }) => {
  const source = vmlLineDocx(to, true, options);
  const renderOptions = { browser, outputs: ['pdf'] as const, reviewProfile: 'final' as const, commentProfile: 'hidden' as const };
  const result = await renderDocxArtifacts(source, renderOptions);
  expect(result.renderReport.unsupportedContent).toEqual([expect.objectContaining({ contentType: 'Other' })]);
  await expect(renderDocxArtifacts(source, { ...renderOptions, unsupportedContent: 'strict' }))
    .rejects.toMatchObject({ code: 'resource_policy_failure' });
});
