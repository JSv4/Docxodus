import { expect, test } from '@playwright/test';
import {
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFString,
  StandardFonts,
} from 'pdf-lib';
import {
  PDF_RASTER_CONTRACT,
  PDF_RASTER_CONTRACT_SHA256,
  inspectPdf,
  popplerRasterArguments,
} from './visual-parity/pdf.js';

async function semanticFixture(): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const first = pdf.addPage([612, 792]);
  first.setCropBox(10, 20, 580, 740);
  first.drawText('SELECTABLE ALPHA', { x: 72, y: 700, size: 12, font });
  const uri = pdf.context.register(pdf.context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Link'),
    Rect: [72, 675, 240, 695],
    Border: [0, 0, 0],
    A: {
      Type: PDFName.of('Action'),
      S: PDFName.of('URI'),
      URI: PDFString.of('https://example.invalid/contract'),
    },
  }));

  const second = pdf.addPage([792, 612]);
  second.drawText('SELECTABLE OMEGA', { x: 72, y: 520, size: 12, font });
  const internal = pdf.context.register(pdf.context.obj({
    Type: PDFName.of('Annot'),
    Subtype: PDFName.of('Link'),
    Rect: [72, 640, 240, 660],
    Border: [0, 0, 0],
    Dest: [second.ref, PDFName.of('Fit')],
  }));
  first.node.set(PDFName.of('Annots'), pdf.context.obj([uri, internal]));
  return new Uint8Array(await pdf.save({ useObjectStreams: false }));
}

test.describe('generated-PDF visual parity helpers', () => {
  test('one frozen Poppler argument contract fixes DPI, RGB PNG, antialiasing, and MediaBox', () => {
    expect(Object.isFrozen(PDF_RASTER_CONTRACT)).toBe(true);
    expect(PDF_RASTER_CONTRACT).toMatchObject({
      tool: 'pdftoppm',
      dpi: 96,
      format: 'png',
      colorModel: 'rgb',
      pageBox: 'media',
      fontAntialiasing: true,
      vectorAntialiasing: true,
      thinLineMode: 'none',
    });
    expect(PDF_RASTER_CONTRACT_SHA256).toMatch(/^[0-9a-f]{64}$/);

    const args = popplerRasterArguments('/tmp/source.pdf', '/tmp/pages/render');
    expect(Object.isFrozen(args)).toBe(true);
    expect(args).toEqual([
      '-r', '96',
      '-png',
      '-freetype', 'yes',
      '-thinlinemode', 'none',
      '-aa', 'yes',
      '-aaVector', 'yes',
      '-f', '1',
      '-l', '33',
      '-forcenum',
      '-q',
      '/tmp/source.pdf',
      '/tmp/pages/render',
    ]);
    expect(args).not.toContain('-gray');
    expect(args).not.toContain('-mono');
    expect(args).not.toContain('-cropbox');
  });

  test('applies UserUnit and page rotation without discarding nonzero box origins', async () => {
    const pdf = await PDFDocument.create();
    const rotated90 = pdf.addPage([100, 200]);
    rotated90.setMediaBox(10, 20, 30, 40);
    rotated90.setCropBox(12, 22, 20, 30);
    rotated90.node.set(PDFName.of('UserUnit'), PDFNumber.of(2));
    rotated90.node.set(PDFName.of('Rotate'), PDFNumber.of(90));
    const rotated270 = pdf.addPage([100, 200]);
    rotated270.setMediaBox(5, 7, 20, 30);
    rotated270.node.set(PDFName.of('UserUnit'), PDFNumber.of(3));
    rotated270.node.set(PDFName.of('Rotate'), PDFNumber.of(270));

    const inspection = await inspectPdf(new Uint8Array(await pdf.save({ useObjectStreams: false })));

    expect(inspection.pages[0]).toMatchObject({
      userUnit: 2,
      rotation: 90,
      mediaBox: { x: 20, y: 40, width: 80, height: 60 },
      cropBox: { x: 24, y: 44, width: 60, height: 40 },
    });
    expect(inspection.pages[1]).toMatchObject({
      userUnit: 3,
      rotation: 270,
      mediaBox: { x: 15, y: 21, width: 90, height: 60 },
    });
  });

  test('inspects actual per-page boxes, selectable text, exact link targets, and hashes', async () => {
    const bytes = await semanticFixture();
    const inspection = await inspectPdf(bytes, {
      requireSelectableText: true,
      requiredText: ['SELECTABLE ALPHA', 'SELECTABLE OMEGA'],
      forbiddenText: ['REDACTED SECRET'],
      requiredHyperlinkTargets: [
        { kind: 'url', value: 'https://example.invalid/contract', representation: 'source' },
        { kind: 'destination', pageNumber: 2 },
      ],
    });

    expect(inspection.pdfSha256).toMatch(/^[0-9a-f]{64}$/);
    expect(inspection.pageCount).toBe(2);
    expect(inspection.pages.map((page) => page.mediaBox)).toEqual([
      { x: 0, y: 0, width: 612, height: 792 },
      { x: 0, y: 0, width: 792, height: 612 },
    ]);
    expect(inspection.pages[0].cropBox).toEqual({ x: 10, y: 20, width: 580, height: 740 });
    expect(inspection.pages[1].cropBox).toEqual({ x: 0, y: 0, width: 792, height: 612 });
    expect(inspection.pages.every((page) => /^[0-9a-f]{64}$/.test(page.textSha256))).toBe(true);
    expect(inspection.linkAnnotations).toBe(2);
    expect(inspection.pages[0].hyperlinks[0].target).toEqual({
      kind: 'url',
      value: 'https://example.invalid/contract',
      unsafeValue: 'https://example.invalid/contract',
    });
    expect(inspection.pages[0].hyperlinks[1].target).toMatchObject({
      kind: 'destination',
      pageNumber: 2,
    });
    expect(inspection.semantics).toMatchObject({
      passed: true,
      selectableText: { status: 'verified', missingFragments: [] },
      hyperlinks: {
        status: 'verified',
        annotationCount: 2,
        supportedCount: 2,
        unsupportedCount: 0,
        missingTargets: [],
      },
    });
  });

  test('reports missing text and hyperlink expectations independently of raster metrics', async () => {
    const inspection = await inspectPdf(await semanticFixture(), {
      requiredText: ['NOT PRESENT'],
      forbiddenText: ['SELECTABLE ALPHA'],
      requiredHyperlinkTargets: [{ kind: 'url', value: 'https://missing.invalid/' }],
    });

    expect(inspection.semantics.passed).toBe(false);
    expect(inspection.semantics.selectableText).toMatchObject({
      status: 'failed',
      missingFragments: ['NOT PRESENT'],
      presentForbiddenFragments: ['SELECTABLE ALPHA'],
    });
    expect(inspection.semantics.hyperlinks).toMatchObject({
      status: 'failed',
      missingTargets: [{ kind: 'url', value: 'https://missing.invalid/' }],
    });
  });
});
