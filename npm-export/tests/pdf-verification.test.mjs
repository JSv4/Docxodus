import assert from "node:assert/strict";
import { describe, test } from "node:test";
import {
  PDFArray,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  decodePDFRawStream,
} from "pdf-lib";
import { restorePageGeometry, verifyPdf } from "../dist/pdf.js";

const EXPECTED_PAGE = Object.freeze({
  pageNumber: 1,
  pageInSection: 1,
  width: 612,
  height: 792,
  sectionIndex: 0,
  pageName: "docxodus-section-0",
});

function tagForStructuralSmokeTest(pdf) {
  pdf.catalog.set(PDFName.of("MarkInfo"), pdf.context.obj({ Marked: true }));
  pdf.catalog.set(PDFName.of("StructTreeRoot"), pdf.context.obj({
    Type: "StructTreeRoot",
    K: [],
  }));
}

async function savePdf(configure) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([612, 792]);
  tagForStructuralSmokeTest(pdf);
  configure?.(pdf, page);
  return new Uint8Array(await pdf.save({ useObjectStreams: false }));
}

describe("PDF physical-page verification", () => {
  // Chromium prints a requested 595 x 842.5pt page as 594.95996 x 841.91998pt and paints from the
  // top edge through a flip about that quantized height (issue #1125).
  async function chromiumShapedPdf(width, height, { annotate = false } = {}) {
    return savePdf((pdf, page) => {
      page.setMediaBox(0, 0, width, height);
      page.setCropBox(0, 0, width, height);
      const stream = pdf.context.register(pdf.context.stream(`q .24 0 0 -.24 0 ${height} cm 100 100 50 50 re f Q`));
      page.node.set(PDFName.of("Contents"), stream);
      if (annotate) {
        const target = pdf.context.obj([page.ref, PDFName.of("XYZ"), 0, height - 10, 0]);
        const link = pdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [10, height - 40, 60, height - 20], Dest: target });
        page.node.set(PDFName.of("Annots"), pdf.context.obj([pdf.context.register(link)]));
      }
    });
  }

  async function contentText(bytes) {
    const pdf = await PDFDocument.load(bytes);
    const contents = pdf.getPages()[0].node.Contents();
    const streams = contents instanceof PDFArray ? contents.asArray().map((ref) => pdf.context.lookup(ref)) : [contents];
    return streams.map((stream) => Buffer.from(stream instanceof PDFRawStream ? decodePDFRawStream(stream).decode() : stream.getContents()).toString("latin1")).join("\n");
  }

  test("restores Chromium's quantized page size and keeps content on the top edge", async () => {
    const expected = [{ ...EXPECTED_PAGE, width: 595, height: 842.5 }];
    const quantized = await chromiumShapedPdf(594.95996, 841.91998, { annotate: true });
    // The delivered size must be the finalized size, not a bounded approximation of it.
    await assert.rejects(verifyPdf(quantized, expected, quantized.byteLength),
      (error) => error?.code === "output_verification_failure" && /MediaBox/.test(error.message));

    const restored = await restorePageGeometry(quantized, expected, 1 << 24);
    const verified = await verifyPdf(restored, expected, restored.byteLength);
    assert.match(verified.digest, /^[0-9a-f]{64}$/);
    const page = (await PDFDocument.load(restored)).getPages()[0];
    assert.deepEqual(page.getMediaBox(), { x: 0, y: 0, width: 595, height: 842.5 });
    assert.deepEqual(page.getCropBox(), { x: 0, y: 0, width: 595, height: 842.5 });

    // The 0.58002pt of height is added below the content: a translation ahead of Chromium's flip
    // keeps every mark, link rectangle and destination at its distance from the top edge.
    const dy = 842.5 - 841.91998;
    const text = await contentText(restored);
    assert.match(text, new RegExp(`^1 0 0 1 0 ${dy.toFixed(6)} cm`));
    assert.ok(text.indexOf("cm\n") < text.indexOf(".24 0 0 -.24"));
    const link = page.node.Annots().lookup(0);
    const rect = link.lookup(PDFName.of("Rect")).asArray().map((n) => n.asNumber());
    assert.deepEqual(rect, [10, 842.5 - 40, 60, 842.5 - 20]);
    assert.equal(link.lookup(PDFName.of("Dest")).lookup(3).asNumber(), 842.5 - 10);
  });

  test("shifts named and outline destinations, and parses nothing beyond the admission limit", async () => {
    const expected = [{ ...EXPECTED_PAGE, width: 595, height: 842.5 }];
    const quantized = await savePdf((pdf, page) => {
      page.setMediaBox(0, 0, 594.95996, 841.91998);
      page.setCropBox(0, 0, 594.95996, 841.91998);
      const named = pdf.context.obj([page.ref, PDFName.of("XYZ"), 0, 800, 0]);
      pdf.catalog.set(PDFName.of("Names"), pdf.context.obj({
        Dests: pdf.context.obj({ Names: ["Willow", named] }),
      }));
      const item = pdf.context.register(pdf.context.obj({ Title: "Willow", Dest: [page.ref, PDFName.of("FitH"), 700] }));
      pdf.catalog.set(PDFName.of("Outlines"), pdf.context.obj({ Type: "Outlines", First: item, Last: item, Count: 1 }));
    });
    const dy = 842.5 - 841.91998;
    const restored = await PDFDocument.load(await restorePageGeometry(quantized, expected, 1 << 24));
    const named = restored.catalog.lookup(PDFName.of("Names")).lookup(PDFName.of("Dests")).lookup(PDFName.of("Names"));
    assert.equal(named.lookup(1).lookup(3).asNumber(), 800 + dy);
    const outline = restored.catalog.lookup(PDFName.of("Outlines")).lookup(PDFName.of("First"));
    assert.equal(outline.lookup(PDFName.of("Dest")).lookup(2).asNumber(), 700 + dy);

    const untouched = await restorePageGeometry(quantized, expected, quantized.byteLength - 1);
    assert.equal(Buffer.compare(Buffer.from(untouched), Buffer.from(quantized)), 0);
  });

  test("restores width without moving left-anchored content", async () => {
    const expected = [{ ...EXPECTED_PAGE, width: 626.5, height: 852 }];
    const restored = await restorePageGeometry(await chromiumShapedPdf(625.91998, 852), expected, 1 << 24);
    await verifyPdf(restored, expected, restored.byteLength);
    assert.doesNotMatch(await contentText(restored), /^1 0 0 1 0 /);
  });

  test("checks both physical boxes strictly and leaves non-quantization errors for rejection", async () => {
    for (const box of ['MediaBox', 'CropBox']) {
      for (const axis of ['width', 'height']) {
        const exact = await savePdf((_pdf, page) => {
          page[`set${box}`](0, 0, 612 + (axis === 'width' ? 0.01 : 0), 792 + (axis === 'height' ? 0.01 : 0));
        });
        await verifyPdf(exact, [EXPECTED_PAGE], exact.byteLength);
        const outside = await savePdf((_pdf, page) => {
          page[`set${box}`](0, 0, 612 + (axis === 'width' ? 0.011 : 0), 792 + (axis === 'height' ? 0.011 : 0));
        });
        await assert.rejects(verifyPdf(outside, [EXPECTED_PAGE], outside.byteLength),
          error => error?.code === "output_verification_failure" && error.message.includes(box));
      }
    }
    // A size error beyond Chromium's quantization is a layout error, not something to repair.
    const wrong = await chromiumShapedPdf(611.2, 792);
    const untouched = await restorePageGeometry(wrong, [EXPECTED_PAGE], 1 << 24);
    assert.equal(Buffer.compare(Buffer.from(untouched), Buffer.from(wrong)), 0);
    await assert.rejects(verifyPdf(untouched, [EXPECTED_PAGE], untouched.byteLength),
      error => error?.code === "output_verification_failure" && /MediaBox/.test(error.message));
    // So is a page whose origin moved.
    const shifted = await savePdf((_pdf, page) => {
      page.setMediaBox(0.6, 0, 612, 792);
    });
    await assert.rejects(verifyPdf(await restorePageGeometry(shifted, [EXPECTED_PAGE], 1 << 24), [EXPECTED_PAGE], shifted.byteLength * 2),
      error => error?.code === "output_verification_failure" && /MediaBox/.test(error.message));
  });

  test("accounts for inherited UserUnit and quarter-turn rotation", async () => {
    const bytes = await savePdf((pdf, page) => {
      const parent = page.node.Parent();
      assert.ok(parent);
      page.node.delete(PDFName.of("MediaBox"));
      parent.set(PDFName.of("MediaBox"), pdf.context.obj([0, 0, 396, 306]));
      parent.set(PDFName.of("CropBox"), pdf.context.obj([0, 0, 396, 306]));
      parent.set(PDFName.of("UserUnit"), PDFNumber.of(2));
      parent.set(PDFName.of("Rotate"), PDFNumber.of(90));
    });

    const verified = await verifyPdf(bytes, [EXPECTED_PAGE], bytes.byteLength);
    assert.match(verified.digest, /^[0-9a-f]{64}$/);

    const oversized = await savePdf((pdf, page) => {
      const parent = page.node.Parent();
      page.node.delete(PDFName.of("MediaBox"));
      parent.set(PDFName.of("MediaBox"), pdf.context.obj([0, 0, 396.4, 306]));
      parent.set(PDFName.of("UserUnit"), PDFNumber.of(2));
      parent.set(PDFName.of("Rotate"), PDFNumber.of(90));
    });
    await assert.rejects(verifyPdf(oversized, [EXPECTED_PAGE], oversized.byteLength),
      error => error?.code === "output_verification_failure" && /MediaBox/.test(error.message));
  });

  test("rejects a non-zero effective page-box origin", async () => {
    const bytes = await savePdf((_pdf, page) => {
      page.setMediaBox(1, 0, 612, 792);
      page.setCropBox(1, 0, 612, 792);
    });

    await assert.rejects(
      verifyPdf(bytes, [EXPECTED_PAGE], bytes.byteLength),
      (error) => error?.code === "output_verification_failure"
        && /MediaBox is \[1, 0, 612, 792\]pt/.test(error.message),
    );
  });

  test("rejects malformed UserUnit and Rotate page attributes", async () => {
    const invalidUserUnit = await savePdf((pdf, page) => {
      page.node.set(PDFName.of("UserUnit"), pdf.context.obj(0));
    });
    await assert.rejects(
      verifyPdf(invalidUserUnit, [EXPECTED_PAGE], invalidUserUnit.byteLength),
      (error) => error?.code === "output_verification_failure"
        && /invalid \/UserUnit 0/.test(error.message),
    );

    const invalidRotation = await savePdf((_pdf, page) => {
      page.node.set(PDFName.of("Rotate"), PDFNumber.of(45));
    });
    await assert.rejects(
      verifyPdf(invalidRotation, [EXPECTED_PAGE], invalidRotation.byteLength),
      (error) => error?.code === "output_verification_failure"
        && /invalid \/Rotate 45/.test(error.message),
    );
  });
});
