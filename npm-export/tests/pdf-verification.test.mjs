import assert from "node:assert/strict";
import { describe, test } from "node:test";
import {
  PDFDocument,
  PDFName,
  PDFNumber,
} from "pdf-lib";
import { verifyPdf } from "../dist/pdf.js";

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

  test("preserves cancellation and deadline taxonomy around parser work", async () => {
    const bytes = await savePdf();
    const controller = new AbortController();
    controller.abort(new Error("cancelled by test"));
    await assert.rejects(
      verifyPdf(bytes, [EXPECTED_PAGE], bytes.byteLength, Number.POSITIVE_INFINITY, controller.signal),
      (error) => error?.code === "operation_cancelled"
        && error.phase === "output_verification",
    );
    await assert.rejects(
      verifyPdf(bytes, [EXPECTED_PAGE], bytes.byteLength, performance.now() - 1),
      (error) => error?.code === "readiness_timeout"
        && error.phase === "output_verification",
    );
  });
});
