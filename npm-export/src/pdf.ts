import { createHash } from "node:crypto";
import { PDFDocument } from "pdf-lib";
import type { CompleteRenderReport } from "./contracts.js";
import { exportError } from "./contracts.js";

export interface VerifiedPdf {
  digest: string;
  volatileMetadata: Record<string, string>;
}

const GEOMETRY_TOLERANCE_POINTS = 0.5;

function closeTo(actual: number, expected: number): boolean {
  return Number.isFinite(actual)
    && Math.abs(actual - expected) <= GEOMETRY_TOLERANCE_POINTS;
}

function metadata(pdf: PDFDocument): Record<string, string> {
  const values: Record<string, string | undefined> = {
    title: pdf.getTitle(),
    author: pdf.getAuthor(),
    subject: pdf.getSubject(),
    creator: pdf.getCreator(),
    producer: pdf.getProducer(),
    creationDate: pdf.getCreationDate()?.toISOString(),
    modificationDate: pdf.getModificationDate()?.toISOString(),
  };
  return Object.fromEntries(
    Object.entries(values).filter((entry): entry is [string, string] => entry[1] !== undefined),
  );
}

export async function verifyPdf(
  bytes: Uint8Array,
  expectedPages: CompleteRenderReport["pages"],
): Promise<VerifiedPdf> {
  if (bytes.byteLength < 8
    || new TextDecoder("ascii").decode(bytes.subarray(0, 5)) !== "%PDF-") {
    exportError(
      "output_verification_failure",
      "output_verification",
      "Chromium did not return a recognizable PDF document.",
      "Retry with the pinned Chromium runtime and inspect browser diagnostics.",
    );
  }

  let pdf: PDFDocument;
  try {
    pdf = await PDFDocument.load(new Uint8Array(bytes), {
      ignoreEncryption: false,
      updateMetadata: false,
      throwOnInvalidObject: true,
    });
  } catch (cause) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The generated PDF could not be parsed as a complete document.",
      "Retry with the pinned Chromium runtime and inspect the source document.",
      { cause },
    );
  }

  if (pdf.isEncrypted) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "Chromium unexpectedly produced an encrypted PDF.",
      "Use the supported unencrypted export path.",
    );
  }
  const pages = pdf.getPages();
  if (pages.length !== expectedPages.length) {
    exportError(
      "output_verification_failure",
      "output_verification",
      `PDF page count ${pages.length} does not match finalized layout count ${expectedPages.length}.`,
      "Report the document and Chromium build; no PDF artifact was returned.",
    );
  }

  pages.forEach((page, index) => {
    const expected = expectedPages[index];
    const media = page.getMediaBox();
    const crop = page.getCropBox();
    for (const [boxName, box] of [["MediaBox", media], ["CropBox", crop]] as const) {
      if (!closeTo(box.x, 0) || !closeTo(box.y, 0)
        || !closeTo(box.width, expected.width) || !closeTo(box.height, expected.height)) {
        exportError(
          "output_verification_failure",
          "output_verification",
          `PDF page ${index + 1} ${boxName} is `
            + `[${box.x}, ${box.y}, ${box.width}, ${box.height}]pt; finalized layout requires `
            + `[0, 0, ${expected.width}, ${expected.height}]pt.`,
          "Use CSS page-size printing with zero browser margins.",
        );
      }
    }
  });

  return {
    digest: createHash("sha256").update(bytes).digest("hex"),
    volatileMetadata: metadata(pdf),
  };
}
