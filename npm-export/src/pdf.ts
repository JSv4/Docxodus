import { createHash } from "node:crypto";
import { PDFBool, PDFDict, PDFDocument, PDFName, PDFNumber } from "pdf-lib";
import type { PDFObject, PDFPage } from "pdf-lib";
import type { CompleteRenderReport } from "./contracts.js";
import { exportError } from "./contracts.js";

export interface VerifiedPdf {
  digest: string;
  volatileMetadata: Record<string, string>;
}

const GEOMETRY_TOLERANCE_POINTS = 0.5;
const PDF_WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);

function closeTo(actual: number, expected: number): boolean {
  return Number.isFinite(actual)
    && Math.abs(actual - expected) <= GEOMETRY_TOLERANCE_POINTS;
}

interface PdfRectangle {
  x: number;
  y: number;
  width: number;
  height: number;
}

function inheritedPageNumber(page: PDFPage, name: string): number | undefined {
  let value: PDFObject | undefined;
  try {
    value = page.node.getInheritableAttribute(PDFName.of(name));
  } catch (cause) {
    exportError(
      "output_verification_failure",
      "output_verification",
      `PDF page attribute /${name} could not be resolved through the page tree.`,
      "Retry with the pinned Chromium runtime and reject malformed page dictionaries.",
      { cause },
    );
  }
  if (value === undefined) return undefined;
  let resolved: PDFObject | undefined;
  try {
    resolved = page.node.context.lookup(value);
  } catch (cause) {
    exportError(
      "output_verification_failure",
      "output_verification",
      `PDF page attribute /${name} references an invalid object.`,
      "Retry with the pinned Chromium runtime and reject malformed page dictionaries.",
      { cause },
    );
  }
  if (!(resolved instanceof PDFNumber)) {
    exportError(
      "output_verification_failure",
      "output_verification",
      `PDF page attribute /${name} is not a number.`,
      "Retry with the pinned Chromium runtime and reject malformed page dictionaries.",
    );
  }
  return resolved.asNumber();
}

/**
 * Resolve the physical rectangle represented by a PDF page box.
 *
 * MediaBox, CropBox, and Rotate are inheritable page attributes. pdf-lib resolves the two boxes
 * and Rotate through the page tree; UserUnit is also inheritable in the PDF page model but is not
 * surfaced by its high-level PDFPage API, so read it through the same ancestor walk. UserUnit
 * scales default user-space coordinates, and quarter-turn rotation swaps physical width/height.
 */
function physicalRectangle(
  rectangle: PdfRectangle,
  userUnit: number,
  rotation: number,
): PdfRectangle {
  const swapsAxes = rotation === 90 || rotation === 270;
  return {
    x: rectangle.x * userUnit,
    y: rectangle.y * userUnit,
    width: (swapsAxes ? rectangle.height : rectangle.width) * userUnit,
    height: (swapsAxes ? rectangle.width : rectangle.height) * userUnit,
  };
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
  maximumParserBytes: number,
): Promise<VerifiedPdf> {
  if (!Number.isSafeInteger(maximumParserBytes) || maximumParserBytes <= 0
    || bytes.byteLength > maximumParserBytes) {
    exportError(
      "resource_limit",
      "output_verification",
      `pdfParserExpandedBytes admission failed (${bytes.byteLength} > ${maximumParserBytes}).`,
      "Use a smaller PDF or a reviewed parser-memory limit.",
    );
  }
  const view = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.byteLength < 16 || !/^%PDF-[12]\.[0-9](?:\r\n|\r|\n)/.test(
    view.subarray(0, Math.min(16, view.byteLength)).toString("ascii"),
  )) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "Chromium did not return a recognizable PDF document.",
      "Retry with the pinned Chromium runtime and inspect browser diagnostics.",
    );
  }
  const eof = view.lastIndexOf(Buffer.from("%%EOF", "ascii"));
  if (eof < 0 || view.subarray(eof + 5).some((value) => !PDF_WHITESPACE.has(value))) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The generated PDF is missing a terminal EOF marker or has trailing non-PDF bytes.",
      "Reject truncated and polyglot output; retry with pinned Chromium.",
    );
  }
  const trailerStart = Math.max(0, eof - 128);
  if (!/startxref\s+[0-9]+\s*%%EOF\s*$/.test(view.subarray(trailerStart).toString("ascii"))) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The generated PDF does not end in a canonical startxref/EOF trailer.",
      "Retry with pinned Chromium and reject incomplete byte streams.",
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
  if (expectedPages.length < 1 || pages.length !== expectedPages.length) {
    exportError(
      "output_verification_failure",
      "output_verification",
      `PDF page count ${pages.length} does not match finalized layout count ${expectedPages.length}.`,
      "Report the document and Chromium build; no PDF artifact was returned.",
    );
  }

  pages.forEach((page, index) => {
    const expected = expectedPages[index];
    const rawUserUnit = inheritedPageNumber(page, "UserUnit");
    const userUnit = rawUserUnit ?? 1;
    if (!Number.isFinite(userUnit) || userUnit <= 0 || userUnit > 75_000) {
      exportError(
        "output_verification_failure",
        "output_verification",
        `PDF page ${index + 1} has invalid /UserUnit ${String(userUnit)}.`,
        "Use a finite positive PDF user unit no greater than the PDF specification limit.",
      );
    }
    const rawRotation = inheritedPageNumber(page, "Rotate") ?? 0;
    if (!Number.isFinite(rawRotation) || !Number.isInteger(rawRotation)
      || rawRotation % 90 !== 0) {
      exportError(
        "output_verification_failure",
        "output_verification",
        `PDF page ${index + 1} has invalid /Rotate ${String(rawRotation)}.`,
        "Use a finite integer rotation that is a multiple of 90 degrees.",
      );
    }
    const rotation = ((rawRotation % 360) + 360) % 360;
    let media: PdfRectangle;
    let crop: PdfRectangle;
    try {
      media = physicalRectangle(page.getMediaBox(), userUnit, rotation);
      crop = physicalRectangle(page.getCropBox(), userUnit, rotation);
    } catch (cause) {
      exportError(
        "output_verification_failure",
        "output_verification",
        `PDF page ${index + 1} has a malformed MediaBox or CropBox.`,
        "Retry with the pinned Chromium runtime and reject malformed page dictionaries.",
        { cause },
      );
    }
    for (const [boxName, box] of [["MediaBox", media], ["CropBox", crop]] as const) {
      if (!closeTo(box.x, 0) || !closeTo(box.y, 0)
        || !closeTo(box.width, expected.width) || !closeTo(box.height, expected.height)) {
        exportError(
          "output_verification_failure",
          "output_verification",
          `PDF page ${index + 1} ${boxName} is `
            + `[${box.x}, ${box.y}, ${box.width}, ${box.height}]pt after /UserUnit `
            + `${userUnit} and /Rotate ${rotation}; finalized layout requires `
            + `[0, 0, ${expected.width}, ${expected.height}]pt.`,
          "Use CSS page-size printing with zero browser margins.",
        );
      }
    }
  });

  const markInfo = pdf.catalog.lookupMaybe(PDFName.of("MarkInfo"), PDFDict);
  const marked = markInfo?.lookupMaybe(PDFName.of("Marked"), PDFBool);
  const structureTree = pdf.catalog.lookupMaybe(PDFName.of("StructTreeRoot"), PDFDict);
  if (marked?.asBoolean() !== true || !structureTree) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "Chromium did not produce the required tagged PDF structure.",
      "Use tagged PDF printing with the supported Chromium revision.",
    );
  }

  return {
    digest: createHash("sha256").update(bytes).digest("hex"),
    volatileMetadata: metadata(pdf),
  };
}
