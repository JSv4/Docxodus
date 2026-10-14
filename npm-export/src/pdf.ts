import { createHash } from "node:crypto";
import { PDFArray, PDFBool, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRef } from "pdf-lib";
import type { PDFObject, PDFPage } from "pdf-lib";
import type { CompleteRenderReport } from "./contracts.js";
import { exportError } from "./contracts.js";

export interface VerifiedPdf {
  digest: string;
  volatileMetadata: Record<string, string>;
}

const PAGE_ORIGIN_TOLERANCE_POINTS = 0.5;
// Chromium prints in 300-DPI device units (0.24pt) and rounds the CSS page canvas, so a requested
// 878.5pt page arrives 877.91998pt tall. Bound half a point plus one device unit and float
// serialization. Only `restorePageGeometry` uses this bound, to recognize that quantization; the
// delivered PDF must then match the finalized page size within float serialization (issue #1125).
const CHROMIUM_PAGE_QUANTIZATION_POINTS = 0.75;
const PAGE_SIZE_TOLERANCE_POINTS = 0.01;
const PDF_WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);

function closeTo(actual: number, expected: number, tolerance = PAGE_ORIGIN_TOLERANCE_POINTS): boolean {
  return Number.isFinite(actual)
    && Math.abs(actual - expected) <= tolerance;
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

function quantized(actual: number, expected: number): boolean {
  return Number.isFinite(actual) && Math.abs(actual - expected) <= CHROMIUM_PAGE_QUANTIZATION_POINTS;
}

/** Move a destination array that targets `pageRef` up by `dy` in default user space. */
function shiftDestination(destination: PDFObject | undefined, pageRef: PDFRef, dy: number, pdf: PDFDocument): void {
  const resolved = destination instanceof PDFRef ? pdf.context.lookup(destination) : destination;
  const array = resolved instanceof PDFDict ? resolved.lookupMaybe(PDFName.of("D"), PDFArray) : resolved;
  if (!(array instanceof PDFArray) || array.size() < 2 || array.get(0) !== pageRef) return;
  const kind = array.lookupMaybe(1, PDFName)?.asString();
  // The vertical operands of each explicit destination form (PDF 32000-1 table 151).
  const vertical = kind === "/XYZ" ? [3] : kind === "/FitH" || kind === "/FitBH" ? [2] : kind === "/FitR" ? [3, 5] : [];
  for (const index of vertical) {
    const value = array.lookupMaybe(index, PDFNumber);
    if (value) array.set(index, PDFNumber.of(value.asNumber() + dy));
  }
}

/**
 * Restore the finalized physical page size Chromium quantized while printing.
 *
 * Chromium paints each page from its top edge: the content stream maps CSS space with a flip about
 * the quantized height. Setting MediaBox/CropBox to the finalized size therefore also needs a
 * `0 dy` translation ahead of that content, so every mark keeps its distance from the top edge, and
 * the same shift for annotation rectangles and destinations on the page. Width needs no shift: the
 * content is anchored to the left edge. Pages that are rotated, scaled by /UserUnit, offset, or
 * outside the quantization bound are left for `verifyPdf` to reject.
 */
/** Every explicit destination in the document, wherever PDF 32000-1 lets one live. */
function destinations(pdf: PDFDocument): PDFObject[] {
  // A destination reachable by two routes (an annotation and a name tree) is shifted once.
  const found = new Set<PDFObject>();
  const add = (value: PDFObject | undefined) => {
    if (value === undefined) return;
    const resolved = value instanceof PDFRef ? pdf.context.lookup(value) : value;
    const array = resolved instanceof PDFDict ? resolved.lookupMaybe(PDFName.of("D"), PDFArray) : resolved;
    if (array) found.add(array);
  };
  for (const page of pdf.getPages()) {
    const annotations = page.node.Annots();
    for (let i = 0; annotations && i < annotations.size(); i++) {
      const annotation = annotations.lookupMaybe(i, PDFDict);
      add(annotation?.get(PDFName.of("Dest")));
      add(annotation?.lookupMaybe(PDFName.of("A"), PDFDict)?.get(PDFName.of("D")));
    }
  }
  pdf.catalog.lookupMaybe(PDFName.of("Dests"), PDFDict)?.entries().forEach(([, value]) => add(value));
  const visited = new Set<PDFDict>();
  const nameTree = (node: PDFDict | undefined, depth = 0) => {
    if (!node || visited.has(node) || depth > 64) return;
    visited.add(node);
    const names = node.lookupMaybe(PDFName.of("Names"), PDFArray);
    for (let i = 1; names && i < names.size(); i += 2) add(names.get(i));
    const kids = node.lookupMaybe(PDFName.of("Kids"), PDFArray);
    for (let i = 0; kids && i < kids.size(); i++) nameTree(kids.lookupMaybe(i, PDFDict), depth + 1);
  };
  nameTree(pdf.catalog.lookupMaybe(PDFName.of("Names"), PDFDict)?.lookupMaybe(PDFName.of("Dests"), PDFDict));
  const outline = (item: PDFDict | undefined, depth = 0) => {
    for (let node = item; node && !visited.has(node) && depth <= 64; node = node.lookupMaybe(PDFName.of("Next"), PDFDict)) {
      visited.add(node);
      add(node.get(PDFName.of("Dest")));
      add(node.lookupMaybe(PDFName.of("A"), PDFDict)?.get(PDFName.of("D")));
      outline(node.lookupMaybe(PDFName.of("First"), PDFDict), depth + 1);
    }
  };
  outline(pdf.catalog.lookupMaybe(PDFName.of("Outlines"), PDFDict)?.lookupMaybe(PDFName.of("First"), PDFDict));
  return [...found];
}

export async function restorePageGeometry(
  bytes: Uint8Array,
  expectedPages: CompleteRenderReport["pages"],
  maximumParserBytes: number,
): Promise<Uint8Array> {
  // Parse nothing the verifier would refuse to parse; it reports the admission failure.
  if (!Number.isSafeInteger(maximumParserBytes) || maximumParserBytes <= 0
    || bytes.byteLength > maximumParserBytes) return bytes;
  let pdf: PDFDocument;
  try {
    pdf = await PDFDocument.load(new Uint8Array(bytes), {
      ignoreEncryption: false,
      updateMetadata: false,
      throwOnInvalidObject: true,
    });
  } catch {
    return bytes;
  }
  const pages = pdf.getPages();
  if (pdf.isEncrypted || pages.length !== expectedPages.length) return bytes;
  let changed = false;
  const allDestinations = destinations(pdf);
  pages.forEach((page, index) => {
    const expected = expectedPages[index];
    if (page.getRotation().angle % 360 !== 0
      || page.node.getInheritableAttribute(PDFName.of("UserUnit")) !== undefined) return;
    const media = page.getMediaBox();
    const crop = page.getCropBox();
    const boxes = [media, crop];
    if (boxes.some((box) => box.x !== 0 || box.y !== 0 || box.width !== media.width || box.height !== media.height
      || !quantized(box.width, expected.width) || !quantized(box.height, expected.height))) return;
    if (media.width === expected.width && media.height === expected.height) return;
    const dy = expected.height - media.height;
    page.setMediaBox(0, 0, expected.width, expected.height);
    page.setCropBox(0, 0, expected.width, expected.height);
    for (const name of ["TrimBox", "BleedBox", "ArtBox"]) page.node.delete(PDFName.of(name));
    if (dy !== 0) {
      // Fixed-point operand: content-stream numbers have no exponent form.
      const prefix = pdf.context.register(pdf.context.stream(`1 0 0 1 0 ${dy.toFixed(6)} cm\n`));
      const contents = page.node.get(PDFName.of("Contents"));
      const list = contents instanceof PDFArray
        ? contents
        : contents instanceof PDFRef
          ? (pdf.context.lookup(contents) instanceof PDFArray ? pdf.context.lookup(contents) as PDFArray : pdf.context.obj([contents]))
          : pdf.context.obj([]);
      list.insert(0, prefix);
      page.node.set(PDFName.of("Contents"), list);
      const annotations = page.node.Annots();
      for (let i = 0; annotations && i < annotations.size(); i++) {
        const annotation = annotations.lookup(i, PDFDict);
        for (const key of ["Rect", "QuadPoints"]) {
          const values = annotation.lookupMaybe(PDFName.of(key), PDFArray);
          for (let j = 1; values && j < values.size(); j += 2) {
            const value = values.lookupMaybe(j, PDFNumber);
            if (value) values.set(j, PDFNumber.of(value.asNumber() + dy));
          }
        }
      }
      for (const destination of allDestinations) shiftDestination(destination, page.ref, dy, pdf);
    }
    changed = true;
  });
  if (!changed) return bytes;
  return new Uint8Array(await pdf.save({ useObjectStreams: false, updateFieldAppearances: false }));
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
        || !closeTo(box.width, expected.width, PAGE_SIZE_TOLERANCE_POINTS)
        || !closeTo(box.height, expected.height, PAGE_SIZE_TOLERANCE_POINTS)) {
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
