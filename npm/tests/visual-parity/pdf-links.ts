import type { PdfLinkExpectation, PdfParityCorpusEntry } from './pdf-corpus.js';
import type { PdfHyperlinkAnnotation, PdfInspection } from './pdf.js';

export interface ObservedLinkEvidence {
  annotationPage: number;
  annotationCount: number;
  rectangles: readonly (readonly number[])[];
  kind: string;
  value?: string;
  unsafeValue?: string;
  destinationPage?: number;
  expectedSourceText?: string;
  sourceTextPresent: boolean;
  destinationTextPresent: boolean;
  exactTargetPassed: boolean;
  exactAnnotationCountPassed: boolean;
}

export interface LinkEvidence {
  expected: readonly PdfLinkExpectation[];
  observed: ObservedLinkEvidence[];
  missingOrMismatched: readonly PdfLinkExpectation[];
  unexpectedLogicalLinks: number;
  unsupportedAnnotations: number;
  passed: boolean;
}

interface LocatedAnnotation {
  annotation: PdfHyperlinkAnnotation;
  annotationPage: number;
  pageText: string;
}

function normalized(value: string): string {
  return value.normalize('NFC').replace(/\s+/g, ' ').trim();
}

function comparableText(value: string): string {
  return normalized(value).toLocaleLowerCase('en-US');
}

function sourceRepresentation(annotation: PdfHyperlinkAnnotation): string | undefined {
  return annotation.target.kind === 'url'
    ? annotation.target.unsafeValue ?? annotation.target.value
    : undefined;
}

function logicalTargetKey(annotation: PdfHyperlinkAnnotation, index: number): string {
  const target = annotation.target;
  if (target.kind === 'destination') {
    return `destination\0${target.namedDestination ?? target.value}\0${target.pageNumber}`;
  }
  if (target.kind === 'url') return `url\0${sourceRepresentation(annotation) ?? ''}`;
  return `unsupported\0${index}`;
}

function consecutiveGroups(annotations: readonly LocatedAnnotation[]): LocatedAnnotation[][] {
  const groups: LocatedAnnotation[][] = [];
  let previousKey: string | undefined;
  annotations.forEach((item, index) => {
    const key = logicalTargetKey(item.annotation, index);
    if (key !== previousKey) groups.push([]);
    groups[groups.length - 1].push(item);
    previousKey = key;
  });
  return groups;
}

/**
 * Bind logical links to the manifest in deterministic document order. Chromium currently emits
 * several consecutive rectangles for one internal hyperlink, so the manifest pins both the
 * logical target sequence and exact annotation multiplicity instead of mistaking those rectangles
 * for unexpected links or silently accepting injected duplicates.
 */
export function exactLinkEvidence(
  inspection: PdfInspection,
  entry: Pick<PdfParityCorpusEntry, 'semantics'>,
): LinkEvidence {
  const expected = [...(entry.semantics.links ?? [])];
  const annotations = inspection.pages.flatMap((page) => page.hyperlinks.map((annotation) => ({
    annotation,
    annotationPage: page.pageNumber,
    pageText: comparableText(page.text),
  })));
  const groups = consecutiveGroups(annotations);
  const missingOrMismatched: PdfLinkExpectation[] = [];
  const observed = groups.map((group, index): ObservedLinkEvidence => {
    const first = group[0];
    const expectation = expected[index];
    const target = first.annotation.target;
    const destinationPage = target.kind === 'destination' ? target.pageNumber : undefined;
    const destinationText = expectation?.kind === 'internal'
      ? comparableText(inspection.pages[destinationPage === undefined ? -1 : destinationPage - 1]?.text ?? '')
      : '';
    const sourceTextPresent = expectation !== undefined
      && first.pageText.includes(comparableText(expectation.sourceText));
    const destinationTextPresent = expectation?.kind !== 'internal'
      || destinationText.includes(comparableText(expectation.destinationText));
    const exactTargetPassed = expectation?.kind === 'external'
      ? target.kind === 'url' && sourceRepresentation(first.annotation) === expectation.expectedPdfTarget
      : expectation?.kind === 'internal'
        ? target.kind === 'destination'
          && (target.namedDestination ?? target.value) === expectation.anchor
          && destinationPage !== undefined
        : false;
    const exactAnnotationCountPassed = expectation !== undefined
      && group.length === expectation.expectedPdfAnnotations;
    const rectangles = group.flatMap((item) => item.annotation.rectangle ? [item.annotation.rectangle] : []);
    const rectanglesPassed = rectangles.length === group.length
      && rectangles.every((rectangle) => rectangle.length === 4 && rectangle.every(Number.isFinite));
    const oneSourcePage = group.every((item) => item.annotationPage === first.annotationPage);
    if (!expectation || !sourceTextPresent || !destinationTextPresent || !exactTargetPassed
      || !exactAnnotationCountPassed || !rectanglesPassed || !oneSourcePage) {
      if (expectation) missingOrMismatched.push(expectation);
    }
    return {
      annotationPage: first.annotationPage,
      annotationCount: group.length,
      rectangles,
      kind: target.kind,
      ...(target.kind === 'unsupported' ? {} : { value: target.value }),
      ...(target.kind === 'url' && target.unsafeValue ? { unsafeValue: target.unsafeValue } : {}),
      ...(destinationPage === undefined ? {} : { destinationPage }),
      ...(expectation === undefined ? {} : { expectedSourceText: expectation.sourceText }),
      sourceTextPresent,
      destinationTextPresent,
      exactTargetPassed,
      exactAnnotationCountPassed,
    };
  });
  if (groups.length < expected.length) missingOrMismatched.push(...expected.slice(groups.length));
  const unsupportedAnnotations = annotations
    .filter((item) => item.annotation.target.kind === 'unsupported').length;
  const exactContractRequired = expected.length > 0;
  const unexpectedLogicalLinks = exactContractRequired ? Math.max(0, groups.length - expected.length) : 0;
  return {
    expected,
    observed,
    missingOrMismatched,
    unexpectedLogicalLinks,
    unsupportedAnnotations,
    passed: unsupportedAnnotations === 0
      && (!exactContractRequired
        || (groups.length === expected.length && missingOrMismatched.length === 0)),
  };
}
