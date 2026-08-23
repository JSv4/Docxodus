import { expect, test } from '@playwright/test';
import type { PdfInspection } from './visual-parity/pdf.js';
import { exactLinkEvidence } from './visual-parity/pdf-links.js';

const semantics = {
  requiredText: ['TOC', 'Target'],
  links: [{
    kind: 'internal' as const,
    sourceText: 'TOC Entry',
    anchor: '_Target',
    destinationText: 'Target Heading',
    expectedPdfAnnotations: 1,
  }],
};

function inspection(): PdfInspection {
  return {
    pdfSha256: 'a'.repeat(64),
    pageCount: 2,
    searchableText: 'TOC Entry\nTarget Heading',
    linkAnnotations: 1,
    vectorPathOperations: 0,
    marked: false,
    semantics: {} as PdfInspection['semantics'],
    pages: [
      {
        pageNumber: 1, userUnit: 1, rotation: 0,
        mediaBox: { x: 0, y: 0, width: 612, height: 792 },
        cropBox: { x: 0, y: 0, width: 612, height: 792 },
        text: 'TOC Entry', textSha256: 'b'.repeat(64), constructPathOperations: 0,
        hyperlinks: [{
          rectangle: [1, 2, 3, 4],
          target: { kind: 'destination', value: '_Target', namedDestination: '_Target', pageNumber: 2 },
        }],
      },
      {
        pageNumber: 2, userUnit: 1, rotation: 0,
        mediaBox: { x: 0, y: 0, width: 612, height: 792 },
        cropBox: { x: 0, y: 0, width: 612, height: 792 },
        text: 'Target Heading', textSha256: 'c'.repeat(64), constructPathOperations: 0,
        hyperlinks: [],
      },
    ],
  };
}

test.describe('exact PDF hyperlink evidence', () => {
  test('binds source label, exact target, destination page text, rectangle, and cardinality', () => {
    expect(exactLinkEvidence(inspection(), { semantics })).toMatchObject({
      passed: true,
      unexpectedLogicalLinks: 0,
      missingOrMismatched: [],
    });
  });

  test('rejects swapped/wrong targets, wrong destination text, missing rectangles, and extras', () => {
    const mutations = [
      (value: PdfInspection) => { (value.pages[0].hyperlinks[0].target as any).namedDestination = '_Wrong'; },
      (value: PdfInspection) => { value.pages[1].text = 'Wrong page'; },
      (value: PdfInspection) => { delete (value.pages[0].hyperlinks[0] as any).rectangle; },
      (value: PdfInspection) => { value.pages[0].hyperlinks.push(value.pages[0].hyperlinks[0]); },
    ];
    for (const mutate of mutations) {
      const value = inspection();
      mutate(value);
      expect(exactLinkEvidence(value, { semantics }).passed).toBe(false);
    }
  });

  test('requires the source URI representation instead of a canonicalized equivalent', () => {
    const value = inspection();
    value.pages[0].hyperlinks = [{
      rectangle: [1, 2, 3, 4],
      target: { kind: 'url', value: 'http://example.invalid/', unsafeValue: 'http://example.invalid/' },
    }];
    const external = {
      requiredText: ['Example'],
      links: [{
        kind: 'external' as const,
        sourceText: 'TOC Entry',
        relationshipId: 'rId1',
        exactTarget: 'http://example.invalid',
        expectedPdfTarget: 'http://example.invalid',
        expectedPdfAnnotations: 1,
      }],
    };
    expect(exactLinkEvidence(value, { semantics: external }).passed).toBe(false);
  });
});
