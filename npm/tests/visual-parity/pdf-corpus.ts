import type { CommentProfile, ReviewProfile } from "../../src/export-browser.js";
import type { VisualDisposition } from "./corpus.js";

/**
 * The generated-PDF parity corpus is deliberately separate from the broader HTML-vs-LibreOffice
 * visual corpus. Its entries are inputs to a strict release gate, so fixture identity, provenance,
 * review/comment projection, searchable text, and link semantics are all reviewable data.
 */
export const PDF_PARITY_CORPUS_SCHEMA_VERSION = 1 as const;

export const REQUIRED_PDF_PARITY_CATEGORIES = [
  "legal-contract",
  "tables",
  "columns",
  "footnotes",
  "endnotes",
  "headers-footers",
  "images",
  "charts",
  "revisions",
  "comments",
  "mixed-orientation",
  "internal-links",
  "external-links",
] as const;

export type PdfParityCategory = typeof REQUIRED_PDF_PARITY_CATEGORIES[number];

export interface TrackedGeneratorProvenance {
  path: string;
  sha256: string;
  gitBlob: string;
  rationale: string;
}

export interface TrackedFixtureSource {
  kind: "tracked-fixture";
  path: string;
  /** Hash of the exact DOCX bytes. Unlike a comparison to HEAD, this survives fixture edits. */
  sha256: string;
  /** Git blob identity makes repository-history inspection direct without replacing SHA-256. */
  gitBlob: string;
  provenance: {
    origin: "repository-authored" | "open-xml-powertools";
    introducedBy: string;
    rationale: string;
    generator?: TrackedGeneratorProvenance;
  };
}

export interface PdfParityProfileExpectation {
  candidate: {
    reviewProfile: ReviewProfile;
    commentProfile: CommentProfile;
  };
  reference: {
    /**
     * LibreOffice follows saved redline state. Revision cases therefore accept a temporary copy
     * for the reference only; the candidate must still receive the original bytes through the
     * supported export API so its `final` projection is exercised.
     */
    revisionProjection: "source" | "accepted";
    commentProjection: "hidden";
  };
  rationale: string;
}

export interface InternalPdfLinkExpectation {
  kind: "internal";
  sourceText: string;
  anchor: string;
  destinationText: string;
  /** Chromium emits one logical internal link as this exact consecutive annotation group. */
  expectedPdfAnnotations: number;
}

export interface ExternalPdfLinkExpectation {
  kind: "external";
  sourceText: string;
  relationshipId: string;
  exactTarget: string;
  /** Exact URI representation delivered in the PDF after Chromium URL serialization. */
  expectedPdfTarget: string;
  expectedPdfAnnotations: number;
}

export type PdfLinkExpectation = InternalPdfLinkExpectation | ExternalPdfLinkExpectation;

export interface PdfSemanticExpectation {
  /** Tokens which must be present in independently extracted PDF text. */
  requiredText: readonly string[];
  /** Profile-hidden or final-view content which must not leak into extracted PDF text. */
  forbiddenText?: readonly string[];
  /** Exact targets to inspect independently of raster comparison. */
  links?: readonly PdfLinkExpectation[];
}

export interface PdfParityCorpusEntry {
  id: string;
  categories: readonly PdfParityCategory[];
  rationale: string;
  source: TrackedFixtureSource;
  profiles: PdfParityProfileExpectation;
  semantics: PdfSemanticExpectation;
  /** Attribution of the dominant raster residual; semantic failures always gate separately. */
  disposition: VisualDisposition;
}

export interface PdfParityCorpusManifest {
  schemaVersion: typeof PDF_PARITY_CORPUS_SCHEMA_VERSION;
  rationale: string;
  cases: readonly PdfParityCorpusEntry[];
}

const VP_GENERATOR = {
  path: "TestFiles/VP/make-vp-fixtures.py",
  sha256: "9f3d1cb820f50fa6b6b1d3cbb4252f818f13b282c28a2461022d8806241a368e",
  gitBlob: "2b7ff34d3193c9857656db0cf523bc41bcebee13",
  rationale: "Repository generator fixes ZIP timestamps and member order and emits byte-identical fixtures.",
} as const satisfies TrackedGeneratorProvenance;

const FINAL_HIDDEN_SOURCE = {
  candidate: { reviewProfile: "final", commentProfile: "hidden" },
  reference: { revisionProjection: "source", commentProjection: "hidden" },
  rationale: "Compare the supported final/hidden export profile against the same source view.",
} as const satisfies PdfParityProfileExpectation;

const FINAL_HIDDEN_ACCEPTED_REFERENCE = {
  candidate: { reviewProfile: "final", commentProfile: "hidden" },
  reference: { revisionProjection: "accepted", commentProjection: "hidden" },
  rationale: "The candidate receives original revision markup; only the reference copy is accepted to force the same final view.",
} as const satisfies PdfParityProfileExpectation;

const LEGAL_CONTRACT_INTERNAL_LINKS = [
  ["1. Definitions", "_Toc400000001", "Definitions"],
  ["2. Services", "_Toc400000002", "Services"],
  ["2.1 Statements of Work", "_Toc400000003", "Statements of Work"],
  ["2.2 Change Orders", "_Toc400000004", "Change Orders"],
  ["3. Fees and Payment", "_Toc400000005", "Fees and Payment"],
  ["3.1 Fees", "_Toc400000006", "Fees"],
  ["3.2 Invoicing; Late Payment", "_Toc400000007", "Invoicing; Late Payment"],
  ["4. Term and Termination", "_Toc400000008", "Term and Termination"],
  ["5. Confidentiality", "_Toc400000009", "Confidentiality"],
  ["6. Limitation of Liability", "_Toc400000010", "Limitation of Liability"],
  ["7. General Provisions", "_Toc400000011", "General Provisions"],
].map(([sourceText, anchor, destinationText]) => ({
  kind: "internal" as const,
  sourceText,
  anchor,
  destinationText,
  expectedPdfAnnotations: 5,
}));

/**
 * Ten cases / a small page set rather than the full diagnostic corpus. Several cases are
 * intentionally multi-purpose: the legal agreement and the endnote exercise tables, while the
 * mixed-section fixture covers both running stories and portrait/landscape transitions.
 */
export const PDF_PARITY_CORPUS = {
  schemaVersion: PDF_PARITY_CORPUS_SCHEMA_VERSION,
  rationale: "Compact, hash-pinned generated-PDF release corpus for issue #443.",
  cases: [
    {
      id: "pdf-legal-contract",
      categories: ["legal-contract", "tables", "internal-links"],
      rationale: "Realistic three-page agreement with a signature table and eleven internal TOC destinations.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/VP/VP004-Legal-Contract.docx",
        sha256: "5ad0115afae50d06da1fc7cb21aa61714026fa59b9fa67e454741622d2689346",
        gitBlob: "0c805ef6ff5a7c3c5f2c456932dee05df5bed07d",
        provenance: {
          origin: "repository-authored",
          introducedBy: "1d1332261891d16bed32480c4e77c4a1d1a1b3b7",
          rationale: "Authored specifically for the visual-parity legal-document corpus in issue #400.",
          generator: VP_GENERATOR,
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: {
        requiredText: ["MASTER SERVICES AGREEMENT", "Meridian Consulting Group LLC", "ATLAS MANUFACTURING CORPORATION"],
        links: LEGAL_CONTRACT_INTERNAL_LINKS,
      },
      disposition: {
        kind: "renderer-bug",
        rationale: "The existing visual corpus attributes accumulated clause-position residuals to Docxodus; PDF output must not regress them.",
        reference: "https://github.com/JSv4/Docxodus/issues/415",
      },
    },
    {
      id: "pdf-two-column-section",
      categories: ["columns"],
      rationale: "Continuous transition from a one-column title to a two-column body.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/VP/VP003-Two-Column-Section.docx",
        sha256: "e51b73ae672efcf92d9312e14fade03ff95be2b89a1c06950d98e961e8a61600",
        gitBlob: "2fca8bdfebcc2a1ba6a435876d1699e5f054529b",
        provenance: {
          origin: "repository-authored",
          introducedBy: "1d1332261891d16bed32480c4e77c4a1d1a1b3b7",
          rationale: "Authored specifically to isolate continuous-section column flow in issue #400.",
          generator: VP_GENERATOR,
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: { requiredText: ["The Docxodus Gazette", "A Two-Column Layout Exercise"] },
      disposition: {
        kind: "renderer-bug",
        rationale: "Column fill and split geometry remain renderer-owned in the existing ratchet.",
        reference: "https://github.com/JSv4/Docxodus/issues/413",
      },
    },
    {
      id: "pdf-footnote",
      categories: ["footnotes"],
      rationale: "Footnote reference, separator, note text, and bottom-of-page placement.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/CA/CA008-Footnote-Reference.docx",
        sha256: "a3a43411bf1dfa9b8b3a0907a9966a9f856484564a3f4672fca6be1c2e5f4a81",
        gitBlob: "ba2dcc3a3a5734a56c4a137f2436e10af26d59d9",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "7e8fb1dca18072c19414c7e99851c3d1e8ded95c",
          rationale: "Tracked Open-Xml-PowerTools comparison fixture retained in the repository test corpus.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: { requiredText: ["provides a way.", "This is a test."] },
      disposition: {
        kind: "environment",
        rationale: "Placement is close; the known residual is shared substitute-font rasterization.",
      },
    },
    {
      id: "pdf-endnote-table",
      categories: ["endnotes", "tables"],
      rationale: "End-of-document note flow with a real table inside the endnote.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/WC/WC036-Endnote-With-Table-Before.docx",
        sha256: "52dd6af5071df72afd7b3207076d23b119f5b43be30a72ab15bf00531fea62b1",
        gitBlob: "18dc4e9db268909a79c58dda8aec403dd13dbf42",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "7e8fb1dca18072c19414c7e99851c3d1e8ded95c",
          rationale: "Tracked Open-Xml-PowerTools comparison fixture retained in the repository test corpus.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: { requiredText: ["Video provides a powerful way", "Aaa", "Iii"] },
      disposition: {
        kind: "unattributed",
        rationale: "Endnote flow changed recently; keep the first generated-PDF baseline strict until its residual is re-triaged.",
        reference: "https://github.com/JSv4/Docxodus/issues/414",
      },
    },
    {
      id: "pdf-mixed-orientation-running-stories",
      categories: ["headers-footers", "mixed-orientation"],
      rationale: "Landscape-to-portrait section transition with first/even/default headers and footers.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/DB001-Sections.docx",
        sha256: "e4175a97187da72a48ee42efc74a16674f9767d687ade34cc6337ae999cd8479",
        gitBlob: "46b4acc9c46829af19569ebc69ab2da97cc69c16",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "38d18c3d3422e47592db4fd96011b246e97097e0",
          rationale: "Imported with the original Open-Xml-PowerTools section/running-story fixtures.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: {
        requiredText: ["This is a section that is in landscape mode.", "This is in the following section."],
      },
      disposition: {
        kind: "environment",
        rationale: "Page dimensions agree in the existing corpus; remaining deltas are shared-font line metrics.",
      },
    },
    {
      id: "pdf-raster-image",
      categories: ["images"],
      rationale: "Embedded PNG decode, physical sizing, placement, and surrounding selectable text.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/HC042-Image-Png.docx",
        sha256: "c71ca76022a332a4f4aaf2b6db77937c8345b1c9765efed324c9749b9843fd05",
        gitBlob: "b6e52cc56a868117d76b75a0c1ec3c741d89e871",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "38d18c3d3422e47592db4fd96011b246e97097e0",
          rationale: "Imported with the original Open-Xml-PowerTools HTML conversion fixtures.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: { requiredText: ["Video provides a powerful way to help you prove your point."] },
      disposition: {
        kind: "environment",
        rationale: "The image extent is correct; the existing residual is surrounding text line-box behavior.",
        reference: "https://github.com/JSv4/Docxodus/issues/404",
      },
    },
    {
      id: "pdf-chart",
      categories: ["charts"],
      rationale: "Cached-data clustered chart exercises SVG-to-PDF vector and label preservation.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/HC043-Chart.docx",
        sha256: "fc56de3bbe6ab73d8cd289f9dd6c884d57fcf3e7160be5960796d1a5e809e248",
        gitBlob: "a9221c7285d90f4cc3b9355d36445e9aae4c541f",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "38d18c3d3422e47592db4fd96011b246e97097e0",
          rationale: "Imported with the original Open-Xml-PowerTools HTML conversion fixtures.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: { requiredText: ["Series 1", "Category 1"] },
      disposition: {
        kind: "environment",
        rationale: "The supported clustered chart is close; its remaining residual is label rasterization.",
      },
    },
    {
      id: "pdf-final-revisions",
      categories: ["revisions"],
      rationale: "The supported API must project a real deleted run to final view before PDF output.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/FA/RevTracking/001-DeletedRun.docx",
        sha256: "b450ee098b2010955a6de2ed5ed0a86ee1d65a39e08f2c377eeee8040505778a",
        gitBlob: "636af5ab41ded2fe1946ab6948864abb1612c615",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "135af9043c733a44cb262d8c75fe6aa91cf3697f",
          rationale: "Added to Open-Xml-PowerTools for formatting/revision processing coverage.",
        },
      },
      profiles: FINAL_HIDDEN_ACCEPTED_REFERENCE,
      semantics: {
        requiredText: ["Video provides a", "way to help you prove your point."],
        forbiddenText: ["powerful"],
      },
      disposition: {
        kind: "environment",
        rationale: "With identical final-view content, the existing residual is heading/font rasterization rather than revision semantics.",
        reference: "https://github.com/JSv4/Docxodus/issues/404",
      },
    },
    {
      id: "pdf-hidden-comments",
      categories: ["comments"],
      rationale: "Ten dense, overlapping, point, and threaded comments prove the explicit hidden profile does not leak note bodies.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/DD/DD002-DenseComments.docx",
        sha256: "a4271f1e693994f5d7e6b05d843939f457726ccf8de561e6baf69deb33c3192c",
        gitBlob: "c6834c45f95f81adaa34d46747338e5fe69053a8",
        provenance: {
          origin: "repository-authored",
          introducedBy: "5418d00905082c2860d3798d41cd32e2adca6dc7",
          rationale: "Authored for the repository's dense comment-range and threading regression suite.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: {
        requiredText: ["Master Agreement", "These recitals bind the parties named below"],
        forbiddenText: ["scope of the recitals", "agree, keep it as drafted"],
      },
      disposition: {
        kind: "unattributed",
        rationale: "New PDF corpus case; strict until the first measured raster run is reviewed.",
      },
    },
    {
      id: "pdf-external-hyperlink",
      categories: ["external-links"],
      rationale: "One inert external relationship with an exact URI, validated independently from pixels.",
      source: {
        kind: "tracked-fixture",
        path: "TestFiles/HC023-Hyperlink.docx",
        sha256: "85de8c3dcc381f49df150b06398a06b030c9423380407983be29e3d666c12927",
        gitBlob: "ba2b2449e8247e648f922ff2a2493f55aa37e654",
        provenance: {
          origin: "open-xml-powertools",
          introducedBy: "38d18c3d3422e47592db4fd96011b246e97097e0",
          rationale: "Imported with the original Open-Xml-PowerTools HTML conversion fixtures.",
        },
      },
      profiles: FINAL_HIDDEN_SOURCE,
      semantics: {
        requiredText: ["Following is a hyperlink.", "EricWhite.com"],
        links: [{
          kind: "external",
          sourceText: "EricWhite.com",
          relationshipId: "rId4",
          exactTarget: "http://www.ericwhite.com",
          expectedPdfTarget: "http://www.ericwhite.com/",
          expectedPdfAnnotations: 1,
        }],
      },
      disposition: {
        kind: "unattributed",
        rationale: "Semantic link preservation gates independently; raster residual starts strict until measured.",
      },
    },
  ],
} as const satisfies PdfParityCorpusManifest;
