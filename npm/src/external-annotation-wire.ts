/**
 * Wire-shape readers for the external annotation family.
 *
 * The engine serializes these results with PascalCase property names (and older builds with
 * camelCase), so every consumer has to normalize them once. The main-thread entry points in
 * `core.ts` and the Web Worker (`docxodus.worker.ts`) both read through here, which is what
 * lets the worker hand back the same typed objects the main thread would without a second
 * copy of the mapping.
 */

import type {
  AnnotationLabel,
  ExternalAnnotationSet,
  ExternalAnnotationValidationResult,
  OpenContractDocExport,
  OpenContractsAnnotation,
  OpenContractsRelationship,
  OpenContractsSinglePageAnnotation,
  PawlsPage,
  TextSpan,
} from "./types.js";

/* eslint-disable @typescript-eslint/no-explicit-any */

function readPawlsPage(p: any): PawlsPage {
  return {
    page: {
      width: p.Page?.Width ?? p.page?.width,
      height: p.Page?.Height ?? p.page?.height,
      index: p.Page?.Index ?? p.page?.index,
    },
    tokens: (p.Tokens || p.tokens || []).map((t: any) => ({
      x: t.X ?? t.x,
      y: t.Y ?? t.y,
      width: t.Width ?? t.width,
      height: t.Height ?? t.height,
      text: t.Text ?? t.text,
    })),
  };
}

function readAnnotationJson(
  json: any,
): TextSpan | Record<string, OpenContractsSinglePageAnnotation> | undefined {
  if (!json) return undefined;

  // A text span carries offsets; anything else is a page-keyed dictionary.
  if (json.Start !== undefined || json.start !== undefined) {
    return {
      id: json.Id ?? json.id,
      start: json.Start ?? json.start,
      end: json.End ?? json.end,
      text: json.Text ?? json.text,
    };
  }

  const result: Record<string, OpenContractsSinglePageAnnotation> = {};
  for (const [key, value] of Object.entries(json)) {
    const v = value as any;
    result[key] = {
      bounds: {
        top: v.Bounds?.Top ?? v.bounds?.top,
        bottom: v.Bounds?.Bottom ?? v.bounds?.bottom,
        left: v.Bounds?.Left ?? v.bounds?.left,
        right: v.Bounds?.Right ?? v.bounds?.right,
      },
      tokensJsons: (v.TokensJsons || v.tokensJsons || []).map((t: any) => ({
        pageIndex: t.PageIndex ?? t.pageIndex,
        tokenIndex: t.TokenIndex ?? t.tokenIndex,
      })),
      rawText: v.RawText ?? v.rawText,
    };
  }
  return result;
}

function readAnnotation(a: any): OpenContractsAnnotation {
  return {
    id: a.Id ?? a.id,
    annotationLabel: a.AnnotationLabel ?? a.annotationLabel,
    rawText: a.RawText ?? a.rawText,
    page: a.Page ?? a.page,
    annotationJson: readAnnotationJson(a.AnnotationJson ?? a.annotationJson),
    parentId: a.ParentId ?? a.parentId,
    annotationType: a.AnnotationType ?? a.annotationType,
    structural: a.Structural ?? a.structural,
  };
}

function readRelationship(r: any): OpenContractsRelationship {
  return {
    id: r.Id ?? r.id,
    relationshipLabel: r.RelationshipLabel ?? r.relationshipLabel,
    sourceAnnotationIds: r.SourceAnnotationIds ?? r.sourceAnnotationIds ?? [],
    targetAnnotationIds: r.TargetAnnotationIds ?? r.targetAnnotationIds ?? [],
    structural: r.Structural ?? r.structural,
  };
}

function readLabel(l: any): AnnotationLabel {
  return {
    id: l.Id ?? l.id,
    color: l.Color ?? l.color,
    description: l.Description ?? l.description ?? "",
    icon: l.Icon ?? l.icon ?? "",
    text: l.Text ?? l.text,
    labelType: l.LabelType ?? l.labelType ?? "text",
  };
}

function readLabels(raw: any): Record<string, AnnotationLabel> {
  const labels: Record<string, AnnotationLabel> = {};
  for (const [key, value] of Object.entries(raw || {})) labels[key] = readLabel(value);
  return labels;
}

/** Read an OpenContracts document export from the engine's parsed JSON. */
export function readOpenContractExport(parsed: any): OpenContractDocExport {
  return {
    title: parsed.Title ?? parsed.title,
    content: parsed.Content ?? parsed.content,
    description: parsed.Description ?? parsed.description,
    pageCount: parsed.PageCount ?? parsed.pageCount,
    pawlsFileContent: (parsed.PawlsFileContent || parsed.pawlsFileContent || []).map(readPawlsPage),
    docLabels: parsed.DocLabels ?? parsed.docLabels ?? [],
    labelledText: (parsed.LabelledText || parsed.labelledText || []).map(readAnnotation),
    relationships: (parsed.Relationships || parsed.relationships)?.map(readRelationship),
  };
}

/** Read an external annotation set — an OpenContracts export plus identity and label tables. */
export function readExternalAnnotationSet(parsed: any): ExternalAnnotationSet {
  return {
    documentId: parsed.DocumentId ?? parsed.documentId,
    documentHash: parsed.DocumentHash ?? parsed.documentHash,
    createdAt: parsed.CreatedAt ?? parsed.createdAt,
    updatedAt: parsed.UpdatedAt ?? parsed.updatedAt,
    version: parsed.Version ?? parsed.version,
    ...readOpenContractExport(parsed),
    textLabels: readLabels(parsed.TextLabels || parsed.textLabels),
    docLabelDefinitions: readLabels(parsed.DocLabelDefinitions || parsed.docLabelDefinitions),
  };
}

/** Read the result of validating an annotation set against a document. */
export function readExternalAnnotationValidation(parsed: any): ExternalAnnotationValidationResult {
  return {
    isValid: parsed.IsValid ?? parsed.isValid,
    hashMismatch: parsed.HashMismatch ?? parsed.hashMismatch,
    issues: (parsed.Issues || parsed.issues || []).map((i: any) => ({
      annotationId: i.AnnotationId ?? i.annotationId,
      issueType: i.IssueType ?? i.issueType,
      description: i.Description ?? i.description,
      expectedText: i.ExpectedText ?? i.expectedText,
      actualText: i.ActualText ?? i.actualText,
    })),
  };
}

/** The engine wraps projected HTML in a small JSON envelope; unwrap it. */
export function readProjectedHtml(parsed: any): string {
  return parsed.Html ?? parsed.html;
}
