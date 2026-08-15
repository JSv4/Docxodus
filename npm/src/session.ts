// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

import type {
  AnchorInfo,
  AnchorRef,
  AnchorTargetRef,
  AnnotationUpdate,
  BlockMetadata,
  BookmarkInfo,
  BulkEditResult,
  CharSpan,
  CommentListEntry,
  ContentControlFillOptions,
  ContentControlInfo,
  CrossBlockMatch,
  DiffEntry,
  DocumentAnnotation,
  DocumentRange,
  DocxodusWasmExports,
  DocxSessionProjection,
  DocxSessionSettings,
  EditError,
  EditResult,
  EditSummary,
  FillOptions,
  FindOptions,
  FormatOp,
  FormattingInspection,
  HeaderFooterKind,
  HyperlinkInfo,
  HyperlinkKind,
  ImageCapabilities,
  ImageDimensions,
  ImageInsertOptions,
  ImageOccurrence,
  FloatingImageLayout,
  InlineSpan,
  NumberFormat,
  PageNumberField,
  PageNumberingOp,
  PageCitation,
  PageCitationRequest,
  PageMapRegistrationResult,
  PageMapStatus,
  ParagraphBorderEdge,
  ParagraphFormatOp,
  TableBorderSpec,
  TableInsertOptions,
  TableMetadataResult,
  TableCellResolutionResult,
  TableMergeContent,
  TableRowOptions,
  TableShadingScope,
  ListFormat,
  GrepOptions,
  ListMembership,
  MutationBatchFailure,
  MutationBatchChangeSet,
  MutationBatchMode,
  MutationBatchPreviewOptions,
  MutationBatchPreviewStep,
  MutationBatchResult,
  MutationBatchStep,
  MutationBatchStepResult,
  MutationPreconditions,
  ReplaceOptions,
  RevisionListEntry,
  SectionInfo,
  StyleInfo,
  TemplatePlaceholder,
  TextMatch,
} from "./types.js";
import type { PageMap } from "./pagination.js";
import { ContextBoundary, DiffFormat, PlaceholderKinds, ProjectionDepth, ProjectionScopes, TrackedChangeMode } from "./types.js";

function mutationBatchChangeSet<T>(
  before: readonly T[],
  after: readonly T[],
  key: (value: T) => string,
): MutationBatchChangeSet<T> {
  const beforeGroups = new Map<string, number[]>();
  const afterGroups = new Map<string, number[]>();
  const group = (items: readonly T[], target: Map<string, number[]>): void => {
    items.forEach((item, index) => {
      const identity = key(item);
      const indices = target.get(identity) ?? [];
      indices.push(index);
      target.set(identity, indices);
    });
  };
  group(before, beforeGroups);
  group(after, afterGroups);

  const beforeMatched = before.map(() => false);
  const afterMatched = after.map(() => false);
  const modified = after.map(() => false);
  for (const [identity, afterIndices] of afterGroups) {
    const beforeIndices = beforeGroups.get(identity) ?? [];
    for (const afterIndex of afterIndices) {
      const beforeIndex = beforeIndices.find(index =>
        !beforeMatched[index] && JSON.stringify(before[index]) === JSON.stringify(after[afterIndex]));
      if (beforeIndex === undefined) continue;
      beforeMatched[beforeIndex] = true;
      afterMatched[afterIndex] = true;
    }
    const remainingBefore = beforeIndices.filter(index => !beforeMatched[index]);
    const remainingAfter = afterIndices.filter(index => !afterMatched[index]);
    const modifiedCount = Math.min(remainingBefore.length, remainingAfter.length);
    for (let index = 0; index < modifiedCount; index++) {
      beforeMatched[remainingBefore[index]!] = true;
      afterMatched[remainingAfter[index]!] = true;
      modified[remainingAfter[index]!] = true;
    }
  }
  return {
    added: after.filter((_, index) => !afterMatched[index]),
    removed: before.filter((_, index) => !beforeMatched[index]),
    modified: after.filter((_, index) => modified[index]),
  };
}

/**
 * Stateful in-memory DOCX editing session keyed by markdown-projection anchor ids.
 * Mirror of the .NET `DocxSession` surface. See
 * `docs/architecture/docx_mutation_api.md` for the surface contract,
 * anchor lifecycle, error catalog, and supported markdown subset.
 *
 * Sessions are not eligible for JS-side garbage collection — call {@link close}
 * (or use a `using` block under TypeScript 5.2+) when done.
 */
export class DocxSession {
  private readonly handle: number;
  private readonly wasm: DocxodusWasmExports["DocxSessionBridge"];

  /** @internal */
  constructor(handle: number, wasm: DocxodusWasmExports["DocxSessionBridge"]) {
    this.handle = handle;
    this.wasm = wasm;
  }

  // ─── View ────────────────────────────────────────────────────────────

  project(): DocxSessionProjection {
    return JSON.parse(this.wasm.Project(this.handle)) as DocxSessionProjection;
  }

  /** Monotonic document version (0 at open; +1 per committed mutation/undo/redo). */
  getVersion(): number {
    return (JSON.parse(this.wasm.GetVersion(this.handle)) as { version: number }).version;
  }

  /** Register a browser-materialized PageMap without changing the document version. */
  registerPageMap(pageMap: PageMap, expectedRendererFingerprint?: string): PageMapRegistrationResult {
    return JSON.parse(this.wasm.RegisterPageMap(
      this.handle,
      JSON.stringify(pageMap),
      expectedRendererFingerprint ?? "",
    )) as PageMapRegistrationResult;
  }

  getPageMapStatus(request?: PageCitationRequest): PageMapStatus {
    return JSON.parse(this.wasm.GetPageMapStatus(
      this.handle,
      request ? JSON.stringify(request) : "",
    )) as PageMapStatus;
  }

  getPageCitation(anchorId: string, request: PageCitationRequest): PageCitation {
    return JSON.parse(this.wasm.GetPageCitation(
      this.handle,
      anchorId,
      JSON.stringify(request),
    )) as PageCitation;
  }

  /** Evaluate optimistic guards without mutating or advancing the version. */
  checkPreconditions(preconditions: MutationPreconditions): EditResult {
    return JSON.parse(
      this.wasm.CheckPreconditions(this.handle, JSON.stringify(preconditions)),
    ) as EditResult;
  }

  /**
   * Guard any synchronous mutation. WASM calls are synchronous and single-threaded, so the
   * check and callback form one uninterrupted client-side operation. Prefer a method's native
   * `preconditions` option where it has one (notably replaceTextRange's match-count guard).
   */
  runWithPreconditions(
    preconditions: MutationPreconditions,
    mutation: () => EditResult,
  ): EditResult {
    const checked = this.checkPreconditions(preconditions);
    return checked.success ? mutation() : checked;
  }

  /**
   * Execute synchronous mutations atomically by default. Atomic success is one undo/version
   * unit; any failed or thrown step restores the exact package and history checkpoint.
   */
  executeBatch(
    steps: readonly MutationBatchStep[],
    mode: MutationBatchMode = "atomic",
  ): MutationBatchResult {
    if (mode !== "atomic" && mode !== "best_effort") {
      throw new RangeError(`unknown mutation batch mode: ${String(mode)}`);
    }
    const baseVersion = this.getVersion();
    const observationWarnings: string[] = [];
    const inspect = <T>(label: string, read: () => T, fallback: T): T => {
      try { return read(); } catch (error) {
        observationWarnings.push(`${label} unavailable: ${error instanceof Error ? error.message : String(error)}`);
        return fallback;
      }
    };
    const beforeRevisions = inspect("Revision delta inspection", () => this.listRevisions(), []);
    const beforeComments = inspect("Comment delta inspection", () => this.listComments(), []);
    const beforeAnnotations = inspect("Annotation delta inspection", () => this.listAnnotations(), []);
    const complete = (result: {
      mode: MutationBatchMode;
      status: "ok" | "failed" | "partial";
      success: boolean;
      rolledBack: boolean;
      steps: readonly MutationBatchStepResult[];
      failure?: MutationBatchFailure;
    }): MutationBatchResult => {
      try {
        const revisionChanges = mutationBatchChangeSet(
          beforeRevisions,
          inspect("Revision delta inspection", () => this.listRevisions(), beforeRevisions),
          revision => revision.id,
        );
        const commentChanges = mutationBatchChangeSet(
          beforeComments,
          inspect("Comment delta inspection", () => this.listComments(), beforeComments),
          comment => comment.anchorId,
        );
        const annotationChanges = mutationBatchChangeSet(
          beforeAnnotations,
          inspect("Annotation delta inspection", () => this.listAnnotations(), beforeAnnotations),
          annotation => annotation.id ?? "",
        );
        const resultVersion = inspect(
          "Result version inspection", () => this.getVersion(), baseVersion,
        );
        const warnings: string[] = [...observationWarnings];
        if ([...revisionChanges.added, ...revisionChanges.modified]
          .some(revision => revision.date !== undefined && revision.date !== null)) {
          warnings.push("Tracked-revision date attributes may use the execution clock; compare revision ids, authors, types, text, and anchors across separate executions.");
        }
        if ([...commentChanges.added, ...commentChanges.modified]
          .some(comment => comment.date !== undefined && comment.date !== null)) {
          warnings.push("Comment date attributes may be generated from the execution clock; supply dates explicitly when byte-identical replay is required.");
        }
        // Same predicate as the .NET receipt (`annotation.Created.HasValue`): the warning is
        // about an execution CLOCK, so an annotation added with no created timestamp is
        // deterministic and must not raise it on one surface and not the other.
        if (annotationChanges.added
          .some(annotation => annotation.created !== undefined && annotation.created !== null)) {
          warnings.push("Auto-generated annotation ids or creation timestamps are execution metadata; supply id and created explicitly when byte-identical replay is required.");
        }
        if (result.steps.some(step => step.results.some(edit => edit.created.length > 0))) {
          warnings.push("Created anchors and related OOXML ids may be generated independently on replay; preview/apply equivalence is semantic and packageHash or anchor ids may differ.");
        }
        if (mode === "best_effort" && !result.success) {
          warnings.push("Best-effort execution retains every successful step despite later failures.");
        }
        // null, never "": an absent hash must not compare equal to another absent hash, or a
        // naive `preview.packageHash === applied.packageHash` replay assertion passes vacuously.
        let packageHash: string | null = null;
        if (!this.wasm.GetPackageContentHash) {
          warnings.push("This WASM bundle predates package equivalence hashes; packageHash is unavailable.");
        } else {
          try { packageHash = this.wasm.GetPackageContentHash(this.handle); } catch (error) {
            warnings.push(`Package equivalence hash unavailable: ${error instanceof Error ? error.message : String(error)}`);
          }
        }
        return {
          ...result,
          preview: false,
          baseVersion,
          resultVersion,
          packageHash,
          revisionChanges,
          commentChanges,
          annotationChanges,
          warnings,
          html: null,
        };
      } catch (error) {
        return {
          ...result,
          preview: false,
          baseVersion,
          resultVersion: inspect("Result version inspection", () => this.getVersion(), baseVersion),
          packageHash: null,
          revisionChanges: { added: [], removed: [], modified: [] },
          commentChanges: { added: [], removed: [], modified: [] },
          annotationChanges: { added: [], removed: [], modified: [] },
          warnings: [
            ...observationWarnings,
            `Batch receipt enrichment unavailable: ${error instanceof Error ? error.message : String(error)}`,
          ],
          html: null,
        };
      }
    };
    const internalFailure = (value: unknown): EditResult => ({
      success: false,
      error: { code: "internal_error", message: value instanceof Error ? value.message : String(value) },
      created: [], removed: [], modified: [],
    });
    const run = (step: MutationBatchStep): readonly EditResult[] => {
      try {
        const value = step.mutation();
        const results = Array.isArray(value) ? value : [value];
        if (results.length === 0 || results.some(result =>
          result === null || typeof result !== "object" || typeof result.success !== "boolean")) {
          return [internalFailure("batch mutation returned no valid edit results")];
        }
        return results;
      } catch (error) {
        return [internalFailure(error)];
      }
    };
    const failureOf = (
      step: MutationBatchStepResult,
      rolledBack: boolean,
    ): MutationBatchFailure => ({
      index: step.index,
      tool: step.tool,
      action: step.action,
      error: step.results.find(result => !result.success)?.error
        ?? { code: "internal_error", message: "batch step failed without an error" },
      rolledBack,
    });

    const preflightOne = (step: MutationBatchStep): EditError | undefined => {
      try { return step.preflight?.(); } catch (error) { return internalFailure(error).error; }
    };
    if (mode === "atomic") {
      const preflight = steps.map(preflightOne);
      const failedPreflight = preflight.findIndex(error => error !== undefined);
      if (failedPreflight >= 0) {
        const source = steps[failedPreflight]!;
        const failed: MutationBatchStepResult = {
          index: failedPreflight, tool: source.tool, action: source.action,
          success: false, rolledBack: true,
          results: [{ success: false, error: preflight[failedPreflight]!, created: [], removed: [], modified: [] }],
        };
        return complete({ mode, status: "failed", success: false, rolledBack: true,
          steps: [failed], failure: failureOf(failed, true) });
      }

      const transaction = this.wasm.BeginTransaction(this.handle);
      const completed: MutationBatchStepResult[] = [];
      try {
        for (let index = 0; index < steps.length; index++) {
          const source = steps[index]!;
          const results = run(source);
          const step: MutationBatchStepResult = {
            index, tool: source.tool, action: source.action,
            success: results.every(result => result.success), rolledBack: false, results,
          };
          completed.push(step);
          if (!step.success) {
            this.wasm.RollbackTransaction(transaction);
            const rolledBack = completed.map(value => ({ ...value, rolledBack: true }));
            const failed = rolledBack[rolledBack.length - 1]!;
            return complete({ mode, status: "failed", success: false, rolledBack: true,
              steps: rolledBack, failure: failureOf(failed, true) });
          }
        }
        this.wasm.CommitTransaction(transaction);
        return complete({ mode, status: "ok", success: true, rolledBack: false, steps: completed });
      } catch (error) {
        try { this.wasm.RollbackTransaction(transaction); } catch { /* preserve the original */ }
        throw error;
      }
    }

    // Preserve sequential best-effort semantics: a later preflight can observe state created by
    // an earlier successful step, so run it immediately before that step rather than up front.
    const completed: MutationBatchStepResult[] = steps.map((source, index) => {
      const preflight = preflightOne(source);
      const results = preflight
        ? [{ success: false, error: preflight, created: [], removed: [], modified: [] }]
        : run(source);
      return {
        index, tool: source.tool, action: source.action,
        success: results.every(result => result.success), rolledBack: false, results,
      };
    });
    const failed = completed.find(step => !step.success);
    return complete({
      mode,
      status: failed ? (completed.some(step => step.success) ? "partial" : "failed") : "ok",
      success: failed === undefined,
      rolledBack: false,
      steps: completed,
      failure: failed ? failureOf(failed, false) : undefined,
    });
  }

  /**
   * Execute the same callback batch algorithm against a complete isolated package clone.
   * Callbacks receive the shadow session explicitly; mutate that argument. The live session's
   * package, caches, version, configuration, and undo/redo history are never execution targets.
   */
  previewBatch(
    steps: readonly MutationBatchPreviewStep[],
    mode: MutationBatchMode = "atomic",
    options?: MutationBatchPreviewOptions,
  ): MutationBatchResult {
    if (mode !== "atomic" && mode !== "best_effort") {
      throw new RangeError(`unknown mutation batch mode: ${String(mode)}`);
    }
    const htmlMode = options?.html ?? "none";
    if (htmlMode !== "none" && htmlMode !== "scoped" && htmlMode !== "full") {
      throw new RangeError(`unknown preview HTML mode: ${String(htmlMode)}`);
    }
    if (!this.wasm.OpenPreviewSession) {
      throw new Error("This WASM bundle does not support isolated mutation previews.");
    }

    const shadow = new DocxSession(this.wasm.OpenPreviewSession(this.handle), this.wasm);
    try {
      const result = shadow.executeBatch(
        steps.map(step => ({
          tool: step.tool,
          action: step.action,
          mutation: () => step.mutation(shadow),
          preflight: step.preflight ? () => step.preflight!(shadow) : undefined,
        })),
        mode,
      );
      const warnings = [...result.warnings];
      let html: string | null = null;
      // A rendered document always starts with '<'; a leading '{' is the bridge's error object.
      const unwrapRendered = (rendered: string): string | null => {
        if (rendered.trimStart().startsWith("{")) {
          const envelope = JSON.parse(rendered) as { error?: string };
          if (envelope.error) {
            warnings.push(`Preview HTML could not be generated: ${envelope.error}`);
            return null;
          }
        }
        return rendered;
      };
      // Preview HTML MUST come from the façade's preview profile (DocxSessionOps.RenderPreview*),
      // not the editor's authoring profile: the editor render hides comments, annotations and
      // headers/footers, so routing a preview through it would show this surface a materially
      // different document than the stdio/Python/MCP surfaces show for the identical batch.
      const legacyProfileWarning =
        "This WASM bundle predates the shared preview HTML profile; preview HTML omits comments, " +
        "annotations and headers/footers and may differ from other surfaces.";
      try {
        if (htmlMode === "scoped") {
          if (!options?.htmlAnchorId) {
            warnings.push("Scoped HTML was requested without htmlAnchorId; no HTML was generated.");
          } else if (this.wasm.RenderPreviewBlockHtml) {
            html = unwrapRendered(
              this.wasm.RenderPreviewBlockHtml(shadow.handle, options.htmlAnchorId));
          } else {
            warnings.push(legacyProfileWarning);
            html = shadow.renderBlock(options.htmlAnchorId);
          }
        } else if (htmlMode === "full") {
          if (this.wasm.RenderPreviewHtml) {
            html = unwrapRendered(this.wasm.RenderPreviewHtml(shadow.handle));
          } else {
            warnings.push(legacyProfileWarning);
            html = unwrapRendered(this.wasm.RenderHtmlForReview
              ? this.wasm.RenderHtmlForReview(shadow.handle, "docx-", false, false, 1, true)
              : this.wasm.RenderHtml(shadow.handle, "docx-", false, false, 1));
          }
        }
      } catch (error) {
        warnings.push(`Preview HTML could not be generated: ${error instanceof Error ? error.message : String(error)}`);
      }
      return { ...result, preview: true, warnings, html };
    } finally {
      shadow.close();
    }
  }

  /**
   * Project a slice of the document keyed off an anchor — useful for showing
   * one section to an LLM at a time without paying the cost of projecting the
   * whole document.
   *
   * - `ProjectionDepth.SelfOnly` — just the addressed block (one paragraph,
   *   row, etc.).
   * - `ProjectionDepth.Subtree` — the block + descendants (e.g. a table with
   *   all its rows/cells, but no following content).
   * - `ProjectionDepth.SubtreeAndFollowingSiblings` (default) — for headings
   *   this returns the whole section (heading + content up to the next same-
   *   or-higher heading); for non-headings it behaves like `Subtree`.
   *
   * @see docs/architecture/docx_mutation_api.md
   */
  projectAnchor(
    anchorId: string,
    depth: ProjectionDepth = ProjectionDepth.SubtreeAndFollowingSiblings,
    citation?: PageCitationRequest,
  ): DocxSessionProjection {
    return JSON.parse(
      citation
        ? this.wasm.ProjectAnchorWithCitations(this.handle, anchorId, depth, JSON.stringify(citation))
        : this.wasm.ProjectAnchor(this.handle, anchorId, depth),
    ) as DocxSessionProjection;
  }

  /**
   * Render a single block to faithful HTML from the live session — the editor's
   * incremental per-block re-render after an edit. Resolves against the in-memory
   * document (no Save round-trip). `anchorId` is a block anchor (`kind:scope:unid`)
   * or the bare unid carried by a `data-anchor` attribute. Returns the block's HTML
   * element (no `<html>`/`<head>` wrapper).
   */
  renderBlock(
    anchorId: string,
    options?: { cssPrefix?: string; fabricateClasses?: boolean },
  ): string {
    const html = this.wasm.RenderBlockHtml(
      this.handle,
      anchorId,
      options?.cssPrefix ?? "docx-",
      options?.fabricateClasses ?? false,
    );
    // Rendered HTML always begins with '<'; a leading '{' signals an error object.
    if (html.charCodeAt(0) === 0x7b /* '{' */) {
      const err = JSON.parse(html) as { error?: string };
      throw new Error(`renderBlock failed: ${err.error ?? "unknown error"}`);
    }
    return html;
  }

  // ─── Tier A: text CRUD ───────────────────────────────────────────────

  replaceText(
    anchorId: string,
    markdown: string,
    preconditions?: MutationPreconditions,
  ): EditResult {
    const apply = () => JSON.parse(
      this.wasm.ReplaceText(this.handle, anchorId, markdown),
    ) as EditResult;
    return preconditions
      ? this.runWithPreconditions(
          { ...preconditions, anchorId: preconditions.anchorId ?? anchorId }, apply)
      : apply();
  }

  deleteBlock(anchorId: string, preconditions?: MutationPreconditions): EditResult {
    const apply = () => JSON.parse(this.wasm.DeleteBlock(this.handle, anchorId)) as EditResult;
    return preconditions
      ? this.runWithPreconditions(
          { ...preconditions, anchorId: preconditions.anchorId ?? anchorId }, apply)
      : apply();
  }

  /** Reorder one top-level paragraph/heading/list/table block relative to another. */
  moveBlock(sourceAnchorId: string, targetAnchorId: string, position: "before" | "after"): EditResult {
    return JSON.parse(
      this.wasm.MoveBlock(this.handle, sourceAnchorId, targetAnchorId, position),
    ) as EditResult;
  }

  /**
   * Delete every top-level block-level sibling between `fromAnchorId` (inclusive)
   * and `toAnchorIdExclusive` (exclusive). Both anchors must share a direct
   * parent and live in the same package part. Returns a single `EditResult`
   * whose `removed` lists every anchor that was deleted.
   *
   * Records ONE undo snapshot — `undo()` restores the entire range.
   *
   * @see docs/architecture/docx_mutation_api.md#deleterange
   */
  deleteRange(fromAnchorId: string, toAnchorIdExclusive: string): EditResult {
    return JSON.parse(this.wasm.DeleteRange(this.handle, fromAnchorId, toAnchorIdExclusive)) as EditResult;
  }

  /**
   * Delete a heading and everything below it up to (but not including) the next
   * heading at the same or higher level. The heading anchor must have `kind === "h"`.
   *
   * If the target is the last heading in its parent, the section extends to the
   * end of the parent (heading + everything after).
   *
   * @see docs/architecture/docx_mutation_api.md#deletesection
   */
  deleteSection(headingAnchorId: string): EditResult {
    return JSON.parse(this.wasm.DeleteSection(this.handle, headingAnchorId)) as EditResult;
  }

  // ─── Tier B: structural ──────────────────────────────────────────────

  insertParagraph(anchorId: string, position: "before" | "after", markdown: string): EditResult {
    return JSON.parse(this.wasm.InsertParagraph(this.handle, anchorId, position, markdown)) as EditResult;
  }

  splitParagraph(anchorId: string, characterOffset: number): EditResult {
    return JSON.parse(this.wasm.SplitParagraph(this.handle, anchorId, characterOffset)) as EditResult;
  }

  mergeParagraphs(firstAnchorId: string, secondAnchorId: string): EditResult {
    return JSON.parse(this.wasm.MergeParagraphs(this.handle, firstAnchorId, secondAnchorId)) as EditResult;
  }

  /**
   * Insert an empty paragraph carrying a bottom border — an S-1-style horizontal rule —
   * before/after the block. `rule` styles the line (default: a single ≈1.5pt black rule).
   */
  insertHorizontalRule(
    anchorId: string,
    position: "before" | "after",
    rule?: ParagraphBorderEdge,
  ): EditResult {
    const ruleJson = rule ? JSON.stringify(rule) : "";
    return JSON.parse(
      this.wasm.InsertHorizontalRule(this.handle, anchorId, position, ruleJson),
    ) as EditResult;
  }

  /**
   * Insert a `rows`×`cols` table before/after the block. `options` controls borders, row-major
   * cell markdown, and cell alignment. The returned `EditResult.created` lists canonical `tc`
   * anchors (row-major), so each cell can then be addressed to fill/format.
   */
  insertTable(
    anchorId: string,
    position: "before" | "after",
    rows: number,
    cols: number,
    options?: TableInsertOptions,
  ): EditResult {
    const optionsJson = options ? JSON.stringify(options) : "";
    return JSON.parse(
      this.wasm.InsertTable(this.handle, anchorId, position, rows, cols, optionsJson),
    ) as EditResult;
  }

  /** Resolve a canonical `tbl` anchor to explicit table/row/column/cell identities. */
  getTableMetadata(tableAnchorId: string): TableMetadataResult {
    return JSON.parse(this.wasm.GetTableMetadata(this.handle, tableAnchorId)) as TableMetadataResult;
  }

  /** Resolve a canonical `tc` anchor to its zero-based table-grid coordinate and spans. */
  resolveTableCellAnchor(cellAnchorId: string): TableCellResolutionResult {
    return JSON.parse(
      this.wasm.ResolveTableCellAnchor(this.handle, cellAnchorId),
    ) as TableCellResolutionResult;
  }

  /** Resolve a zero-based table-grid coordinate to the physical `tc` covering it. */
  resolveTableCellCoordinate(
    tableAnchorId: string,
    rowIndex: number,
    columnIndex: number,
  ): TableCellResolutionResult {
    return JSON.parse(
      this.wasm.ResolveTableCellCoordinate(this.handle, tableAnchorId, rowIndex, columnIndex),
    ) as TableCellResolutionResult;
  }

  /**
   * Table row/column editing, addressed by the canonical `tc` anchor returned from
   * {@link insertTable}'s `created` or table metadata. Insert clones the reference row/column's
   * widths and starts empty (`created` lists new `tc` anchors); delete of the last row/column removes
   * the whole table. All four are grid-aware: inserting across a merge extends it, deleting
   * through one narrows it, and deleting a vertical merge's lead row promotes the next row to
   * carry it — the grid is never left ragged.
   */
  insertTableRow(cellAnchorId: string, position: "before" | "after"): EditResult {
    return JSON.parse(this.wasm.InsertTableRow(this.handle, cellAnchorId, position)) as EditResult;
  }

  insertTableColumn(cellAnchorId: string, position: "before" | "after"): EditResult {
    return JSON.parse(this.wasm.InsertTableColumn(this.handle, cellAnchorId, position)) as EditResult;
  }

  deleteTableRow(cellAnchorId: string): EditResult {
    return JSON.parse(this.wasm.DeleteTableRow(this.handle, cellAnchorId)) as EditResult;
  }

  deleteTableColumn(cellAnchorId: string): EditResult {
    return JSON.parse(this.wasm.DeleteTableColumn(this.handle, cellAnchorId)) as EditResult;
  }

  /**
   * Merge the rectangle of cells anchored at `cellAnchorId` running `rowSpan` rows down ×
   * `colSpan` cells right (Word's *Merge Cells*): `w:gridSpan` for the horizontal extent,
   * `w:vMerge` restart/continue for the vertical one. The rectangle must tile the same whole grid
   * columns in every row it covers and must not clip a vertical merge entering from above or
   * continuing below — a partial overlap fails with `invalid_table_merge` instead of tearing the
   * grid. `content` decides what happens to the absorbed cells' content (default `"append"`).
   */
  mergeCells(
    cellAnchorId: string,
    rowSpan: number,
    colSpan: number,
    content: TableMergeContent = "append",
  ): EditResult {
    return JSON.parse(
      this.wasm.MergeCells(this.handle, cellAnchorId, rowSpan, colSpan, content),
    ) as EditResult;
  }

  /**
   * Split the merged cell at `cellAnchorId` back into unit cells, dropping its `w:gridSpan` and
   * `w:vMerge` markup and restoring one cell per grid column (each taking its `w:tblGrid` width).
   * Addressing a vertical-merge continuation unmerges the whole run. A cell with no merge markup
   * fails with `invalid_table_merge`.
   */
  unmergeCells(cellAnchorId: string): EditResult {
    return JSON.parse(this.wasm.UnmergeCells(this.handle, cellAnchorId)) as EditResult;
  }

  /**
   * Table styling, addressed by a canonical `tc` anchor — the post-insert counterpart of
   * {@link insertTable}'s options (issue #315 Stage A). `setColumnWidths` retunes `w:tblGrid` +
   * every row's cell width (one positive twip value per column) and pins the table to fixed
   * layout, exactly as inserting with explicit `columnWidths` would.
   */
  setColumnWidths(cellAnchorId: string, widthsTwips: number[]): EditResult {
    return JSON.parse(
      this.wasm.SetColumnWidths(this.handle, cellAnchorId, JSON.stringify(widthsTwips)),
    ) as EditResult;
  }

  /**
   * Set the table-level borders (`w:tblPr/w:tblBorders`) of the table containing the anchor.
   * Only the edges named by `spec.scope` are written; the rest are left untouched. Style
   * `"none"` removes the targeted edges. Omitting `spec` writes a thin single border all round.
   */
  setTableBorders(cellAnchorId: string, spec?: TableBorderSpec): EditResult {
    return JSON.parse(
      this.wasm.SetTableBorders(this.handle, cellAnchorId, spec ? JSON.stringify(spec) : ""),
    ) as EditResult;
  }

  /**
   * Shade the cell containing the anchor — or, with scope `"row"`, every cell of its row
   * (header-row banding). `fill` is a hex RRGGBB triplet (leading '#' tolerated) or `"auto"`;
   * `null` removes the shading.
   */
  setCellShading(
    cellAnchorId: string,
    fill: string | null,
    scope: TableShadingScope = "cell",
  ): EditResult {
    return JSON.parse(
      this.wasm.SetCellShading(this.handle, cellAnchorId, fill ?? "", scope),
    ) as EditResult;
  }

  /**
   * Mark (or unmark) the row containing the anchor as a repeating header row
   * (`w:trPr/w:tblHeader`), so a multi-page table re-shows it on every page. Word only honors
   * the flag on a run of rows starting at the table's first row.
   */
  setRepeatHeaderRow(cellAnchorId: string, repeat: boolean): EditResult {
    return JSON.parse(
      this.wasm.SetRepeatHeaderRow(this.handle, cellAnchorId, repeat),
    ) as EditResult;
  }

  /** Apply row layout options to the row containing the canonical cell anchor. */
  setTableRowOptions(cellAnchorId: string, options: TableRowOptions): EditResult {
    return JSON.parse(
      this.wasm.SetTableRowOptions(
        this.handle,
        cellAnchorId,
        options.repeatHeader ?? null,
        options.allowBreakAcrossPages ?? null,
        options.heightTwips ?? null,
        options.heightRule ?? "atLeast",
      ),
    ) as EditResult;
  }

  // ─── Headers / footers / page numbers ────────────────────────────────

  /**
   * Set the running header story for the section that owns `anchorId` (any body block in that
   * section) to `markdown`. Creates the header part + `w:headerReference` if the story of `kind`
   * doesn't exist yet, else replaces its content. The created header-paragraph anchors (scope
   * `hdr{N}`) come back in `EditResult.created` — insert a page number into one with
   * {@link insertPageNumberField}. `"first"` sets the section's title-page flag; `"even"` sets the
   * document's even/odd-headers flag.
   */
  setHeaderText(anchorId: string, kind: HeaderFooterKind, markdown: string): EditResult {
    return JSON.parse(this.wasm.SetHeaderText(this.handle, anchorId, kind, markdown)) as EditResult;
  }

  /** Set the running footer story for the section that owns `anchorId` — see {@link setHeaderText};
   * the created footer-paragraph anchors (scope `ftr{N}`) come back in `EditResult.created`. */
  setFooterText(anchorId: string, kind: HeaderFooterKind, markdown: string): EditResult {
    return JSON.parse(this.wasm.SetFooterText(this.handle, anchorId, kind, markdown)) as EditResult;
  }

  /**
   * Append a page-number field to the paragraph `anchorId` — typically a header/footer paragraph
   * returned by {@link setFooterText}/{@link setHeaderText}. `"currentPage"` emits a PAGE field,
   * `"totalPages"` a NUMPAGES field (native complex field with a cached result). Center it by
   * setting the paragraph alignment ({@link setParagraphFormat}). Returns the paragraph anchor in
   * `modified`.
   *
   * `format` writes the field's own `\*` general-formatting switch (`PAGE \* roman` → `i, ii, iii`).
   * Omitting it — the default — emits a plain field, which is what Word inserts and what follows the
   * SECTION's format ({@link setPageNumbering}). Prefer the section setting for ordinary page
   * numbering: a switch here overrides it for this one field and keeps overriding it if the section
   * later changes.
   */
  insertPageNumberField(
    anchorId: string,
    field: PageNumberField = "currentPage",
    format?: NumberFormat
  ): EditResult {
    return JSON.parse(
      this.wasm.InsertPageNumberField(this.handle, anchorId, field, format ?? "")
    ) as EditResult;
  }

  /**
   * Set the page-numbering properties (`w:pgNumType`) of the section that owns `anchorId` (any body
   * block in that section) — Word's *Format Page Numbers…* dialog: which number the section starts
   * at and which format its pages use. Omitted fields on `op` are left unchanged, so the start can
   * be set without disturbing the format and vice versa. Creates the element, and a trailing
   * `w:sectPr`, if absent.
   *
   * Applying values the section already has is a successful no-op that does NOT consume undo
   * history — safe to call from a dropdown's change handler.
   */
  setPageNumbering(anchorId: string, op: PageNumberingOp): EditResult {
    return JSON.parse(
      this.wasm.SetPageNumbering(this.handle, anchorId, JSON.stringify(op))
    ) as EditResult;
  }

  /**
   * Remove the section's page-numbering start/format: it reverts to continuing the previous
   * section's numbering in Word's default `1, 2, 3`. Chapter-numbering attributes
   * (`w:chapStyle`/`w:chapSep`) are preserved. A section with nothing to clear is a no-op.
   */
  clearPageNumbering(anchorId: string): EditResult {
    return JSON.parse(this.wasm.ClearPageNumbering(this.handle, anchorId)) as EditResult;
  }

  /**
   * Make the `kind` header/footer stories of the section that owns `anchorId` actually RENDER:
   * `"first"` sets `w:titlePg`, `"even"` sets the document-global `w:evenAndOddHeaders`;
   * `"default"` needs no flag and succeeds as a no-op. Idempotent.
   *
   * {@link setHeaderText}/{@link setFooterText} set these flags while writing content, which covers
   * authoring a story from scratch — but NOT a document that already carries a first/even reference
   * with the flag absent (Word leaves exactly that behind when "Different first page" is switched
   * off). Editing such a story through the text ops otherwise yields header content that is present
   * but invisible. Note the `"even"` caveat from {@link setHeaderText}: the flag is document-global
   * and governs footers too.
   */
  ensureHeaderFooterVisible(anchorId: string, kind: HeaderFooterKind): EditResult {
    return JSON.parse(this.wasm.EnsureHeaderFooterVisible(this.handle, anchorId, kind)) as EditResult;
  }

  // ─── Footnotes / endnotes ────────────────────────────────────────────

  /**
   * Create a footnote whose body is `markdown` and cite it from the body paragraph `anchorId`, at
   * `characterOffset` characters into that paragraph's text (0 = before all text, text length =
   * after all of it). On a document with no footnotes yet this also creates the footnotes part,
   * Word's two reserved separator notes, the `FootnoteText`/`FootnoteReference` styles and the
   * `w:footnotePr` settings declaration; otherwise the existing part is reused. The note id is
   * allocated above every id already used in the package, so non-contiguous ids can't collide.
   *
   * Returns the created note anchors in `EditResult.created` — the definition (kind `fn`) and its
   * paragraphs (kind `p`, scope `fn`) — so the note can immediately be edited with
   * {@link replaceText} or removed with {@link deleteBlock} (which also drops the body reference).
   *
   * Body paragraphs only: Word does not allow a note reference inside a header/footer story or
   * inside another note, so a non-body anchor fails with `anchorWrongKind`.
   */
  insertFootnote(anchorId: string, characterOffset: number, markdown: string): EditResult {
    return JSON.parse(
      this.wasm.InsertFootnote(this.handle, anchorId, characterOffset, markdown),
    ) as EditResult;
  }

  /** Create an endnote — see {@link insertFootnote}; writes the endnotes part and a
   * `w:endnoteReference`, and the created definition anchor has kind `en`. */
  insertEndnote(anchorId: string, characterOffset: number, markdown: string): EditResult {
    return JSON.parse(
      this.wasm.InsertEndnote(this.handle, anchorId, characterOffset, markdown),
    ) as EditResult;
  }

  // ─── Comments (issue #300) ───────────────────────────────────────────

  /**
   * Add a **native Word comment** (real `w:comment` markup, visible in Word/Google
   * Docs/LibreOffice's Reviewing pane — not the {@link addAnnotation} overlay) on the body
   * paragraph `anchorId`. `span` selects the commented character range; `null` comments the
   * whole block. On a document with no comments yet this also creates the comments part and
   * the `CommentText`/`CommentReference` styles. `opts.date` (ISO-8601) is written only when
   * provided, keeping output deterministic by default.
   *
   * Returns the created definition anchor (kind `cmt`) and its paragraph anchors (kind `p`,
   * scope `cmt`) in `EditResult.created`, so the comment can immediately be edited with
   * {@link updateComment} or removed with {@link removeComment}.
   *
   * Body paragraphs only (Word has no comments-on-comments); a non-body anchor fails with
   * `anchor_wrong_kind`, a zero-length span with `empty_comment_span`.
   */
  addComment(
    anchorId: string,
    span: CharSpan | null,
    author: string,
    markdown: string,
    opts?: { initials?: string; date?: string },
  ): EditResult {
    const spanJson = span ? JSON.stringify(span) : "";
    return JSON.parse(
      this.wasm.AddComment(
        this.handle,
        anchorId,
        spanJson,
        author,
        opts?.initials ?? "",
        opts?.date ?? "",
        markdown,
      ),
    ) as EditResult;
  }

  /** Add a native Word comment around the exact live extent of a tracked revision returned by
   * {@link listRevisions}. Accepting/rejecting the revision keeps the comment and either leaves
   * its range on surviving text or collapses it to a point. */
  addCommentToRevision(
    revisionId: string,
    author: string,
    markdown: string,
    opts?: { initials?: string; date?: string },
  ): EditResult {
    return JSON.parse(
      this.wasm.AddCommentToRevision(
        this.handle,
        revisionId,
        author,
        opts?.initials ?? "",
        opts?.date ?? "",
        markdown,
      ),
    ) as EditResult;
  }

  /**
   * Add a native Word reply to `parentCommentAnchorId`. It adds an adjacent reference and
   * inherits the thread root's range through `w15:paraIdParent`; the required
   * `commentsExtended.xml` and `commentsIds.xml` parts are created when the parent was flat.
   */
  addCommentReply(
    parentCommentAnchorId: string,
    author: string,
    markdown: string,
    opts?: { initials?: string; date?: string },
  ): EditResult {
    return JSON.parse(
      this.wasm.AddCommentReply(
        this.handle,
        parentCommentAnchorId,
        author,
        opts?.initials ?? "",
        opts?.date ?? "",
        markdown,
      ),
    ) as EditResult;
  }

  /** Replace a comment's body text, addressed by its definition anchor (kind `cmt`); the
   * comment's author/initials/date are preserved, as is the last paragraph's `w14:paraId`
   * (Word's reply-threading key). */
  updateComment(commentAnchorId: string, markdown: string): EditResult {
    return JSON.parse(this.wasm.UpdateComment(this.handle, commentAnchorId, markdown)) as EditResult;
  }

  /** Resolve or reopen one comment (`false` reopens it). Flat comments are upgraded with the
   * paraId-keyed metadata Word uses, and the operation is undoable. */
  setCommentResolved(commentAnchorId: string, resolved: boolean): EditResult {
    return JSON.parse(
      this.wasm.SetCommentResolved(this.handle, commentAnchorId, resolved),
    ) as EditResult;
  }

  /** Remove a comment: the definition, its body marker triple everywhere in the package, and
   * any `commentsExtended`/`commentsIds` threading entries keyed by it. */
  removeComment(commentAnchorId: string): EditResult {
    return JSON.parse(this.wasm.RemoveComment(this.handle, commentAnchorId)) as EditResult;
  }

  /** The document's native Word comments in comments-part order. */
  listComments(): CommentListEntry[] {
    return JSON.parse(this.wasm.ListComments(this.handle)) as CommentListEntry[];
  }

  listHyperlinks(scopes: ProjectionScopes = ProjectionScopes.All): HyperlinkInfo[] {
    return JSON.parse(this.wasm.ListHyperlinks(this.handle, scopes)) as HyperlinkInfo[];
  }

  addHyperlink(anchorId: string, span: CharSpan, kind: HyperlinkKind, target: string): EditResult {
    return JSON.parse(this.wasm.AddHyperlink(
      this.handle, anchorId, span.start, span.length, kind, target,
    )) as EditResult;
  }

  updateHyperlink(hyperlinkId: string, kind: HyperlinkKind, target: string): EditResult {
    return JSON.parse(this.wasm.UpdateHyperlink(this.handle, hyperlinkId, kind, target)) as EditResult;
  }

  removeHyperlink(hyperlinkId: string): EditResult {
    return JSON.parse(this.wasm.RemoveHyperlink(this.handle, hyperlinkId)) as EditResult;
  }

  /** Versioned operational facts for native image inspection/mutation in this runtime. */
  getImageCapabilities(): ImageCapabilities {
    return JSON.parse(this.wasm.GetImageCapabilities()) as ImageCapabilities;
  }

  listImages(scopes: ProjectionScopes = ProjectionScopes.All): ImageOccurrence[] {
    return JSON.parse(this.wasm.ListImages(this.handle, scopes)) as ImageOccurrence[];
  }

  insertImage(anchorId: string, characterOffset: number, bytes: Uint8Array,
    options: ImageInsertOptions = {}): EditResult {
    return JSON.parse(this.wasm.InsertImage(this.handle, anchorId, characterOffset,
      imageBytesToBase64(bytes), JSON.stringify(options))) as EditResult;
  }

  replaceImage(imageId: string, bytes: Uint8Array): EditResult {
    return JSON.parse(this.wasm.ReplaceImage(
      this.handle, imageId, imageBytesToBase64(bytes))) as EditResult;
  }

  setImageDimensions(imageId: string, dimensions: ImageDimensions): EditResult {
    return JSON.parse(this.wasm.SetImageDimensions(
      this.handle, imageId, JSON.stringify(dimensions))) as EditResult;
  }

  setImageMetadata(imageId: string, altText: string | null, title: string | null): EditResult {
    return JSON.parse(this.wasm.SetImageMetadata(
      this.handle, imageId, altText, title)) as EditResult;
  }

  setImageFloatingLayout(imageId: string, layout: FloatingImageLayout): EditResult {
    return JSON.parse(this.wasm.SetImageFloatingLayout(
      this.handle, imageId, JSON.stringify(layout))) as EditResult;
  }

  removeImage(imageId: string): EditResult {
    return JSON.parse(this.wasm.RemoveImage(this.handle, imageId)) as EditResult;
  }

  /** Native Word structured-document tags, in outer-before-inner story order. */
  listContentControls(scopes: ProjectionScopes = ProjectionScopes.All): ContentControlInfo[] {
    return JSON.parse(this.wasm.ListContentControls(this.handle, scopes)) as ContentControlInfo[];
  }

  fillContentControlText(anchorId: string, text: string,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.FillContentControlText(
      this.handle, anchorId, text, JSON.stringify(options))) as EditResult;
  }

  fillContentControlRichText(anchorId: string, markdown: string,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.FillContentControlRichText(
      this.handle, anchorId, markdown, JSON.stringify(options))) as EditResult;
  }

  setContentControlChecked(anchorId: string, isChecked: boolean,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.SetContentControlChecked(
      this.handle, anchorId, isChecked, JSON.stringify(options))) as EditResult;
  }

  setContentControlDate(anchorId: string, value: string | Date, displayText?: string,
    options: ContentControlFillOptions = {}): EditResult {
    const timestamp = value instanceof Date ? value.toISOString() : value;
    return JSON.parse(this.wasm.SetContentControlDate(
      this.handle, anchorId, timestamp, displayText ?? null, JSON.stringify(options))) as EditResult;
  }

  selectContentControlItem(anchorId: string, value: string,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.SelectContentControlItem(
      this.handle, anchorId, value, JSON.stringify(options))) as EditResult;
  }

  fillContentControlPicture(anchorId: string, bytes: Uint8Array,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.FillContentControlPicture(
      this.handle, anchorId, imageBytesToBase64(bytes), JSON.stringify(options))) as EditResult;
  }

  addRepeatingSectionItem(sectionAnchorId: string, afterItemAnchorId?: string,
    options: ContentControlFillOptions = {}): EditResult {
    return JSON.parse(this.wasm.AddRepeatingSectionItem(
      this.handle, sectionAnchorId, afterItemAnchorId ?? "", JSON.stringify(options))) as EditResult;
  }

  removeRepeatingSectionItem(itemAnchorId: string): EditResult {
    return JSON.parse(this.wasm.RemoveRepeatingSectionItem(
      this.handle, itemAnchorId)) as EditResult;
  }

  listBookmarks(scopes: ProjectionScopes = ProjectionScopes.All): BookmarkInfo[] {
    return JSON.parse(this.wasm.ListBookmarks(this.handle, scopes)) as BookmarkInfo[];
  }

  addBookmark(name: string, range: DocumentRange): EditResult {
    return JSON.parse(this.wasm.AddBookmark(this.handle, name,
      range.startAnchorId, range.startOffset, range.endAnchorId, range.endOffset)) as EditResult;
  }

  renameBookmark(name: string, newName: string): EditResult {
    return JSON.parse(this.wasm.RenameBookmark(this.handle, name, newName)) as EditResult;
  }

  moveBookmark(name: string, range: DocumentRange): EditResult {
    return JSON.parse(this.wasm.MoveBookmark(this.handle, name,
      range.startAnchorId, range.startOffset, range.endAnchorId, range.endOffset)) as EditResult;
  }

  removeBookmark(name: string): EditResult {
    return JSON.parse(this.wasm.RemoveBookmark(this.handle, name)) as EditResult;
  }

  // ─── Tracked revisions (issue #318) ──────────────────────────────────

  /** Markup-native tracked-revision listing, in document order across body, headers,
   * footers, footnotes, and endnotes. Ids are stable while the underlying markup
   * exists and address {@link acceptRevision}/{@link rejectRevision}; authors/dates
   * are the markup's own (no accept/reject re-diff). */
  listRevisions(): RevisionListEntry[] {
    return JSON.parse(this.wasm.ListRevisions(this.handle)) as RevisionListEntry[];
  }

  /** Accept ONE revision by the id {@link listRevisions} reported — insertions keep
   * their content, deletions are carried out, a move materializes at its destination,
   * a format change keeps the new properties. Undoable. */
  acceptRevision(revisionId: string): EditResult {
    return JSON.parse(this.wasm.AcceptRevision(this.handle, revisionId)) as EditResult;
  }

  /** Reject ONE revision by id — the inverse of {@link acceptRevision}: insertions are
   * removed, deleted content is restored, a move stays at its source, a format change
   * restores the stored old properties. Undoable. */
  rejectRevision(revisionId: string): EditResult {
    return JSON.parse(this.wasm.RejectRevision(this.handle, revisionId)) as EditResult;
  }

  /**
   * Accept every live revision as one undoable session mutation.
   *
   * Fails closed: an unsupported, malformed, or ambiguous registry entry aborts the whole
   * operation (`revisionUnsupported`/`revisionMalformed`/`revisionAmbiguous`) and nothing is
   * mutated. There is no force mode — call {@link listRevisions} and read each entry's
   * `diagnostic` to see what blocks it.
   */
  acceptAllRevisions(): EditResult {
    return JSON.parse(this.wasm.AcceptAllRevisions(this.handle)) as EditResult;
  }

  /** Reject every live revision as one undoable session mutation. Fails closed exactly like
   * {@link acceptAllRevisions}. */
  rejectAllRevisions(): EditResult {
    return JSON.parse(this.wasm.RejectAllRevisions(this.handle)) as EditResult;
  }

  // ─── Tier C: formatting ──────────────────────────────────────────────

  applyFormat(anchorId: string, span: CharSpan | null, op: FormatOp): EditResult {
    const spanJson = span ? JSON.stringify(span) : "";
    return JSON.parse(this.wasm.ApplyFormat(this.handle, anchorId, spanJson, JSON.stringify(op))) as EditResult;
  }

  /**
   * Convenience: find `substring` in the anchor's flat text and apply `op` to the
   * first occurrence. Eliminates the offset-arithmetic trap from #138 — caller passes
   * the visible text they want formatted, the WASM-side resolves it to a CharSpan.
   */
  applyFormatBySubstring(anchorId: string, substring: string, op: FormatOp): EditResult {
    return JSON.parse(
      this.wasm.ApplyFormatBySubstring(this.handle, anchorId, substring, JSON.stringify(op))
    ) as EditResult;
  }

  /**
   * Convenience: apply `op` to the exact span of a {@link TextMatch} (typically from
   * {@link grep}). The match's `enclosingAnchor.id` + `span` address one specific
   * occurrence even when several identical needles share the same block.
   */
  applyFormatToMatch(match: TextMatch, op: FormatOp): EditResult {
    const span: CharSpan = { start: match.span.start, length: match.span.length };
    return this.applyFormat(match.enclosingAnchor.id, span, op);
  }

  setParagraphStyle(anchorId: string, styleId: string): EditResult {
    return JSON.parse(this.wasm.SetParagraphStyle(this.handle, anchorId, styleId)) as EditResult;
  }

  /** Set paragraph alignment / indent / page-break-before (omitted fields are left unchanged). */
  setParagraphFormat(anchorId: string, op: ParagraphFormatOp): EditResult {
    return JSON.parse(this.wasm.SetParagraphFormat(this.handle, anchorId, JSON.stringify(op))) as EditResult;
  }

  setListLevel(anchorId: string, levelDelta: number): EditResult {
    return JSON.parse(this.wasm.SetListLevel(this.handle, anchorId, levelDelta)) as EditResult;
  }

  removeListMembership(anchorId: string): EditResult {
    return JSON.parse(this.wasm.RemoveListMembership(this.handle, anchorId)) as EditResult;
  }

  /** Make the paragraph a bullet/numbered list item, or remove list membership ("none"). */
  applyListFormat(anchorId: string, kind: ListFormat): EditResult {
    return JSON.parse(this.wasm.ApplyListFormat(this.handle, anchorId, kind)) as EditResult;
  }

  /** Apply one list format across the contiguous sibling run from `firstAnchorId` to
   * `lastAnchorId` inclusive (either document order). Every member shares one `w:num`
   * instance so the numbering sequence stays intact; the whole range is a single undo step. */
  applyListFormatRange(firstAnchorId: string, lastAnchorId: string, kind: ListFormat): EditResult {
    return JSON.parse(
      this.wasm.ApplyListFormatRange(this.handle, firstAnchorId, lastAnchorId, kind),
    ) as EditResult;
  }

  /** Restart the anchored list item's numbering at `value` — Word's *Set Numbering Value…*.
   * Writes a `w:startOverride` on a dedicated `w:num` instance and repoints the anchored item
   * plus every following member of its sequence, so a mid-list restart splits the sequence
   * exactly like Word (earlier items keep their numbers, the tail continues from `value`). */
  setListStartOverride(anchorId: string, value: number): EditResult {
    return JSON.parse(this.wasm.SetListStartOverride(this.handle, anchorId, value)) as EditResult;
  }

  /** Remove the numbering restart from the anchored item's whole sequence (the inverse of
   * {@link setListStartOverride}); the sequence reverts to the definition's own start. A
   * sequence with no override at the item's level is a successful no-op. */
  clearListStartOverride(anchorId: string): EditResult {
    return JSON.parse(this.wasm.ClearListStartOverride(this.handle, anchorId)) as EditResult;
  }

  // ─── Tier D: cell content ────────────────────────────────────────────

  replaceCellContent(cellAnchorId: string, markdown: string): EditResult {
    return JSON.parse(this.wasm.ReplaceCellContent(this.handle, cellAnchorId, markdown)) as EditResult;
  }

  // ─── Raw escape hatch ────────────────────────────────────────────────

  readonly raw = {
    getXml: (anchorId: string): string => this.wasm.RawGetXml(this.handle, anchorId),
    insertXml: (anchorId: string, position: "before" | "after", xml: string): EditResult =>
      JSON.parse(this.wasm.RawInsertXml(this.handle, anchorId, position, xml)) as EditResult,
    replaceXml: (anchorId: string, xml: string): EditResult =>
      JSON.parse(this.wasm.RawReplaceXml(this.handle, anchorId, xml)) as EditResult,
  };

  // ─── Search ──────────────────────────────────────────────────────────

  /**
   * Searches the flat text of every paragraph/heading/list-item in scope for
   * matches of `pattern`, returning them in document order with the run
   * fragments each match spans. Lets callers rewrite a match in place while
   * preserving each fragment's formatting (bold/italic/hyperlink/etc.).
   *
   * `pattern` is a regular expression — use plain string equivalents wrapped
   * in `^` / `$` or pass literal text escaped via a helper.
   *
   * @see docs/architecture/docx_mutation_api.md#grep
   */
  grep(pattern: string, options?: GrepOptions): TextMatch[] {
    return JSON.parse(this.wasm.Grep(this.handle, pattern, options ? JSON.stringify(options) : "")) as TextMatch[];
  }

  /**
   * Like {@link grep}, but lets a single match span adjacent block-level
   * siblings (paragraphs/headings/list items) under the same parent. Block
   * boundaries appear in the matched text as `\n`, so `^`/`$` with the
   * Multiline flag anchor at boundaries and `.` won't cross unless Singleline
   * is set.
   *
   * Matches never cross OOXML package parts, container boundaries (body →
   * table cell), or non-paragraph siblings (a table between two paragraphs
   * breaks the run). Returned superset of {@link grep}: single-block matches
   * still appear with one slice. Filter `slices.length > 1` for cross-block only.
   *
   * @see docs/architecture/docx_mutation_api.md#grepcrossblock
   */
  grepCrossBlock(pattern: string, options?: GrepOptions): CrossBlockMatch[] {
    return JSON.parse(
      this.wasm.GrepCrossBlock(this.handle, pattern, options ? JSON.stringify(options) : "")
    ) as CrossBlockMatch[];
  }

  /**
   * Finds every literal occurrence of `find` in the anchor's flat text and
   * replaces it with `replace`, preserving the surrounding run formatting that
   * the match didn't touch. Returns one `EditResult` per attempted match.
   *
   * Run-formatting contract: the replacement text inherits the formatting of
   * the FIRST run the match spanned. Middle/trailing runs keep their `w:rPr`
   * but lose the slice of text the match consumed.
   *
   * @see docs/architecture/docx_mutation_api.md#replacetextrange
   */
  replaceTextRange(anchorId: string, find: string, replace: string, options?: ReplaceOptions): EditResult[] {
    return JSON.parse(
      this.wasm.ReplaceTextRange(this.handle, anchorId, find, replace, options ? JSON.stringify(options) : "")
    ) as EditResult[];
  }

  /**
   * Replaces a specific Grep match in place — addresses the exact span by
   * `enclosingAnchor.id` + `span.{start,length}`, so identical needles in the
   * same paragraph (the template-fill case where five `[___]` placeholders
   * each get a different value) don't collide.
   */
  replaceMatch(match: TextMatch, replace: string): EditResult {
    return JSON.parse(
      this.wasm.ReplaceTextAtSpan(this.handle, match.enclosingAnchor.id, match.span.start, match.span.length, replace)
    ) as EditResult;
  }

  /**
   * Helper for {@link fillPlaceholders} `coalesceWhitespaceAroundEmptyFill` path —
   * mirrors the .NET `ReplaceMatchCoalescingNeighbors` rules. Inspects the chars
   * immediately surrounding the match via `match.contextBefore` / `contextAfter`
   * (so the option requires `contextChars >= 1`, the default) and expands the
   * deletion span to absorb whitespace / leading-space-before-punctuation /
   * matched-brackets where the patterns match. Falls back to literal-delete
   * when no neighbor pattern applies.
   *
   * Note: with `boundary: ContextBoundary.Bracket`, neighbor brackets are not
   * captured in context, so the bracket-coalesce rule won't fire on the JS side.
   * The .NET implementation reads flat text directly and handles that case;
   * callers who care should leave `boundary` at the default `Char`.
   */
  private replaceMatchCoalescingNeighbors(match: TextMatch): EditResult {
    // Fold NBSP / narrow NBSP / thin space to ASCII space so e.g. an NBSP on
    // either side still gets treated as whitespace by the rules below.
    const fold = (c: string | undefined): string | undefined => {
      if (c === " " || c === " " || c === " ") return " ";
      return c;
    };
    const l = fold(match.contextBefore.length > 0 ? match.contextBefore[match.contextBefore.length - 1] : undefined);
    const r = fold(match.contextAfter.length > 0 ? match.contextAfter[0] : undefined);

    const isSpace = (c: string | undefined): boolean => c === " " || c === "\t";
    const isClauseTerm = (c: string | undefined): boolean =>
      c === "." || c === "," || c === ";" || c === ":" || c === "!" || c === "?";
    const isOpen = (c: string | undefined): boolean => c === "(" || c === "[" || c === "{";
    const isClose = (c: string | undefined): boolean => c === ")" || c === "]" || c === "}";

    let extendLeft = 0;
    let extendRight = 0;
    if (isSpace(l) && isSpace(r)) {
      extendRight = 1;
    } else if (isSpace(l) && isClauseTerm(r)) {
      extendLeft = 1;
    } else if (isOpen(l) && isClose(r)) {
      extendLeft = 1;
      extendRight = 1;
    }

    if (extendLeft === 0 && extendRight === 0) {
      return this.replaceMatch(match, "");
    }

    return JSON.parse(
      this.wasm.ReplaceTextAtSpan(
        this.handle,
        match.enclosingAnchor.id,
        match.span.start - extendLeft,
        match.span.length + extendLeft + extendRight,
        "",
      ),
    ) as EditResult;
  }

  /**
   * Replace the bracketed portion of a `TextMatch` with `newInner`, preserving any
   * prefix or suffix outside the brackets. Designed for `findPlaceholders` matches
   * like `$[___]` where the regex `\$?\[…\]` captures a leading `$`:
   * `replaceInner(match, "0.20")` yields `$0.20`, not `0.20`.
   *
   * Returns `MalformedMarkdown` if the match text does not contain balanced brackets.
   */
  replaceInner(match: TextMatch, newInner: string): EditResult {
    return JSON.parse(this.wasm.ReplaceInner(
      this.handle,
      match.text,
      match.enclosingAnchor.id,
      match.span.start,
      match.span.length,
      newInner,
    )) as EditResult;
  }

  /**
   * Picker-driven template fill. For every placeholder matching `options.kinds`,
   * calls `picker`; if the picker returns a non-null string, the placeholder is
   * replaced (with optional `$`-prefix preservation). Iterates until no more
   * placeholders match (or `maxPasses` is reached, or a pass makes zero changes)
   * — handles nested brackets that surface only after the inner ones are stripped.
   *
   * The TypeScript implementation mirrors the .NET `DocxSession.FillPlaceholders`
   * exactly.
   *
   * The picker is invoked synchronously by this loop on the JS side (it does
   * NOT run inside the WASM module). Async pickers are not supported: returning
   * a `Promise` will cause a `TypeError` at runtime inside the `$`-prefix
   * preservation branch (`Promise.startsWith is not a function`). For async
   * data, pre-build a lookup map before calling and have the picker read from
   * it synchronously.
   */
  fillPlaceholders(
    picker: (p: TemplatePlaceholder) => string | null | undefined,
    options?: FillOptions,
  ): BulkEditResult {
    const opts = options ?? {};
    // Default Kinds = All so the picker is invoked for every kind the doc contains.
    // Callers that want to ignore AlternativeClause matches should narrow this to
    // `PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction`.
    const kinds = opts.kinds ?? PlaceholderKinds.All;
    const scope = opts.scope ?? 1; // Body
    const maxPasses = opts.maxPasses ?? 8;
    const preserveDollarPrefix = opts.preserveDollarPrefix ?? true;
    const contextChars = opts.contextChars ?? 80;
    const boundary = opts.boundary ?? ContextBoundary.Char;
    const coalesceEmpty = opts.coalesceWhitespaceAroundEmptyFill ?? false;

    if (maxPasses <= 0) {
      throw new RangeError("FillOptions.maxPasses must be > 0");
    }

    let filled = 0;
    let workPasses = 0;
    const errors: EditError[] = [];
    const unfilled: TemplatePlaceholder[] = [];
    const seenSkipKeys = new Set<string>();

    for (let pass = 1; pass <= maxPasses; pass++) {
      const placeholders = this.findPlaceholders(kinds, scope, contextChars, boundary)
        .sort((a, b) => {
          const cmp = b.match.enclosingAnchor.id.localeCompare(a.match.enclosingAnchor.id);
          if (cmp !== 0) return cmp;
          return b.match.span.start - a.match.span.start;
        });
      if (placeholders.length === 0) break;

      let passChanges = 0;
      for (const p of placeholders) {
        const pick = picker(p);
        if (pick == null) {
          const key = `${p.match.enclosingAnchor.id}:${p.match.span.start}:${p.match.span.length}`;
          if (!seenSkipKeys.has(key)) {
            seenSkipKeys.add(key);
            unfilled.push(p);
          }
          continue;
        }

        let replacement = pick;
        if (preserveDollarPrefix && p.match.text.startsWith("$") && !replacement.startsWith("$")) {
          replacement = "$" + replacement;
        }

        const r = coalesceEmpty && replacement.length === 0
          ? this.replaceMatchCoalescingNeighbors(p.match)
          : this.replaceMatch(p.match, replacement);
        if (r.success) {
          filled++;
          passChanges++;
        } else if (r.error) {
          errors.push(r.error);
        }
      }

      if (passChanges > 0) workPasses = pass;
      if (passChanges === 0) break;
    }

    const stillPresent = this.findPlaceholders(kinds, scope).length;

    return {
      filled,
      skipped: unfilled.length,
      stillPresent,
      passes: workPasses,
      unfilled,
      errors,
    };
  }

  /**
   * Enumerate template placeholders in the document. Thin classifier over
   * {@link grep}: distinguishes `[___]` value blanks (`blank_fill`),
   * `[bracketed alternative clauses]` (`alternative_clause`), and
   * `[insert X]` / `[*italic hint*]` instructions (`instruction`).
   *
   * Combine kinds with bitwise OR: `PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction`.
   * Default is `PlaceholderKinds.All`; default scope is body only (1).
   *
   * @see docs/architecture/docx_mutation_api.md#findplaceholders
   */
  findPlaceholders(
    kinds: number = PlaceholderKinds.All,
    scope: number = 1,
    contextChars: number = 80,
    boundary: number = ContextBoundary.Char,
    citation?: PageCitationRequest,
  ): TemplatePlaceholder[] {
    const json = citation
      ? this.wasm.FindPlaceholdersWithCitations(
          this.handle, kinds, scope, contextChars, boundary, JSON.stringify(citation),
        )
      : this.wasm.FindPlaceholders(this.handle, kinds, scope, contextChars, boundary);
    return JSON.parse(json) as TemplatePlaceholder[];
  }

  /**
   * Returns a snapshot of edit-state introspection signals — placeholder counts,
   * underscore-run leftovers, footnote/comment counts. Useful for "am I done?"
   * verification at the end of an edit pipeline.
   */
  getEditSummary(): EditSummary {
    return JSON.parse(this.wasm.GetEditSummary(this.handle)) as EditSummary;
  }

  /**
   * Discoverability alias for {@link findPlaceholders}. Same return shape.
   */
  remainingPlaceholders(kinds: number = PlaceholderKinds.All): TemplatePlaceholder[] {
    return JSON.parse(this.wasm.RemainingPlaceholders(this.handle, kinds)) as TemplatePlaceholder[];
  }

  /**
   * Diff the document's current projection against the projection captured at
   * session construction time.
   *
   * Requires `captureInitialProjection: true` in {@link DocxSessionSettings}
   * (the default). Throws if not enabled.
   *
   * The return type depends on `format`:
   * - `DiffFormat.Json` (default) — structured anchor-keyed `DiffEntry[]`.
   * - `DiffFormat.Unified` — `patch(1)`-compatible unified-diff text;
   *   empty string when nothing has changed.
   * - `DiffFormat.SideBySide` — two-column human-review text
   *   (`diff -y` style).
   */
  getDiff(format?: typeof DiffFormat.Json): DiffEntry[];
  getDiff(format: typeof DiffFormat.Unified | typeof DiffFormat.SideBySide): string;
  getDiff(format: number = DiffFormat.Json): DiffEntry[] | string {
    const raw = this.wasm.GetDiff(this.handle, format);
    if (format === DiffFormat.Json) {
      return JSON.parse(raw) as DiffEntry[];
    }
    return raw;
  }

  // ─── Annotation-based anchor discovery (#132) ────────────────────────

  /**
   * Resolves an annotation's range to the block-level markdown anchors covering
   * it, in document order. The bridge between Docxodus' read-side annotation API
   * and the write-side session: an agent that wants to edit "the indemnification
   * clause" looks the annotation up by id and gets the anchors it can hand to
   * {@link replaceText} / {@link deleteBlock} / {@link raw}. Returns an empty
   * list when the id is unknown or its bookmark is missing.
   *
   * v1 returns the enclosing block anchors — every paragraph/heading/list-item/
   * cell/row/table whose subtree overlaps the bookmark range. Filter by
   * `kind === "p" | "h" | "li"` when you want only text-bearing blocks.
   *
   * @see docs/architecture/docx_mutation_api.md#findbyannotation
   */
  findByAnnotation(annotationId: string, citation?: PageCitationRequest): AnchorTargetRef[] {
    const json = citation
      ? this.wasm.FindByAnnotationWithCitations(this.handle, annotationId, JSON.stringify(citation))
      : this.wasm.FindByAnnotation(this.handle, annotationId);
    return JSON.parse(json) as AnchorTargetRef[];
  }

  /**
   * Finds every annotation whose `labelId` matches and resolves each of their
   * ranges. The result is keyed by annotation id so callers can disambiguate
   * when the same label is applied to multiple regions (three "WARRANTY"
   * annotations on different paragraphs become three entries). Annotations
   * whose bookmark resolves to no anchors are omitted from the result.
   */
  findByLabel(labelId: string, citation?: PageCitationRequest): Record<string, AnchorTargetRef[]> {
    const json = citation
      ? this.wasm.FindByLabelWithCitations(this.handle, labelId, JSON.stringify(citation))
      : this.wasm.FindByLabel(this.handle, labelId);
    return JSON.parse(json) as Record<string, AnchorTargetRef[]>;
  }

  /**
   * Resolves any bookmark in the main document part (Docxodus-managed or
   * user-authored) to the block-level anchors covering its range, in document
   * order. Empty when the bookmark name is unknown. Use this for raw bookmark
   * names that didn't come from the annotation system.
   */
  findByBookmark(bookmarkName: string, citation?: PageCitationRequest): AnchorTargetRef[] {
    const json = citation
      ? this.wasm.FindByBookmarkWithCitations(this.handle, bookmarkName, JSON.stringify(citation))
      : this.wasm.FindByBookmark(this.handle, bookmarkName);
    return JSON.parse(json) as AnchorTargetRef[];
  }

  // ─── Text/kind-based anchor discovery (#171) ─────────────────────────

  /**
   * True when `anchorId` resolves to a live element in the current session.
   * Cheap existence probe — use it to guard an anchor obtained from an earlier
   * projection before handing it to a mutation (anchors can be invalidated by
   * intervening edits; see the anchor lifecycle table in the mutation docs).
   */
  exists(anchorId: string): boolean {
    return this.wasm.Exists(this.handle, anchorId);
  }

  /**
   * Find the first block-level anchor (in document order) whose flat text
   * contains `needle`, or `null` when nothing matches. `options` tune case /
   * whitespace handling and narrow the search by kind or scope. For all
   * matches use {@link findAllByText}.
   */
  findByText(needle: string, options?: FindOptions): AnchorTargetRef | null {
    return JSON.parse(
      this.wasm.FindByText(this.handle, needle, options ? JSON.stringify(options) : ""),
    ) as AnchorTargetRef | null;
  }

  /**
   * Like {@link findByText} but returns every matching anchor in document
   * order (empty when nothing matches).
   */
  findAllByText(needle: string, options?: FindOptions): AnchorTargetRef[] {
    return JSON.parse(
      this.wasm.FindAllByText(this.handle, needle, options ? JSON.stringify(options) : ""),
    ) as AnchorTargetRef[];
  }

  /**
   * Find every block-level anchor whose flat text matches the regular
   * expression `pattern`, in document order. `regexOptions` uses the numeric
   * layout of .NET `RegexOptions` (e.g. `1` = IgnoreCase); `options` is the
   * same shape as {@link findByText} (its `ignoreCase` composes with the regex
   * flag). Defaults to `regexOptions = 0` (none).
   */
  findByRegex(pattern: string, regexOptions = 0, options?: FindOptions): AnchorTargetRef[] {
    return JSON.parse(
      this.wasm.FindByRegex(this.handle, pattern, regexOptions, options ? JSON.stringify(options) : ""),
    ) as AnchorTargetRef[];
  }

  /**
   * Return every anchor of the given `kind` — one of `"p"`, `"h"`, `"li"`,
   * `"tbl"`, `"tr"`, `"tc"`, `"col"`, `"sdt"`, `"sec"`, `"fn"`, `"en"`,
   * `"cmt"`, `"unk"` — in document order. The token set is exact and anything
   * outside it throws; in particular the row and cell tokens are `"tr"` and
   * `"tc"`, never `"row"`/`"cell"`. `"img"` and `"drw"` parse but are reserved:
   * the projection never assigns them, so they always return empty — address
   * images through the image surface instead. Reads the projection's anchor
   * index directly — no text scan. Pass `scope` (e.g. `"body"`) to restrict to
   * a single part; omit it to span all scopes.
   */
  findByKind(kind: string, scope?: string, citation?: PageCitationRequest): AnchorTargetRef[] {
    const json = citation
      ? this.wasm.FindByKindWithCitations(
          this.handle, kind, scope ?? "", JSON.stringify(citation),
        )
      : this.wasm.FindByKind(this.handle, kind, scope ?? "");
    return JSON.parse(json) as AnchorTargetRef[];
  }

  /**
   * Look up a single anchor's preview info — `{ id, kind, scope, textPreview }`.
   * Returns null when the anchor id is unknown.
   *
   * For iterating many anchors at once, prefer reading `textPreview` directly
   * off the {@link MarkdownProjection.anchorIndex} entries (cheaper — no extra
   * WASM round trip), or use {@link getAnchorInfos} for batched lookups.
   */
  getAnchorInfo(anchorId: string): AnchorInfo | null {
    const raw = this.wasm.GetAnchorInfo(this.handle, anchorId);
    return JSON.parse(raw) as AnchorInfo | null;
  }

  /**
   * Bulk variant of {@link getAnchorInfo}: takes an array of anchor ids,
   * returns a record where each unknown id maps to `null`.
   */
  getAnchorInfos(anchorIds: readonly string[]): Record<string, AnchorInfo | null> {
    const raw = this.wasm.GetAnchorInfos(this.handle, JSON.stringify(anchorIds));
    return JSON.parse(raw) as Record<string, AnchorInfo | null>;
  }

  /**
   * Resolve block-level metadata (style id+name, outline level, list membership,
   * formatting probe) for an anchor. Returns null when the anchor doesn't exist.
   */
  getBlockMetadata(anchorId: string): BlockMetadata | null {
    const raw = this.wasm.GetBlockMetadata(this.handle, anchorId);
    return JSON.parse(raw) as BlockMetadata | null;
  }

  /**
   * Bulk variant of {@link getBlockMetadata}. Unknown ids map to null;
   * duplicates are deduped.
   */
  getBlockMetadatas(anchorIds: readonly string[]): Record<string, BlockMetadata | null> {
    const raw = this.wasm.GetBlockMetadatas(this.handle, JSON.stringify(anchorIds));
    return JSON.parse(raw) as Record<string, BlockMetadata | null>;
  }

  /**
   * Resolve the numbering facts for a list-item paragraph; returns null when
   * the anchor has no w:numPr.
   */
  getListMembership(anchorId: string): ListMembership | null {
    const raw = this.wasm.GetListMembership(this.handle, anchorId);
    return JSON.parse(raw) as ListMembership | null;
  }

  /**
   * Resolve page-layout info for the w:sectPr that governs an anchor.
   * Returns null for anchors outside the body part.
   */
  getSectionInfo(anchorId: string): SectionInfo | null {
    const raw = this.wasm.GetSectionInfo(this.handle, anchorId);
    return JSON.parse(raw) as SectionInfo | null;
  }

  /** Enumerate the document's explicit style catalog with resolved high-signal properties. */
  listStyles(): StyleInfo[] {
    return JSON.parse(this.wasm.ListStyles(this.handle)) as StyleInfo[];
  }

  /** Inspect direct and effective paragraph/run formatting for one paragraph anchor. */
  getFormatting(anchorId: string): FormattingInspection | null {
    return JSON.parse(this.wasm.GetFormatting(this.handle, anchorId)) as FormattingInspection | null;
  }

  /** Enumerate text-bearing runs as mutation-compatible anchor/span pairs. */
  listInlineSpans(anchorId: string): InlineSpan[] {
    return JSON.parse(this.wasm.ListInlineSpans(this.handle, anchorId)) as InlineSpan[];
  }

  /**
   * Enumerates every annotation persisted in the document. Lets an agent prime
   * itself with "here are the labeled regions you can target" before committing
   * to a specific id.
   */
  listAnnotations(): DocumentAnnotation[] {
    return JSON.parse(this.wasm.ListAnnotations(this.handle)) as DocumentAnnotation[];
  }

  // ─── Annotation write surface ────────────────────────────────────────

  /**
   * Annotate a range inside `anchorId`. When `span` is `null`/`undefined`
   * the annotation wraps every inline run of the block. When
   * `annotation.id` is `undefined`, a 16-char hex id is auto-generated and
   * returned in `EditResult.annotationId`.
   */
  addAnnotation(
    anchorId: string,
    span: CharSpan | null,
    annotation: DocumentAnnotation,
  ): EditResult {
    const spanJson = span ? JSON.stringify(span) : "";
    return JSON.parse(
      this.wasm.AddAnnotation(this.handle, anchorId, spanJson, JSON.stringify(annotation)),
    ) as EditResult;
  }

  removeAnnotation(annotationId: string): EditResult {
    return JSON.parse(this.wasm.SessionRemoveAnnotation(this.handle, annotationId)) as EditResult;
  }

  updateAnnotation(annotationId: string, update: AnnotationUpdate): EditResult {
    return JSON.parse(
      this.wasm.UpdateAnnotation(this.handle, annotationId, JSON.stringify(update)),
    ) as EditResult;
  }

  moveAnnotation(
    annotationId: string,
    newAnchorId: string,
    newSpan: CharSpan | null,
  ): EditResult {
    const spanJson = newSpan ? JSON.stringify(newSpan) : "";
    return JSON.parse(
      this.wasm.MoveAnnotation(this.handle, annotationId, newAnchorId, spanJson),
    ) as EditResult;
  }

  // ─── Lifecycle ───────────────────────────────────────────────────────

  /**
   * Switch how subsequent mutations are recorded (issue #304). Session configuration,
   * not a document mutation: not undoable, and already-applied markup is never touched.
   */
  setTrackedChanges(mode: TrackedChangeMode): void {
    this.wasm.SetTrackedChanges(this.handle, mode);
  }

  /** Author stamped on subsequent tracked-change markup; `null` restores the "docxodus" default. */
  setRevisionAuthor(author: string | null): void {
    this.wasm.SetRevisionAuthor(this.handle, author ?? "");
  }

  undo(): boolean {
    return this.wasm.Undo(this.handle);
  }

  redo(): boolean {
    return this.wasm.Redo(this.handle);
  }

  save(): Uint8Array {
    return this.wasm.Save(this.handle);
  }

  close(): void {
    this.wasm.CloseSession(this.handle);
  }

  // TypeScript 5.2+ disposable protocol
  [Symbol.dispose]?(): void {
    this.close();
  }
}

function imageBytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  }
  return globalThis.btoa(binary);
}

/**
 * Opens a new {@link DocxSession} over the supplied DOCX bytes.
 * The returned session holds its document in WASM memory until you call
 * {@link DocxSession.close} (or it is disposed).
 */
export function openDocxSession(
  bytes: Uint8Array,
  wasmExports: DocxodusWasmExports,
  settings?: DocxSessionSettings,
): DocxSession {
  const bridge = wasmExports.DocxSessionBridge;
  const handle = bridge.OpenSession(bytes, settings ? JSON.stringify(settings) : "");
  return new DocxSession(handle, bridge);
}

export type { AnchorInfo, AnchorRef, AnchorTargetRef, BlockSlice, CharSpan, CommentListEntry, CrossBlockMatch, DocumentAnnotation, DocxSessionProjection, DocxSessionSettings, EditError, EditErrorCode, EditResult, FindOptions, FormatOp, FormattingInspection, GrepOptions, InlineSpan, MarkdownPatch, MutationBatchChangeSet, MutationBatchFailure, MutationBatchMode, MutationBatchPreviewOptions, MutationBatchPreviewStep, MutationBatchResult, MutationBatchStep, MutationBatchStepResult, MutationPreconditions, PageCitation, PageCitationRequest, PageMapRegistrationResult, PageMapStatus, ParagraphFormatting, PlaceholderKind, PreconditionFailure, PreconditionTarget, ReplaceOptions, RunFormatting, RunFormattingInfo, RunFragment, StyleInfo, TableStyleFormatting, TemplatePlaceholder, TextMatch, TextRangePrecondition } from "./types.js";
export { ContextBoundary, PlaceholderKinds } from "./types.js";
