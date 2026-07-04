// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

import type {
  AnchorInfo,
  AnchorRef,
  AnchorTargetRef,
  AnnotationUpdate,
  BlockMetadata,
  BulkEditResult,
  CharSpan,
  CrossBlockMatch,
  DiffEntry,
  DocumentAnnotation,
  DocxodusWasmExports,
  DocxSessionProjection,
  DocxSessionSettings,
  EditError,
  EditResult,
  EditSummary,
  FillOptions,
  FindOptions,
  FormatOp,
  ParagraphBorderEdge,
  ParagraphFormatOp,
  TableInsertOptions,
  ListFormat,
  GrepOptions,
  ListMembership,
  ReplaceOptions,
  SectionInfo,
  TemplatePlaceholder,
  TextMatch,
} from "./types.js";
import { ContextBoundary, DiffFormat, PlaceholderKinds, ProjectionDepth } from "./types.js";

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
  ): DocxSessionProjection {
    return JSON.parse(
      this.wasm.ProjectAnchor(this.handle, anchorId, depth),
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

  replaceText(anchorId: string, markdown: string): EditResult {
    return JSON.parse(this.wasm.ReplaceText(this.handle, anchorId, markdown)) as EditResult;
  }

  deleteBlock(anchorId: string): EditResult {
    return JSON.parse(this.wasm.DeleteBlock(this.handle, anchorId)) as EditResult;
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
   * cell markdown, and cell alignment. The returned `EditResult.created` lists the cell-paragraph
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

  /**
   * Table row/column editing, addressed by a cell-paragraph anchor (e.g. one returned from
   * {@link insertTable}'s `created`). Insert clones the reference row/column's widths and starts
   * empty (`created` lists the new cell-paragraph anchors); delete of the last row/column removes
   * the whole table. v1 assumes a rectangular grid (no horizontal cell merges).
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

    const isSpace = (c: string | undefined) => c === " " || c === "\t";
    const isClauseTerm = (c: string | undefined) => c === "." || c === "," || c === ";" || c === ":" || c === "!" || c === "?";
    const isOpen = (c: string | undefined) => c === "(" || c === "[" || c === "{";
    const isClose = (c: string | undefined) => c === ")" || c === "]" || c === "}";

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
  ): TemplatePlaceholder[] {
    return JSON.parse(
      this.wasm.FindPlaceholders(this.handle, kinds, scope, contextChars, boundary),
    ) as TemplatePlaceholder[];
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
  findByAnnotation(annotationId: string): AnchorTargetRef[] {
    return JSON.parse(this.wasm.FindByAnnotation(this.handle, annotationId)) as AnchorTargetRef[];
  }

  /**
   * Finds every annotation whose `labelId` matches and resolves each of their
   * ranges. The result is keyed by annotation id so callers can disambiguate
   * when the same label is applied to multiple regions (three "WARRANTY"
   * annotations on different paragraphs become three entries). Annotations
   * whose bookmark resolves to no anchors are omitted from the result.
   */
  findByLabel(labelId: string): Record<string, AnchorTargetRef[]> {
    return JSON.parse(this.wasm.FindByLabel(this.handle, labelId)) as Record<string, AnchorTargetRef[]>;
  }

  /**
   * Resolves any bookmark in the main document part (Docxodus-managed or
   * user-authored) to the block-level anchors covering its range, in document
   * order. Empty when the bookmark name is unknown. Use this for raw bookmark
   * names that didn't come from the annotation system.
   */
  findByBookmark(bookmarkName: string): AnchorTargetRef[] {
    return JSON.parse(this.wasm.FindByBookmark(this.handle, bookmarkName)) as AnchorTargetRef[];
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
   * Return every anchor of the given `kind` (`"p"`, `"h"`, `"li"`, `"tbl"`,
   * `"row"`, `"cell"`, …), in document order. Reads the projection's anchor
   * index directly — no text scan. Pass `scope` (e.g. `"body"`) to restrict to
   * a single part; omit it to span all scopes.
   */
  findByKind(kind: string, scope?: string): AnchorTargetRef[] {
    return JSON.parse(
      this.wasm.FindByKind(this.handle, kind, scope ?? ""),
    ) as AnchorTargetRef[];
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

export type { AnchorInfo, AnchorRef, AnchorTargetRef, BlockSlice, CharSpan, CrossBlockMatch, DocumentAnnotation, DocxSessionProjection, DocxSessionSettings, EditError, EditErrorCode, EditResult, FindOptions, FormatOp, GrepOptions, MarkdownPatch, PlaceholderKind, ReplaceOptions, RunFormatting, RunFragment, TemplatePlaceholder, TextMatch } from "./types.js";
export { ContextBoundary, PlaceholderKinds } from "./types.js";
