#nullable enable

using System.Text.RegularExpressions;

namespace Docxodus.Internal;

/// <summary>
/// Per-operation facade that combines <see cref="SessionRegistry"/> lookup,
/// the corresponding <see cref="DocxSession"/> call, and JSON serialization
/// via <see cref="DocxSessionJson"/>. Every transport — the WASM JSExport
/// bridge and the stdio NDJSON host — funnels into these methods, so the
/// wire format and per-op semantics live in exactly one place.
/// </summary>
internal static class DocxSessionOps
{
    private static MutationPreconditions? ForTarget(MutationPreconditions? preconditions, string? anchorId) =>
        preconditions is not null && preconditions.AnchorId is null && anchorId is not null
            ? preconditions with { AnchorId = anchorId }
            : preconditions;

    private static string Mutate(
        int handle,
        MutationPreconditions? preconditions,
        string? targetAnchorId,
        System.Func<DocxSession, EditResult> mutation)
    {
        var session = SessionRegistry.Get(handle);
        return DocxSessionJson.Serialize(
            session.ExecuteMutation(ForTarget(preconditions, targetAnchorId), mutation));
    }

    // ─── Lifecycle ──────────────────────────────────────────────────────

    public static int OpenSession(byte[] bytes, DocxSessionSettings? settings) =>
        SessionRegistry.OpenSession(bytes, settings);

    /// <summary>Mint a complete blank DOCX (a "New document" seed) as bytes.</summary>
    public static byte[] CreateBlankDocx() => DocxSession.CreateBlankDocxBytes();

    public static void CloseSession(int handle) => SessionRegistry.CloseSession(handle);

    public static byte[] Save(int handle) => SessionRegistry.Get(handle).Save();

    /// <summary>Save with an explicit per-call override of the session's open-time
    /// <see cref="DocxSessionSettings.PersistAnchorIds"/> — see <see cref="DocxSession.Save(bool)"/>.
    /// <c>true</c> keeps the projector's Unid bookkeeping in the written bytes (a later open over
    /// them resolves the same anchor ids); <c>false</c> strips it even from a session opened with
    /// <c>PersistAnchorIds = true</c> (a clean deliverable save from an anchor-stable session).</summary>
    public static byte[] Save(int handle, bool persistAnchorIds) =>
        SessionRegistry.Get(handle).Save(persistAnchorIds);

    /// <summary>
    /// Save KEEPING the projector's Unid bookkeeping — see <see cref="DocxSession.Save(bool)"/>.
    /// Exists for the in-browser editor's remount, which re-renders these bytes and needs the
    /// anchors to survive the hop. The output is ~6x larger than the document and is not meant to
    /// be handed to a user; a save-to-disk wants <see cref="Save(int)"/>.
    /// </summary>
    public static byte[] SaveWithAnchorIds(int handle) =>
        SessionRegistry.Get(handle).Save(persistAnchorIds: true);

    public static long GetVersion(int handle) => SessionRegistry.Get(handle).Version;

    public static string GetVersionJson(int handle) =>
        DocxSessionJson.SerializeVersion(GetVersion(handle));

    public static string RegisterPageMap(
        int handle, PageMap pageMap, string? expectedRendererFingerprint = null) =>
        DocxSessionJson.SerializePageMapRegistration(
            SessionRegistry.Get(handle).RegisterPageMap(pageMap, expectedRendererFingerprint));

    public static string GetPageMapStatus(int handle, PageCitationRequest? request = null) =>
        DocxSessionJson.SerializePageMapStatus(SessionRegistry.Get(handle).GetPageMapStatus(request));

    public static string GetPageCitation(int handle, string anchorId, PageCitationRequest request) =>
        DocxSessionJson.SerializePageCitation(
            SessionRegistry.Get(handle).GetPageCitation(anchorId, request));

    /// <summary>Read-only optimistic guard evaluation for dry runs and transport diagnostics.</summary>
    public static string CheckPreconditions(int handle, MutationPreconditions? preconditions)
    {
        var error = SessionRegistry.Get(handle).EvaluatePreconditions(preconditions);
        return DocxSessionJson.Serialize(error is null
            ? new EditResult { Success = true }
            : new EditResult { Success = false, Error = error });
    }

    public static string ExecuteBatch(
        int handle,
        MutationBatchMode mode,
        System.Collections.Generic.IEnumerable<MutationBatchStep> steps) =>
        DocxSessionJson.SerializeMutationBatchResult(
            SessionRegistry.Get(handle).ExecuteBatch(steps, mode));

    /// <summary>
    /// Execute a serialized/handle-addressed batch against a complete isolated clone. The step
    /// factory receives only the temporary shadow handle, which makes accidentally targeting the
    /// live handle impossible at this central transport seam. The shadow is disposed on every
    /// return/throw path; process abandonment is also safe because the live package was never a
    /// mutation target.
    /// </summary>
    public static string PreviewBatch(
        int liveHandle,
        MutationBatchMode mode,
        System.Func<int, System.Collections.Generic.IEnumerable<MutationBatchStep>> shadowSteps,
        MutationBatchPreviewOptions? options = null)
    {
        DocxSession.ValidatePreviewOptions(options);
        var shadowHandle = SessionRegistry.CloneSessionForPreview(liveHandle);
        try
        {
            var shadow = SessionRegistry.Get(shadowHandle);
            return DocxSessionJson.SerializeMutationBatchResult(
                shadow.FinalizePreviewResult(
                    shadow.ExecuteBatch(shadowSteps(shadowHandle), mode),
                    options));
        }
        finally
        {
            SessionRegistry.CloseSession(shadowHandle);
        }
    }

    public static string GetPackageContentHash(int handle) =>
        SessionRegistry.Get(handle).GetPackageContentHash();

    /// <summary>
    /// Return the stable semantic changes between the package opened for this session and its current
    /// logical state. Requires baseline capture at open time (enabled by default).
    /// </summary>
    public static string GetSemanticChanges(int handle) =>
        SessionRegistry.Get(handle).GetSemanticChangesJson();

    /// <summary>Run the canonical deliverable-verification policy on the current logical package.</summary>
    public static string VerifyDeliverable(int handle) =>
        SessionRegistry.Get(handle).VerifyDeliverableJson();

    /// <summary>
    /// Render a preview shadow to the SAME complete-document profile
    /// <see cref="DocxSession.PreviewBatch"/> uses (<see cref="HtmlConversionOps.PreviewDocumentOptions"/>).
    /// Exists so the callback-shaped npm preview — which drives its shadow from JS and therefore
    /// cannot reuse the typed core's render call — does not have to restate the profile and drift
    /// from it. Use this, never <see cref="RenderHtml"/>, for preview HTML: RenderHtml is the
    /// EDITOR's authoring view (comments and annotations off, headers/footers tied to pagination),
    /// which answers a different question.
    /// </summary>
    public static string RenderPreviewHtml(int handle) =>
        HtmlConversionOps.ConvertToHtml(
            SessionRegistry.Get(handle), HtmlConversionOps.PreviewDocumentOptions());

    /// <summary>Scoped counterpart of <see cref="RenderPreviewHtml"/>; see
    /// <see cref="HtmlConversionOps.PreviewBlockOptions"/>.</summary>
    public static string RenderPreviewBlockHtml(int handle, string anchorId) =>
        HtmlConversionOps.RenderBlockHtml(
            SessionRegistry.Get(handle), anchorId, HtmlConversionOps.PreviewBlockOptions());

    public static DocxSessionTransaction BeginTransaction(int handle) =>
        SessionRegistry.Get(handle).BeginTransaction();

    internal static MutationBatchStep SerializedBatchStep(
        string tool,
        string action,
        System.Func<string> mutation,
        System.Func<EditError?>? preflight = null) => new(
            tool,
            action,
            _ => DocxSessionJson.DeserializeEditResults(mutation()),
            preflight is null ? null : _ => preflight());

    // ─── Projection + discovery ─────────────────────────────────────────

    public static string Project(int handle) =>
        DocxSessionJson.SerializeProjection(SessionRegistry.Get(handle).Project());

    /// <summary>
    /// Ordered top-level render units per scope container — what an incremental
    /// renderer diffs its DOM against after a structural op. See <see cref="RenderPlan"/>.
    /// </summary>
    public static string ListBlocks(int handle) =>
        DocxSessionJson.SerializeRenderPlan(SessionRegistry.Get(handle).ListBlocks());

    public static string ListRenderedBlocks(int handle, bool renderTrackedChanges) =>
        DocxSessionJson.SerializeRenderPlan(
            SessionRegistry.Get(handle).ListBlocks(renderTrackedChanges));

    /// <summary>Citation-ordered footnotes/endnotes — the id↔ordinal authority a client
    /// renumbering rendered note chrome walks. See <see cref="NoteListEntry"/>.</summary>
    public static string ListNotes(int handle, bool endnotes) =>
        DocxSessionJson.SerializeNoteList(SessionRegistry.Get(handle).ListNotes(endnotes));

    /// <summary>
    /// The anchor index alone (no markdown emission or marshaling) — the editor's
    /// per-op anchor-map refresh. Same <c>{"anchorIndex":{…}}</c> shape as
    /// <see cref="Project"/>, an order of magnitude cheaper on a large document.
    /// </summary>
    public static string ListAnchors(int handle) =>
        DocxSessionJson.SerializeAnchorIndex(SessionRegistry.Get(handle).AnchorIndex());

    public static string ProjectAnchor(int handle, string anchorId, ProjectionDepth depth,
        PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeProjection(
            SessionRegistry.Get(handle).ProjectAnchor(anchorId, depth, citationRequest));

    /// <summary>
    /// Render a single block from the live session to faithful HTML — the editor's
    /// incremental per-block re-render. Resolves against the live document (no Save /
    /// byte round-trip). Returns the block's HTML element (no html/head wrapper).
    /// </summary>
    public static string RenderBlockHtml(int handle, string anchorId, string cssPrefix,
        bool fabricateClasses, bool renderTrackedChanges = false) =>
        HtmlConversionOps.RenderBlockHtml(SessionRegistry.Get(handle), anchorId,
            EditorBlockRenderOptions(cssPrefix, fabricateClasses, renderTrackedChanges));

    /// <summary>
    /// Batch single-block render: N anchors, one throwaway document, one converter run.
    /// Returns a JSON object mapping each anchor id to its HTML (null for an anchor that
    /// failed to resolve). See <see cref="HtmlConversionOps.RenderBlocksHtml(DocxSession, System.Collections.Generic.IReadOnlyList{string}, HtmlConversionOptions)"/>.
    /// </summary>
    public static string RenderBlocksHtml(int handle, string anchorIdsJson, string cssPrefix,
        bool fabricateClasses, bool renderTrackedChanges = false) =>
        HtmlConversionOps.RenderBlocksHtml(handle, anchorIdsJson,
            EditorBlockRenderOptions(cssPrefix, fabricateClasses, renderTrackedChanges));

    /// <summary>
    /// The block-render option profile for the editor's incremental swaps. Must agree
    /// with <see cref="RenderHtml"/>'s full-render profile wherever a setting affects
    /// WITHIN-BLOCK output — footnote/endnote citation markers in particular: with
    /// RenderFootnotesAndEndnotes off, a re-rendered citing paragraph silently loses
    /// its citation marker from the DOM.
    /// </summary>
    private static HtmlConversionOptions EditorBlockRenderOptions(
        string cssPrefix, bool fabricateClasses, bool renderTrackedChanges) =>
        new HtmlConversionOptions
        {
            CssClassPrefix = cssPrefix ?? "docx-",
            FabricateCssClasses = fabricateClasses,
            RenderFootnotesAndEndnotes = true,
            RenderTrackedChanges = renderTrackedChanges,
        };

    /// <summary>
    /// Render the live session's current state to a complete anchor-stamped HTML document —
    /// the editor's full re-render (remount) without round-tripping the saved bytes through
    /// the transport. The option profile matches the editor's ConvertDocxToHtmlComplete call
    /// (comments/footnotes/annotations off, headers-and-footers tied to pagination), so a
    /// remount through this path renders byte-identically to the bytes path.
    /// </summary>
    public static string RenderHtml(int handle, string cssPrefix, bool fabricateClasses,
        bool paginated, double scale, bool renderTrackedChanges = false) =>
        HtmlConversionOps.ConvertToHtml(SessionRegistry.Get(handle), new HtmlConversionOptions
        {
            CssClassPrefix = cssPrefix ?? "docx-",
            FabricateCssClasses = fabricateClasses,
            PaginationMode = paginated ? 1 : 0,
            PaginationScale = scale,
            RenderHeadersAndFooters = paginated,
            // Footnotes/endnotes are document CONTENT, not an editing affordance: a document that
            // has them must show them, and each note renders exactly once so its paragraphs are
            // uniquely addressable (AssignAnchorUnids stamps the note parts). Must stay in step
            // with the editor's first-paint profile in npm/src/editor.ts `completeArgs`, which the
            // remount output is required to match byte-for-byte.
            RenderFootnotesAndEndnotes = true,
            RenderTrackedChanges = renderTrackedChanges,
            StampAnchors = true,
        });

    public static string Grep(int handle, string pattern, RegexOptions regexOpts,
        ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary,
        PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeMatches(
            SessionRegistry.Get(handle).Grep(
                pattern, regexOpts, scope, contextChars, whitespace, boundary, citationRequest));

    public static string GrepCrossBlock(int handle, string pattern, RegexOptions regexOpts,
        ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary,
        PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeCrossBlockMatches(
            SessionRegistry.Get(handle).GrepCrossBlock(
                pattern, regexOpts, scope, contextChars, whitespace, boundary, citationRequest));

    public static string FindPlaceholders(int handle, PlaceholderKinds kinds, ProjectionScopes scope,
        int contextChars, ContextBoundary boundary, PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializePlaceholders(
            SessionRegistry.Get(handle).FindPlaceholders(
                kinds, scope, contextChars, boundary, citationRequest));

    public static string FindByAnnotation(
        int handle, string annotationId, PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeAnchorTargets(
            SessionRegistry.Get(handle).FindByAnnotation(annotationId, citationRequest));

    public static string FindByLabel(
        int handle, string labelId, PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeAnchorTargetMap(
            SessionRegistry.Get(handle).FindByLabel(labelId, citationRequest));

    public static string FindByBookmark(
        int handle, string bookmarkName, PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeAnchorTargets(
            SessionRegistry.Get(handle).FindByBookmark(bookmarkName, citationRequest));

    public static string ListAnnotations(int handle) =>
        DocxSessionJson.SerializeAnnotations(SessionRegistry.Get(handle).ListAnnotations());

    public static bool Exists(int handle, string anchorId) =>
        SessionRegistry.Get(handle).Exists(anchorId);

    public static string GetAnchorInfo(int handle, string anchorId) =>
        DocxSessionJson.SerializeAnchorInfoOrNull(SessionRegistry.Get(handle).GetAnchorInfo(anchorId));

    public static string GetAnchorInfos(int handle, System.Collections.Generic.IEnumerable<string> anchorIds) =>
        DocxSessionJson.SerializeAnchorInfoMap(SessionRegistry.Get(handle).GetAnchorInfos(anchorIds));

    public static string GetBlockMetadata(int handle, string anchorId) =>
        DocxSessionJson.SerializeBlockMetadataOrNull(SessionRegistry.Get(handle).GetBlockMetadata(anchorId));

    public static string GetBlockMetadatas(int handle, System.Collections.Generic.IEnumerable<string> anchorIds) =>
        DocxSessionJson.SerializeBlockMetadataMap(SessionRegistry.Get(handle).GetBlockMetadatas(anchorIds));

    public static string GetListMembership(int handle, string anchorId) =>
        DocxSessionJson.SerializeListMembershipOrNull(SessionRegistry.Get(handle).GetListMembership(anchorId));

    public static string GetSectionInfo(int handle, string anchorId) =>
        DocxSessionJson.SerializeSectionInfoOrNull(SessionRegistry.Get(handle).GetSectionInfo(anchorId));

    public static string ListStyles(int handle) =>
        DocxSessionJson.SerializeStyles(SessionRegistry.Get(handle).ListStyles());

    public static string GetFormatting(int handle, string anchorId) =>
        DocxSessionJson.SerializeFormattingInspectionOrNull(SessionRegistry.Get(handle).GetFormatting(anchorId));

    public static string ListInlineSpans(int handle, string anchorId) =>
        DocxSessionJson.SerializeInlineSpans(SessionRegistry.Get(handle).ListInlineSpans(anchorId));

    public static string FindByText(int handle, string needle, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargetOrNull(SessionRegistry.Get(handle).FindByText(needle, options));

    public static string FindAllByText(int handle, string needle, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindAllByText(needle, options));

    public static string FindByRegex(int handle, string pattern, RegexOptions regexOptions, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindByRegex(pattern, regexOptions, options));

    public static string FindByKind(
        int handle, string kind, string? scope, PageCitationRequest? citationRequest = null) =>
        DocxSessionJson.SerializeAnchorTargets(
            SessionRegistry.Get(handle).FindByKind(kind, scope, citationRequest));

    public static string GetEditSummary(int handle) =>
        DocxSessionJson.SerializeEditSummary(SessionRegistry.Get(handle).GetEditSummary());

    public static string RemainingPlaceholders(int handle, PlaceholderKinds kinds) =>
        DocxSessionJson.SerializePlaceholders(SessionRegistry.Get(handle).RemainingPlaceholders(kinds));

    public static string GetDiff(int handle, DiffFormat format) =>
        SessionRegistry.Get(handle).GetDiff(format);

    // ─── Tier A: text mutations ─────────────────────────────────────────

    public static string ReplaceText(int handle, string anchorId, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.ReplaceText(anchorId, markdown));

    public static string DeleteBlock(int handle, string anchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.DeleteBlock(anchorId));

    public static string DeleteRange(int handle, string fromAnchorId, string toAnchorIdExclusive,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, fromAnchorId,
            s => s.DeleteRange(fromAnchorId, toAnchorIdExclusive));

    public static string DeleteSection(int handle, string headingAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, headingAnchorId, s => s.DeleteSection(headingAnchorId));

    public static string ReplaceTextRange(int handle, string anchorId, string find, string replace,
        ReplaceOptions? options, MutationPreconditions? preconditions = null)
    {
        if (preconditions is not null)
            options = (options ?? new ReplaceOptions()) with
            {
                Preconditions = ForTarget(preconditions, anchorId),
            };
        return DocxSessionJson.SerializeEditResults(
            SessionRegistry.Get(handle).ReplaceTextRange(anchorId, find, replace, options));
    }

    public static string ReplaceTextAtSpan(int handle, string anchorId, int spanStart, int spanLength,
        string replace, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.ReplaceTextAtSpan(anchorId, spanStart, spanLength, replace));

    /// <summary>
    /// Bracket-aware variant of <see cref="ReplaceTextAtSpan"/>. Parses the brackets out
    /// of <paramref name="matchText"/> and substitutes <paramref name="newInner"/> for the
    /// bracketed portion, preserving any prefix/suffix outside the brackets (so a match
    /// like <c>$[___]</c> + <c>"0.20"</c> produces <c>$0.20</c>, not <c>0.20</c>).
    /// Returns a <c>MalformedMarkdown</c> EditResult if the match has no balanced
    /// brackets. Mirrors <see cref="DocxSession.ReplaceInner(TextMatch, string)"/>;
    /// transport-side because reconstructing a <see cref="TextMatch"/> from wire fields
    /// (Fragments, ContextBefore, …) would be wasteful.
    /// </summary>
    public static string ReplaceInner(int handle, string matchText, string anchorId,
        int spanStart, int spanLength, string newInner, MutationPreconditions? preconditions = null)
    {
        return Mutate(handle, preconditions, anchorId, s =>
        {
            int lb = matchText.IndexOf('[');
            int rb = matchText.LastIndexOf(']');
            if (lb < 0 || rb <= lb)
                return new EditResult
                {
                    Success = false,
                    Error = new EditError(EditErrorCode.MalformedMarkdown,
                        $"match text has no balanced brackets: '{matchText}'", anchorId),
                };
            var prefix = matchText[..lb];
            var suffix = matchText[(rb + 1)..];
            return s.ReplaceTextAtSpan(anchorId, spanStart, spanLength, prefix + newInner + suffix);
        });
    }

    // ─── Tier B: structural ─────────────────────────────────────────────

    public static string MoveBlock(int handle, string sourceAnchorId, string targetAnchorId,
        Position position, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, sourceAnchorId,
            s => s.MoveBlock(sourceAnchorId, targetAnchorId, position));

    /// <summary>The blocks <paramref name="sourceAnchorId"/> may legally move next to, and on which
    /// side — what a drag UI gates its drop targets on, so it never offers a drop the engine
    /// will refuse.</summary>
    public static string ValidMoveTargets(int handle, string sourceAnchorId) =>
        DocxSessionJson.SerializeMoveTargets(
            SessionRegistry.Get(handle).ValidMoveTargets(sourceAnchorId));

    public static string InsertParagraph(int handle, string anchorId, Position position, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertParagraph(anchorId, position, markdown));

    public static string SplitParagraph(int handle, string anchorId, int characterOffset,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.SplitParagraph(anchorId, characterOffset));

    public static string MergeParagraphs(int handle, string firstAnchorId, string secondAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, firstAnchorId,
            s => s.MergeParagraphs(firstAnchorId, secondAnchorId));

    public static string InsertHorizontalRule(int handle, string anchorId, Position position, string ruleJson,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.InsertHorizontalRule(
            anchorId, position, string.IsNullOrEmpty(ruleJson) ? null : ParseRuleEdge(ruleJson)));

    private static ParagraphBorderEdge? ParseRuleEdge(string json)
    {
        // The rule JSON is itself a border-edge object; reuse the named-property parser by
        // wrapping it under a known key.
        using var d = System.Text.Json.JsonDocument.Parse($"{{\"e\":{json}}}");
        return DocxSessionJson.ParseBorderEdge(d.RootElement, "e");
    }

    // ─── Headers / footers / page numbers ───────────────────────────────

    public static string SetHeaderText(int handle, string anchorId, HeaderFooterKind kind, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetHeaderText(anchorId, kind, markdown));

    public static string SetFooterText(int handle, string anchorId, HeaderFooterKind kind, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetFooterText(anchorId, kind, markdown));

    public static string InsertPageNumberField(
        int handle, string anchorId, PageNumberField field, NumberFormat? format = null,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertPageNumberField(anchorId, field, format));

    public static string EnsureHeaderFooterVisible(int handle, string anchorId, HeaderFooterKind kind,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.EnsureHeaderFooterVisible(anchorId, kind));

    public static string SetPageNumbering(int handle, string anchorId, PageNumberingOp op,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetPageNumbering(anchorId, op));

    public static string ClearPageNumbering(int handle, string anchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.ClearPageNumbering(anchorId));

    // ─── Reference fields (issue #607) ──────────────────────────────────

    public static string InsertTableOfContents(
        int handle, string anchorId, Position pos, TableOfContentsOptions? options = null,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertTableOfContents(anchorId, pos, options));

    public static string InsertTableOfFigures(
        int handle, string anchorId, Position pos, TableOfFiguresOptions? options = null,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertTableOfFigures(anchorId, pos, options));

    public static string InsertTableOfAuthorities(
        int handle, string anchorId, Position pos, TableOfAuthoritiesOptions? options = null,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertTableOfAuthorities(anchorId, pos, options));

    // ─── Footnotes / endnotes ───────────────────────────────────────────

    public static string InsertFootnote(int handle, string anchorId, int characterOffset, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertFootnote(anchorId, characterOffset, markdown));

    public static string InsertEndnote(int handle, string anchorId, int characterOffset, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertEndnote(anchorId, characterOffset, markdown));

    // ─── Comments (issue #300) ──────────────────────────────────────────

    public static string AddComment(int handle, string anchorId, CharSpan? span, string author,
        string? initials, string? dateIso, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.AddComment(
            anchorId, span, author, markdown, initials, DocxSessionJson.ParseCommentDate(dateIso)));

    public static string AddCommentToRevision(int handle, string revisionId, string author,
        string? initials, string? dateIso, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.AddCommentToRevision(
            revisionId, author, markdown, initials, DocxSessionJson.ParseCommentDate(dateIso)));

    public static string AddCommentReply(int handle, string parentCommentAnchorId, string author,
        string? initials, string? dateIso, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, parentCommentAnchorId, s => s.AddCommentReply(
            parentCommentAnchorId, author, markdown, initials, DocxSessionJson.ParseCommentDate(dateIso)));

    public static string UpdateComment(int handle, string commentAnchorId, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, commentAnchorId, s => s.UpdateComment(commentAnchorId, markdown));

    public static string SetCommentResolved(int handle, string commentAnchorId, bool resolved,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, commentAnchorId,
            s => s.SetCommentResolved(commentAnchorId, resolved));

    public static string RemoveComment(int handle, string commentAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, commentAnchorId, s => s.RemoveComment(commentAnchorId));

    public static string ListComments(int handle) =>
        DocxSessionJson.SerializeCommentList(SessionRegistry.Get(handle).ListComments());

    // ─── Hyperlinks / bookmarks (issue #451) ───────────────────────────

    public static string ListHyperlinks(int handle, ProjectionScopes scopes = ProjectionScopes.All) =>
        DocxSessionJson.SerializeHyperlinks(SessionRegistry.Get(handle).ListHyperlinks(scopes));

    /// <summary>Insert a REF-field internal cross-reference (issue #545); see
    /// <see cref="DocxSession.InsertCrossReference"/>. Options arrive typed — each transport
    /// parses its own <c>{referenceNumber, hyperlink, includePosition}</c> wire object.</summary>
    public static string InsertCrossReference(int handle, string anchorId, int characterOffset,
        string bookmarkName, CrossReferenceOptions? options = null,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.InsertCrossReference(anchorId, characterOffset, bookmarkName, options));

    public static string AddHyperlink(int handle, string anchorId, int start, int length,
        string kind, string target)
    {
        var session = SessionRegistry.Get(handle);
        if (!TryParseHyperlinkTarget(kind, target, out var parsed))
            return InvalidHyperlinkKind(kind, anchorId);
        return DocxSessionJson.Serialize(session.AddHyperlink(anchorId,
            new CharSpan(start, length), parsed));
    }

    public static string UpdateHyperlink(int handle, string hyperlinkId, string kind, string target)
    {
        var session = SessionRegistry.Get(handle);
        if (!TryParseHyperlinkTarget(kind, target, out var parsed))
            return InvalidHyperlinkKind(kind);
        return DocxSessionJson.Serialize(session.UpdateHyperlink(hyperlinkId, parsed));
    }

    public static string RemoveHyperlink(int handle, string hyperlinkId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveHyperlink(hyperlinkId));

    public static string ListBookmarks(int handle, ProjectionScopes scopes = ProjectionScopes.All) =>
        DocxSessionJson.SerializeBookmarks(SessionRegistry.Get(handle).ListBookmarks(scopes));

    public static string AddBookmark(int handle, string name, string startAnchorId, int startOffset,
        string endAnchorId, int endOffset) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).AddBookmark(name,
            new DocumentRange(startAnchorId, startOffset, endAnchorId, endOffset)));

    public static string RenameBookmark(int handle, string name, string newName) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RenameBookmark(name, newName));

    public static string MoveBookmark(int handle, string name, string startAnchorId, int startOffset,
        string endAnchorId, int endOffset) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).MoveBookmark(name,
            new DocumentRange(startAnchorId, startOffset, endAnchorId, endOffset)));

    public static string RemoveBookmark(int handle, string name) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveBookmark(name));

    private static bool TryParseHyperlinkTarget(string kind, string target,
        out HyperlinkTarget parsed)
    {
        if (string.Equals(kind, "internal", StringComparison.OrdinalIgnoreCase))
        {
            parsed = HyperlinkTarget.Internal(target);
            return true;
        }
        if (string.Equals(kind, "external", StringComparison.OrdinalIgnoreCase))
        {
            parsed = HyperlinkTarget.External(target);
            return true;
        }
        parsed = null!;
        return false;
    }

    private static string InvalidHyperlinkKind(string kind, string? anchorId = null) =>
        DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidHyperlinkTarget,
            $"unknown hyperlink target kind '{kind}'; expected 'internal' or 'external'", anchorId));

    // ─── Native images (issue #453) ────────────────────────────────────

    public static string GetImageCapabilities() =>
        DocxSessionJson.SerializeImageCapabilities(DocxSession.GetImageCapabilities());

    public static string ListImages(int handle, ProjectionScopes scopes = ProjectionScopes.All) =>
        DocxSessionJson.SerializeImages(SessionRegistry.Get(handle).ListImages(scopes));

    public static string InsertImage(int handle, string anchorId, int characterOffset,
        string imageBase64, string optionsJson)
    {
        if (!TryDecodeImageBase64(imageBase64, anchorId, out var bytes, out var error)) return error!;
        try
        {
            return DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertImage(
                anchorId, characterOffset, bytes!, DocxSessionJson.ParseImageInsertOptions(optionsJson)));
        }
        catch (System.Exception ex) when (ex is System.Text.Json.JsonException or System.ArgumentException
            or System.OverflowException)
        {
            return DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidImageLayout,
                $"invalid image options JSON: {ex.Message}", anchorId));
        }
    }

    public static string ReplaceImage(int handle, string imageId, string imageBase64)
    {
        if (!TryDecodeImageBase64(imageBase64, null, out var bytes, out var error)) return error!;
        return DocxSessionJson.Serialize(SessionRegistry.Get(handle).ReplaceImage(imageId, bytes!));
    }

    public static string SetImageDimensions(int handle, string imageId, string dimensionsJson)
    {
        try
        {
            var dimensions = DocxSessionJson.ParseImageDimensions(dimensionsJson);
            return DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetImageDimensions(imageId,
                dimensions.Width, dimensions.Height, dimensions.PreserveAspect));
        }
        catch (System.Exception ex) when (ex is System.Text.Json.JsonException or System.ArgumentException
            or System.OverflowException)
        {
            return DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                $"invalid image dimensions JSON: {ex.Message}"));
        }
    }

    public static string SetImageMetadata(int handle, string imageId, string? altText, string? title) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetImageMetadata(imageId, altText, title));

    public static string SetImageFloatingLayout(int handle, string imageId, string layoutJson)
    {
        try
        {
            return DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetImageFloatingLayout(
                imageId, DocxSessionJson.ParseFloatingImageLayout(layoutJson)));
        }
        catch (System.Exception ex) when (ex is System.Text.Json.JsonException or System.ArgumentException
            or System.OverflowException)
        {
            return DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidImageLayout,
                $"invalid floating layout JSON: {ex.Message}"));
        }
    }

    public static string RemoveImage(int handle, string imageId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveImage(imageId));

    // ─── Native content controls (issue #452) ─────────────────────────

    public static string ListContentControls(int handle,
        ProjectionScopes scopes = ProjectionScopes.All) =>
        DocxSessionJson.SerializeContentControls(
            SessionRegistry.Get(handle).ListContentControls(scopes));

    public static string FillContentControlText(int handle, string anchorId, string text,
        string optionsJson) => ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).FillContentControlText(anchorId, text, options));

    public static string FillContentControlRichText(int handle, string anchorId, string markdown,
        string optionsJson) => ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).FillContentControlRichText(anchorId, markdown, options));

    public static string SetContentControlChecked(int handle, string anchorId, bool isChecked,
        string optionsJson) => ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).SetContentControlChecked(anchorId, isChecked, options));

    public static string SetContentControlDate(int handle, string anchorId, string value,
        string? displayText, string optionsJson)
    {
        if (!System.DateTimeOffset.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.RoundtripKind, out var parsed))
            return DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidContentControlValue,
                "date value must be an ISO-8601 timestamp", anchorId));
        return ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).SetContentControlDate(anchorId, parsed, displayText, options));
    }

    public static string SelectContentControlItem(int handle, string anchorId, string value,
        string optionsJson) => ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).SelectContentControlItem(anchorId, value, options));

    public static string FillContentControlPicture(int handle, string anchorId,
        string imageBase64, string optionsJson)
    {
        if (!TryDecodeImageBase64(imageBase64, anchorId, out var bytes, out var error)) return error!;
        return ContentControlOptions(anchorId, optionsJson, options =>
            SessionRegistry.Get(handle).FillContentControlPicture(anchorId, bytes!, options));
    }

    public static string AddRepeatingSectionItem(int handle, string sectionAnchorId,
        string? afterItemAnchorId, string optionsJson) =>
        ContentControlOptions(sectionAnchorId, optionsJson, options =>
            SessionRegistry.Get(handle).AddRepeatingSectionItem(
                sectionAnchorId, afterItemAnchorId, options));

    public static string RemoveRepeatingSectionItem(int handle, string itemAnchorId) =>
        DocxSessionJson.Serialize(
            SessionRegistry.Get(handle).RemoveRepeatingSectionItem(itemAnchorId));

    private static string ContentControlOptions(string anchorId, string optionsJson,
        System.Func<ContentControlFillOptions, EditResult> action)
    {
        try
        {
            return DocxSessionJson.Serialize(action(
                DocxSessionJson.ParseContentControlFillOptions(optionsJson)));
        }
        catch (System.Exception ex) when (ex is System.Text.Json.JsonException
            or System.ArgumentException)
        {
            return DocxSessionJson.Serialize(EditResult.Fail(
                EditErrorCode.InvalidContentControlValue,
                $"invalid content-control options JSON: {ex.Message}", anchorId));
        }
    }

    private static bool TryDecodeImageBase64(string? base64, string? anchorId,
        out byte[]? bytes, out string? error)
    {
        bytes = null;
        error = null;
        if (string.IsNullOrEmpty(base64))
        {
            error = DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidImageData,
                "image base64 is empty", anchorId));
            return false;
        }
        long maxEncodedLength = ((DocxSession.MaxImageInputBytes + 2) / 3) * 4;
        if (base64.Length > maxEncodedLength)
        {
            error = DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.ImageTooLarge,
                $"encoded image exceeds the {DocxSession.MaxImageInputBytes}-byte runtime limit", anchorId));
            return false;
        }
        try { bytes = System.Convert.FromBase64String(base64); return true; }
        catch (System.FormatException)
        {
            error = DocxSessionJson.Serialize(EditResult.Fail(EditErrorCode.InvalidImageData,
                "imageBase64 is not valid base64", anchorId));
            return false;
        }
    }

    // ─── Tier C: formatting ─────────────────────────────────────────────

    public static string ApplyFormat(int handle, string anchorId, CharSpan? span, FormatOp op,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.ApplyFormat(anchorId, span, op));

    public static string ApplyFormatBySubstring(int handle, string anchorId, string substring, FormatOp op,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId,
            s => s.ApplyFormatToSubstring(anchorId, substring, op));

    public static string SetParagraphStyle(int handle, string anchorId, string styleId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetParagraphStyle(anchorId, styleId));

    public static string SetParagraphFormat(int handle, string anchorId, ParagraphFormatOp op,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetParagraphFormat(anchorId, op));

    public static string SetListLevel(int handle, string anchorId, int levelDelta,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetListLevel(anchorId, levelDelta));

    public static string RemoveListMembership(int handle, string anchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.RemoveListMembership(anchorId));

    public static string ApplyListFormat(int handle, string anchorId, ListFormat kind,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.ApplyListFormat(anchorId, kind));

    public static string ApplyListFormatRange(int handle, string firstAnchorId, string lastAnchorId, ListFormat kind,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, firstAnchorId,
            s => s.ApplyListFormatRange(firstAnchorId, lastAnchorId, kind));

    public static string SetListStartOverride(int handle, string anchorId, int value,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.SetListStartOverride(anchorId, value));

    public static string ClearListStartOverride(int handle, string anchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.ClearListStartOverride(anchorId));

    // ─── Tier D: tables ─────────────────────────────────────────────────

    public static string GetTableMetadata(int handle, string tableAnchorId) =>
        DocxSessionJson.SerializeTableMetadataResult(
            SessionRegistry.Get(handle).GetTableMetadata(tableAnchorId));

    public static string ResolveTableCellAnchor(int handle, string cellAnchorId) =>
        DocxSessionJson.SerializeTableCellResolutionResult(
            SessionRegistry.Get(handle).ResolveTableCellAnchor(cellAnchorId));

    public static string ResolveTableCellCoordinate(
        int handle, string tableAnchorId, int rowIndex, int columnIndex) =>
        DocxSessionJson.SerializeTableCellResolutionResult(
            SessionRegistry.Get(handle).ResolveTableCellCoordinate(tableAnchorId, rowIndex, columnIndex));

    public static string ReplaceCellContent(int handle, string cellAnchorId, string markdown,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId,
            s => s.ReplaceCellContent(cellAnchorId, markdown));

    public static string InsertTable(int handle, string anchorId, Position position, int rows, int cols,
        string optionsJson, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.InsertTable(
            anchorId, position, rows, cols, DocxSessionJson.ParseTableInsertOptions(optionsJson)));

    public static string InsertTableRow(int handle, string cellAnchorId, Position position,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.InsertTableRow(cellAnchorId, position));

    public static string InsertTableColumn(int handle, string cellAnchorId, Position position,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.InsertTableColumn(cellAnchorId, position));

    public static string DeleteTableRow(int handle, string cellAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.DeleteTableRow(cellAnchorId));

    public static string DeleteTableColumn(int handle, string cellAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.DeleteTableColumn(cellAnchorId));

    // ─── Cell merge / unmerge (issue #340 Stage B) ──────────────────────

    /// <summary><paramref name="content"/> is "append" (default) | "discard" | "reject" —
    /// what happens to the content of the cells the merge absorbs.</summary>
    public static string MergeCells(int handle, string cellAnchorId, int rowSpan, int colSpan,
        string? content, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.MergeCells(cellAnchorId, rowSpan, colSpan,
            new TableMergeOptions { Content = DocxSessionJson.ParseTableMergeContent(content) }));

    public static string UnmergeCells(int handle, string cellAnchorId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.UnmergeCells(cellAnchorId));

    // ─── Table styling (issue #315 Stage A) ─────────────────────────────

    /// <summary><paramref name="widthsJson"/> is a JSON array of per-column twip widths
    /// (one positive value per column, left→right).</summary>
    public static string SetColumnWidths(int handle, string cellAnchorId, string widthsJson,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId,
            s => s.SetColumnWidths(cellAnchorId, DocxSessionJson.ParseIntArray(widthsJson)));

    /// <summary><paramref name="specJson"/> is a TableBorderSpec object
    /// ({ scope?: "all"|"outside"|"inside", style?, size?, color? }); "" uses the defaults.</summary>
    public static string SetTableBorders(int handle, string cellAnchorId, string specJson,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId,
            s => s.SetTableBorders(cellAnchorId, DocxSessionJson.ParseTableBorderSpec(specJson)));

    /// <summary><paramref name="fill"/> is a hex RRGGBB triplet or "auto"; "" clears the shading.
    /// <paramref name="scope"/> is "cell" | "row".</summary>
    public static string SetCellShading(int handle, string cellAnchorId, string fill, string scope,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.SetCellShading(
            cellAnchorId, string.IsNullOrEmpty(fill) ? null : fill,
            DocxSessionJson.ParseTableShadingScope(scope)));

    public static string SetRepeatHeaderRow(int handle, string cellAnchorId, bool repeat,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId,
            s => s.SetRepeatHeaderRow(cellAnchorId, repeat));

    public static string SetTableRowOptions(int handle, string cellAnchorId, bool? repeatHeader,
        bool? allowBreakAcrossPages, int? heightTwips, string? heightRule,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, cellAnchorId, s => s.SetTableRowOptions(cellAnchorId,
            new TableRowOptions
            {
                RepeatHeader = repeatHeader,
                AllowBreakAcrossPages = allowBreakAcrossPages,
                HeightTwips = heightTwips,
                HeightRule = DocxSessionJson.ParseTableRowHeightRule(heightRule),
            }));

    // ─── Raw escape hatch ───────────────────────────────────────────────

    public static string RawGetXml(int handle, string anchorId) =>
        SessionRegistry.Get(handle).Raw.GetXml(anchorId);

    public static string RawInsertXml(int handle, string anchorId, Position position, string xml,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.Raw.InsertXml(anchorId, position, xml));

    public static string RawReplaceXml(int handle, string anchorId, string xml,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.Raw.ReplaceXml(anchorId, xml));

    // ─── Tier E: annotations ────────────────────────────────────────────

    public static string AddAnnotation(int handle, string anchorId, CharSpan? span,
        string annotationJson, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, anchorId, s => s.AddAnnotation(
            anchorId, span, DocxSessionJson.DeserializeAnnotation(annotationJson)));

    public static string RemoveAnnotation(int handle, string annotationId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.RemoveAnnotation(annotationId));

    public static string UpdateAnnotation(int handle, string annotationId, string updateJson,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.UpdateAnnotation(
            annotationId, DocxSessionJson.DeserializeAnnotationUpdate(updateJson)));

    public static string MoveAnnotation(int handle, string annotationId, string newAnchorId,
        CharSpan? newSpan, MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, newAnchorId,
            s => s.MoveAnnotation(annotationId, newAnchorId, newSpan));

    // ─── Tracked revisions (issue #318) ─────────────────────────────────

    /// <summary>Markup-native revision listing — stable ids, true authors/dates, no
    /// accept/reject re-diff. See <see cref="DocxSession.ListRevisions"/>.</summary>
    public static string ListRevisions(int handle) =>
        DocxSessionJson.SerializeRevisionList(SessionRegistry.Get(handle).ListRevisions());

    public static string AcceptRevision(int handle, string revisionId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.AcceptRevision(revisionId));

    public static string RejectRevision(int handle, string revisionId,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.RejectRevision(revisionId));

    public static string AcceptAllRevisions(int handle,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.AcceptAllRevisions());

    public static string RejectAllRevisions(int handle,
        MutationPreconditions? preconditions = null) =>
        Mutate(handle, preconditions, null, s => s.RejectAllRevisions());

    // ─── Undo / Redo ────────────────────────────────────────────────────

    public static bool Undo(int handle) => SessionRegistry.Get(handle).Undo();

    public static bool Redo(int handle) => SessionRegistry.Get(handle).Redo();

    public static string UndoChecked(int handle, MutationPreconditions? preconditions) =>
        Mutate(handle, preconditions, null, s => s.Undo()
            ? new EditResult { Success = true }
            : EditResult.Fail(EditErrorCode.NothingToUndo, "nothing to undo"));

    public static string RedoChecked(int handle, MutationPreconditions? preconditions) =>
        Mutate(handle, preconditions, null, s => s.Redo()
            ? new EditResult { Success = true }
            : EditResult.Fail(EditErrorCode.NothingToRedo, "nothing to redo"));

    // ─── Session configuration (issue #304) ─────────────────────────────

    public static void SetTrackedChanges(int handle, TrackedChangeMode mode) =>
        SessionRegistry.Get(handle).SetTrackedChanges(mode);

    public static void SetRevisionAuthor(int handle, string? author) =>
        SessionRegistry.Get(handle).SetRevisionAuthor(author);

    public static string GetTrackedChanges(int handle)
    {
        var s = SessionRegistry.Get(handle);
        return "{\"trackedChanges\":" + DocxSessionJson.JsonString(DocxSessionJson.TrackedChangeModeName(s.TrackedChanges))
            + ",\"revisionAuthor\":" + (s.RevisionAuthor is null ? "null" : DocxSessionJson.JsonString(s.RevisionAuthor))
            + "}";
    }
}
