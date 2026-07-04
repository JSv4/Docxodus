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
    // ─── Lifecycle ──────────────────────────────────────────────────────

    public static int OpenSession(byte[] bytes, DocxSessionSettings? settings) =>
        SessionRegistry.OpenSession(bytes, settings);

    /// <summary>Mint a complete blank DOCX (a "New document" seed) as bytes.</summary>
    public static byte[] CreateBlankDocx() => DocxSession.CreateBlankDocxBytes();

    public static void CloseSession(int handle) => SessionRegistry.CloseSession(handle);

    public static byte[] Save(int handle) => SessionRegistry.Get(handle).Save();

    // ─── Projection + discovery ─────────────────────────────────────────

    public static string Project(int handle) =>
        DocxSessionJson.SerializeProjection(SessionRegistry.Get(handle).Project());

    public static string ProjectAnchor(int handle, string anchorId, ProjectionDepth depth) =>
        DocxSessionJson.SerializeProjection(SessionRegistry.Get(handle).ProjectAnchor(anchorId, depth));

    /// <summary>
    /// Render a single block from the live session to faithful HTML — the editor's
    /// incremental per-block re-render. Resolves against the live document (no Save /
    /// byte round-trip). Returns the block's HTML element (no html/head wrapper).
    /// </summary>
    public static string RenderBlockHtml(int handle, string anchorId, string cssPrefix, bool fabricateClasses) =>
        HtmlConversionOps.RenderBlockHtml(SessionRegistry.Get(handle), anchorId,
            new HtmlConversionOptions { CssClassPrefix = cssPrefix ?? "docx-", FabricateCssClasses = fabricateClasses });

    public static string Grep(int handle, string pattern, RegexOptions regexOpts,
        ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary) =>
        DocxSessionJson.SerializeMatches(
            SessionRegistry.Get(handle).Grep(pattern, regexOpts, scope, contextChars, whitespace, boundary));

    public static string GrepCrossBlock(int handle, string pattern, RegexOptions regexOpts,
        ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary) =>
        DocxSessionJson.SerializeCrossBlockMatches(
            SessionRegistry.Get(handle).GrepCrossBlock(pattern, regexOpts, scope, contextChars, whitespace, boundary));

    public static string FindPlaceholders(int handle, PlaceholderKinds kinds, ProjectionScopes scope,
        int contextChars, ContextBoundary boundary) =>
        DocxSessionJson.SerializePlaceholders(
            SessionRegistry.Get(handle).FindPlaceholders(kinds, scope, contextChars, boundary));

    public static string FindByAnnotation(int handle, string annotationId) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindByAnnotation(annotationId));

    public static string FindByLabel(int handle, string labelId) =>
        DocxSessionJson.SerializeAnchorTargetMap(SessionRegistry.Get(handle).FindByLabel(labelId));

    public static string FindByBookmark(int handle, string bookmarkName) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindByBookmark(bookmarkName));

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

    public static string FindByText(int handle, string needle, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargetOrNull(SessionRegistry.Get(handle).FindByText(needle, options));

    public static string FindAllByText(int handle, string needle, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindAllByText(needle, options));

    public static string FindByRegex(int handle, string pattern, RegexOptions regexOptions, FindOptions? options) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindByRegex(pattern, regexOptions, options));

    public static string FindByKind(int handle, string kind, string? scope) =>
        DocxSessionJson.SerializeAnchorTargets(SessionRegistry.Get(handle).FindByKind(kind, scope));

    public static string GetEditSummary(int handle) =>
        DocxSessionJson.SerializeEditSummary(SessionRegistry.Get(handle).GetEditSummary());

    public static string RemainingPlaceholders(int handle, PlaceholderKinds kinds) =>
        DocxSessionJson.SerializePlaceholders(SessionRegistry.Get(handle).RemainingPlaceholders(kinds));

    public static string GetDiff(int handle, DiffFormat format) =>
        SessionRegistry.Get(handle).GetDiff(format);

    // ─── Tier A: text mutations ─────────────────────────────────────────

    public static string ReplaceText(int handle, string anchorId, string markdown) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).ReplaceText(anchorId, markdown));

    public static string DeleteBlock(int handle, string anchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).DeleteBlock(anchorId));

    public static string DeleteRange(int handle, string fromAnchorId, string toAnchorIdExclusive) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).DeleteRange(fromAnchorId, toAnchorIdExclusive));

    public static string DeleteSection(int handle, string headingAnchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).DeleteSection(headingAnchorId));

    public static string ReplaceTextRange(int handle, string anchorId, string find, string replace,
        ReplaceOptions? options) =>
        DocxSessionJson.SerializeEditResults(
            SessionRegistry.Get(handle).ReplaceTextRange(anchorId, find, replace, options));

    public static string ReplaceTextAtSpan(int handle, string anchorId, int spanStart, int spanLength,
        string replace) =>
        DocxSessionJson.Serialize(
            SessionRegistry.Get(handle).ReplaceTextAtSpan(anchorId, spanStart, spanLength, replace));

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
        int spanStart, int spanLength, string newInner)
    {
        int lb = matchText.IndexOf('[');
        int rb = matchText.LastIndexOf(']');
        if (lb < 0 || rb <= lb)
            return DocxSessionJson.Serialize(new EditResult
            {
                Success = false,
                Error = new EditError(EditErrorCode.MalformedMarkdown,
                    $"match text has no balanced brackets: '{matchText}'", anchorId),
            });
        var prefix = matchText[..lb];
        var suffix = matchText[(rb + 1)..];
        return DocxSessionJson.Serialize(
            SessionRegistry.Get(handle).ReplaceTextAtSpan(anchorId, spanStart, spanLength, prefix + newInner + suffix));
    }

    // ─── Tier B: structural ─────────────────────────────────────────────

    public static string InsertParagraph(int handle, string anchorId, Position position, string markdown) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertParagraph(anchorId, position, markdown));

    public static string SplitParagraph(int handle, string anchorId, int characterOffset) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).SplitParagraph(anchorId, characterOffset));

    public static string MergeParagraphs(int handle, string firstAnchorId, string secondAnchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).MergeParagraphs(firstAnchorId, secondAnchorId));

    public static string InsertHorizontalRule(int handle, string anchorId, Position position, string ruleJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertHorizontalRule(
            anchorId, position,
            string.IsNullOrEmpty(ruleJson) ? null : ParseRuleEdge(ruleJson)));

    private static ParagraphBorderEdge? ParseRuleEdge(string json)
    {
        // The rule JSON is itself a border-edge object; reuse the named-property parser by
        // wrapping it under a known key.
        using var d = System.Text.Json.JsonDocument.Parse($"{{\"e\":{json}}}");
        return DocxSessionJson.ParseBorderEdge(d.RootElement, "e");
    }

    // ─── Tier C: formatting ─────────────────────────────────────────────

    public static string ApplyFormat(int handle, string anchorId, CharSpan? span, FormatOp op) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).ApplyFormat(anchorId, span, op));

    public static string ApplyFormatBySubstring(int handle, string anchorId, string substring, FormatOp op) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).ApplyFormatToSubstring(anchorId, substring, op));

    public static string SetParagraphStyle(int handle, string anchorId, string styleId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetParagraphStyle(anchorId, styleId));

    public static string SetParagraphFormat(int handle, string anchorId, ParagraphFormatOp op) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetParagraphFormat(anchorId, op));

    public static string SetListLevel(int handle, string anchorId, int levelDelta) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).SetListLevel(anchorId, levelDelta));

    public static string RemoveListMembership(int handle, string anchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveListMembership(anchorId));

    public static string ApplyListFormat(int handle, string anchorId, ListFormat kind) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).ApplyListFormat(anchorId, kind));

    // ─── Tier D: tables ─────────────────────────────────────────────────

    public static string ReplaceCellContent(int handle, string cellAnchorId, string markdown) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).ReplaceCellContent(cellAnchorId, markdown));

    public static string InsertTable(int handle, string anchorId, Position position, int rows, int cols, string optionsJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertTable(
            anchorId, position, rows, cols, DocxSessionJson.ParseTableInsertOptions(optionsJson)));

    public static string InsertTableRow(int handle, string cellAnchorId, Position position) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertTableRow(cellAnchorId, position));

    public static string InsertTableColumn(int handle, string cellAnchorId, Position position) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).InsertTableColumn(cellAnchorId, position));

    public static string DeleteTableRow(int handle, string cellAnchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).DeleteTableRow(cellAnchorId));

    public static string DeleteTableColumn(int handle, string cellAnchorId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).DeleteTableColumn(cellAnchorId));

    // ─── Raw escape hatch ───────────────────────────────────────────────

    public static string RawGetXml(int handle, string anchorId) =>
        SessionRegistry.Get(handle).Raw.GetXml(anchorId);

    public static string RawInsertXml(int handle, string anchorId, Position position, string xml) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).Raw.InsertXml(anchorId, position, xml));

    public static string RawReplaceXml(int handle, string anchorId, string xml) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).Raw.ReplaceXml(anchorId, xml));

    // ─── Tier E: annotations ────────────────────────────────────────────

    public static string AddAnnotation(int handle, string anchorId, CharSpan? span,
        string annotationJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).AddAnnotation(
            anchorId, span, DocxSessionJson.DeserializeAnnotation(annotationJson)));

    public static string RemoveAnnotation(int handle, string annotationId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveAnnotation(annotationId));

    public static string UpdateAnnotation(int handle, string annotationId, string updateJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).UpdateAnnotation(
            annotationId, DocxSessionJson.DeserializeAnnotationUpdate(updateJson)));

    public static string MoveAnnotation(int handle, string annotationId, string newAnchorId,
        CharSpan? newSpan) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).MoveAnnotation(
            annotationId, newAnchorId, newSpan));

    // ─── Undo / Redo ────────────────────────────────────────────────────

    public static bool Undo(int handle) => SessionRegistry.Get(handle).Undo();

    public static bool Redo(int handle) => SessionRegistry.Get(handle).Redo();
}
