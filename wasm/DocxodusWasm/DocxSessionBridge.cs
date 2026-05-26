#nullable enable

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus;
using Docxodus.Internal;

namespace DocxodusWasm;

/// <summary>
/// JSExport bridge for <see cref="DocxSession"/>. Sessions live on the .NET heap
/// and persist across JSExport calls — keyed by an integer handle returned from
/// <see cref="OpenSession"/>. JS-side code must call <see cref="CloseSession"/>
/// when done; sessions are not eligible for GC otherwise.
///
/// All wire-format work — JSON serialization, settings/format-op parsing, the
/// handle pool — lives in <see cref="DocxSessionOps"/> / <see cref="DocxSessionJson"/>
/// / <see cref="SessionRegistry"/>. This file is a thin JSExport-attributed
/// shell so the WASM and stdio NDJSON transports stay byte-for-byte identical.
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class DocxSessionBridge
{
    [JSExport]
    public static int OpenSession(byte[] bytes, string settingsJson) =>
        DocxSessionOps.OpenSession(bytes, DocxSessionJson.ParseSettings(settingsJson));

    [JSExport]
    public static void CloseSession(int handle) => DocxSessionOps.CloseSession(handle);

    [JSExport]
    public static string Project(int handle) => DocxSessionOps.Project(handle);

    [JSExport]
    public static string ReplaceText(int h, string anchor, string md) =>
        DocxSessionOps.ReplaceText(h, anchor, md);

    [JSExport]
    public static string DeleteBlock(int h, string anchor) =>
        DocxSessionOps.DeleteBlock(h, anchor);

    /// <summary>
    /// Bridge for <see cref="DocxSession.DeleteRange"/>. Deletes every top-level
    /// block-level sibling between <paramref name="fromAnchorId"/> (inclusive) and
    /// <paramref name="toAnchorIdExclusive"/> (exclusive). Both anchors must share a
    /// direct parent and live in the same package part. Returns a single EditResult.
    /// </summary>
    [JSExport]
    public static string DeleteRange(int h, string fromAnchorId, string toAnchorIdExclusive) =>
        DocxSessionOps.DeleteRange(h, fromAnchorId, toAnchorIdExclusive);

    /// <summary>
    /// Bridge for <see cref="DocxSession.DeleteSection"/>. Deletes a heading and
    /// every sibling below it up to (but not including) the next heading at the
    /// same or higher level. <paramref name="headingAnchorId"/> must address a
    /// heading-kind anchor (<c>h</c>).
    /// </summary>
    [JSExport]
    public static string DeleteSection(int h, string headingAnchorId) =>
        DocxSessionOps.DeleteSection(h, headingAnchorId);

    [JSExport]
    public static string InsertParagraph(int h, string anchor, string posStr, string md) =>
        DocxSessionOps.InsertParagraph(h, anchor, DocxSessionJson.ParsePos(posStr), md);

    [JSExport]
    public static string SplitParagraph(int h, string anchor, int offset) =>
        DocxSessionOps.SplitParagraph(h, anchor, offset);

    [JSExport]
    public static string MergeParagraphs(int h, string first, string second) =>
        DocxSessionOps.MergeParagraphs(h, first, second);

    [JSExport]
    public static string ApplyFormat(int h, string anchor, string spanJson, string opJson)
    {
        CharSpan? span = null;
        if (!string.IsNullOrEmpty(spanJson))
        {
            using var doc = JsonDocument.Parse(spanJson);
            span = new CharSpan(
                doc.RootElement.GetProperty("start").GetInt32(),
                doc.RootElement.GetProperty("length").GetInt32());
        }
        return DocxSessionOps.ApplyFormat(h, anchor, span, DocxSessionJson.ParseFormatOp(opJson));
    }

    /// <summary>
    /// Bridge for the substring-targeted <see cref="DocxSession.ApplyFormat(string, string, FormatOp)"/>
    /// overload. Lets JS callers say "bold the substring 'foo' in this paragraph" without
    /// computing offsets — the overload finds the first occurrence and converts to a CharSpan.
    /// </summary>
    [JSExport]
    public static string ApplyFormatBySubstring(int h, string anchor, string substring, string opJson) =>
        DocxSessionOps.ApplyFormatBySubstring(h, anchor, substring, DocxSessionJson.ParseFormatOp(opJson));

    [JSExport]
    public static string SetParagraphStyle(int h, string anchor, string styleId) =>
        DocxSessionOps.SetParagraphStyle(h, anchor, styleId);

    [JSExport]
    public static string SetListLevel(int h, string anchor, int delta) =>
        DocxSessionOps.SetListLevel(h, anchor, delta);

    [JSExport]
    public static string RemoveListMembership(int h, string anchor) =>
        DocxSessionOps.RemoveListMembership(h, anchor);

    [JSExport]
    public static string ReplaceCellContent(int h, string anchor, string md) =>
        DocxSessionOps.ReplaceCellContent(h, anchor, md);

    [JSExport]
    public static string RawGetXml(int h, string anchor) => DocxSessionOps.RawGetXml(h, anchor);

    [JSExport]
    public static string RawInsertXml(int h, string anchor, string posStr, string xml) =>
        DocxSessionOps.RawInsertXml(h, anchor, DocxSessionJson.ParsePos(posStr), xml);

    [JSExport]
    public static string RawReplaceXml(int h, string anchor, string xml) =>
        DocxSessionOps.RawReplaceXml(h, anchor, xml);

    /// <summary>
    /// Bridge for <see cref="DocxSession.Grep"/>. <paramref name="optionsJson"/>
    /// accepts <c>{regexOptions?: number, scope?: number, contextChars?: number,
    /// whitespace?: number, boundary?: number}</c>; numeric values follow the .NET
    /// <see cref="System.Text.RegularExpressions.RegexOptions"/>, <see cref="ProjectionScopes"/>,
    /// <see cref="WhitespaceMode"/>, and <see cref="ContextBoundary"/> flag layouts.
    /// Missing fields use sensible defaults (no options, body-only, 80 chars of
    /// context, preserve whitespace, char-boundary).
    /// </summary>
    [JSExport]
    public static string Grep(int h, string pattern, string optionsJson)
    {
        ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
        return DocxSessionOps.Grep(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GrepCrossBlock"/>. Same <paramref name="optionsJson"/>
    /// shape as <see cref="Grep"/>; returns a JSON array of CrossBlockMatch records (each
    /// carries <c>enclosingAnchors[]</c> + <c>slices[]</c>).
    /// </summary>
    [JSExport]
    public static string GrepCrossBlock(int h, string pattern, string optionsJson)
    {
        ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
        return DocxSessionOps.GrepCrossBlock(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceTextRange"/>. <paramref name="optionsJson"/>
    /// accepts <c>{ignoreCase?: boolean, maxReplacements?: number}</c>. Returns a
    /// JSON array of EditResult — one per attempted match.
    /// </summary>
    [JSExport]
    public static string ReplaceTextRange(int h, string anchor, string find, string replace, string optionsJson)
    {
        ReplaceOptions? opts = null;
        if (!string.IsNullOrEmpty(optionsJson))
        {
            using var doc = JsonDocument.Parse(optionsJson);
            var root = doc.RootElement;
            opts = new ReplaceOptions
            {
                IgnoreCase = DocxSessionJson.TryGetBool(root, "ignoreCase", false),
                MaxReplacements = root.TryGetProperty("maxReplacements", out var mr) && mr.ValueKind == JsonValueKind.Number
                    ? mr.GetInt32() : (int?)null,
            };
        }
        return DocxSessionOps.ReplaceTextRange(h, anchor, find, replace, opts);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceTextAtSpan"/> — the span-addressable
    /// variant that lets JS callers replace a specific Grep match (by its EnclosingAnchor
    /// id + Span coordinates) instead of every occurrence of its text.
    /// </summary>
    [JSExport]
    public static string ReplaceTextAtSpan(int h, string anchor, int spanStart, int spanLength, string replace) =>
        DocxSessionOps.ReplaceTextAtSpan(h, anchor, spanStart, spanLength, replace);

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceInner(TextMatch, string)"/>. Takes the
    /// match's text (so the shared core can locate the brackets) plus anchor+span (so
    /// it can dispatch to <see cref="DocxSession.ReplaceTextAtSpan"/>). Bracket
    /// parsing happens transport-side rather than serializing a full <see cref="TextMatch"/>
    /// — the existing wire shape already carries text + anchor + span via Grep results,
    /// so callers don't need anything they don't already have.
    /// </summary>
    [JSExport]
    public static string ReplaceInner(int h, string matchText, string anchor, int spanStart, int spanLength, string newInner) =>
        DocxSessionOps.ReplaceInner(h, matchText, anchor, spanStart, spanLength, newInner);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindPlaceholders"/>. <paramref name="kinds"/>
    /// uses the numeric layout of <see cref="PlaceholderKinds"/> (BlankFill=1,
    /// AlternativeClause=2, Instruction=4, All=7); 0 returns nothing. <paramref name="scope"/>
    /// uses the <see cref="ProjectionScopes"/> flag layout. Returns a JSON array of placeholders.
    /// </summary>
    [JSExport]
    public static string FindPlaceholders(int h, int kinds, int scope, int contextChars, int boundary) =>
        DocxSessionOps.FindPlaceholders(h, (PlaceholderKinds)kinds, (ProjectionScopes)scope, contextChars, (ContextBoundary)boundary);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByAnnotation"/>. Returns a JSON array of
    /// <see cref="AnchorTarget"/> records (each <c>{id, kind, scope, unid, partUri}</c>);
    /// empty array when the id is unknown.
    /// </summary>
    [JSExport]
    public static string FindByAnnotation(int h, string annotationId) =>
        DocxSessionOps.FindByAnnotation(h, annotationId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByLabel"/>. Returns a JSON object keyed by
    /// annotation id; each value is the same AnchorTarget array shape as
    /// <see cref="FindByAnnotation"/>.
    /// </summary>
    [JSExport]
    public static string FindByLabel(int h, string labelId) =>
        DocxSessionOps.FindByLabel(h, labelId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByBookmark"/>. Same return shape as
    /// <see cref="FindByAnnotation"/>; accepts any bookmark name (Docxodus-managed or
    /// user-authored).
    /// </summary>
    [JSExport]
    public static string FindByBookmark(int h, string bookmarkName) =>
        DocxSessionOps.FindByBookmark(h, bookmarkName);

    /// <summary>
    /// Bridge for <see cref="DocxSession.ListAnnotations"/>. Returns a JSON array of
    /// annotation records — id/labelId/label/color/author/created/bookmarkName/
    /// annotatedText. Metadata, page info, and the unused-by-agents fields are omitted
    /// to keep the wire format compact; callers needing them can use the .NET API.
    /// </summary>
    [JSExport]
    public static string ListAnnotations(int h) => DocxSessionOps.ListAnnotations(h);

    /// <summary>Bridge for <see cref="DocxSession.Exists"/>. Returns true/false.</summary>
    [JSExport]
    public static bool Exists(int h, string anchorId) => DocxSessionOps.Exists(h, anchorId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetAnchorInfo"/>. Returns a JSON object
    /// <c>{id, kind, scope, textPreview}</c> or the literal <c>null</c> if the
    /// anchor is not found.
    /// </summary>
    [JSExport]
    public static string GetAnchorInfo(int h, string anchorId) => DocxSessionOps.GetAnchorInfo(h, anchorId);

    /// <summary>
    /// Bulk variant of <see cref="GetAnchorInfo"/>. Takes a JSON array of anchor
    /// ids and returns a JSON object keyed by id; each value is the AnchorInfo
    /// shape or <c>null</c> for unknown ids.
    /// </summary>
    [JSExport]
    public static string GetAnchorInfos(int h, string anchorIdsJson)
    {
        string[] ids;
        try
        {
            ids = JsonSerializer.Deserialize<string[]>(
                anchorIdsJson, DocxodusJsonContext.Default.StringArray) ?? System.Array.Empty<string>();
        }
        catch (JsonException)
        {
            return "{\"error\":\"malformed anchor id array\"}";
        }
        return DocxSessionOps.GetAnchorInfos(h, ids);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByText"/>. Returns a single AnchorTarget
    /// JSON object (first match in document order) or the literal <c>null</c> if no
    /// anchor contains the needle. <paramref name="optionsJson"/> accepts
    /// <c>{ignoreCase?, ignoreWhitespace?, kindFilter?, scopeFilter?}</c>.
    /// </summary>
    [JSExport]
    public static string FindByText(int h, string needle, string optionsJson) =>
        DocxSessionOps.FindByText(h, needle, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindAllByText"/>. Same options shape as
    /// <see cref="FindByText"/>; returns the full AnchorTarget array in document order.
    /// </summary>
    [JSExport]
    public static string FindAllByText(int h, string needle, string optionsJson) =>
        DocxSessionOps.FindAllByText(h, needle, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByRegex"/>. <paramref name="regexOptions"/>
    /// uses the numeric layout of <see cref="System.Text.RegularExpressions.RegexOptions"/>;
    /// <paramref name="optionsJson"/> matches the <see cref="FindByText"/> shape.
    /// </summary>
    [JSExport]
    public static string FindByRegex(int h, string pattern, int regexOptions, string optionsJson) =>
        DocxSessionOps.FindByRegex(h, pattern, (RegexOptions)regexOptions, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByKind"/>. <paramref name="scope"/> may be
    /// empty/null to match all scopes. No text scan — reads the AnchorIndex directly.
    /// </summary>
    [JSExport]
    public static string FindByKind(int h, string kind, string scope) =>
        DocxSessionOps.FindByKind(h, kind, string.IsNullOrEmpty(scope) ? null : scope);

    [JSExport]
    public static bool Undo(int h) => DocxSessionOps.Undo(h);

    [JSExport]
    public static bool Redo(int h) => DocxSessionOps.Redo(h);

    [JSExport]
    public static byte[] Save(int h) => DocxSessionOps.Save(h);

    // ─── Helpers ────────────────────────────────────────────────────────

    private static FindOptions? ParseFindOptions(string optionsJson)
    {
        if (string.IsNullOrEmpty(optionsJson)) return null;
        using var doc = JsonDocument.Parse(optionsJson);
        return DocxSessionJson.ParseFindOptions(doc.RootElement);
    }

    private static void ParseGrepOptions(string optionsJson, out RegexOptions regexOpts,
        out ProjectionScopes scope, out int contextChars, out WhitespaceMode whitespace,
        out ContextBoundary boundary)
    {
        regexOpts = RegexOptions.None;
        scope = ProjectionScopes.Body;
        contextChars = 80;
        whitespace = WhitespaceMode.Preserve;
        boundary = ContextBoundary.Char;
        if (string.IsNullOrEmpty(optionsJson)) return;
        using var doc = JsonDocument.Parse(optionsJson);
        var root = doc.RootElement;
        if (root.TryGetProperty("regexOptions", out var ro) && ro.ValueKind == JsonValueKind.Number)
            regexOpts = (RegexOptions)ro.GetInt32();
        if (root.TryGetProperty("scope", out var s) && s.ValueKind == JsonValueKind.Number)
            scope = (ProjectionScopes)s.GetInt32();
        if (root.TryGetProperty("contextChars", out var c) && c.ValueKind == JsonValueKind.Number)
            contextChars = c.GetInt32();
        if (root.TryGetProperty("whitespace", out var w) && w.ValueKind == JsonValueKind.Number)
            whitespace = (WhitespaceMode)w.GetInt32();
        if (root.TryGetProperty("boundary", out var b) && b.ValueKind == JsonValueKind.Number)
            boundary = (ContextBoundary)b.GetInt32();
    }
}
