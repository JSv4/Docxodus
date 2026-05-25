using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text;
using System.Text.Json;
using Docxodus;

namespace DocxodusWasm;

/// <summary>
/// JSExport bridge for <see cref="DocxSession"/>. Sessions live on the .NET heap
/// and persist across JSExport calls — keyed by an integer handle returned from
/// <see cref="OpenSession"/>. JS-side code must call <see cref="CloseSession"/>
/// when done; sessions are not eligible for GC otherwise.
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class DocxSessionBridge
{
    private static readonly ConcurrentDictionary<int, DocxSession> _sessions = new();
    private static int _nextId;

    [JSExport]
    public static int OpenSession(byte[] bytes, string settingsJson)
    {
        var settings = ParseSettings(settingsJson);
        var session = new DocxSession(bytes, settings);
        var id = System.Threading.Interlocked.Increment(ref _nextId);
        _sessions[id] = session;
        return id;
    }

    [JSExport]
    public static void CloseSession(int handle)
    {
        if (_sessions.TryRemove(handle, out var s)) s.Dispose();
    }

    [JSExport]
    public static string Project(int handle) => SerializeProjection(Get(handle).Project());

    [JSExport]
    public static string ReplaceText(int h, string anchor, string md) =>
        Serialize(Get(h).ReplaceText(anchor, md));

    [JSExport]
    public static string DeleteBlock(int h, string anchor) =>
        Serialize(Get(h).DeleteBlock(anchor));

    [JSExport]
    public static string InsertParagraph(int h, string anchor, string posStr, string md) =>
        Serialize(Get(h).InsertParagraph(anchor, ParsePos(posStr), md));

    [JSExport]
    public static string SplitParagraph(int h, string anchor, int offset) =>
        Serialize(Get(h).SplitParagraph(anchor, offset));

    [JSExport]
    public static string MergeParagraphs(int h, string first, string second) =>
        Serialize(Get(h).MergeParagraphs(first, second));

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
        var op = ParseFormatOp(opJson);
        return Serialize(Get(h).ApplyFormat(anchor, span, op));
    }

    /// <summary>
    /// Bridge for the substring-targeted <see cref="DocxSession.ApplyFormat(string, string, FormatOp)"/>
    /// overload. Lets JS callers say "bold the substring 'foo' in this paragraph" without
    /// computing offsets — the overload finds the first occurrence and converts to a CharSpan.
    /// </summary>
    [JSExport]
    public static string ApplyFormatBySubstring(int h, string anchor, string substring, string opJson) =>
        Serialize(Get(h).ApplyFormatToSubstring(anchor, substring, ParseFormatOp(opJson)));

    [JSExport]
    public static string SetParagraphStyle(int h, string anchor, string styleId) =>
        Serialize(Get(h).SetParagraphStyle(anchor, styleId));

    [JSExport]
    public static string SetListLevel(int h, string anchor, int delta) =>
        Serialize(Get(h).SetListLevel(anchor, delta));

    [JSExport]
    public static string RemoveListMembership(int h, string anchor) =>
        Serialize(Get(h).RemoveListMembership(anchor));

    [JSExport]
    public static string ReplaceCellContent(int h, string anchor, string md) =>
        Serialize(Get(h).ReplaceCellContent(anchor, md));

    [JSExport]
    public static string RawGetXml(int h, string anchor) => Get(h).Raw.GetXml(anchor);

    [JSExport]
    public static string RawInsertXml(int h, string anchor, string posStr, string xml) =>
        Serialize(Get(h).Raw.InsertXml(anchor, ParsePos(posStr), xml));

    [JSExport]
    public static string RawReplaceXml(int h, string anchor, string xml) =>
        Serialize(Get(h).Raw.ReplaceXml(anchor, xml));

    /// <summary>
    /// Bridge for <see cref="DocxSession.Grep"/>. <paramref name="optionsJson"/>
    /// accepts <c>{regexOptions?: number, scope?: number, contextChars?: number}</c>;
    /// numeric values follow the .NET <see cref="System.Text.RegularExpressions.RegexOptions"/>
    /// and <see cref="ProjectionScopes"/> flag layouts. Missing fields use sensible
    /// defaults (no options, body-only, 40 chars of context).
    /// </summary>
    [JSExport]
    public static string Grep(int h, string pattern, string optionsJson)
    {
        var regexOpts = System.Text.RegularExpressions.RegexOptions.None;
        var scope = ProjectionScopes.Body;
        var contextChars = 40;
        var whitespace = WhitespaceMode.Preserve;
        if (!string.IsNullOrEmpty(optionsJson))
        {
            using var doc = JsonDocument.Parse(optionsJson);
            var root = doc.RootElement;
            if (root.TryGetProperty("regexOptions", out var ro) && ro.ValueKind == JsonValueKind.Number)
                regexOpts = (System.Text.RegularExpressions.RegexOptions)ro.GetInt32();
            if (root.TryGetProperty("scope", out var s) && s.ValueKind == JsonValueKind.Number)
                scope = (ProjectionScopes)s.GetInt32();
            if (root.TryGetProperty("contextChars", out var c) && c.ValueKind == JsonValueKind.Number)
                contextChars = c.GetInt32();
            if (root.TryGetProperty("whitespace", out var w) && w.ValueKind == JsonValueKind.Number)
                whitespace = (WhitespaceMode)w.GetInt32();
        }
        return SerializeMatches(Get(h).Grep(pattern, regexOpts, scope, contextChars, whitespace));
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GrepCrossBlock"/>. Same <paramref name="optionsJson"/>
    /// shape as <see cref="Grep"/>; returns a JSON array of CrossBlockMatch records (each
    /// carries <c>enclosingAnchors[]</c> + <c>slices[]</c>).
    /// </summary>
    [JSExport]
    public static string GrepCrossBlock(int h, string pattern, string optionsJson)
    {
        var regexOpts = System.Text.RegularExpressions.RegexOptions.None;
        var scope = ProjectionScopes.Body;
        var contextChars = 40;
        var whitespace = WhitespaceMode.Preserve;
        if (!string.IsNullOrEmpty(optionsJson))
        {
            using var doc = JsonDocument.Parse(optionsJson);
            var root = doc.RootElement;
            if (root.TryGetProperty("regexOptions", out var ro) && ro.ValueKind == JsonValueKind.Number)
                regexOpts = (System.Text.RegularExpressions.RegexOptions)ro.GetInt32();
            if (root.TryGetProperty("scope", out var s) && s.ValueKind == JsonValueKind.Number)
                scope = (ProjectionScopes)s.GetInt32();
            if (root.TryGetProperty("contextChars", out var c) && c.ValueKind == JsonValueKind.Number)
                contextChars = c.GetInt32();
            if (root.TryGetProperty("whitespace", out var w) && w.ValueKind == JsonValueKind.Number)
                whitespace = (WhitespaceMode)w.GetInt32();
        }
        return SerializeCrossBlockMatches(Get(h).GrepCrossBlock(pattern, regexOpts, scope, contextChars, whitespace));
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
                IgnoreCase = TryGetBool(root, "ignoreCase", false),
                MaxReplacements = root.TryGetProperty("maxReplacements", out var mr) && mr.ValueKind == JsonValueKind.Number
                    ? mr.GetInt32() : (int?)null,
            };
        }
        var results = Get(h).ReplaceTextRange(anchor, find, replace, opts);
        return SerializeEditResults(results);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceTextAtSpan"/> — the span-addressable
    /// variant that lets JS callers replace a specific Grep match (by its EnclosingAnchor
    /// id + Span coordinates) instead of every occurrence of its text.
    /// </summary>
    [JSExport]
    public static string ReplaceTextAtSpan(int h, string anchor, int spanStart, int spanLength, string replace) =>
        Serialize(Get(h).ReplaceTextAtSpan(anchor, spanStart, spanLength, replace));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindPlaceholders"/>. <paramref name="kinds"/>
    /// uses the numeric layout of <see cref="PlaceholderKinds"/> (BlankFill=1,
    /// AlternativeClause=2, Instruction=4, All=7); 0 returns nothing. <paramref name="scope"/>
    /// uses the <see cref="ProjectionScopes"/> flag layout. Returns a JSON array of placeholders.
    /// </summary>
    [JSExport]
    public static string FindPlaceholders(int h, int kinds, int scope)
    {
        var placeholders = Get(h).FindPlaceholders((PlaceholderKinds)kinds, (ProjectionScopes)scope);
        return SerializePlaceholders(placeholders);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByAnnotation"/>. Returns a JSON array of
    /// <see cref="AnchorTarget"/> records (each <c>{id, kind, scope, unid, partUri}</c>);
    /// empty array when the id is unknown.
    /// </summary>
    [JSExport]
    public static string FindByAnnotation(int h, string annotationId) =>
        SerializeAnchorTargets(Get(h).FindByAnnotation(annotationId));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByLabel"/>. Returns a JSON object keyed by
    /// annotation id; each value is the same AnchorTarget array shape as
    /// <see cref="FindByAnnotation"/>.
    /// </summary>
    [JSExport]
    public static string FindByLabel(int h, string labelId) =>
        SerializeAnchorTargetMap(Get(h).FindByLabel(labelId));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByBookmark"/>. Same return shape as
    /// <see cref="FindByAnnotation"/>; accepts any bookmark name (Docxodus-managed or
    /// user-authored).
    /// </summary>
    [JSExport]
    public static string FindByBookmark(int h, string bookmarkName) =>
        SerializeAnchorTargets(Get(h).FindByBookmark(bookmarkName));

    /// <summary>
    /// Bridge for <see cref="DocxSession.ListAnnotations"/>. Returns a JSON array of
    /// annotation records — id/labelId/label/color/author/created/bookmarkName/
    /// annotatedText. Metadata, page info, and the unused-by-agents fields are omitted
    /// to keep the wire format compact; callers needing them can use the .NET API.
    /// </summary>
    [JSExport]
    public static string ListAnnotations(int h) =>
        SerializeAnnotations(Get(h).ListAnnotations());

    [JSExport]
    public static bool Undo(int h) => Get(h).Undo();

    [JSExport]
    public static bool Redo(int h) => Get(h).Redo();

    [JSExport]
    public static byte[] Save(int h) => Get(h).Save();

    // ─── Helpers ────────────────────────────────────────────────────────

    private static DocxSession Get(int handle)
    {
        if (!_sessions.TryGetValue(handle, out var s))
            throw new ArgumentException($"unknown session handle: {handle}");
        return s;
    }

    private static Position ParsePos(string s) =>
        string.Equals(s, "before", StringComparison.OrdinalIgnoreCase) ? Position.Before : Position.After;

    private static DocxSessionSettings ParseSettings(string settingsJson)
    {
        if (string.IsNullOrEmpty(settingsJson)) return new DocxSessionSettings();
        using var doc = JsonDocument.Parse(settingsJson);
        var root = doc.RootElement;
        int undoDepth = TryGetInt(root, "undoDepth", 50);
        bool validateRawOps = TryGetBool(root, "validateRawOps", false);
        var trackedStr = TryGetString(root, "trackedChanges", "accept");
        var tracked = trackedStr switch
        {
            "render_inline" => TrackedChangeMode.RenderInline,
            "strip_deletions" => TrackedChangeMode.StripDeletions,
            _ => TrackedChangeMode.Accept,
        };
        var revisionAuthor = TryGetString(root, "revisionAuthor", null);
        bool persistAnchorIds = TryGetBool(root, "persistAnchorIds", false);
        bool smartQuotes = TryGetBool(root, "smartQuotes", false);
        return new DocxSessionSettings
        {
            UndoDepth = undoDepth,
            ValidateRawOps = validateRawOps,
            TrackedChanges = tracked,
            RevisionAuthor = revisionAuthor,
            PersistAnchorIds = persistAnchorIds,
            SmartQuotes = smartQuotes,
        };
    }

    private static FormatOp ParseFormatOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new FormatOp();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        return new FormatOp
        {
            Bold = TryGetBoolNullable(root, "bold"),
            Italic = TryGetBoolNullable(root, "italic"),
            Underline = TryGetBoolNullable(root, "underline"),
            Strike = TryGetBoolNullable(root, "strike"),
            Code = TryGetBoolNullable(root, "code"),
            Color = TryGetString(root, "color", null),
            RunStyle = TryGetString(root, "runStyle", null),
        };
    }

    private static int TryGetInt(JsonElement root, string name, int fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;
    private static bool TryGetBool(JsonElement root, string name, bool fallback) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : fallback;
    private static bool? TryGetBoolNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : (bool?)null;
    private static string? TryGetString(JsonElement root, string name, string? fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : fallback;

    private static string Serialize(EditResult r)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"success\":").Append(r.Success ? "true" : "false");
        if (r.Error is not null)
        {
            sb.Append(",\"error\":{")
              .Append("\"code\":\"").Append(EnumToSnake(r.Error.Code)).Append('"')
              .Append(",\"message\":").Append(JsonString(r.Error.Message));
            if (r.Error.AnchorId is not null)
                sb.Append(",\"anchorId\":").Append(JsonString(r.Error.AnchorId));
            sb.Append('}');
        }
        sb.Append(",\"created\":"); AppendAnchorArray(sb, r.Created);
        sb.Append(",\"removed\":"); AppendAnchorArray(sb, r.Removed);
        sb.Append(",\"modified\":"); AppendAnchorArray(sb, r.Modified);
        if (r.Patch is not null)
        {
            sb.Append(",\"patch\":{")
              .Append("\"scopeAnchorId\":").Append(JsonString(r.Patch.ScopeAnchorId))
              .Append(",\"markdown\":").Append(JsonString(r.Patch.Markdown))
              .Append('}');
        }
        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendAnchorArray(StringBuilder sb, System.Collections.Generic.IReadOnlyList<Anchor> anchors)
    {
        sb.Append('[');
        for (int i = 0; i < anchors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var a = anchors[i];
            sb.Append('{')
              .Append("\"id\":").Append(JsonString(a.Id))
              .Append(",\"kind\":").Append(JsonString(a.Kind))
              .Append(",\"scope\":").Append(JsonString(a.Scope))
              .Append(",\"unid\":").Append(JsonString(a.Unid))
              .Append('}');
        }
        sb.Append(']');
    }

    private static string SerializePlaceholders(System.Collections.Generic.IReadOnlyList<TemplatePlaceholder> placeholders)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < placeholders.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var p = placeholders[i];
            sb.Append("{\"kind\":\"").Append(p.Kind switch
            {
                PlaceholderKind.BlankFill => "blank_fill",
                PlaceholderKind.AlternativeClause => "alternative_clause",
                PlaceholderKind.Instruction => "instruction",
                _ => "unknown",
            }).Append('"');
            if (p.Hint is not null)
                sb.Append(",\"hint\":").Append(JsonString(p.Hint));
            sb.Append(",\"match\":");
            AppendMatch(sb, p.Match);
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static void AppendMatch(StringBuilder sb, TextMatch m)
    {
        sb.Append("{\"text\":").Append(JsonString(m.Text))
          .Append(",\"enclosingAnchor\":{")
          .Append("\"id\":").Append(JsonString(m.EnclosingAnchor.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(m.EnclosingAnchor.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(m.EnclosingAnchor.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(m.EnclosingAnchor.Anchor.Unid))
          .Append('}')
          .Append(",\"span\":{\"start\":").Append(m.Span.Start).Append(",\"length\":").Append(m.Span.Length).Append('}')
          .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
          .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
          .Append(",\"fragments\":[");
        for (int f = 0; f < m.Fragments.Count; f++)
        {
            if (f > 0) sb.Append(',');
            var fr = m.Fragments[f];
            sb.Append("{\"unid\":").Append(JsonString(fr.Unid))
              .Append(",\"text\":").Append(JsonString(fr.Text))
              .Append(",\"spanInElement\":{\"start\":").Append(fr.SpanInElement.Start)
              .Append(",\"length\":").Append(fr.SpanInElement.Length).Append('}')
              .Append(",\"formatting\":{")
              .Append("\"bold\":").Append(fr.Formatting.Bold ? "true" : "false")
              .Append(",\"italic\":").Append(fr.Formatting.Italic ? "true" : "false")
              .Append(",\"underline\":").Append(fr.Formatting.Underline ? "true" : "false")
              .Append(",\"strike\":").Append(fr.Formatting.Strike ? "true" : "false")
              .Append(",\"code\":").Append(fr.Formatting.Code ? "true" : "false")
              .Append("}}");
        }
        sb.Append(']');
        // Groups omitted from placeholder serialization (rarely useful for this surface).
        sb.Append('}');
    }

    private static string SerializeEditResults(System.Collections.Generic.IReadOnlyList<EditResult> results)
    {
        var sb = new StringBuilder(256);
        sb.Append('[');
        for (int i = 0; i < results.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(Serialize(results[i]));
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static string SerializeMatches(System.Collections.Generic.IReadOnlyList<TextMatch> matches)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < matches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var m = matches[i];
            sb.Append("{\"text\":").Append(JsonString(m.Text))
              .Append(",\"enclosingAnchor\":");
            AppendAnchor(sb, m.EnclosingAnchor);
            sb.Append(",\"span\":{\"start\":").Append(m.Span.Start).Append(",\"length\":").Append(m.Span.Length).Append('}')
              .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
              .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
              .Append(",\"groups\":");
            AppendStringArray(sb, m.Groups);
            sb.Append(",\"fragments\":");
            AppendFragments(sb, m.Fragments);
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static string SerializeCrossBlockMatches(System.Collections.Generic.IReadOnlyList<CrossBlockMatch> matches)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < matches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var m = matches[i];
            sb.Append("{\"text\":").Append(JsonString(m.Text))
              .Append(",\"enclosingAnchors\":[");
            for (int a = 0; a < m.EnclosingAnchors.Count; a++)
            {
                if (a > 0) sb.Append(',');
                AppendAnchor(sb, m.EnclosingAnchors[a]);
            }
            sb.Append(']')
              .Append(",\"slices\":[");
            for (int sIdx = 0; sIdx < m.Slices.Count; sIdx++)
            {
                if (sIdx > 0) sb.Append(',');
                var slice = m.Slices[sIdx];
                sb.Append("{\"anchor\":");
                AppendAnchor(sb, slice.Anchor);
                sb.Append(",\"spanInBlock\":{\"start\":").Append(slice.SpanInBlock.Start)
                  .Append(",\"length\":").Append(slice.SpanInBlock.Length).Append('}')
                  .Append(",\"fragments\":");
                AppendFragments(sb, slice.Fragments);
                sb.Append('}');
            }
            sb.Append(']')
              .Append(",\"contextBefore\":").Append(JsonString(m.ContextBefore))
              .Append(",\"contextAfter\":").Append(JsonString(m.ContextAfter))
              .Append(",\"groups\":");
            AppendStringArray(sb, m.Groups);
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static void AppendAnchor(StringBuilder sb, AnchorTarget t) =>
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Anchor.Unid))
          .Append('}');

    private static void AppendStringArray(StringBuilder sb, System.Collections.Generic.IReadOnlyList<string> items)
    {
        sb.Append('[');
        for (int i = 0; i < items.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(items[i]));
        }
        sb.Append(']');
    }

    private static void AppendFragments(StringBuilder sb, System.Collections.Generic.IReadOnlyList<RunFragment> fragments)
    {
        sb.Append('[');
        for (int f = 0; f < fragments.Count; f++)
        {
            if (f > 0) sb.Append(',');
            var fr = fragments[f];
            sb.Append("{\"unid\":").Append(JsonString(fr.Unid))
              .Append(",\"text\":").Append(JsonString(fr.Text))
              .Append(",\"spanInElement\":{\"start\":").Append(fr.SpanInElement.Start)
              .Append(",\"length\":").Append(fr.SpanInElement.Length).Append('}')
              .Append(",\"formatting\":{")
              .Append("\"bold\":").Append(fr.Formatting.Bold ? "true" : "false")
              .Append(",\"italic\":").Append(fr.Formatting.Italic ? "true" : "false")
              .Append(",\"underline\":").Append(fr.Formatting.Underline ? "true" : "false")
              .Append(",\"strike\":").Append(fr.Formatting.Strike ? "true" : "false")
              .Append(",\"code\":").Append(fr.Formatting.Code ? "true" : "false");
            if (fr.Formatting.Color is not null)
                sb.Append(",\"color\":").Append(JsonString(fr.Formatting.Color));
            if (fr.Formatting.HyperlinkUrl is not null)
                sb.Append(",\"hyperlinkUrl\":").Append(JsonString(fr.Formatting.HyperlinkUrl));
            if (fr.Formatting.RunStyle is not null)
                sb.Append(",\"runStyle\":").Append(JsonString(fr.Formatting.RunStyle));
            sb.Append("}}");
        }
        sb.Append(']');
    }

    private static string SerializeProjection(MarkdownProjection p)
    {
        var sb = new StringBuilder(p.Markdown.Length + 200);
        sb.Append("{\"markdown\":").Append(JsonString(p.Markdown));
        sb.Append(",\"anchorIndex\":{");
        bool first = true;
        foreach (var kv in p.AnchorIndex)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(":{")
              .Append("\"partUri\":").Append(JsonString(kv.Value.PartUri))
              .Append(",\"unid\":").Append(JsonString(kv.Value.Unid))
              .Append(",\"kind\":").Append(JsonString(kv.Value.Anchor.Kind))
              .Append(",\"scope\":").Append(JsonString(kv.Value.Anchor.Scope))
              .Append('}');
        }
        sb.Append("}}");
        return sb.ToString();
    }

    private static string JsonString(string s)
    {
        var sb = new StringBuilder(s.Length + 2);
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    private static string EnumToSnake(EditErrorCode code)
    {
        var s = code.ToString();
        var sb = new StringBuilder(s.Length + 4);
        for (int i = 0; i < s.Length; i++)
        {
            if (i > 0 && char.IsUpper(s[i])) sb.Append('_');
            sb.Append(char.ToLowerInvariant(s[i]));
        }
        return sb.ToString();
    }

    private static string SerializeAnchorTargets(System.Collections.Generic.IReadOnlyList<AnchorTarget> targets)
    {
        var sb = new StringBuilder(targets.Count * 128 + 2);
        sb.Append('[');
        for (int i = 0; i < targets.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendAnchorTarget(sb, targets[i]);
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static string SerializeAnchorTargetMap(
        System.Collections.Generic.IReadOnlyDictionary<string, System.Collections.Generic.IReadOnlyList<AnchorTarget>> map)
    {
        var sb = new StringBuilder(256);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeAnchorTargets(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    private static void AppendAnchorTarget(StringBuilder sb, AnchorTarget t)
    {
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Unid))
          .Append(",\"partUri\":").Append(JsonString(t.PartUri))
          .Append('}');
    }

    private static string SerializeAnnotations(System.Collections.Generic.IReadOnlyList<DocumentAnnotation> anns)
    {
        var sb = new StringBuilder(anns.Count * 200 + 2);
        sb.Append('[');
        for (int i = 0; i < anns.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var a = anns[i];
            sb.Append("{\"id\":").Append(JsonString(a.Id ?? string.Empty))
              .Append(",\"labelId\":").Append(JsonString(a.LabelId ?? string.Empty))
              .Append(",\"label\":").Append(JsonString(a.Label ?? string.Empty))
              .Append(",\"color\":").Append(JsonString(a.Color ?? string.Empty))
              .Append(",\"bookmarkName\":").Append(JsonString(a.BookmarkName ?? string.Empty));
            if (a.Author is not null)
                sb.Append(",\"author\":").Append(JsonString(a.Author));
            if (a.Created.HasValue)
                sb.Append(",\"created\":").Append(JsonString(a.Created.Value.ToString("o")));
            if (a.AnnotatedText is not null)
                sb.Append(",\"annotatedText\":").Append(JsonString(a.AnnotatedText));
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }
}
