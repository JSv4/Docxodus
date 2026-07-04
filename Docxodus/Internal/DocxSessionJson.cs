#nullable enable

using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace Docxodus.Internal;

/// <summary>
/// Shared JSON serialization + parsing for the DocxSession bridge wire format.
/// Both the WASM JSExport bridge and the stdio NDJSON host emit and consume the
/// shapes defined here, so the TypeScript and Python clients see identical JSON.
/// All output is camelCase; clients that prefer snake_case (e.g. Python dataclasses)
/// convert during deserialization.
/// </summary>
internal static class DocxSessionJson
{
    // ─── Parsers ────────────────────────────────────────────────────────

    public static Position ParsePos(string s) =>
        string.Equals(s, "before", System.StringComparison.OrdinalIgnoreCase) ? Position.Before : Position.After;

    public static DocxSessionSettings ParseSettings(string settingsJson)
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
        bool captureInitialProjection = TryGetBool(root, "captureInitialProjection", true);
        var projectionSettings = root.TryGetProperty("projectionSettings", out var ps) && ps.ValueKind == JsonValueKind.Object
            ? ParseProjectionSettings(ps)
            : new WmlToMarkdownConverterSettings();
        return new DocxSessionSettings
        {
            UndoDepth = undoDepth,
            ValidateRawOps = validateRawOps,
            TrackedChanges = tracked,
            RevisionAuthor = revisionAuthor,
            PersistAnchorIds = persistAnchorIds,
            SmartQuotes = smartQuotes,
            CaptureInitialProjection = captureInitialProjection,
            ProjectionSettings = projectionSettings,
        };
    }

    /// <summary>
    /// Parse a JSON object into <see cref="WmlToMarkdownConverterSettings"/>. Mirrors
    /// the <c>MarkdownProjectionSettings</c> TS interface and the
    /// <c>MarkdownProjectionSettingsDto</c> WASM DTO — numeric enum fields use the
    /// same flag/value layout as the .NET enums. Unknown / missing fields fall back
    /// to <see cref="WmlToMarkdownConverterSettings"/> defaults.
    /// </summary>
    public static WmlToMarkdownConverterSettings ParseProjectionSettings(JsonElement root)
    {
        var settings = new WmlToMarkdownConverterSettings();
        if (root.ValueKind != JsonValueKind.Object) return settings;
        if (root.TryGetProperty("scopes", out var sc) && sc.ValueKind == JsonValueKind.Number)
            settings.Scopes = (ProjectionScopes)sc.GetInt32();
        if (root.TryGetProperty("headingLevelOffset", out var hl) && hl.ValueKind == JsonValueKind.Number)
            settings.HeadingLevelOffset = hl.GetInt32();
        if (root.TryGetProperty("anchorMode", out var am) && am.ValueKind == JsonValueKind.Number)
            settings.AnchorMode = (AnchorRenderMode)am.GetInt32();
        if (root.TryGetProperty("tableMode", out var tm) && tm.ValueKind == JsonValueKind.Number)
            settings.TableMode = (TableRenderMode)tm.GetInt32();
        if (root.TryGetProperty("tableInlineCellMax", out var tic) && tic.ValueKind == JsonValueKind.Number)
            settings.TableInlineCellMax = tic.GetInt32();
        if (root.TryGetProperty("trackedChanges", out var tc) && tc.ValueKind == JsonValueKind.Number)
            settings.TrackedChanges = (TrackedChangeMode)tc.GetInt32();
        if (root.TryGetProperty("resolveNumbering", out var rn) && (rn.ValueKind == JsonValueKind.True || rn.ValueKind == JsonValueKind.False))
            settings.ResolveNumbering = rn.GetBoolean();
        if (root.TryGetProperty("emptyParagraphs", out var ep) && ep.ValueKind == JsonValueKind.Number)
            settings.EmptyParagraphs = (EmptyParagraphMode)ep.GetInt32();
        if (root.TryGetProperty("anchorIdRendering", out var air) && air.ValueKind == JsonValueKind.Number)
            settings.AnchorIdRendering = (AnchorIdRendering)air.GetInt32();
        return settings;
    }

    public static FormatOp ParseFormatOp(string json)
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
            VertAlign = TryGetString(root, "vertAlign", null),
            FontSizePts = TryGetDoubleNullable(root, "fontSizePts"),
            FontFamily = TryGetString(root, "fontFamily", null),
        };
    }

    /// <summary>
    /// Parse a ParagraphFormatOp wire object: { alignment?: "left"|"center"|"right"|"justify",
    /// indentDelta?: int (twips), pageBreakBefore?: bool }. Missing fields leave that property
    /// unchanged.
    /// </summary>
    public static ParagraphFormatOp ParseParagraphFormatOp(string json)
    {
        if (string.IsNullOrEmpty(json)) return new ParagraphFormatOp();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        ParagraphAlignment? align = TryGetString(root, "alignment", null)?.ToLowerInvariant() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" or "both" => ParagraphAlignment.Justify,
            _ => null,
        };
        int? indentDelta = root.TryGetProperty("indentDelta", out var d) && d.ValueKind == JsonValueKind.Number
            ? d.GetInt32()
            : null;
        return new ParagraphFormatOp
        {
            Alignment = align,
            IndentDelta = indentDelta,
            PageBreakBefore = TryGetBoolNullable(root, "pageBreakBefore"),
            TopBorder = ParseBorderEdge(root, "topBorder"),
            BottomBorder = ParseBorderEdge(root, "bottomBorder"),
            ClearBorders = TryGetBoolNullable(root, "clearBorders"),
        };
    }

    /// <summary>
    /// Parse a <see cref="ParagraphBorderEdge"/> from a named object property
    /// ({ style?, size?, color?, space? }), or null when the property is absent/not an object.
    /// </summary>
    public static ParagraphBorderEdge? ParseBorderEdge(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out var e) || e.ValueKind != JsonValueKind.Object) return null;
        return new ParagraphBorderEdge
        {
            Style = TryGetString(e, "style", null),
            Size = e.TryGetProperty("size", out var s) && s.ValueKind == JsonValueKind.Number ? s.GetInt32() : null,
            Color = TryGetString(e, "color", null),
            Space = e.TryGetProperty("space", out var sp) && sp.ValueKind == JsonValueKind.Number ? sp.GetInt32() : null,
        };
    }

    /// <summary>
    /// Parse a <see cref="TableInsertOptions"/> wire object:
    /// { borderless?: bool, cellContents?: string[], cellAlignment?: "left"|"center"|"right"|"justify",
    ///   columnWidths?: number[] (twips, one per column) }.
    /// </summary>
    public static TableInsertOptions ParseTableInsertOptions(string json)
    {
        if (string.IsNullOrEmpty(json)) return new TableInsertOptions();
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        List<string>? cells = null;
        if (root.TryGetProperty("cellContents", out var cc) && cc.ValueKind == JsonValueKind.Array)
        {
            cells = new List<string>(cc.GetArrayLength());
            foreach (var item in cc.EnumerateArray())
                cells.Add(item.ValueKind == JsonValueKind.String ? (item.GetString() ?? string.Empty) : string.Empty);
        }
        ParagraphAlignment? align = TryGetString(root, "cellAlignment", null)?.ToLowerInvariant() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" or "both" => ParagraphAlignment.Justify,
            _ => null,
        };
        List<int>? widths = null;
        if (root.TryGetProperty("columnWidths", out var cw) && cw.ValueKind == JsonValueKind.Array)
        {
            widths = new List<int>(cw.GetArrayLength());
            foreach (var item in cw.EnumerateArray())
                widths.Add(item.ValueKind == JsonValueKind.Number ? item.GetInt32() : 0);
        }
        return new TableInsertOptions
        {
            Borderless = TryGetBool(root, "borderless", false),
            CellContents = cells,
            CellAlignment = align,
            ColumnWidths = widths,
        };
    }

    /// <summary>Parse a list-format kind string ("bullet" | "decimal"/"number" | "none").</summary>
    public static ListFormat ParseListFormat(string? kind) => kind?.ToLowerInvariant() switch
    {
        "bullet" => ListFormat.Bullet,
        "decimal" or "number" or "numbered" => ListFormat.Decimal,
        _ => ListFormat.None,
    };

    public static FindOptions? ParseFindOptions(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return null;
        var scopes = (ProjectionScopes)TryGetInt(root, "scopes", (int)ProjectionScopes.All);
        return new FindOptions
        {
            IgnoreCase = TryGetBool(root, "ignoreCase", false),
            IgnoreWhitespace = TryGetBool(root, "ignoreWhitespace", false),
            KindFilter = TryGetString(root, "kindFilter", null),
            Scopes = scopes,
            ScopeFilter = TryGetString(root, "scopeFilter", null),
        };
    }

    public static int TryGetInt(JsonElement root, string name, int fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;

    public static bool TryGetBool(JsonElement root, string name, bool fallback) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : fallback;

    public static bool? TryGetBoolNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : (bool?)null;

    public static string? TryGetString(JsonElement root, string name, string? fallback) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : fallback;

    public static double? TryGetDoubleNullable(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetDouble() : (double?)null;

    // ─── Serializers ────────────────────────────────────────────────────

    public static string Serialize(EditResult r)
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
        if (r.AnnotationId is not null)
            sb.Append(",\"annotationId\":").Append(JsonString(r.AnnotationId));
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

    public static string SerializeEditResults(IReadOnlyList<EditResult> results)
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

    public static void AppendAnchorArray(StringBuilder sb, IReadOnlyList<Anchor> anchors)
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

    private static string KindToString(PlaceholderKind kind) => kind switch
    {
        PlaceholderKind.BlankFill => "blank_fill",
        PlaceholderKind.AlternativeClause => "alternative_clause",
        PlaceholderKind.Instruction => "instruction",
        _ => "unknown",
    };

    public static string SerializeEditSummary(EditSummary summary)
    {
        var sb = new StringBuilder(1024);
        sb.Append("{\"totalAnchors\":").Append(summary.TotalAnchors)
          .Append(",\"remainingPlaceholders\":").Append(SerializePlaceholders(summary.RemainingPlaceholders))
          .Append(",\"bareUnderscoreRuns\":").Append(SerializeMatches(summary.BareUnderscoreRuns))
          .Append(",\"footnoteCount\":").Append(summary.FootnoteCount)
          .Append(",\"inlineFootnoteRefCount\":").Append(summary.InlineFootnoteRefCount)
          .Append(",\"commentCount\":").Append(summary.CommentCount)
          .Append('}');
        return sb.ToString();
    }

    public static string SerializePlaceholders(IReadOnlyList<TemplatePlaceholder> placeholders)
    {
        var sb = new StringBuilder(512);
        sb.Append('[');
        for (int i = 0; i < placeholders.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var p = placeholders[i];
            sb.Append("{\"kind\":\"").Append(KindToString(p.Kind)).Append('"');
            sb.Append(",\"alternativeKinds\":[");
            for (int a = 0; a < p.AlternativeKinds.Count; a++)
            {
                if (a > 0) sb.Append(',');
                sb.Append('"').Append(KindToString(p.AlternativeKinds[a])).Append('"');
            }
            sb.Append(']');
            if (p.Hint is not null)
                sb.Append(",\"hint\":").Append(JsonString(p.Hint));
            sb.Append(",\"match\":");
            AppendMatch(sb, p.Match);
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    public static void AppendMatch(StringBuilder sb, TextMatch m)
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

    public static string SerializeMatches(IReadOnlyList<TextMatch> matches)
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

    public static string SerializeCrossBlockMatches(IReadOnlyList<CrossBlockMatch> matches)
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

    public static void AppendAnchor(StringBuilder sb, AnchorTarget t) =>
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Anchor.Unid))
          .Append('}');

    public static void AppendStringArray(StringBuilder sb, IReadOnlyList<string> items)
    {
        sb.Append('[');
        for (int i = 0; i < items.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(items[i]));
        }
        sb.Append(']');
    }

    public static void AppendFragments(StringBuilder sb, IReadOnlyList<RunFragment> fragments)
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

    public static string SerializeProjection(MarkdownProjection p)
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
              .Append(",\"textPreview\":").Append(JsonString(kv.Value.TextPreview));
            if (kv.Value.AutoNumberPrefix is { } prefix)
                sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
            sb.Append('}');
        }
        sb.Append("}}");
        return sb.ToString();
    }

    public static string JsonString(string s)
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

    public static string EnumToSnake(EditErrorCode code)
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

    public static string SerializeAnchorTargets(IReadOnlyList<AnchorTarget> targets)
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

    public static string SerializeAnchorTargetMap(
        IReadOnlyDictionary<string, IReadOnlyList<AnchorTarget>> map)
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

    public static void AppendAnchorTarget(StringBuilder sb, AnchorTarget t)
    {
        sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
          .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
          .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
          .Append(",\"unid\":").Append(JsonString(t.Unid))
          .Append(",\"partUri\":").Append(JsonString(t.PartUri))
          .Append(",\"textPreview\":").Append(JsonString(t.TextPreview));
        if (t.AutoNumberPrefix is { } prefix)
            sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
        sb.Append('}');
    }

    public static string SerializeAnchorTargetOrNull(AnchorTarget? target)
    {
        if (target is null) return "null";
        var sb = new StringBuilder(128);
        AppendAnchorTarget(sb, target);
        return sb.ToString();
    }

    public static string SerializeAnchorInfoMap(IReadOnlyDictionary<string, AnchorInfo?> map)
    {
        var sb = new StringBuilder(map.Count * 100 + 2);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeAnchorInfoOrNull(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeAnchorInfoOrNull(AnchorInfo? info)
    {
        if (info is null) return "null";
        var sb = new StringBuilder(128);
        sb.Append("{\"id\":").Append(JsonString(info.Id))
          .Append(",\"kind\":").Append(JsonString(info.Kind))
          .Append(",\"scope\":").Append(JsonString(info.Scope))
          .Append(",\"textPreview\":").Append(JsonString(info.TextPreview));
        if (info.AutoNumberPrefix is { } prefix)
            sb.Append(",\"autoNumberPrefix\":").Append(JsonString(prefix));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeBlockMetadataOrNull(BlockMetadata? meta)
    {
        if (meta is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"anchorId\":").Append(JsonString(meta.AnchorId))
          .Append(",\"kind\":").Append(JsonString(meta.Kind))
          .Append(",\"scope\":").Append(JsonString(meta.Scope));
        if (meta.StyleId is not null)
            sb.Append(",\"styleId\":").Append(JsonString(meta.StyleId));
        if (meta.StyleName is not null)
            sb.Append(",\"styleName\":").Append(JsonString(meta.StyleName));
        if (meta.OutlineLevel.HasValue)
            sb.Append(",\"outlineLevel\":").Append(meta.OutlineLevel.Value);
        if (meta.List is not null)
            sb.Append(",\"list\":").Append(SerializeListMembershipOrNull(meta.List));
        sb.Append(",\"hasInlineFormatting\":").Append(meta.HasInlineFormatting ? "true" : "false");
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeBlockMetadataMap(System.Collections.Generic.IReadOnlyDictionary<string, BlockMetadata?> map)
    {
        var sb = new StringBuilder(map.Count * 200 + 2);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeBlockMetadataOrNull(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeListMembershipOrNull(ListMembership? list)
    {
        if (list is null) return "null";
        var sb = new StringBuilder(128);
        sb.Append("{\"numId\":").Append(list.NumId)
          .Append(",\"abstractNumId\":").Append(list.AbstractNumId)
          .Append(",\"level\":").Append(list.Level)
          .Append(",\"format\":").Append(JsonString(NumberFormatToString(list.Format)))
          .Append(",\"isAutoNumbered\":").Append(list.IsAutoNumbered ? "true" : "false")
          .Append(",\"fromStyle\":").Append(list.FromStyle ? "true" : "false");
        if (list.StartOverride.HasValue)
            sb.Append(",\"startOverride\":").Append(list.StartOverride.Value);
        if (list.GeneratedLabel is not null)
            sb.Append(",\"generatedLabel\":").Append(JsonString(list.GeneratedLabel));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeSectionInfoOrNull(SectionInfo? info)
    {
        if (info is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"sectionUnid\":").Append(JsonString(info.SectionUnid))
          .Append(",\"pageWidthTwips\":").Append(info.PageWidthTwips)
          .Append(",\"pageHeightTwips\":").Append(info.PageHeightTwips)
          .Append(",\"landscape\":").Append(info.Landscape ? "true" : "false")
          .Append(",\"marginTopTwips\":").Append(info.MarginTopTwips)
          .Append(",\"marginBottomTwips\":").Append(info.MarginBottomTwips)
          .Append(",\"marginLeftTwips\":").Append(info.MarginLeftTwips)
          .Append(",\"marginRightTwips\":").Append(info.MarginRightTwips)
          .Append(",\"columns\":").Append(info.Columns)
          .Append(",\"headerPartUris\":[");
        for (int i = 0; i < info.HeaderPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.HeaderPartUris[i]));
        }
        sb.Append("],\"footerPartUris\":[");
        for (int i = 0; i < info.FooterPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.FooterPartUris[i]));
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string NumberFormatToString(NumberFormat f) => f switch
    {
        NumberFormat.Decimal => "decimal",
        NumberFormat.UpperLetter => "upperLetter",
        NumberFormat.LowerLetter => "lowerLetter",
        NumberFormat.UpperRoman => "upperRoman",
        NumberFormat.LowerRoman => "lowerRoman",
        NumberFormat.Bullet => "bullet",
        _ => "decimal",
    };

    public static string SerializeAnnotations(IReadOnlyList<DocumentAnnotation> anns)
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
            if (a.Metadata is { Count: > 0 })
            {
                sb.Append(",\"metadata\":{");
                bool firstMeta = true;
                foreach (var kv in a.Metadata)
                {
                    if (!firstMeta) sb.Append(',');
                    firstMeta = false;
                    sb.Append(JsonString(kv.Key)).Append(':').Append(JsonString(kv.Value ?? string.Empty));
                }
                sb.Append('}');
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    // ─── Deserializers ──────────────────────────────────────────────────

    /// <summary>
    /// Parses a camelCase annotation JSON object into a <see cref="DocumentAnnotation"/>
    /// using only <see cref="JsonDocument"/> — safe in trimmed WASM Release builds
    /// where reflection-based <see cref="JsonSerializer"/> is disabled. Both the
    /// WASM bridge and the stdio NDJSON host route annotation writes through here,
    /// so wire shape stays unified across transports.
    /// </summary>
    public static DocumentAnnotation DeserializeAnnotation(string json)
    {
        if (string.IsNullOrEmpty(json))
            throw new System.ArgumentException("annotation JSON is null or empty");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new System.ArgumentException("annotation JSON must be an object");

        var annotation = new DocumentAnnotation
        {
            Id            = TryGetString(root, "id", string.Empty) ?? string.Empty,
            LabelId       = TryGetString(root, "labelId", string.Empty) ?? string.Empty,
            Label         = TryGetString(root, "label", string.Empty) ?? string.Empty,
            Color         = TryGetString(root, "color", string.Empty) ?? string.Empty,
            Author        = TryGetString(root, "author", string.Empty) ?? string.Empty,
            BookmarkName  = TryGetString(root, "bookmarkName", string.Empty) ?? string.Empty,
            AnnotatedText = TryGetString(root, "annotatedText", string.Empty) ?? string.Empty,
            Metadata      = new System.Collections.Generic.Dictionary<string, string>(),
        };

        if (TryGetDateTime(root, "created", out var created))
            annotation.Created = created;

        if (TryGetIntNullable(root, "startPage", out var startPage))
            annotation.StartPage = startPage;

        if (TryGetIntNullable(root, "endPage", out var endPage))
            annotation.EndPage = endPage;

        // Honour explicit pageInfoStale if supplied; default already true on the model.
        if (root.TryGetProperty("pageInfoStale", out var stale) &&
            (stale.ValueKind == JsonValueKind.True || stale.ValueKind == JsonValueKind.False))
            annotation.PageInfoStale = stale.GetBoolean();

        if (TryGetDateTime(root, "pageInfoComputedAt", out var computedAt))
            annotation.PageInfoComputedAt = computedAt;

        if (root.TryGetProperty("metadata", out var meta) && meta.ValueKind == JsonValueKind.Object)
        {
            foreach (var kv in meta.EnumerateObject())
                if (kv.Value.ValueKind == JsonValueKind.String)
                    annotation.Metadata[kv.Name] = kv.Value.GetString()!;
        }

        return annotation;
    }

    /// <summary>
    /// Parses a camelCase annotation-update JSON object into an <see cref="AnnotationUpdate"/>
    /// using only <see cref="JsonDocument"/> — trim-safe under the WASM Release build.
    /// <see cref="AnnotationUpdate.MetadataPatch"/> honours explicit JSON nulls (a null
    /// value means "remove this key"), while missing fields leave the existing annotation
    /// value unchanged.
    /// </summary>
    public static AnnotationUpdate DeserializeAnnotationUpdate(string json)
    {
        if (string.IsNullOrEmpty(json))
            throw new System.ArgumentException("annotation update JSON is null or empty");

        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new System.ArgumentException("annotation update JSON must be an object");

        System.Collections.Generic.Dictionary<string, string?>? patch = null;
        if (root.TryGetProperty("metadataPatch", out var mp) && mp.ValueKind == JsonValueKind.Object)
        {
            patch = new System.Collections.Generic.Dictionary<string, string?>();
            foreach (var kv in mp.EnumerateObject())
                patch[kv.Name] = kv.Value.ValueKind == JsonValueKind.Null ? null : kv.Value.GetString();
        }

        return new AnnotationUpdate
        {
            LabelId       = TryGetString(root, "labelId", null),
            Label         = TryGetString(root, "label", null),
            Color         = TryGetString(root, "color", null),
            Author        = TryGetString(root, "author", null),
            MetadataPatch = patch,
        };
    }

    private static bool TryGetDateTime(JsonElement root, string name, out System.DateTime value)
    {
        if (root.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String
            && prop.TryGetDateTime(out var dt))
        {
            value = dt;
            return true;
        }
        value = default;
        return false;
    }

    private static bool TryGetIntNullable(JsonElement root, string name, out int? value)
    {
        if (root.TryGetProperty(name, out var prop))
        {
            if (prop.ValueKind == JsonValueKind.Number && prop.TryGetInt32(out var n))
            {
                value = n;
                return true;
            }
            if (prop.ValueKind == JsonValueKind.Null)
            {
                value = null;
                return true;
            }
        }
        value = null;
        return false;
    }
}
