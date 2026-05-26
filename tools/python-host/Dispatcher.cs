#nullable enable

using System;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.PyHost;

/// <summary>
/// Op-name → <see cref="DocxSessionOps"/> routing. Each case parses the
/// <c>args</c> JsonElement into the primitive/value-type arguments the Ops
/// facade expects and returns the JSON fragment to embed as the response
/// <c>result</c>.
///
/// Op names mirror the snake_case Python API. The argument keys on the wire
/// are camelCase to match the existing WASM bridge serialization, so the same
/// JSON shapes are interchangeable between TypeScript and Python clients —
/// the Python wrapper normalizes camelCase to snake_case on the decode side.
/// </summary>
internal static class Dispatcher
{
    public static string Dispatch(string op, JsonElement args) => op switch
    {
        "ping" => Ping(),
        "open_session" => OpenSession(args),
        "close_session" => CloseSession(args),
        "save" => Save(args),
        "project" => DocxSessionOps.Project(Handle(args)),

        "replace_text" => DocxSessionOps.ReplaceText(Handle(args), Str(args, "anchorId"), Str(args, "markdown")),
        "delete_block" => DocxSessionOps.DeleteBlock(Handle(args), Str(args, "anchorId")),
        "delete_range" => DocxSessionOps.DeleteRange(
            Handle(args), Str(args, "fromAnchorId"), Str(args, "toAnchorIdExclusive")),
        "delete_section" => DocxSessionOps.DeleteSection(
            Handle(args), Str(args, "headingAnchorId")),
        "replace_text_range" => DocxSessionOps.ReplaceTextRange(
            Handle(args), Str(args, "anchorId"), Str(args, "find"), Str(args, "replace"), ParseReplaceOptions(args)),
        "replace_text_at_span" => DocxSessionOps.ReplaceTextAtSpan(
            Handle(args), Str(args, "anchorId"), Int(args, "spanStart"), Int(args, "spanLength"), Str(args, "replace")),
        "replace_inner" => DocxSessionOps.ReplaceInner(
            Handle(args), Str(args, "matchText"), Str(args, "anchorId"),
            Int(args, "spanStart"), Int(args, "spanLength"), Str(args, "newInner")),

        "insert_paragraph" => DocxSessionOps.InsertParagraph(
            Handle(args), Str(args, "anchorId"), ParsePos(args, "position"), Str(args, "markdown")),
        "split_paragraph" => DocxSessionOps.SplitParagraph(
            Handle(args), Str(args, "anchorId"), Int(args, "characterOffset")),
        "merge_paragraphs" => DocxSessionOps.MergeParagraphs(
            Handle(args), Str(args, "firstAnchorId"), Str(args, "secondAnchorId")),

        "apply_format" => DocxSessionOps.ApplyFormat(
            Handle(args), Str(args, "anchorId"), ParseOptionalSpan(args, "span"), ParseFormatOp(args, "op")),
        "apply_format_by_substring" => DocxSessionOps.ApplyFormatBySubstring(
            Handle(args), Str(args, "anchorId"), Str(args, "substring"), ParseFormatOp(args, "op")),
        "set_paragraph_style" => DocxSessionOps.SetParagraphStyle(
            Handle(args), Str(args, "anchorId"), Str(args, "styleId")),
        "set_list_level" => DocxSessionOps.SetListLevel(
            Handle(args), Str(args, "anchorId"), Int(args, "levelDelta")),
        "remove_list_membership" => DocxSessionOps.RemoveListMembership(
            Handle(args), Str(args, "anchorId")),

        "replace_cell_content" => DocxSessionOps.ReplaceCellContent(
            Handle(args), Str(args, "cellAnchorId"), Str(args, "markdown")),

        "raw_get_xml" => JsonString(DocxSessionOps.RawGetXml(Handle(args), Str(args, "anchorId"))),
        "raw_insert_xml" => DocxSessionOps.RawInsertXml(
            Handle(args), Str(args, "anchorId"), ParsePos(args, "position"), Str(args, "xml")),
        "raw_replace_xml" => DocxSessionOps.RawReplaceXml(
            Handle(args), Str(args, "anchorId"), Str(args, "xml")),

        "grep" => Grep(args, crossBlock: false),
        "grep_cross_block" => Grep(args, crossBlock: true),

        "find_placeholders" => DocxSessionOps.FindPlaceholders(
            Handle(args),
            (PlaceholderKinds)IntOptional(args, "kinds", (int)PlaceholderKinds.All),
            (ProjectionScopes)IntOptional(args, "scope", (int)ProjectionScopes.Body),
            IntOptional(args, "contextChars", 80),
            (ContextBoundary)IntOptional(args, "boundary", (int)ContextBoundary.Char)),
        "find_by_annotation" => DocxSessionOps.FindByAnnotation(Handle(args), Str(args, "annotationId")),
        "find_by_label" => DocxSessionOps.FindByLabel(Handle(args), Str(args, "labelId")),
        "find_by_bookmark" => DocxSessionOps.FindByBookmark(Handle(args), Str(args, "bookmarkName")),
        "list_annotations" => DocxSessionOps.ListAnnotations(Handle(args)),

        "exists" => DocxSessionOps.Exists(Handle(args), Str(args, "anchorId")) ? "true" : "false",
        "get_anchor_info" => DocxSessionOps.GetAnchorInfo(Handle(args), Str(args, "anchorId")),
        "get_anchor_infos" => DocxSessionOps.GetAnchorInfos(Handle(args), ParseAnchorIdArray(args)),
        "find_by_text" => DocxSessionOps.FindByText(Handle(args), Str(args, "needle"), ParseFindOptions(args)),
        "find_all_by_text" => DocxSessionOps.FindAllByText(Handle(args), Str(args, "needle"), ParseFindOptions(args)),
        "find_by_regex" => DocxSessionOps.FindByRegex(
            Handle(args), Str(args, "pattern"),
            (RegexOptions)IntOptional(args, "regexOptions", 0),
            ParseFindOptions(args)),
        "find_by_kind" => DocxSessionOps.FindByKind(
            Handle(args), Str(args, "kind"),
            args.ValueKind == JsonValueKind.Object && args.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.String
                ? sc.GetString() : null),

        "undo" => DocxSessionOps.Undo(Handle(args)) ? "true" : "false",
        "redo" => DocxSessionOps.Redo(Handle(args)) ? "true" : "false",

        _ => throw new UnknownOpException(op),
    };

    private static string Ping()
    {
        var version = typeof(DocxSession).Assembly.GetName().Version?.ToString() ?? "0.0.0";
        var sb = new System.Text.StringBuilder(96);
        sb.Append("{\"pong\":true,\"version\":");
        sb.Append(DocxSessionJson.JsonString(version));
        sb.Append(",\"dotnet\":");
        sb.Append(DocxSessionJson.JsonString(Environment.Version.ToString()));
        sb.Append(",\"sessions\":");
        sb.Append(SessionRegistry.Count);
        sb.Append('}');
        return sb.ToString();
    }

    private static string OpenSession(JsonElement args)
    {
        var b64 = Str(args, "docxB64");
        var bytes = Convert.FromBase64String(b64);
        DocxSessionSettings? settings = null;
        if (args.TryGetProperty("settings", out var s) && s.ValueKind == JsonValueKind.Object)
        {
            settings = DocxSessionJson.ParseSettings(s.GetRawText());
        }
        var handle = DocxSessionOps.OpenSession(bytes, settings);
        return handle.ToString(CultureInfo.InvariantCulture);
    }

    private static string CloseSession(JsonElement args)
    {
        DocxSessionOps.CloseSession(Handle(args));
        return "null";
    }

    private static string Save(JsonElement args)
    {
        var bytes = DocxSessionOps.Save(Handle(args));
        return "{\"docxB64\":" + DocxSessionJson.JsonString(Convert.ToBase64String(bytes)) + "}";
    }

    private static string Grep(JsonElement args, bool crossBlock)
    {
        var pattern = Str(args, "pattern");
        var regexOpts = (RegexOptions)IntOptional(args, "regexOptions", 0);
        var scope = (ProjectionScopes)IntOptional(args, "scope", (int)ProjectionScopes.Body);
        var contextChars = IntOptional(args, "contextChars", 80);
        var whitespace = (WhitespaceMode)IntOptional(args, "whitespace", (int)WhitespaceMode.Preserve);
        var boundary = (ContextBoundary)IntOptional(args, "boundary", (int)ContextBoundary.Char);
        return crossBlock
            ? DocxSessionOps.GrepCrossBlock(Handle(args), pattern, regexOpts, scope, contextChars, whitespace, boundary)
            : DocxSessionOps.Grep(Handle(args), pattern, regexOpts, scope, contextChars, whitespace, boundary);
    }

    // ─── Arg helpers ────────────────────────────────────────────────────

    private static int Handle(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("handle", out var h) || h.ValueKind != JsonValueKind.Number)
            throw new FormatException("args missing numeric \"handle\"");
        return h.GetInt32();
    }

    private static string Str(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.String)
            throw new FormatException($"args missing string \"{name}\"");
        return v.GetString()!;
    }

    private static int Int(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Number)
            throw new FormatException($"args missing number \"{name}\"");
        return v.GetInt32();
    }

    private static int IntOptional(JsonElement args, string name, int fallback)
    {
        if (args.ValueKind != JsonValueKind.Object) return fallback;
        return args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : fallback;
    }

    private static Position ParsePos(JsonElement args, string name) =>
        DocxSessionJson.ParsePos(Str(args, name));

    private static CharSpan? ParseOptionalSpan(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object) return null;
        if (!args.TryGetProperty(name, out var s) || s.ValueKind != JsonValueKind.Object) return null;
        return new CharSpan(
            s.TryGetProperty("start", out var st) && st.ValueKind == JsonValueKind.Number ? st.GetInt32() : 0,
            s.TryGetProperty("length", out var ln) && ln.ValueKind == JsonValueKind.Number ? ln.GetInt32() : 0);
    }

    private static FormatOp ParseFormatOp(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var op) || op.ValueKind != JsonValueKind.Object)
            return new FormatOp();
        return DocxSessionJson.ParseFormatOp(op.GetRawText());
    }

    private static string[] ParseAnchorIdArray(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("anchorIds", out var arr) || arr.ValueKind != JsonValueKind.Array)
            throw new FormatException("args missing array \"anchorIds\"");
        var result = new string[arr.GetArrayLength()];
        int i = 0;
        foreach (var el in arr.EnumerateArray())
        {
            if (el.ValueKind != JsonValueKind.String)
                throw new FormatException("\"anchorIds\" entries must be strings");
            result[i++] = el.GetString()!;
        }
        return result;
    }

    private static FindOptions? ParseFindOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("options", out var o) || o.ValueKind != JsonValueKind.Object)
            return null;
        return DocxSessionJson.ParseFindOptions(o);
    }

    private static ReplaceOptions? ParseReplaceOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("options", out var o) || o.ValueKind != JsonValueKind.Object)
            return null;
        return new ReplaceOptions
        {
            IgnoreCase = DocxSessionJson.TryGetBool(o, "ignoreCase", false),
            MaxReplacements = o.TryGetProperty("maxReplacements", out var mr) && mr.ValueKind == JsonValueKind.Number
                ? mr.GetInt32() : (int?)null,
        };
    }

    private static string JsonString(string s) => DocxSessionJson.JsonString(s);
}
