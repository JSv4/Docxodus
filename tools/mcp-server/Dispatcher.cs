#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>
/// Tool name (+ <c>action</c> discriminator, for the grouped-intent tools) → Docxodus API
/// routing. Every method here parses the <c>arguments</c> object of an MCP <c>tools/call</c>
/// request and returns the JSON fragment to embed as the tool result's text content — mirrors
/// <c>tools/python-host/Dispatcher.cs</c>'s op-routing style, just keyed by (tool, action)
/// instead of a flat op-name string, because the tool surface here groups many Docxodus ops
/// under one MCP tool per functional category (see docs/architecture/docx_agent_server.md).
///
/// Argument problems, unknown sessions, and underlying exceptions all become
/// <see cref="McpToolException"/> here — <c>Program.cs</c> catches them uniformly and reports
/// them as an MCP tool result with <c>isError: true</c> rather than a JSON-RPC protocol error,
/// per the MCP convention that business-level tool failures belong in the result, not the
/// envelope.
/// </summary>
internal static class Dispatcher
{
    public static string Call(SessionStore store, string tool, JsonElement args) => tool switch
    {
        "docxodus_open" => Open(store, args),
        "docxodus_save" => Save(store, args),
        "docxodus_close" => Close(store, args),
        "docxodus_get_content" => GetContent(store, args),
        "docxodus_preview" => Preview(store, args),
        "docxodus_pagination" => Pagination(store, args),
        "docxodus_search" => Search(store, args),
        "docxodus_edit" => Edit(store, args),
        "docxodus_format" => Format(store, args),
        "docxodus_create" => Create(store, args),
        "docxodus_list" => ListTool(store, args),
        "docxodus_comment" => Comment(store, args),
        "docxodus_links" => Links(store, args),
        "docxodus_images" => Images(store, args),
        "docxodus_annotate" => Annotate(store, args),
        "docxodus_track_changes" => TrackChanges(store, args),
        "docxodus_mutations" => Mutations(store, args),
        "docxodus_table" => Table(store, args),
        _ => throw new McpToolException($"unknown tool: {tool}"),
    };

    // ─── Lifecycle ──────────────────────────────────────────────────────

    private static string Open(SessionStore store, JsonElement args)
    {
        // Resolve THROUGH the store: this is where an out-of-scope location is rejected, and the
        // canonical form it returns is what the session records for a later write-back.
        var location = store.Documents.Resolve(Str(args, "path"));
        var bytes = store.Documents.Read(location);

        var tracked = OptStr(args, "trackedChanges") switch
        {
            "render_inline" => TrackedChangeMode.RenderInline,
            "strip_deletions" => TrackedChangeMode.StripDeletions,
            _ => TrackedChangeMode.Accept,
        };
        // Defaults come from the settings object, not repeated literals, so this surface cannot
        // drift from the .NET default the way the hardcoded undoDepth of 50 had.
        var settingDefaults = new DocxSessionSettings();
        var settings = new DocxSessionSettings
        {
            TrackedChanges = tracked,
            RevisionAuthor = OptStr(args, "revisionAuthor"),
            UndoDepth = IntOpt(args, "undoDepth", settingDefaults.UndoDepth),
            UndoMemoryBudgetBytes = LongOpt(
                args, "undoMemoryBudgetBytes", settingDefaults.UndoMemoryBudgetBytes),
            PersistAnchorIds = BoolOpt(args, "persistAnchorIds", false),
        };

        var session = store.Open(bytes, location, settings);
        return $"{{\"sessionId\":{JsonRpcIo.JsonString(session.Id)},\"path\":{JsonRpcIo.JsonString(location)}}}";
    }

    private static string Save(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        // An explicit destination is re-resolved (so it is scope-checked like any other location);
        // the recorded one already is.
        var destination = OptStr(args, "path") is { } requested
            ? store.Documents.Resolve(requested)
            : session.Location
              ?? throw new McpToolException("session was not opened from a location; pass \"path\" explicitly");

        // Tri-state: absent → the session's open-time PersistAnchorIds; explicit true/false
        // overrides it for this save only (true = anchor-stable checkpoint, false = clean
        // deliverable from a session that was opened anchor-stable).
        var bytes = OptBool(args, "persistAnchorIds") is { } persist
            ? DocxSessionOps.Save(session.Handle, persist)
            : DocxSessionOps.Save(session.Handle);
        store.Documents.Write(destination, bytes);
        return $"{{\"path\":{JsonRpcIo.JsonString(destination)},\"bytesWritten\":{bytes.Length}}}";
    }

    private static string Close(SessionStore store, JsonElement args)
    {
        store.Close(Str(args, "sessionId"));
        return "{\"closed\":true}";
    }

    // ─── Content ────────────────────────────────────────────────────────

    private static string GetContent(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var format = Str(args, "format");
        var anchorId = OptStr(args, "anchorId");
        var citation = DocxSessionJson.ParsePageCitationRequest(args);

        switch (format)
        {
            case "markdown":
                return anchorId is null
                    ? DocxSessionOps.Project(session.Handle)
                    : DocxSessionOps.ProjectAnchor(session.Handle, anchorId,
                        ProjectionDepth.SubtreeAndFollowingSiblings, citation);

            case "text":
            {
                var projectionJson = anchorId is null
                    ? DocxSessionOps.Project(session.Handle)
                    : DocxSessionOps.ProjectAnchor(session.Handle, anchorId,
                        ProjectionDepth.SubtreeAndFollowingSiblings, citation);
                using var doc = JsonDocument.Parse(projectionJson);
                var markdown = doc.RootElement.GetProperty("markdown").GetString() ?? string.Empty;
                return $"{{\"text\":{JsonRpcIo.JsonString(StripMarkdownSyntax(markdown))}}}";
            }

            case "html":
                return $"{{\"html\":{JsonRpcIo.JsonString(
                    anchorId is null
                        ? DocxSessionOps.RenderHtml(session.Handle, "docx-", false, false, 1.0,
                            renderTrackedChanges: true)
                        : DocxSessionOps.RenderBlockHtml(session.Handle, anchorId, "docx-", false,
                            renderTrackedChanges: true))}}}";

            case "blocks":
            {
                var projectionJson = DocxSessionOps.Project(session.Handle);
                using var doc = JsonDocument.Parse(projectionJson);
                var ids = new List<string>();
                foreach (var prop in doc.RootElement.GetProperty("anchorIndex").EnumerateObject())
                    ids.Add(prop.Name);
                return $"{{\"blocks\":{DocxSessionOps.GetBlockMetadatas(session.Handle, ids)}}}";
            }

            case "info":
            {
                var editSummary = DocxSessionOps.GetEditSummary(session.Handle);
                string sectionInfo = "null";
                if (anchorId is not null)
                {
                    sectionInfo = DocxSessionOps.GetSectionInfo(session.Handle, anchorId);
                }
                else
                {
                    var projectionJson = DocxSessionOps.Project(session.Handle);
                    using var doc = JsonDocument.Parse(projectionJson);
                    foreach (var prop in doc.RootElement.GetProperty("anchorIndex").EnumerateObject())
                    {
                        var kind = prop.Value.GetProperty("kind").GetString();
                        var scope = prop.Value.GetProperty("scope").GetString();
                        if (scope != "body" || kind is not ("p" or "h" or "li")) continue;
                        sectionInfo = DocxSessionOps.GetSectionInfo(session.Handle, prop.Name);
                        break;
                    }
                }
                return $"{{\"version\":{DocxSessionOps.GetVersion(session.Handle)},\"editSummary\":{editSummary},\"sectionInfo\":{sectionInfo}}}";
            }

            case "version":
                return DocxSessionOps.GetVersionJson(session.Handle);

            case "check_preconditions":
                return DocxSessionOps.CheckPreconditions(
                    session.Handle, ParsePreconditions(args, OptStr(args, "anchorId")));

            case "styles":
                return $"{{\"styles\":{DocxSessionOps.ListStyles(session.Handle)}}}";

            case "formatting":
                if (anchorId is null)
                    throw new McpToolException("formatting requires anchorId");
                return $"{{\"formatting\":{DocxSessionOps.GetFormatting(session.Handle, anchorId)}}}";

            case "spans":
                if (anchorId is null)
                    throw new McpToolException("spans requires anchorId");
                return $"{{\"spans\":{DocxSessionOps.ListInlineSpans(session.Handle, anchorId)}}}";

            default:
                throw new McpToolException($"unknown format: {format}");
        }
    }

    /// <summary>Render for the inline preview widget (MCP Apps / ChatGPT). Same converter
    /// profile as get_content's html format; the html field is lifted out of the model-visible
    /// result by <see cref="UiResources.WrapToolResult"/>, so its size here is harmless.</summary>
    private static string Preview(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var anchorId = OptStr(args, "anchorId");
        var citationRequest = DocxSessionJson.ParsePageCitationRequest(args);
        var citationJson = anchorId is not null && citationRequest is not null
            ? DocxSessionOps.GetPageCitation(session.Handle, anchorId, citationRequest)
            : null;
        var hasPhysicalCitation = false;
        if (citationJson is not null)
        {
            using var citationDoc = JsonDocument.Parse(citationJson);
            hasPhysicalCitation = citationDoc.RootElement.GetProperty("availability").GetString()
                == "available";
        }
        // A cited preview carries paginated converter staging so the browser widget can project
        // the exact registered page geometry without inventing layout. Ordinary preview remains
        // the established lightweight continuous/block render.
        var html = hasPhysicalCitation
            ? DocxSessionOps.RenderHtml(session.Handle, "docx-", false, true, 1.0)
            : anchorId is null
                ? DocxSessionOps.RenderHtml(session.Handle, "docx-", false, false, 1.0)
                : DocxSessionOps.RenderBlockHtml(session.Handle, anchorId, "docx-", false);
        return $"{{\"sessionId\":{JsonRpcIo.JsonString(session.Id)}"
            + (anchorId is null ? "" : $",\"anchorId\":{JsonRpcIo.JsonString(anchorId)}")
            + (citationJson is null ? "" : $",\"citation\":{citationJson}")
            + (hasPhysicalCitation
                ? ",\"pageNavigation\":\"available_registered_map\""
                : ",\"pageNavigation\":\"unavailable_continuous_preview\"")
            + $",\"html\":{JsonRpcIo.JsonString(html)}}}";
    }

    private static string Pagination(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return Str(args, "action") switch
        {
            "register" => DocxSessionOps.RegisterPageMap(
                session.Handle,
                DocxSessionJson.ParsePageMap(Object(args, "pageMap")),
                OptStr(args, "expectedRendererFingerprint")),
            "status" => DocxSessionOps.GetPageMapStatus(
                session.Handle, DocxSessionJson.ParsePageCitationRequest(args)),
            "cite" => DocxSessionOps.GetPageCitation(
                session.Handle, Str(args, "anchorId"),
                DocxSessionJson.ParsePageCitationRequest(args)
                    ?? throw new McpToolException("cite requires citation {documentVersion, rendererFingerprint}")),
            _ => throw new McpToolException("unknown pagination action"),
        };
    }

    /// <summary>
    /// Best-effort plain-text approximation of a markdown projection: unescapes the projector's
    /// backslash-escaped punctuation, then strips ATX heading markers, emphasis/strike/code
    /// delimiters, and link syntax. Not a full markdown parser — good enough for an agent that
    /// wants prose without markup, not for round-tripping. Use format "markdown" for anything
    /// that needs to survive a write-back.
    /// </summary>
    private static string StripMarkdownSyntax(string markdown)
    {
        var s = Regex.Replace(markdown, @"\\([-*_`~\\\[\]()#+.!])", "$1");
        s = Regex.Replace(s, @"(?m)^(\s*)#{1,6}\s+", "$1");
        s = Regex.Replace(s, @"(?m)^(\s*)[-*]\s+", "$1");
        s = Regex.Replace(s, @"(?m)^(\s*)\d+\.\s+", "$1");
        s = Regex.Replace(s, @"\*\*(.+?)\*\*", "$1");
        s = Regex.Replace(s, @"~~(.+?)~~", "$1");
        s = Regex.Replace(s, @"`(.+?)`", "$1");
        s = Regex.Replace(s, @"\*(.+?)\*", "$1");
        s = Regex.Replace(s, @"\[([^\]]+)\]\([^)]+\)", "$1");
        return s;
    }

    // ─── Search ─────────────────────────────────────────────────────────

    private static string Search(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var mode = Str(args, "mode");
        var query = Str(args, "query");
        var caseSensitive = BoolOpt(args, "caseSensitive", false);
        var contextChars = IntOpt(args, "contextChars", 80);
        var scope = ParseSearchScope(OptStr(args, "scope"));
        var maxResults = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("maxResults", out var mr) && mr.ValueKind == JsonValueKind.Number
            ? mr.GetInt32() : (int?)null;
        var citation = DocxSessionJson.ParsePageCitationRequest(args);

        string matchesJson = mode switch
        {
            "text" => DocxSessionOps.Grep(
                session.Handle, Regex.Escape(query), caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase,
                scope, contextChars, WhitespaceMode.Preserve, ContextBoundary.Char, citation),
            "regex" => DocxSessionOps.Grep(
                session.Handle, query, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase,
                scope, contextChars, WhitespaceMode.Preserve, ContextBoundary.Char, citation),
            "kind" => DocxSessionOps.FindByKind(session.Handle, query, null, citation),
            "annotation" => DocxSessionOps.FindByAnnotation(session.Handle, query, citation),
            "bookmark" => DocxSessionOps.FindByBookmark(session.Handle, query, citation),
            _ => throw new McpToolException($"unknown search mode: {mode}"),
        };

        if (maxResults is null) return $"{{\"matches\":{matchesJson}}}";

        using var doc = JsonDocument.Parse(matchesJson);
        var items = new List<string>();
        foreach (var el in doc.RootElement.EnumerateArray())
        {
            if (items.Count >= maxResults) break;
            items.Add(el.GetRawText());
        }
        return "{\"matches\":[" + string.Join(",", items) + "]}";
    }

    /// <summary>Translate the MCP search vocabulary onto the engine's flag set. Omitted scope
    /// intentionally remains body-only: widening the historical default would make an existing
    /// query start returning repeated running-story text from every header/footer part.</summary>
    private static ProjectionScopes ParseSearchScope(string? scope) => scope switch
    {
        null or "body" => ProjectionScopes.Body,
        "headers" => ProjectionScopes.Headers,
        "footers" => ProjectionScopes.Footers,
        "header_footer" => ProjectionScopes.Headers | ProjectionScopes.Footers,
        "all" => ProjectionScopes.All,
        _ => throw new McpToolException($"unknown search scope: {scope}"),
    };

    // ─── Edit ───────────────────────────────────────────────────────────

    private static string Edit(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunEditAction(session, Str(args, "action"), args);
    }

    /// <summary>Shared by <see cref="Edit"/> and <see cref="Mutations"/> (batched steps route
    /// through the same per-tool functions so there's exactly one place each action's argument
    /// parsing lives).</summary>
    private static string RunEditAction(DocSession session, string action, JsonElement args)
    {
        var preconditions = ParsePreconditions(args, MutationTarget(args));
        if (action == "undo")
            return preconditions is null
                ? BoolResult(DocxSessionOps.Undo(session.Handle))
                : DocxSessionOps.UndoChecked(session.Handle, preconditions);
        if (action == "redo")
            return preconditions is null
                ? BoolResult(DocxSessionOps.Redo(session.Handle))
                : DocxSessionOps.RedoChecked(session.Handle, preconditions);
        return Guarded(session, preconditions, () => action switch
        {
        "insert_paragraph" => DocxSessionOps.InsertParagraph(
            session.Handle, Str(args, "anchorId"), ParsePos(args), Str(args, "markdown")),
        "replace_text" => DocxSessionOps.ReplaceText(session.Handle, Str(args, "anchorId"), Str(args, "markdown")),
        "replace_text_range" => DocxSessionOps.ReplaceTextRange(
            session.Handle, Str(args, "anchorId"), Str(args, "find"), Str(args, "replace"),
            new ReplaceOptions { IgnoreCase = !BoolOpt(args, "caseSensitive", false) }, preconditions),
        "delete_block" => DocxSessionOps.DeleteBlock(session.Handle, Str(args, "anchorId")),
        "move_block" => DocxSessionOps.MoveBlock(
            session.Handle, Str(args, "sourceAnchorId"), Str(args, "targetAnchorId"), ParsePos(args)),
        "delete_range" => DocxSessionOps.DeleteRange(
            session.Handle, Str(args, "fromAnchorId"), Str(args, "toAnchorIdExclusive")),
        "delete_section" => DocxSessionOps.DeleteSection(session.Handle, Str(args, "headingAnchorId")),
        "split_paragraph" => DocxSessionOps.SplitParagraph(
            session.Handle, Str(args, "anchorId"), Int(args, "characterOffset")),
        "merge_paragraphs" => DocxSessionOps.MergeParagraphs(
            session.Handle, Str(args, "anchorId"), Str(args, "secondAnchorId")),
        _ => throw new McpToolException($"unknown docxodus_edit action: {action}"),
        });
    }

    private static bool IsMutatingEditAction(string action) => action is not ("undo" or "redo");

    private static string BoolResult(bool value) => value ? "{\"success\":true}" : "{\"success\":false}";

    // ─── Format ─────────────────────────────────────────────────────────

    private static string Format(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunFormatAction(session, Str(args, "action"), args);
    }

    private static string RunFormatAction(DocSession session, string action, JsonElement args) =>
        Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
        "apply_format" => DocxSessionOps.ApplyFormat(
            session.Handle, Str(args, "anchorId"), ParseSpan(args, "span"), ParseFormatOp(args)),
        "apply_format_by_substring" => DocxSessionOps.ApplyFormatBySubstring(
            session.Handle, Str(args, "anchorId"), Str(args, "substring"), ParseFormatOp(args)),
        "set_paragraph_style" => DocxSessionOps.SetParagraphStyle(
            session.Handle, Str(args, "anchorId"), Str(args, "styleId")),
        "set_paragraph_format" => DocxSessionOps.SetParagraphFormat(
            session.Handle, Str(args, "anchorId"), ParseParagraphFormatOp(args)),
        "set_list_level" => DocxSessionOps.SetListLevel(
            session.Handle, Str(args, "anchorId"), Int(args, "levelDelta")),
        "remove_list_membership" => DocxSessionOps.RemoveListMembership(session.Handle, Str(args, "anchorId")),
        "apply_list_format" => DocxSessionOps.ApplyListFormat(
            session.Handle, Str(args, "anchorId"), DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        _ => throw new McpToolException($"unknown docxodus_format action: {action}"),
        });

    private static FormatOp ParseFormatOp(JsonElement args) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty("format", out var f) && f.ValueKind == JsonValueKind.Object
            ? DocxSessionJson.ParseFormatOp(f.GetRawText())
            : new FormatOp();

    private static ParagraphFormatOp ParseParagraphFormatOp(JsonElement args) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty("paragraphFormat", out var f) && f.ValueKind == JsonValueKind.Object
            ? DocxSessionJson.ParseParagraphFormatOp(f.GetRawText())
            : new ParagraphFormatOp();

    // ─── Create ─────────────────────────────────────────────────────────

    private static string Create(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunCreateAction(session, Str(args, "action"), args);
    }

    private static string RunCreateAction(DocSession session, string action, JsonElement args) =>
        Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
        "insert_paragraph" => DocxSessionOps.InsertParagraph(
            session.Handle, Str(args, "anchorId"), ParsePos(args), Str(args, "markdown")),
        "insert_heading" => DocxSessionOps.InsertParagraph(
            session.Handle, Str(args, "anchorId"), ParsePos(args),
            new string('#', Math.Clamp(Int(args, "level"), 1, 6)) + " " + Str(args, "text")),
        "insert_table" => DocxSessionOps.InsertTable(
            session.Handle, Str(args, "anchorId"), ParsePos(args),
            Int(args, "rows"), Int(args, "columns"), BuildTableInsertOptionsJson(args)),
        "insert_horizontal_rule" => DocxSessionOps.InsertHorizontalRule(
            session.Handle, Str(args, "anchorId"), ParsePos(args), BuildRuleEdgeJson(args)),
        "insert_footnote" => DocxSessionOps.InsertFootnote(
            session.Handle, Str(args, "anchorId"), Int(args, "characterOffset"), Str(args, "markdown")),
        "insert_endnote" => DocxSessionOps.InsertEndnote(
            session.Handle, Str(args, "anchorId"), Int(args, "characterOffset"), Str(args, "markdown")),
        "insert_page_number_field" => DocxSessionOps.InsertPageNumberField(
            session.Handle, Str(args, "anchorId"),
            OptStr(args, "field") == "total_pages" ? PageNumberField.TotalPages : PageNumberField.CurrentPage,
            DocxSessionJson.ParseNumberFormatOrNull(OptStr(args, "numberFormat"))),
        "set_header_text" => DocxSessionOps.SetHeaderText(
            session.Handle, Str(args, "bodyAnchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind")), Str(args, "markdown")),
        "set_footer_text" => DocxSessionOps.SetFooterText(
            session.Handle, Str(args, "bodyAnchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind")), Str(args, "markdown")),
        "ensure_header_footer_visible" => DocxSessionOps.EnsureHeaderFooterVisible(
            session.Handle, Str(args, "bodyAnchorId"),
            DocxSessionJson.ParseHeaderFooterKind(Str(args, "kind"))),
        _ => throw new McpToolException($"unknown docxodus_create action: {action}"),
        });

    private static string BuildTableInsertOptionsJson(JsonElement args)
    {
        var opts = new Dictionary<string, object?>
        {
            ["borderless"] = BoolOpt(args, "borderless", false),
            ["cellAlignment"] = OptStr(args, "cellAlignment"),
        };
        if (args.ValueKind == JsonValueKind.Object && args.TryGetProperty("cellContents", out var cc) && cc.ValueKind == JsonValueKind.Array)
            opts["cellContents"] = cc;
        if (args.ValueKind == JsonValueKind.Object && args.TryGetProperty("columnWidths", out var cw) && cw.ValueKind == JsonValueKind.Array)
            opts["columnWidths"] = cw;
        return JsonSerializer.Serialize(opts);
    }

    private static string BuildRuleEdgeJson(JsonElement args)
    {
        var style = OptStr(args, "ruleStyle");
        return style is null ? "" : JsonSerializer.Serialize(new { style });
    }

    // ─── List ───────────────────────────────────────────────────────────

    private static string ListTool(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunListAction(session, Str(args, "action"), args);
    }

    private static string RunListAction(DocSession session, string action, JsonElement args)
    {
        if (action == "get_membership")
            return DocxSessionOps.GetListMembership(session.Handle, Str(args, "anchorId"));
        return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
        "apply_format" => DocxSessionOps.ApplyListFormat(
            session.Handle, Str(args, "anchorId"), DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        "apply_format_range" => DocxSessionOps.ApplyListFormatRange(
            session.Handle, Str(args, "firstAnchorId"), Str(args, "lastAnchorId"),
            DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        "set_level" => DocxSessionOps.SetListLevel(session.Handle, Str(args, "anchorId"), Int(args, "levelDelta")),
        "set_start" => DocxSessionOps.SetListStartOverride(
            session.Handle, Str(args, "anchorId"), Int(args, "startValue")),
        "clear_start" => DocxSessionOps.ClearListStartOverride(session.Handle, Str(args, "anchorId")),
        "remove" => DocxSessionOps.RemoveListMembership(session.Handle, Str(args, "anchorId")),
        _ => throw new McpToolException($"unknown docxodus_list action: {action}"),
        });
    }

    private static bool IsMutatingListAction(string action) => action != "get_membership";

    // ─── Comment (native Word comments, issue #300) ────────────────────

    private static string Comment(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunCommentAction(session, Str(args, "action"), args);
    }

    private static string RunCommentAction(DocSession session, string action, JsonElement args)
    {
        if (action == "list")
            return $"{{\"comments\":{DocxSessionOps.ListComments(session.Handle)}}}";
        return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
        "add" => AddComment(session, args),
        "reply" => DocxSessionOps.AddCommentReply(
            session.Handle, Str(args, "commentAnchorId"), Str(args, "author"),
            OptStr(args, "initials"), OptStr(args, "date"), OptStr(args, "markdown") ?? ""),
        "update" => DocxSessionOps.UpdateComment(
            session.Handle, Str(args, "commentAnchorId"), Str(args, "markdown")),
        "resolve" => DocxSessionOps.SetCommentResolved(
            session.Handle, Str(args, "commentAnchorId"), BoolOpt(args, "resolved", true)),
        "remove" => DocxSessionOps.RemoveComment(session.Handle, Str(args, "commentAnchorId")),
        _ => throw new McpToolException($"unknown docxodus_comment action: {action}"),
        });
    }

    private static bool IsMutatingCommentAction(string action) => action != "list";

    // ─── Native hyperlinks / bookmarks (issue #451) ───────────────────

    private static string Links(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunLinksAction(session, Str(args, "action"), args);
    }

    private static string RunLinksAction(DocSession session, string action, JsonElement args) => action switch
    {
        "list_hyperlinks" => $"{{\"hyperlinks\":{DocxSessionOps.ListHyperlinks(session.Handle, ParseLinkScopes(OptStr(args, "scope")))}}}",
        "add_hyperlink" => DocxSessionOps.AddHyperlink(session.Handle, Str(args, "anchorId"),
            Int(args, "startOffset"), Int(args, "length"), Str(args, "kind"), Str(args, "target")),
        "update_hyperlink" => DocxSessionOps.UpdateHyperlink(session.Handle,
            Str(args, "hyperlinkId"), Str(args, "kind"), Str(args, "target")),
        "remove_hyperlink" => DocxSessionOps.RemoveHyperlink(session.Handle, Str(args, "hyperlinkId")),
        "list_bookmarks" => $"{{\"bookmarks\":{DocxSessionOps.ListBookmarks(session.Handle, ParseLinkScopes(OptStr(args, "scope")))}}}",
        "add_bookmark" => BookmarkRangeAction(session, args, move: false),
        "move_bookmark" => BookmarkRangeAction(session, args, move: true),
        "rename_bookmark" => DocxSessionOps.RenameBookmark(session.Handle, Str(args, "name"), Str(args, "newName")),
        "remove_bookmark" => DocxSessionOps.RemoveBookmark(session.Handle, Str(args, "name")),
        _ => throw new McpToolException($"unknown docxodus_links action: {action}"),
    };

    private static string BookmarkRangeAction(DocSession session, JsonElement args, bool move) =>
        move
            ? DocxSessionOps.MoveBookmark(session.Handle, Str(args, "name"),
                Str(args, "startAnchorId"), Int(args, "startOffset"),
                Str(args, "endAnchorId"), Int(args, "endOffset"))
            : DocxSessionOps.AddBookmark(session.Handle, Str(args, "name"),
                Str(args, "startAnchorId"), Int(args, "startOffset"),
                Str(args, "endAnchorId"), Int(args, "endOffset"));

    private static ProjectionScopes ParseLinkScopes(string? scope) => scope switch
    {
        null or "all" => ProjectionScopes.All,
        "body" => ProjectionScopes.Body,
        "headers" => ProjectionScopes.Headers,
        "footers" => ProjectionScopes.Footers,
        "footnotes" => ProjectionScopes.Footnotes,
        "endnotes" => ProjectionScopes.Endnotes,
        "comments" => ProjectionScopes.Comments,
        _ => throw new McpToolException($"unknown link scope: {scope}"),
    };

    private static bool IsMutatingLinksAction(string action) =>
        action is not ("list_hyperlinks" or "list_bookmarks");

    // ─── Native images (issue #453) ───────────────────────────────────

    private static string Images(SessionStore store, JsonElement args)
    {
        var action = Str(args, "action");
        if (action == "capabilities")
            return $"{{\"capabilities\":{DocxSessionOps.GetImageCapabilities()}}}";
        var session = Session(store, args);
        return RunImagesAction(session, action, args);
    }

    private static string RunImagesAction(DocSession session, string action, JsonElement args) =>
        action switch
        {
            "list" => $"{{\"images\":{DocxSessionOps.ListImages(session.Handle,
                ParseLinkScopes(OptStr(args, "scope")))}}}",
            "insert" => DocxSessionOps.InsertImage(session.Handle, Str(args, "anchorId"),
                Int(args, "characterOffset"), Str(args, "imageBase64"), RawObjectOrEmpty(args, "options")),
            "replace" => DocxSessionOps.ReplaceImage(session.Handle,
                Str(args, "imageId"), Str(args, "imageBase64")),
            "set_dimensions" => DocxSessionOps.SetImageDimensions(session.Handle,
                Str(args, "imageId"), RawObject(args, "dimensions")),
            "set_metadata" => SetImageMetadata(session, args),
            "set_floating_layout" => DocxSessionOps.SetImageFloatingLayout(session.Handle,
                Str(args, "imageId"), RawObject(args, "layout")),
            "remove" => DocxSessionOps.RemoveImage(session.Handle, Str(args, "imageId")),
            _ => throw new McpToolException($"unknown docxodus_images action: {action}"),
        };

    private static bool IsMutatingImagesAction(string action) =>
        action is not ("capabilities" or "list");

    private static string SetImageMetadata(DocSession session, JsonElement args)
    {
        if (!args.TryGetProperty("altText", out var alt) || alt.ValueKind is not (JsonValueKind.String or JsonValueKind.Null)
            || !args.TryGetProperty("title", out var title) || title.ValueKind is not (JsonValueKind.String or JsonValueKind.Null))
            throw new McpToolException("docxodus_images set_metadata requires altText and title as string or null");
        return DocxSessionOps.SetImageMetadata(session.Handle, Str(args, "imageId"),
            alt.ValueKind == JsonValueKind.Null ? null : alt.GetString(),
            title.ValueKind == JsonValueKind.Null ? null : title.GetString());
    }

    private static string AddComment(DocSession session, JsonElement args)
    {
        var anchorId = OptStr(args, "anchorId");
        var revisionId = OptStr(args, "revisionId");
        var hasSpan = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("span", out _);
        if ((anchorId is null) == (revisionId is null) || (revisionId is not null && hasSpan))
            throw new McpToolException(
                "docxodus_comment add requires exactly one target: anchorId (with optional span) or revisionId");

        return revisionId is not null
            ? DocxSessionOps.AddCommentToRevision(
                session.Handle, revisionId, Str(args, "author"), OptStr(args, "initials"),
                OptStr(args, "date"), OptStr(args, "markdown") ?? "")
            : DocxSessionOps.AddComment(
                session.Handle, anchorId!, ParseSpan(args, "span"), Str(args, "author"),
                OptStr(args, "initials"), OptStr(args, "date"), OptStr(args, "markdown") ?? "");
    }

    // ─── Annotate (annotation overlay) ─────────────────────────────────

    private static string Annotate(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var action = Str(args, "action");
        if (action == "list")
            return $"{{\"annotations\":{DocxSessionOps.ListAnnotations(session.Handle)}}}";
        if (action == "find")
            return $"{{\"anchors\":{DocxSessionOps.FindByAnnotation(session.Handle, Str(args, "query"))}}}";
        return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
            "add" => AddAnnotation(session, args),
            "update" => DocxSessionOps.UpdateAnnotation(
                session.Handle, Str(args, "annotationId"),
                args.TryGetProperty("update", out var u) && u.ValueKind == JsonValueKind.Object
                    ? u.GetRawText() : "{}"),
            "remove" => DocxSessionOps.RemoveAnnotation(session.Handle, Str(args, "annotationId")),
            "move" => DocxSessionOps.MoveAnnotation(
                session.Handle, Str(args, "annotationId"), Str(args, "newAnchorId"), ParseSpan(args, "newSpan")),
            _ => throw new McpToolException($"unknown docxodus_annotate action: {action}"),
        });
    }

    private static string AddAnnotation(DocSession session, JsonElement args)
    {
        var annotationJson = JsonSerializer.Serialize(new
        {
            id = OptStr(args, "annotationId") ?? "",
            labelId = OptStr(args, "labelId") ?? "",
            label = OptStr(args, "label") ?? "",
            color = OptStr(args, "color") ?? "#FFEB3B",
            author = OptStr(args, "author") ?? "",
        });
        return DocxSessionOps.AddAnnotation(
            session.Handle, Str(args, "anchorId"), ParseSpan(args, "span"), annotationJson);
    }

    // ─── Track changes ──────────────────────────────────────────────────

    private static string TrackChanges(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var action = Str(args, "action");
        return RunTrackChangesAction(session, action, args);
    }

    private static string RunTrackChangesAction(DocSession session, string action, JsonElement args)
    {
        switch (action)
        {
            case "list":
            {
                // Markup-native listing (issue #318): read w:ins/w:del/move/format markup
                // straight off the live session — stable per-revision ids, the markup's
                // true authors/dates, and none of the ~seconds-long accept-all/reject-all
                // re-diff the old listing paid on large documents.
                var revisionsJson = "{\"revisions\":" + DocxSessionOps.ListRevisions(session.Handle) + "}";
                return FilterRevisions(revisionsJson, OptStr(args, "author"), OptStr(args, "changeType"),
                    OptStr(args, "family"), OptStr(args, "resolutionStatus"), OptStr(args, "partUri"));
            }
            case "accept":
                return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () =>
                    DocxSessionOps.AcceptRevision(session.Handle, Str(args, "revisionId")));
            case "reject":
                return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () =>
                    DocxSessionOps.RejectRevision(session.Handle, Str(args, "revisionId")));
            case "accept_all":
                return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () =>
                    DocxSessionOps.AcceptAllRevisions(session.Handle));
            case "reject_all":
                return Guarded(session, ParsePreconditions(args, MutationTarget(args)), () =>
                    DocxSessionOps.RejectAllRevisions(session.Handle));
            case "set_mode":
            {
                var modeStr = Str(args, "mode");
                var mode = modeStr switch
                {
                    "accept" => TrackedChangeMode.Accept,
                    "render_inline" => TrackedChangeMode.RenderInline,
                    "strip_deletions" => TrackedChangeMode.StripDeletions,
                    _ => throw new McpToolException($"unknown trackedChanges mode: {modeStr}"),
                };
                DocxSessionOps.SetTrackedChanges(session.Handle, mode);
                if (OptStr(args, "revisionAuthor") is { } author)
                    DocxSessionOps.SetRevisionAuthor(session.Handle, author.Length == 0 ? null : author);
                var state = DocxSessionOps.GetTrackedChanges(session.Handle);
                return "{\"success\":true," + state.Substring(1);
            }
            default:
                throw new McpToolException($"unknown docxodus_track_changes action: {action}");
        }
    }

    private static string FilterRevisions(string revisionsJson, string? author, string? changeType,
        string? family, string? resolutionStatus, string? partUri)
    {
        if (author is null && changeType is null && family is null
            && resolutionStatus is null && partUri is null) return revisionsJson;
        using var doc = JsonDocument.Parse(revisionsJson);
        if (!doc.RootElement.TryGetProperty("revisions", out var revisions) || revisions.ValueKind != JsonValueKind.Array)
            return revisionsJson;

        var kept = new List<string>();
        foreach (var r in revisions.EnumerateArray())
        {
            if (author is not null
                && (!r.TryGetProperty("author", out var a) || !string.Equals(a.GetString(), author, StringComparison.OrdinalIgnoreCase)))
                continue;
            if (changeType is not null
                && (!r.TryGetProperty("type", out var t) || !string.Equals(t.GetString(), changeType, StringComparison.OrdinalIgnoreCase)))
                continue;
            if (family is not null
                && (!r.TryGetProperty("family", out var f) || !string.Equals(f.GetString(), family, StringComparison.OrdinalIgnoreCase)))
                continue;
            if (resolutionStatus is not null
                && (!r.TryGetProperty("resolutionStatus", out var status)
                    || !string.Equals(status.GetString(), resolutionStatus, StringComparison.OrdinalIgnoreCase)))
                continue;
            if (partUri is not null
                && (!r.TryGetProperty("partUri", out var part)
                    || !string.Equals(part.GetString(), partUri, StringComparison.Ordinal)))
                continue;
            kept.Add(r.GetRawText());
        }
        return "{\"revisions\":[" + string.Join(",", kept) + "]}";
    }

    // ─── Mutations (batch) ──────────────────────────────────────────────

    private static string Mutations(SessionStore store, JsonElement args)
    {
        var liveSession = Session(store, args);
        var mode = args.TryGetProperty("mode", out _)
            ? Str(args, "mode")
            : "atomic";
        if (mode is not ("atomic" or "best_effort" or "apply" or "preview"))
            throw new McpToolException($"unknown docxodus_mutations mode: {mode}");
        if (!args.TryGetProperty("steps", out var stepsEl) || stepsEl.ValueKind != JsonValueKind.Array)
            throw new McpToolException("docxodus_mutations requires an array \"steps\"");

        var preview = mode == "preview" || BoolOpt(args, "preview", false);
        if (mode == "apply" && preview)
            throw new McpToolException("docxodus_mutations preview cannot be combined with deprecated mode 'apply'");
        var policyName = mode == "preview"
            ? OptStr(args, "previewPolicy") ?? "atomic"
            : mode;
        if (policyName is not ("atomic" or "best_effort" or "apply"))
            throw new McpToolException($"unknown docxodus_mutations previewPolicy: {policyName}");
        var coreMode = policyName == "atomic"
            ? MutationBatchMode.Atomic
            : MutationBatchMode.BestEffort;
        var htmlMode = OptStr(args, "previewHtml") switch
        {
            null or "none" => MutationPreviewHtmlMode.None,
            "scoped" => MutationPreviewHtmlMode.Scoped,
            "full" => MutationPreviewHtmlMode.Full,
            var value => throw new McpToolException($"unknown docxodus_mutations previewHtml: {value}"),
        };
        if (!preview && htmlMode != MutationPreviewHtmlMode.None)
            throw new McpToolException("previewHtml is only valid for a preview batch");

        if (preview)
        {
            return DocxSessionOps.PreviewBatch(
                liveSession.Handle,
                coreMode,
                shadowHandle =>
                {
                    var shadow = new DocSession
                    {
                        Id = liveSession.Id,
                        Handle = shadowHandle,
                    };
                    var batchCheck = Check(shadow, ParsePreconditions(args, MutationTarget(args)));
                    if (batchCheck is not null)
                    {
                        return new[]
                        {
                            DocxSessionOps.SerializedBatchStep(
                                "docxodus_mutations",
                                "preconditions",
                                () => batchCheck),
                        };
                    }
                    return BuildMutationBatchSteps(shadow, stepsEl, legacyApply: false);
                },
                new MutationBatchPreviewOptions
                {
                    HtmlMode = htmlMode,
                    HtmlAnchorId = OptStr(args, "previewAnchorId"),
                });
        }

        var liveBatchCheck = Check(liveSession, ParsePreconditions(args, MutationTarget(args)));
        if (liveBatchCheck is not null) return liveBatchCheck;
        var liveSteps = BuildMutationBatchSteps(liveSession, stepsEl, legacyApply: mode == "apply");
        return DocxSessionOps.ExecuteBatch(liveSession.Handle, coreMode, liveSteps);
    }

    private static IReadOnlyList<MutationBatchStep> BuildMutationBatchSteps(
        DocSession session,
        JsonElement stepsEl,
        bool legacyApply)
    {
        var result = new List<MutationBatchStep>();
        foreach (var step in stepsEl.EnumerateArray())
        {
            var stepTool = step.TryGetProperty("tool", out var toolEl) && toolEl.ValueKind == JsonValueKind.String
                ? toolEl.GetString()! : throw new McpToolException("mutation step missing string \"tool\"");
            var stepArgs = step.TryGetProperty("args", out var a) && a.ValueKind == JsonValueKind.Object
                ? a : throw new McpToolException("mutation step missing object \"args\"");
            var action = stepArgs.TryGetProperty("action", out var actEl) && actEl.ValueKind == JsonValueKind.String
                ? actEl.GetString()! : throw new McpToolException("mutation step args missing string \"action\"");

            var actionError = ValidateMutationBatchAction(stepTool, action);
            if (legacyApply && actionError is not null)
                throw new McpToolException(actionError.Message);
            // Step preconditions are evaluated by the core batch preflight: all against the
            // batch-start state for atomic mode, immediately before each step for best-effort.
            // Remove them from the actual dispatch so a valid atomic preflight is not evaluated
            // a second time against state changed by an earlier step in the same batch.
            var mutationArgs = WithoutProperty(stepArgs, "preconditions");
            result.Add(DocxSessionOps.SerializedBatchStep(
                stepTool,
                action,
                () => stepTool switch
                {
                    "docxodus_edit" => RunEditAction(session, action, mutationArgs),
                    "docxodus_format" => RunFormatAction(session, action, mutationArgs),
                    "docxodus_create" => RunCreateAction(session, action, mutationArgs),
                    "docxodus_table" => RunTableAction(session, action, mutationArgs),
                    "docxodus_list" => RunListAction(session, action, mutationArgs),
                    "docxodus_comment" => RunCommentAction(session, action, mutationArgs),
                    "docxodus_links" => RunLinksAction(session, action, mutationArgs),
                    "docxodus_images" => RunImagesAction(session, action, mutationArgs),
                    "docxodus_track_changes" => RunTrackChangesAction(session, action, mutationArgs),
                    _ => throw new McpToolException($"docxodus_mutations does not accept \"{stepTool}\" as a step"),
                },
                () => ValidateMutationBatchStep(session, stepTool, action, stepArgs)));
        }
        return result;
    }

    private static EditError? ValidateMutationBatchAction(string tool, string action)
    {
        bool known = tool switch
        {
            "docxodus_edit" => action is "insert_paragraph" or "replace_text" or "replace_text_range"
                or "delete_block" or "move_block" or "delete_range" or "delete_section"
                or "split_paragraph" or "merge_paragraphs",
            "docxodus_format" => action is "apply_format" or "apply_format_by_substring"
                or "set_paragraph_style" or "set_paragraph_format" or "set_list_level"
                or "remove_list_membership" or "apply_list_format",
            "docxodus_create" => action is "insert_paragraph" or "insert_heading" or "insert_table"
                or "insert_horizontal_rule" or "insert_footnote" or "insert_endnote"
                or "insert_page_number_field" or "set_header_text" or "set_footer_text"
                or "ensure_header_footer_visible",
            "docxodus_table" => action is "insert" or "insert_row" or "insert_column"
                or "delete_row" or "delete_column" or "replace_cell_content" or "merge_cells"
                or "unmerge_cells" or "set_column_widths" or "set_borders" or "set_shading"
                or "set_repeat_header_row" or "set_row_options",
            "docxodus_list" => action is "apply_format" or "apply_format_range" or "set_level"
                or "set_start" or "clear_start" or "remove",
            "docxodus_comment" => action is "add" or "reply" or "resolve" or "update" or "remove",
            "docxodus_links" => action is "add_hyperlink" or "update_hyperlink" or "remove_hyperlink"
                or "add_bookmark" or "move_bookmark" or "rename_bookmark" or "remove_bookmark",
            "docxodus_images" => action is "insert" or "replace" or "set_dimensions"
                or "set_metadata" or "set_floating_layout" or "remove",
            "docxodus_track_changes" => action is "accept" or "reject"
                or "accept_all" or "reject_all",
            _ => false,
        };
        return known ? null : new EditError(
            EditErrorCode.InvalidBatchStep,
            $"unsupported or read-only batch action: {tool}/{action}");
    }

    private static EditError? ValidateMutationBatchStep(
        DocSession session,
        string tool,
        string action,
        JsonElement args)
    {
        var actionError = ValidateMutationBatchAction(tool, action);
        if (actionError is not null) return actionError;

        try
        {
            ValidateMutationBatchArguments(tool, action, args);
            var failure = Check(session, ParsePreconditions(args, MutationTarget(args)));
            if (failure is null) return null;
            return DocxSessionJson.DeserializeEditResults(failure).FirstOrDefault()?.Error
                ?? new EditError(EditErrorCode.PreconditionFailed,
                    "batch step precondition failed");
        }
        catch (Exception ex) when (ex is McpToolException
            or ArgumentException or FormatException or JsonException or OverflowException)
        {
            return new EditError(EditErrorCode.InvalidBatchStep, ex.Message);
        }
    }

    /// <summary>
    /// Parse every syntactic input that a mutation action will consume, without invoking the
    /// mutation. This keeps caller-attributable schema/enum errors out of InternalError and lets
    /// atomic mode reject the complete batch before step zero changes the package.
    /// </summary>
    private static void ValidateMutationBatchArguments(string tool, string action, JsonElement args)
    {
        switch ((tool, action))
        {
            case ("docxodus_edit", "insert_paragraph"):
                RequireStrings(args, "anchorId", "markdown");
                ValidateOptionalEnum(args, "position", "before", "after");
                break;
            case ("docxodus_edit", "replace_text"):
                RequireStrings(args, "anchorId", "markdown");
                break;
            case ("docxodus_edit", "replace_text_range"):
                RequireStrings(args, "anchorId", "find", "replace");
                ValidateOptionalBool(args, "caseSensitive");
                break;
            case ("docxodus_edit", "delete_block"):
                RequireStrings(args, "anchorId");
                break;
            case ("docxodus_edit", "move_block"):
                RequireStrings(args, "sourceAnchorId", "targetAnchorId");
                ValidateOptionalEnum(args, "position", "before", "after");
                break;
            case ("docxodus_edit", "delete_range"):
                RequireStrings(args, "fromAnchorId", "toAnchorIdExclusive");
                break;
            case ("docxodus_edit", "delete_section"):
                RequireStrings(args, "headingAnchorId");
                break;
            case ("docxodus_edit", "split_paragraph"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "characterOffset");
                break;
            case ("docxodus_edit", "merge_paragraphs"):
                RequireStrings(args, "anchorId", "secondAnchorId");
                break;

            case ("docxodus_format", "apply_format"):
                RequireStrings(args, "anchorId");
                ValidateOptionalObject(args, "format");
                ValidateOptionalSpan(args, "span");
                _ = ParseFormatOp(args);
                break;
            case ("docxodus_format", "apply_format_by_substring"):
                RequireStrings(args, "anchorId", "substring");
                ValidateOptionalObject(args, "format");
                _ = ParseFormatOp(args);
                break;
            case ("docxodus_format", "set_paragraph_style"):
                RequireStrings(args, "anchorId", "styleId");
                break;
            case ("docxodus_format", "set_paragraph_format"):
                RequireStrings(args, "anchorId");
                ValidateOptionalObject(args, "paragraphFormat");
                _ = ParseParagraphFormatOp(args);
                break;
            case ("docxodus_format", "set_list_level"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "levelDelta");
                break;
            case ("docxodus_format", "remove_list_membership"):
                RequireStrings(args, "anchorId");
                break;
            case ("docxodus_format", "apply_list_format"):
                RequireStrings(args, "anchorId");
                ValidateOptionalListFormat(args);
                break;

            case ("docxodus_create", "insert_paragraph"):
                RequireStrings(args, "anchorId", "markdown");
                ValidateOptionalEnum(args, "position", "before", "after");
                break;
            case ("docxodus_create", "insert_heading"):
                RequireStrings(args, "anchorId", "text");
                RequireNumbers(args, "level");
                ValidateOptionalEnum(args, "position", "before", "after");
                break;
            case ("docxodus_create", "insert_table"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "rows", "columns");
                ValidateOptionalEnum(args, "position", "before", "after");
                ValidateOptionalArray(args, "cellContents");
                ValidateOptionalArray(args, "columnWidths");
                ValidateOptionalEnum(args, "cellAlignment", "left", "center", "right", "justify");
                ValidateOptionalBool(args, "borderless");
                break;
            case ("docxodus_create", "insert_horizontal_rule"):
                RequireStrings(args, "anchorId");
                ValidateOptionalEnum(args, "position", "before", "after");
                ValidateOptionalEnum(args, "ruleStyle", "single", "double", "thick");
                break;
            case ("docxodus_create", "insert_footnote"):
            case ("docxodus_create", "insert_endnote"):
                RequireStrings(args, "anchorId", "markdown");
                RequireNumbers(args, "characterOffset");
                break;
            case ("docxodus_create", "insert_page_number_field"):
                RequireStrings(args, "anchorId");
                ValidateOptionalEnum(args, "field", "current_page", "total_pages");
                ValidateOptionalEnum(args, "numberFormat", "decimal", "upperLetter",
                    "lowerLetter", "upperRoman", "lowerRoman");
                break;
            case ("docxodus_create", "set_header_text"):
            case ("docxodus_create", "set_footer_text"):
                RequireStrings(args, "bodyAnchorId", "kind", "markdown");
                ValidateRequiredEnum(args, "kind", "default", "first", "even");
                break;
            case ("docxodus_create", "ensure_header_footer_visible"):
                RequireStrings(args, "bodyAnchorId", "kind");
                ValidateRequiredEnum(args, "kind", "default", "first", "even");
                break;

            case ("docxodus_table", "insert"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "rows", "columns");
                ValidateOptionalEnum(args, "position", "before", "after");
                ValidateOptionalArray(args, "cellContents");
                ValidateOptionalArray(args, "columnWidths");
                ValidateOptionalEnum(args, "cellAlignment", "left", "center", "right", "justify");
                ValidateOptionalBool(args, "borderless");
                break;
            case ("docxodus_table", "insert_row"):
            case ("docxodus_table", "insert_column"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalEnum(args, "position", "before", "after");
                break;
            case ("docxodus_table", "delete_row"):
            case ("docxodus_table", "delete_column"):
            case ("docxodus_table", "unmerge_cells"):
                RequireStrings(args, "cellAnchorId");
                break;
            case ("docxodus_table", "replace_cell_content"):
                RequireStrings(args, "cellAnchorId", "markdown");
                break;
            case ("docxodus_table", "merge_cells"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalNumber(args, "rowSpan");
                ValidateOptionalNumber(args, "colSpan");
                ValidateOptionalEnum(args, "mergeContent", "append", "discard", "reject");
                break;
            case ("docxodus_table", "set_column_widths"):
                RequireStrings(args, "cellAnchorId");
                _ = RawArray(args, "widths");
                break;
            case ("docxodus_table", "set_borders"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalEnum(args, "borderScope", "all", "outside", "inside");
                ValidateOptionalString(args, "borderStyle");
                ValidateOptionalNumber(args, "borderSize");
                ValidateOptionalString(args, "borderColor");
                break;
            case ("docxodus_table", "set_shading"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalString(args, "fill");
                ValidateOptionalEnum(args, "shadingScope", "cell", "row");
                break;
            case ("docxodus_table", "set_repeat_header_row"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalBool(args, "repeat");
                break;
            case ("docxodus_table", "set_row_options"):
                RequireStrings(args, "cellAnchorId");
                ValidateOptionalBool(args, "repeat");
                ValidateOptionalBool(args, "allowBreakAcrossPages");
                ValidateOptionalNumber(args, "heightTwips");
                ValidateOptionalEnum(args, "heightRule", "auto", "atLeast", "exact");
                break;

            case ("docxodus_list", "apply_format"):
                RequireStrings(args, "anchorId");
                ValidateOptionalListFormat(args);
                break;
            case ("docxodus_list", "apply_format_range"):
                RequireStrings(args, "firstAnchorId", "lastAnchorId");
                ValidateOptionalListFormat(args);
                break;
            case ("docxodus_list", "set_level"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "levelDelta");
                break;
            case ("docxodus_list", "set_start"):
                RequireStrings(args, "anchorId");
                RequireNumbers(args, "startValue");
                break;
            case ("docxodus_list", "clear_start"):
            case ("docxodus_list", "remove"):
                RequireStrings(args, "anchorId");
                break;

            case ("docxodus_comment", "add"):
                ValidateCommentAddArguments(args);
                break;
            case ("docxodus_comment", "reply"):
                RequireStrings(args, "commentAnchorId", "author");
                ValidateOptionalString(args, "initials");
                ValidateOptionalString(args, "date");
                ValidateOptionalString(args, "markdown");
                break;
            case ("docxodus_comment", "update"):
                RequireStrings(args, "commentAnchorId", "markdown");
                break;
            case ("docxodus_comment", "resolve"):
                RequireStrings(args, "commentAnchorId");
                ValidateOptionalBool(args, "resolved");
                break;
            case ("docxodus_comment", "remove"):
                RequireStrings(args, "commentAnchorId");
                break;

            case ("docxodus_links", "add_hyperlink"):
                RequireStrings(args, "anchorId", "kind", "target");
                RequireNumbers(args, "startOffset", "length");
                ValidateRequiredEnum(args, "kind", "external", "internal");
                break;
            case ("docxodus_links", "update_hyperlink"):
                RequireStrings(args, "hyperlinkId", "kind", "target");
                ValidateRequiredEnum(args, "kind", "external", "internal");
                break;
            case ("docxodus_links", "remove_hyperlink"):
                RequireStrings(args, "hyperlinkId");
                break;
            case ("docxodus_links", "add_bookmark"):
            case ("docxodus_links", "move_bookmark"):
                RequireStrings(args, "name", "startAnchorId", "endAnchorId");
                RequireNumbers(args, "startOffset", "endOffset");
                break;
            case ("docxodus_links", "rename_bookmark"):
                RequireStrings(args, "name", "newName");
                break;
            case ("docxodus_links", "remove_bookmark"):
                RequireStrings(args, "name");
                break;

            case ("docxodus_images", "insert"):
                RequireStrings(args, "anchorId", "imageBase64");
                RequireNumbers(args, "characterOffset");
                _ = RawObjectOrEmpty(args, "options");
                break;
            case ("docxodus_images", "replace"):
                RequireStrings(args, "imageId", "imageBase64");
                break;
            case ("docxodus_images", "set_dimensions"):
                RequireStrings(args, "imageId");
                _ = RawObject(args, "dimensions");
                break;
            case ("docxodus_images", "set_metadata"):
                RequireStrings(args, "imageId");
                ValidateNullableString(args, "altText");
                ValidateNullableString(args, "title");
                break;
            case ("docxodus_images", "set_floating_layout"):
                RequireStrings(args, "imageId");
                _ = RawObject(args, "layout");
                break;
            case ("docxodus_images", "remove"):
                RequireStrings(args, "imageId");
                break;

            case ("docxodus_track_changes", "accept"):
            case ("docxodus_track_changes", "reject"):
                RequireStrings(args, "revisionId");
                break;
            case ("docxodus_track_changes", "accept_all"):
            case ("docxodus_track_changes", "reject_all"):
                break;
        }
    }

    private static void ValidateNullableString(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value)
            || value.ValueKind is not (JsonValueKind.String or JsonValueKind.Null))
            throw new McpToolException($"argument \"{name}\" must be a string or null");
    }

    private static void ValidateCommentAddArguments(JsonElement args)
    {
        var anchorId = OptionalStringValue(args, "anchorId");
        var revisionId = OptionalStringValue(args, "revisionId");
        if ((anchorId is null) == (revisionId is null))
            throw new McpToolException(
                "docxodus_comment add requires exactly one target: anchorId or revisionId");
        RequireStrings(args, "author");
        ValidateOptionalString(args, "initials");
        ValidateOptionalString(args, "date");
        ValidateOptionalString(args, "markdown");
        if (revisionId is not null && args.TryGetProperty("span", out _))
            throw new McpToolException("revisionId comment targets cannot include span");
        ValidateOptionalSpan(args, "span");
    }

    private static void RequireStrings(JsonElement args, params string[] names)
    {
        foreach (var name in names) _ = Str(args, name);
    }

    private static void RequireNumbers(JsonElement args, params string[] names)
    {
        foreach (var name in names) _ = Int(args, name);
    }

    private static string? OptionalStringValue(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value)) return null;
        if (value.ValueKind != JsonValueKind.String)
            throw new McpToolException($"argument \"{name}\" must be a string");
        return value.GetString();
    }

    private static void ValidateOptionalString(JsonElement args, string name) =>
        _ = OptionalStringValue(args, name);

    private static void ValidateOptionalBool(JsonElement args, string name)
    {
        if (args.TryGetProperty(name, out var value)
            && value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new McpToolException($"argument \"{name}\" must be a boolean");
    }

    private static void ValidateOptionalNumber(JsonElement args, string name)
    {
        if (args.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Number)
            throw new McpToolException($"argument \"{name}\" must be a number");
        if (args.TryGetProperty(name, out value)) _ = value.GetInt32();
    }

    private static void ValidateOptionalObject(JsonElement args, string name)
    {
        if (args.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"argument \"{name}\" must be an object");
    }

    private static void ValidateOptionalArray(JsonElement args, string name)
    {
        if (args.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Array)
            throw new McpToolException($"argument \"{name}\" must be an array");
    }

    private static void ValidateOptionalSpan(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var span)) return;
        if (span.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"argument \"{name}\" must be an object");
        if (span.TryGetProperty("start", out var start))
        {
            if (start.ValueKind != JsonValueKind.Number)
                throw new McpToolException($"argument \"{name}.start\" must be a number");
            _ = start.GetInt32();
        }
        if (span.TryGetProperty("length", out var length))
        {
            if (length.ValueKind != JsonValueKind.Number)
                throw new McpToolException($"argument \"{name}.length\" must be a number");
            _ = length.GetInt32();
        }
    }

    private static void ValidateOptionalListFormat(JsonElement args) =>
        ValidateOptionalEnum(args, "listFormat", "bullet", "decimal", "lowerLetter",
            "upperLetter", "lowerRoman", "upperRoman", "decimalParenthesis",
            "lowerLetterParenthesis", "upperLetterParenthesis", "lowerRomanParenthesis",
            "upperRomanParenthesis", "none");

    private static void ValidateRequiredEnum(JsonElement args, string name, params string[] values)
    {
        _ = Str(args, name);
        ValidateOptionalEnum(args, name, values);
    }

    private static void ValidateOptionalEnum(JsonElement args, string name, params string[] values)
    {
        var value = OptionalStringValue(args, name);
        if (value is not null && !values.Contains(value, StringComparer.Ordinal))
            throw new McpToolException(
                $"unknown {name}: {value}; expected one of {string.Join(", ", values)}");
    }

    private static JsonElement WithoutProperty(JsonElement source, string propertyName)
    {
        var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(source.GetRawText())!;
        values.Remove(propertyName);
        return JsonSerializer.SerializeToElement(values);
    }

    // ─── Table ──────────────────────────────────────────────────────────

    private static string Table(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunTableAction(session, Str(args, "action"), args);
    }

    private static string RunTableAction(DocSession session, string action, JsonElement args) => action switch
    {
        "get_metadata" => DocxSessionOps.GetTableMetadata(
            session.Handle, Str(args, "tableAnchorId")),
        "resolve_cell_anchor" => DocxSessionOps.ResolveTableCellAnchor(
            session.Handle, Str(args, "cellAnchorId")),
        "resolve_cell_coordinate" => DocxSessionOps.ResolveTableCellCoordinate(
            session.Handle, Str(args, "tableAnchorId"), Int(args, "rowIndex"), Int(args, "columnIndex")),
        _ => Guarded(session, ParsePreconditions(args, MutationTarget(args)), () => action switch
        {
        "insert" => DocxSessionOps.InsertTable(
            session.Handle, Str(args, "anchorId"), ParsePos(args),
            Int(args, "rows"), Int(args, "columns"), BuildTableInsertOptionsJson(args)),
        "insert_row" => DocxSessionOps.InsertTableRow(session.Handle, Str(args, "cellAnchorId"), ParsePos(args)),
        "insert_column" => DocxSessionOps.InsertTableColumn(session.Handle, Str(args, "cellAnchorId"), ParsePos(args)),
        "delete_row" => DocxSessionOps.DeleteTableRow(session.Handle, Str(args, "cellAnchorId")),
        "delete_column" => DocxSessionOps.DeleteTableColumn(session.Handle, Str(args, "cellAnchorId")),
        "replace_cell_content" => DocxSessionOps.ReplaceCellContent(
            session.Handle, Str(args, "cellAnchorId"), Str(args, "markdown")),
        "merge_cells" => DocxSessionOps.MergeCells(
            session.Handle, Str(args, "cellAnchorId"), OptInt(args, "rowSpan") ?? 1,
            OptInt(args, "colSpan") ?? 1, OptStr(args, "mergeContent")),
        "unmerge_cells" => DocxSessionOps.UnmergeCells(session.Handle, Str(args, "cellAnchorId")),
        "set_column_widths" => DocxSessionOps.SetColumnWidths(
            session.Handle, Str(args, "cellAnchorId"), RawArray(args, "widths")),
        "set_borders" => DocxSessionOps.SetTableBorders(
            session.Handle, Str(args, "cellAnchorId"), BuildTableBorderSpecJson(args)),
        "set_shading" => DocxSessionOps.SetCellShading(
            session.Handle, Str(args, "cellAnchorId"), OptStr(args, "fill") ?? "",
            OptStr(args, "shadingScope") ?? "cell"),
        "set_repeat_header_row" => DocxSessionOps.SetRepeatHeaderRow(
            session.Handle, Str(args, "cellAnchorId"), BoolOpt(args, "repeat", true)),
        "set_row_options" => DocxSessionOps.SetTableRowOptions(
            session.Handle, Str(args, "cellAnchorId"), OptBool(args, "repeat"),
            OptBool(args, "allowBreakAcrossPages"), OptInt(args, "heightTwips"),
            OptStr(args, "heightRule")),
        _ => throw new McpToolException($"unknown docxodus_table action: {action}"),
        }),
    };

    /// <summary>The raw JSON text of a required array argument (passed through to the Ops-layer
    /// parser so the wire shape lives in one place).</summary>
    private static string RawArray(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Array)
            throw new McpToolException($"missing required array argument \"{name}\"");
        return v.GetRawText();
    }

    private static string RawObject(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"missing required object argument \"{name}\"");
        return value.GetRawText();
    }

    private static string RawObjectOrEmpty(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var value))
            return "{}";
        if (value.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"optional argument \"{name}\" must be an object when present");
        return value.GetRawText();
    }

    private static string BuildTableBorderSpecJson(JsonElement args)
    {
        var spec = new Dictionary<string, object?>
        {
            ["scope"] = OptStr(args, "borderScope"),
            ["style"] = OptStr(args, "borderStyle"),
            ["color"] = OptStr(args, "borderColor"),
        };
        if (args.ValueKind == JsonValueKind.Object && args.TryGetProperty("borderSize", out var sz) && sz.ValueKind == JsonValueKind.Number)
            spec["size"] = sz.GetInt32();
        return JsonSerializer.Serialize(spec);
    }

    // ─── Arg helpers ────────────────────────────────────────────────────

    private static MutationPreconditions? ParsePreconditions(JsonElement args, string? inferredAnchorId)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("preconditions", out var p)
            || p.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        var parsed = DocxSessionJson.ParseMutationPreconditions(p);
        return parsed is not null && parsed.AnchorId is null && inferredAnchorId is not null
            ? parsed with { AnchorId = inferredAnchorId }
            : parsed;
    }

    private static string? MutationTarget(JsonElement args)
    {
        foreach (var name in new[]
        {
            "anchorId", "cellAnchorId", "sourceAnchorId", "fromAnchorId", "firstAnchorId",
            "headingAnchorId", "bodyAnchorId", "commentAnchorId", "newAnchorId",
        })
        {
            if (args.ValueKind == JsonValueKind.Object
                && args.TryGetProperty(name, out var target)
                && target.ValueKind == JsonValueKind.String)
                return target.GetString();
        }
        return null;
    }

    private static string? Check(DocSession session, MutationPreconditions? preconditions)
    {
        if (preconditions is null) return null;
        var result = DocxSessionOps.CheckPreconditions(session.Handle, preconditions);
        using var doc = JsonDocument.Parse(result);
        return doc.RootElement.GetProperty("success").GetBoolean() ? null : result;
    }

    private static string Guarded(
        DocSession session, MutationPreconditions? preconditions, Func<string> mutation)
    {
        var failure = Check(session, preconditions);
        return failure ?? mutation();
    }

    private static DocSession Session(SessionStore store, JsonElement args) => store.Get(Str(args, "sessionId"));

    private static string Str(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.String)
            throw new McpToolException($"missing required string argument \"{name}\"");
        return v.GetString()!;
    }

    private static string? OptStr(JsonElement args, string name) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
            ? v.GetString() : null;

    private static JsonElement Object(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var value)
            || value.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"missing required object argument \"{name}\"");
        return value;
    }

    private static int Int(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Number)
            throw new McpToolException($"missing required number argument \"{name}\"");
        return v.GetInt32();
    }

    private static int IntOpt(JsonElement args, string name, int fallback) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt32() : fallback;

    private static long LongOpt(JsonElement args, string name, long fallback) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt64() : fallback;

    private static int? OptInt(JsonElement args, string name) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt32() : null;

    private static bool BoolOpt(JsonElement args, string name, bool fallback) =>
        OptBool(args, name) ?? fallback;

    private static bool? OptBool(JsonElement args, string name) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False)
            ? v.GetBoolean() : null;

    private static Position ParsePos(JsonElement args) =>
        DocxSessionJson.ParsePos(OptStr(args, "position") ?? "after");

    private static CharSpan? ParseSpan(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var s) || s.ValueKind != JsonValueKind.Object)
            return null;
        return new CharSpan(
            s.TryGetProperty("start", out var st) && st.ValueKind == JsonValueKind.Number ? st.GetInt32() : 0,
            s.TryGetProperty("length", out var ln) && ln.ValueKind == JsonValueKind.Number ? ln.GetInt32() : 0);
    }
}
