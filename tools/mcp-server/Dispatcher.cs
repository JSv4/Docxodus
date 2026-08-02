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
        "docxodus_search" => Search(store, args),
        "docxodus_edit" => Edit(store, args),
        "docxodus_format" => Format(store, args),
        "docxodus_create" => Create(store, args),
        "docxodus_list" => ListTool(store, args),
        "docxodus_comment" => Comment(store, args),
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
        var settings = new DocxSessionSettings
        {
            TrackedChanges = tracked,
            RevisionAuthor = OptStr(args, "revisionAuthor"),
            UndoDepth = IntOpt(args, "undoDepth", 50),
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

        switch (format)
        {
            case "markdown":
                return anchorId is null
                    ? DocxSessionOps.Project(session.Handle)
                    : DocxSessionOps.ProjectAnchor(session.Handle, anchorId, ProjectionDepth.SubtreeAndFollowingSiblings);

            case "text":
            {
                var projectionJson = anchorId is null
                    ? DocxSessionOps.Project(session.Handle)
                    : DocxSessionOps.ProjectAnchor(session.Handle, anchorId, ProjectionDepth.SubtreeAndFollowingSiblings);
                using var doc = JsonDocument.Parse(projectionJson);
                var markdown = doc.RootElement.GetProperty("markdown").GetString() ?? string.Empty;
                return $"{{\"text\":{JsonRpcIo.JsonString(StripMarkdownSyntax(markdown))}}}";
            }

            case "html":
                return $"{{\"html\":{JsonRpcIo.JsonString(
                    anchorId is null
                        ? DocxSessionOps.RenderHtml(session.Handle, "docx-", false, false, 1.0)
                        : DocxSessionOps.RenderBlockHtml(session.Handle, anchorId, "docx-", false))}}}";

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
                var projectionJson = DocxSessionOps.Project(session.Handle);
                using (var doc = JsonDocument.Parse(projectionJson))
                {
                    foreach (var prop in doc.RootElement.GetProperty("anchorIndex").EnumerateObject())
                    {
                        var kind = prop.Value.GetProperty("kind").GetString();
                        var scope = prop.Value.GetProperty("scope").GetString();
                        if (scope != "body" || kind is not ("p" or "h" or "li")) continue;
                        sectionInfo = DocxSessionOps.GetSectionInfo(session.Handle, prop.Name);
                        break;
                    }
                }
                return $"{{\"editSummary\":{editSummary},\"sectionInfo\":{sectionInfo}}}";
            }

            default:
                throw new McpToolException($"unknown format: {format}");
        }
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
        var maxResults = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("maxResults", out var mr) && mr.ValueKind == JsonValueKind.Number
            ? mr.GetInt32() : (int?)null;

        string matchesJson = mode switch
        {
            "text" => DocxSessionOps.Grep(
                session.Handle, Regex.Escape(query), caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase,
                ProjectionScopes.Body, contextChars, WhitespaceMode.Preserve, ContextBoundary.Char),
            "regex" => DocxSessionOps.Grep(
                session.Handle, query, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase,
                ProjectionScopes.Body, contextChars, WhitespaceMode.Preserve, ContextBoundary.Char),
            "kind" => DocxSessionOps.FindByKind(session.Handle, query, null),
            "annotation" => DocxSessionOps.FindByAnnotation(session.Handle, query),
            "bookmark" => DocxSessionOps.FindByBookmark(session.Handle, query),
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

    // ─── Edit ───────────────────────────────────────────────────────────

    private static string Edit(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunEditAction(session, Str(args, "action"), args);
    }

    /// <summary>Shared by <see cref="Edit"/> and <see cref="Mutations"/> (batched steps route
    /// through the same per-tool functions so there's exactly one place each action's argument
    /// parsing lives).</summary>
    private static string RunEditAction(DocSession session, string action, JsonElement args) => action switch
    {
        "insert_paragraph" => DocxSessionOps.InsertParagraph(
            session.Handle, Str(args, "anchorId"), ParsePos(args), Str(args, "markdown")),
        "replace_text" => DocxSessionOps.ReplaceText(session.Handle, Str(args, "anchorId"), Str(args, "markdown")),
        "replace_text_range" => DocxSessionOps.ReplaceTextRange(
            session.Handle, Str(args, "anchorId"), Str(args, "find"), Str(args, "replace"),
            new ReplaceOptions { IgnoreCase = !BoolOpt(args, "caseSensitive", false) }),
        "delete_block" => DocxSessionOps.DeleteBlock(session.Handle, Str(args, "anchorId")),
        "delete_range" => DocxSessionOps.DeleteRange(
            session.Handle, Str(args, "fromAnchorId"), Str(args, "toAnchorIdExclusive")),
        "delete_section" => DocxSessionOps.DeleteSection(session.Handle, Str(args, "headingAnchorId")),
        "split_paragraph" => DocxSessionOps.SplitParagraph(
            session.Handle, Str(args, "anchorId"), Int(args, "characterOffset")),
        "merge_paragraphs" => DocxSessionOps.MergeParagraphs(
            session.Handle, Str(args, "anchorId"), Str(args, "secondAnchorId")),
        "undo" => BoolResult(DocxSessionOps.Undo(session.Handle)),
        "redo" => BoolResult(DocxSessionOps.Redo(session.Handle)),
        _ => throw new McpToolException($"unknown docxodus_edit action: {action}"),
    };

    private static bool IsMutatingEditAction(string action) => action is not ("undo" or "redo");

    private static string BoolResult(bool value) => value ? "{\"success\":true}" : "{\"success\":false}";

    // ─── Format ─────────────────────────────────────────────────────────

    private static string Format(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunFormatAction(session, Str(args, "action"), args);
    }

    private static string RunFormatAction(DocSession session, string action, JsonElement args) => action switch
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
    };

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

    private static string RunCreateAction(DocSession session, string action, JsonElement args) => action switch
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
        _ => throw new McpToolException($"unknown docxodus_create action: {action}"),
    };

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

    private static string RunListAction(DocSession session, string action, JsonElement args) => action switch
    {
        "apply_format" => DocxSessionOps.ApplyListFormat(
            session.Handle, Str(args, "anchorId"), DocxSessionJson.ParseListFormat(OptStr(args, "listFormat"))),
        "set_level" => DocxSessionOps.SetListLevel(session.Handle, Str(args, "anchorId"), Int(args, "levelDelta")),
        "remove" => DocxSessionOps.RemoveListMembership(session.Handle, Str(args, "anchorId")),
        "get_membership" => DocxSessionOps.GetListMembership(session.Handle, Str(args, "anchorId")),
        _ => throw new McpToolException($"unknown docxodus_list action: {action}"),
    };

    private static bool IsMutatingListAction(string action) => action != "get_membership";

    // ─── Comment (native Word comments, issue #300) ────────────────────

    private static string Comment(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunCommentAction(session, Str(args, "action"), args);
    }

    private static string RunCommentAction(DocSession session, string action, JsonElement args) => action switch
    {
        "add" => DocxSessionOps.AddComment(
            session.Handle, Str(args, "anchorId"), ParseSpan(args, "span"),
            Str(args, "author"), OptStr(args, "initials"), OptStr(args, "date"),
            OptStr(args, "markdown") ?? ""),
        "update" => DocxSessionOps.UpdateComment(
            session.Handle, Str(args, "commentAnchorId"), Str(args, "markdown")),
        "remove" => DocxSessionOps.RemoveComment(session.Handle, Str(args, "commentAnchorId")),
        "list" => $"{{\"comments\":{DocxSessionOps.ListComments(session.Handle)}}}",
        _ => throw new McpToolException($"unknown docxodus_comment action: {action}"),
    };

    private static bool IsMutatingCommentAction(string action) => action != "list";

    // ─── Annotate (annotation overlay) ─────────────────────────────────

    private static string Annotate(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var action = Str(args, "action");
        switch (action)
        {
            case "add":
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
            case "update":
            {
                var updateEl = args.ValueKind == JsonValueKind.Object && args.TryGetProperty("update", out var u) && u.ValueKind == JsonValueKind.Object
                    ? u.GetRawText() : "{}";
                return DocxSessionOps.UpdateAnnotation(session.Handle, Str(args, "annotationId"), updateEl);
            }
            case "remove":
                return DocxSessionOps.RemoveAnnotation(session.Handle, Str(args, "annotationId"));
            case "move":
                return DocxSessionOps.MoveAnnotation(
                    session.Handle, Str(args, "annotationId"), Str(args, "newAnchorId"), ParseSpan(args, "newSpan"));
            case "list":
                return $"{{\"annotations\":{DocxSessionOps.ListAnnotations(session.Handle)}}}";
            case "find":
                return $"{{\"anchors\":{DocxSessionOps.FindByAnnotation(session.Handle, Str(args, "query"))}}}";
            default:
                throw new McpToolException($"unknown docxodus_annotate action: {action}");
        }
    }

    // ─── Track changes ──────────────────────────────────────────────────

    private static string TrackChanges(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var action = Str(args, "action");
        switch (action)
        {
            case "list":
            {
                var bytes = DocxSessionOps.Save(session.Handle);
                var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("session.docx", bytes));
                var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("session.docx", bytes));
                var revisionsJson = DocxDiffOps.GetRevisionsJson(rejected.DocumentByteArray, accepted.DocumentByteArray, null);
                return FilterRevisions(revisionsJson, OptStr(args, "author"), OptStr(args, "changeType"));
            }
            case "accept_all":
            {
                // SaveWithAnchorIds (not Save) so the transformed bytes still carry the
                // PtOpenXml:Unid attributes Rebind's reopen needs to keep anchor ids stable.
                var bytes = DocxSessionOps.SaveWithAnchorIds(session.Handle);
                var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("session.docx", bytes));
                store.Rebind(session, accepted.DocumentByteArray);
                return "{\"success\":true}";
            }
            case "reject_all":
            {
                var bytes = DocxSessionOps.SaveWithAnchorIds(session.Handle);
                var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("session.docx", bytes));
                store.Rebind(session, rejected.DocumentByteArray);
                return "{\"success\":true}";
            }
            default:
                throw new McpToolException($"unknown docxodus_track_changes action: {action}");
        }
    }

    private static string FilterRevisions(string revisionsJson, string? author, string? changeType)
    {
        if (author is null && changeType is null) return revisionsJson;
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
            kept.Add(r.GetRawText());
        }
        return "{\"revisions\":[" + string.Join(",", kept) + "]}";
    }

    // ─── Mutations (batch) ──────────────────────────────────────────────

    private static readonly HashSet<string> BatchableTools = new()
    {
        "docxodus_edit", "docxodus_format", "docxodus_create", "docxodus_table", "docxodus_list",
        "docxodus_comment",
    };

    private static string Mutations(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        var mode = Str(args, "mode");
        if (mode is not ("apply" or "preview"))
            throw new McpToolException($"unknown docxodus_mutations mode: {mode}");
        if (!args.TryGetProperty("steps", out var stepsEl) || stepsEl.ValueKind != JsonValueKind.Array)
            throw new McpToolException("docxodus_mutations requires an array \"steps\"");

        var results = new List<string>();
        var errors = new List<string>();
        int applied = 0;
        int mutatingSteps = 0;

        foreach (var step in stepsEl.EnumerateArray())
        {
            var stepTool = step.TryGetProperty("tool", out var toolEl) && toolEl.ValueKind == JsonValueKind.String
                ? toolEl.GetString()! : throw new McpToolException("mutation step missing string \"tool\"");
            if (!BatchableTools.Contains(stepTool))
                throw new McpToolException($"docxodus_mutations does not accept \"{stepTool}\" as a step (undo/redo and read-only actions are not batchable)");
            var stepArgs = step.TryGetProperty("args", out var a) && a.ValueKind == JsonValueKind.Object
                ? a : throw new McpToolException("mutation step missing object \"args\"");
            var stepAction = stepArgs.TryGetProperty("action", out var actEl) && actEl.ValueKind == JsonValueKind.String
                ? actEl.GetString()! : throw new McpToolException("mutation step args missing string \"action\"");

            bool isMutating = stepTool switch
            {
                "docxodus_edit" => IsMutatingEditAction(stepAction),
                "docxodus_list" => IsMutatingListAction(stepAction),
                "docxodus_comment" => IsMutatingCommentAction(stepAction),
                _ => true, // every docxodus_format/docxodus_create/docxodus_table action mutates
            };
            if (!isMutating)
                throw new McpToolException($"docxodus_mutations does not accept the read-only action \"{stepAction}\" on {stepTool}");

            string resultJson;
            try
            {
                resultJson = stepTool switch
                {
                    "docxodus_edit" => RunEditAction(session, stepAction, stepArgs),
                    "docxodus_format" => RunFormatAction(session, stepAction, stepArgs),
                    "docxodus_create" => RunCreateAction(session, stepAction, stepArgs),
                    "docxodus_table" => RunTableAction(session, stepAction, stepArgs),
                    "docxodus_list" => RunListAction(session, stepAction, stepArgs),
                    "docxodus_comment" => RunCommentAction(session, stepAction, stepArgs),
                    _ => throw new McpToolException($"unreachable: {stepTool}"),
                };
            }
            catch (McpToolException ex)
            {
                errors.Add(JsonRpcIo.JsonString(ex.Message));
                results.Add($"{{\"success\":false,\"error\":{{\"message\":{JsonRpcIo.JsonString(ex.Message)}}}}}");
                continue;
            }

            results.Add(resultJson);
            bool succeeded = true;
            try
            {
                using var rdoc = JsonDocument.Parse(resultJson);
                if (rdoc.RootElement.ValueKind == JsonValueKind.Object
                    && rdoc.RootElement.TryGetProperty("success", out var s) && s.ValueKind == JsonValueKind.False)
                    succeeded = false;
            }
            catch (JsonException) { /* non-EditResult shape (shouldn't happen for batchable tools); assume success */ }

            mutatingSteps++;
            if (succeeded) applied++;
            else
            {
                using var rdoc = JsonDocument.Parse(resultJson);
                errors.Add(rdoc.RootElement.TryGetProperty("error", out var e) ? e.GetRawText() : "\"step failed\"");
            }
        }

        if (mode == "preview")
        {
            for (int i = 0; i < mutatingSteps; i++)
                DocxSessionOps.Undo(session.Handle);
        }

        var status = errors.Count == 0 ? "ok" : applied == 0 ? "failed" : "partial";
        return "{\"status\":\"" + status + "\",\"editsApplied\":" + applied
            + ",\"results\":[" + string.Join(",", results) + "]"
            + ",\"errors\":[" + string.Join(",", errors) + "]}";
    }

    // ─── Table ──────────────────────────────────────────────────────────

    private static string Table(SessionStore store, JsonElement args)
    {
        var session = Session(store, args);
        return RunTableAction(session, Str(args, "action"), args);
    }

    private static string RunTableAction(DocSession session, string action, JsonElement args) => action switch
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
        _ => throw new McpToolException($"unknown docxodus_table action: {action}"),
    };

    // ─── Arg helpers ────────────────────────────────────────────────────

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

    private static int Int(JsonElement args, string name)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Number)
            throw new McpToolException($"missing required number argument \"{name}\"");
        return v.GetInt32();
    }

    private static int IntOpt(JsonElement args, string name, int fallback) =>
        args.ValueKind == JsonValueKind.Object && args.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
            ? v.GetInt32() : fallback;

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
