#nullable enable

using System;
using System.IO;
using System.Text.Json;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the <c>tools/mcp-server</c> tool surface's <see cref="Dispatcher"/>, exercised
/// directly (no stdio transport involved — <c>Program.cs</c> is a thin JSON-RPC wrapper around
/// this). Test IDs follow the <c>MCP###</c> prefix convention. Each test opens a session over a
/// throwaway blank document (<see cref="DocxSession.CreateBlankDocxBytes"/>) written to a temp
/// file, exactly as a real client would via <c>docxodus_open</c>'s <c>path</c> argument.
/// </summary>
public class McpServerDispatcherTests : IDisposable
{
    private readonly string _tempPath;
    private readonly SessionStore _store = new();

    public McpServerDispatcherTests()
    {
        _tempPath = Path.Combine(Path.GetTempPath(), $"mcp-dispatcher-test-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(_tempPath, DocxSession.CreateBlankDocxBytes());
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (File.Exists(_tempPath)) File.Delete(_tempPath);
    }

    private static JsonElement J(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static JsonElement Parse(string json) => J(json);

    private string OpenSession(string? trackedChanges = null)
    {
        var argsJson = trackedChanges is null
            ? $$"""{"path":{{JsonSerializer.Serialize(_tempPath)}}}"""
            : $$"""{"path":{{JsonSerializer.Serialize(_tempPath)}},"trackedChanges":{{JsonSerializer.Serialize(trackedChanges)}}}""";
        var result = Dispatcher.Call(_store, "docxodus_open", J(argsJson));
        return Parse(result).GetProperty("sessionId").GetString()!;
    }

    private static string FirstBodyAnchorId(string sessionId, SessionStore store)
    {
        var content = Dispatcher.Call(store, "docxodus_get_content",
            J($$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"format":"blocks"}"""));
        using var doc = JsonDocument.Parse(content);
        foreach (var prop in doc.RootElement.GetProperty("blocks").EnumerateObject())
            return prop.Name;
        throw new InvalidOperationException("blank document has no addressable blocks");
    }

    // ─── Lifecycle ──────────────────────────────────────────────────────

    [Fact]
    public void MCP001_OpenSaveClose_RoundTrips()
    {
        var sessionId = OpenSession();
        Assert.False(string.IsNullOrEmpty(sessionId));

        var savePath = _tempPath + ".out.docx";
        try
        {
            var saveResult = Dispatcher.Call(_store, "docxodus_save",
                J($$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"path":{{JsonSerializer.Serialize(savePath)}}}"""));
            var saved = Parse(saveResult);
            Assert.Equal(savePath, saved.GetProperty("path").GetString());
            Assert.True(saved.GetProperty("bytesWritten").GetInt32() > 0);
            Assert.True(File.Exists(savePath));
        }
        finally
        {
            if (File.Exists(savePath)) File.Delete(savePath);
        }

        var closeResult = Dispatcher.Call(_store, "docxodus_close", J($$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}}}"""));
        Assert.True(Parse(closeResult).GetProperty("closed").GetBoolean());

        var ex = Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"format":"markdown"}""")));
        Assert.Contains("unknown session_id", ex.Message);
    }

    [Fact]
    public void MCP002_Open_UnreadablePath_ThrowsToolException()
    {
        var missingPath = Path.Combine(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.docx");
        Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_open", J($$"""{"path":{{JsonSerializer.Serialize(missingPath)}}}""")));
    }

    // ─── Content ────────────────────────────────────────────────────────

    [Fact]
    public void MCP010_GetContent_AllFormats_Succeed()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);

        foreach (var format in new[] { "markdown", "html", "text", "blocks", "info" })
        {
            var result = Dispatcher.Call(_store, "docxodus_get_content",
                J($$"""{"sessionId":{{sessionArg}},"format":"{{format}}"}"""));
            Assert.False(string.IsNullOrEmpty(result));
            using var doc = JsonDocument.Parse(result); // must be valid JSON
            Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
        }
    }

    // ─── Edit ───────────────────────────────────────────────────────────

    [Fact]
    public void MCP020_Edit_InsertReplaceUndoRedo()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var insertResult = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_paragraph","anchorId":"{{anchor}}","position":"after","markdown":"Hello world"}""")));
        Assert.True(insertResult.GetProperty("success").GetBoolean());
        var createdAnchor = insertResult.GetProperty("created")[0].GetProperty("id").GetString()!;

        var replaceResult = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{createdAnchor}}","markdown":"Goodbye world"}""")));
        Assert.True(replaceResult.GetProperty("success").GetBoolean());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("Goodbye world", md);

        var undo = Parse(Dispatcher.Call(_store, "docxodus_edit", J($$"""{"sessionId":{{sessionArg}},"action":"undo"}""")));
        Assert.True(undo.GetProperty("success").GetBoolean());
        md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("Hello world", md);

        var redo = Parse(Dispatcher.Call(_store, "docxodus_edit", J($$"""{"sessionId":{{sessionArg}},"action":"redo"}""")));
        Assert.True(redo.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP021_Edit_DeleteBlock_UnknownAnchor_ReturnsFailedEditResult()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var result = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"delete_block","anchorId":"{#p:body:doesnotexist000000000000000000}"}""")));
        Assert.False(result.GetProperty("success").GetBoolean());
        Assert.Equal("anchor_not_found", result.GetProperty("error").GetProperty("code").GetString());
    }

    // ─── Search ─────────────────────────────────────────────────────────

    [Fact]
    public void MCP030_Search_TextMode_FindsInsertedText()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_paragraph","anchorId":"{{anchor}}","position":"after","markdown":"findable needle text"}"""));

        var found = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"text","query":"needle"}""")));
        Assert.True(found.GetProperty("matches").GetArrayLength() > 0);
    }

    [Fact]
    public void MCP031_Search_KindMode_FindsParagraphs()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var found = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"kind","query":"p"}""")));
        Assert.True(found.GetProperty("matches").GetArrayLength() > 0);
    }

    // ─── Format / List ──────────────────────────────────────────────────

    [Fact]
    public void MCP040_Format_ApplyFormat_SetsBold()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"bold me"}"""));

        var result = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""{"sessionId":{{sessionArg}},"action":"apply_format","anchorId":"{{anchor}}","format":{"bold":true} }""")));
        Assert.True(result.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP041_List_ApplyFormatThenSetLevel_ProducesRealNumbering()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"item one"}"""));

        var applied = Parse(Dispatcher.Call(_store, "docxodus_list", J(
            $$"""{"sessionId":{{sessionArg}},"action":"apply_format","anchorId":"{{anchor}}","listFormat":"bullet"}""")));
        Assert.True(applied.GetProperty("success").GetBoolean());

        var membership = Parse(Dispatcher.Call(_store, "docxodus_list", J(
            $$"""{"sessionId":{{sessionArg}},"action":"get_membership","anchorId":"{{anchor}}"}""")));
        Assert.True(membership.GetProperty("isAutoNumbered").GetBoolean());
    }

    // ─── Create / Table ─────────────────────────────────────────────────

    [Fact]
    public void MCP050_Create_InsertHeading()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var result = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_heading","anchorId":"{{anchor}}","position":"after","text":"A Heading","level":2}""")));
        Assert.True(result.GetProperty("success").GetBoolean());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("A Heading", md);
        Assert.Contains("##", md);
    }

    [Fact]
    public void MCP060_Table_InsertRowAndReplaceCellContent()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var inserted = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert","anchorId":"{{anchor}}","position":"after","rows":2,"columns":2}""")));
        Assert.True(inserted.GetProperty("success").GetBoolean());

        // Two Docxodus ops address "the same cell" with two different anchor kinds:
        // ReplaceCellContent requires the "tc" (cell) anchor itself (not returned by InsertTable's
        // Created list — only the cell-paragraph anchors are — so it's found via search), while
        // row/column ops require a "p" (paragraph-inside-the-cell) anchor from Created directly.
        // See the docxodus_table schema note. Use the LAST "p" anchor (a different cell than the
        // one replace_cell_content below rewrites) for insert_row — replacing a cell's content
        // removes and recreates its paragraph, invalidating any anchor into that same cell.
        string? pAnchor = null;
        foreach (var created in inserted.GetProperty("created").EnumerateArray())
        {
            if (created.GetProperty("kind").GetString() == "p") pAnchor = created.GetProperty("id").GetString();
        }
        Assert.NotNull(pAnchor);

        var tcSearch = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"kind","query":"tc"}""")));
        var tcAnchor = tcSearch.GetProperty("matches")[0].GetProperty("id").GetString()!;

        var replaced = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_cell_content","cellAnchorId":"{{tcAnchor}}","markdown":"cell text"}""")));
        Assert.True(replaced.GetProperty("success").GetBoolean());

        var rowAddedJson = Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_row","cellAnchorId":"{{pAnchor}}","position":"after"}"""));
        var rowAdded = Parse(rowAddedJson);
        Assert.True(rowAdded.GetProperty("success").GetBoolean(), rowAddedJson);
    }

    // ─── Comment (annotation overlay) ──────────────────────────────────

    [Fact]
    public void MCP070_Comment_AddListRemove()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"annotate this"}"""));

        var added = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","anchorId":"{{anchor}}","label":"note","labelId":"NOTE","color":"#FFEB3B"}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        var annotationId = added.GetProperty("annotationId").GetString()!;

        var listed = Parse(Dispatcher.Call(_store, "docxodus_comment", J($$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.True(listed.GetProperty("annotations").GetArrayLength() > 0);

        var removed = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove","annotationId":"{{annotationId}}"}""")));
        Assert.True(removed.GetProperty("success").GetBoolean());
    }

    // ─── Track changes ──────────────────────────────────────────────────

    [Fact]
    public void MCP080_TrackChanges_ListThenAcceptAll()
    {
        var sessionId = OpenSession(trackedChanges: "render_inline");
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"tracked edit"}"""));

        var listed = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J($$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.True(listed.GetProperty("revisions").GetArrayLength() > 0);

        var accepted = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J($$"""{"sessionId":{{sessionArg}},"action":"accept_all"}""")));
        Assert.True(accepted.GetProperty("success").GetBoolean());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("tracked edit", md);

        var listedAfterAccept = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J($$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.Equal(0, listedAfterAccept.GetProperty("revisions").GetArrayLength());
    }

    [Fact]
    public void MCP081_TrackChanges_RejectAll_RestoresOriginal()
    {
        var sessionId = OpenSession(trackedChanges: "render_inline");
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"should be reverted"}"""));

        var rejected = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J($$"""{"sessionId":{{sessionArg}},"action":"reject_all"}""")));
        Assert.True(rejected.GetProperty("success").GetBoolean());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.DoesNotContain("should be reverted", md);
    }

    // ─── Mutations (batch) ──────────────────────────────────────────────

    [Fact]
    public void MCP090_Mutations_ApplyMode_AppliesAllSteps()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "apply",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "batched one" } },
                { "tool": "docxodus_format", "args": { "action": "apply_format", "anchorId": "{{anchor}}", "format": { "bold": true } } }
              ]
            }
            """)));
        Assert.Equal("ok", batch.GetProperty("status").GetString());
        Assert.Equal(2, batch.GetProperty("editsApplied").GetInt32());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("batched one", md);
    }

    [Fact]
    public void MCP091_Mutations_PreviewMode_LeavesDocumentUnchanged()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var before = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "preview",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "should not stick" } }
              ]
            }
            """)));
        Assert.Equal("ok", batch.GetProperty("status").GetString());

        var after = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Equal(before, after);
    }

    [Fact]
    public void MCP092_Mutations_RejectsUndoRedoAsSteps()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var ex = Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"apply","steps":[{"tool":"docxodus_edit","args":{"action":"undo"} }]}""")));
        Assert.Contains("undo", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ─── Tool catalog ───────────────────────────────────────────────────

    [Fact]
    public void MCP100_ToolCatalog_HasThirteenDistinctNamedToolsWithValidSchemas()
    {
        Assert.Equal(13, ToolCatalog.Tools.Count);
        var names = new System.Collections.Generic.HashSet<string>();
        foreach (var tool in ToolCatalog.Tools)
        {
            Assert.StartsWith("docxodus_", tool.Name);
            Assert.False(string.IsNullOrWhiteSpace(tool.Description));
            Assert.True(names.Add(tool.Name), $"duplicate tool name: {tool.Name}");
            using var schema = JsonDocument.Parse(tool.InputSchemaJson); // must be valid JSON
            Assert.Equal("object", schema.RootElement.GetProperty("type").GetString());
        }
    }

    // ─── Unknown tool / action ──────────────────────────────────────────

    [Fact]
    public void MCP110_UnknownTool_ThrowsToolException()
    {
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_nonexistent", J("{}")));
    }

    [Fact]
    public void MCP111_UnknownAction_ThrowsToolException()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"not_a_real_action"}""")));
    }
}
