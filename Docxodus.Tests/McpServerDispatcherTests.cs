#nullable enable

using System;
using System.IO;
using System.Linq;
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
    private readonly string _root;
    private readonly string _tempPath;
    private readonly SessionStore _store;

    public McpServerDispatcherTests()
    {
        // Each test class instance gets its own scope root, so the dispatcher runs against a
        // realistically-confined store rather than an unbounded filesystem.
        _root = Path.Combine(Path.GetTempPath(), $"mcp-dispatcher-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));

        _tempPath = Path.Combine(_root, "document.docx");
        File.WriteAllBytes(_tempPath, DocxSession.CreateBlankDocxBytes());
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private static JsonElement J(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static JsonElement Parse(string json) => J(json);

    private string OpenSession(string? trackedChanges = null, bool? persistAnchorIds = null, string? path = null)
    {
        var argsJson = $$"""{"path":{{JsonSerializer.Serialize(path ?? _tempPath)}}""";
        if (trackedChanges is not null)
            argsJson += $$""","trackedChanges":{{JsonSerializer.Serialize(trackedChanges)}}""";
        if (persistAnchorIds is not null)
            argsJson += $$""","persistAnchorIds":{{(persistAnchorIds.Value ? "true" : "false")}}""";
        argsJson += "}";
        var result = Dispatcher.Call(_store, "docxodus_open", J(argsJson));
        return Parse(result).GetProperty("sessionId").GetString()!;
    }

    /// <summary>Insert a paragraph after the document's first block and return the created
    /// paragraph's anchor id — a fresh (randomly-assigned) Unid, which is exactly the kind of
    /// anchor that cannot survive a save→reopen unless the save persists anchor bookkeeping.</summary>
    private string InsertParagraph(string sessionId, string markdown)
    {
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var result = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"insert_paragraph","anchorId":"{{anchor}}","position":"after","markdown":{{JsonSerializer.Serialize(markdown)}}}""")));
        Assert.True(result.GetProperty("success").GetBoolean());
        return result.GetProperty("created")[0].GetProperty("id").GetString()!;
    }

    private void Save(string sessionId, string path, bool? persistAnchorIds = null)
    {
        var argsJson = $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"path":{{JsonSerializer.Serialize(path)}}""";
        if (persistAnchorIds is not null)
            argsJson += $$""","persistAnchorIds":{{(persistAnchorIds.Value ? "true" : "false")}}""";
        argsJson += "}";
        Dispatcher.Call(_store, "docxodus_save", J(argsJson));
    }

    private static JsonElement ReplaceText(SessionStore store, string sessionId, string anchor, string markdown) =>
        Parse(Dispatcher.Call(store, "docxodus_edit", J(
            $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"replace_text","anchorId":"{{anchor}}","markdown":{{JsonSerializer.Serialize(markdown)}}}""")));

    /// <summary>The saved file's main document part XML — where persisted <c>PtOpenXml:Unid</c>
    /// anchor bookkeeping shows up as <c>Unid="…"</c> attributes.</summary>
    private static string SavedDocumentXml(string path)
    {
        using var ms = new MemoryStream(File.ReadAllBytes(path));
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.RootElement!.OuterXml;
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
        var missingPath = Path.Combine(_root, $"does-not-exist-{Guid.NewGuid():N}.docx");
        Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_open", J($$"""{"path":{{JsonSerializer.Serialize(missingPath)}}}""")));
    }

    [Fact]
    public void MCP003_SessionIds_AreUnguessableAndDistinct()
    {
        var first = OpenSession();
        var second = OpenSession();

        Assert.NotEqual(first, second);
        foreach (var id in new[] { first, second })
        {
            Assert.StartsWith("s_", id);
            Assert.Equal(2 + 32, id.Length);            // "s_" + 16 random bytes as hex
            Assert.Matches("^s_[0-9a-f]{32}$", id);
        }
    }

    [Fact]
    public void MCP004_Open_PersistAnchorIds_KeepsCreatedAnchorAcrossSaveReopen()
    {
        var sessionId = OpenSession(persistAnchorIds: true);
        var createdAnchor = InsertParagraph(sessionId, "persist me");

        var savedPath = Path.Combine(_root, "persisted.docx");
        Save(sessionId, savedPath);

        var reopened = OpenSession(path: savedPath);
        var replace = ReplaceText(_store, reopened, createdAnchor, "still addressable");
        Assert.True(replace.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP005_Save_PersistAnchorIdsOverride_KeepsAnchorFromDefaultSession()
    {
        var sessionId = OpenSession();                     // default: anchor ids NOT persisted
        var createdAnchor = InsertParagraph(sessionId, "checkpoint me");

        var savedPath = Path.Combine(_root, "checkpoint.docx");
        Save(sessionId, savedPath, persistAnchorIds: true);

        var reopened = OpenSession(path: savedPath);
        var replace = ReplaceText(_store, reopened, createdAnchor, "still addressable");
        Assert.True(replace.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP006_Open_PersistAnchorIds_GovernsPlainSave()
    {
        var sessionId = OpenSession(persistAnchorIds: true);
        InsertParagraph(sessionId, "bookkeeping should survive");

        var savedPath = Path.Combine(_root, "with-unids.docx");
        Save(sessionId, savedPath);

        Assert.Contains("Unid=", SavedDocumentXml(savedPath));
    }

    [Fact]
    public void MCP007_Save_PersistAnchorIdsFalse_StripsOnPersistTrueSession()
    {
        var sessionId = OpenSession(persistAnchorIds: true);
        InsertParagraph(sessionId, "clean deliverable");

        var savedPath = Path.Combine(_root, "clean.docx");
        Save(sessionId, savedPath, persistAnchorIds: false);

        Assert.DoesNotContain("Unid=", SavedDocumentXml(savedPath));
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

    [Fact]
    public void MCP136_List_ApplyFormatRange_RomanParenthesis_SharedInstance()
    {
        // Issue #313: one apply_format_range call converts a whole contiguous run — with a
        // legal-drafting "(i)" preset — instead of one apply_format call per item.
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var first = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{first}}","markdown":"item one"}"""));
        var second = InsertParagraph(sessionId, "item two");

        var applied = Parse(Dispatcher.Call(_store, "docxodus_list", J(
            $$"""{"sessionId":{{sessionArg}},"action":"apply_format_range","firstAnchorId":"{{first}}","lastAnchorId":"{{second}}","listFormat":"lowerRomanParenthesis"}""")));
        Assert.True(applied.GetProperty("success").GetBoolean());
        Assert.Equal(2, applied.GetProperty("modified").GetArrayLength());

        var secondLi = applied.GetProperty("modified")[1].GetProperty("id").GetString()!;
        var membership = Parse(Dispatcher.Call(_store, "docxodus_list", J(
            $$"""{"sessionId":{{sessionArg}},"action":"get_membership","anchorId":"{{secondLi}}"}""")));
        Assert.True(membership.GetProperty("isAutoNumbered").GetBoolean());
        Assert.Equal("(ii)", membership.GetProperty("generatedLabel").GetString());
    }

    [Fact]
    public void MCP042_Format_SetParagraphFormat_AddsAndClearsBorder()
    {
        // Issue #301: the schema used to only expose clearBorders, so an agent could remove a
        // paragraph border but never add one to an EXISTING paragraph. topBorder/bottomBorder
        // are visible in the rendered HTML (a border-* style on the wrapping div), so that's
        // the black-box signal this test checks — the same one every other MCP04x format test
        // relies on `success` for, just with an observable side effect for a border specifically.
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var applied = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""
            {"sessionId":{{sessionArg}},"action":"set_paragraph_format","anchorId":"{{anchor}}","paragraphFormat":{"bottomBorder":{"style":"single","size":12,"color":"auto"} } }
            """)));
        Assert.True(applied.GetProperty("success").GetBoolean());

        var htmlWithBorder = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"html"}""")))
            .GetProperty("html").GetString()!;
        Assert.Contains("border-bottom", htmlWithBorder);

        var cleared = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_paragraph_format","anchorId":"{{anchor}}","paragraphFormat":{"clearBorders":true} }""")));
        Assert.True(cleared.GetProperty("success").GetBoolean());

        var htmlAfterClear = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"html"}""")))
            .GetProperty("html").GetString()!;
        Assert.DoesNotContain("border-bottom", htmlAfterClear);
    }

    [Fact]
    public void MCP043_Format_SetParagraphFormat_FirstLineIndentAndSpacing()
    {
        // Issue #312: firstLineIndent/hangingIndent + spacing were inexpressible — the nearest
        // reachable op was indentDelta (whole-left-edge shift), a visibly different result.
        // firstLine renders as a text-indent style on the wrapping div, the observable signal.
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var applied = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""
            {"sessionId":{{sessionArg}},"action":"set_paragraph_format","anchorId":"{{anchor}}","paragraphFormat":{"firstLineIndent":720,"spacingBefore":240,"spacingAfter":120,"lineSpacing":360} }
            """)));
        Assert.True(applied.GetProperty("success").GetBoolean());

        var html = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"html"}""")))
            .GetProperty("html").GetString()!;
        Assert.Contains("text-indent: 0.50in", html);

        // Both firstLine and hanging in one op is unrepresentable in w:ind — typed error.
        var both = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""
            {"sessionId":{{sessionArg}},"action":"set_paragraph_format","anchorId":"{{anchor}}","paragraphFormat":{"firstLineIndent":720,"hangingIndent":360} }
            """)));
        Assert.False(both.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_paragraph_format", both.GetProperty("error").GetProperty("code").GetString());
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

    // ─── Annotate (annotation overlay) ─────────────────────────────────

    [Fact]
    public void MCP070_Annotate_AddListRemove()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"annotate this"}"""));

        var added = Parse(Dispatcher.Call(_store, "docxodus_annotate", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","anchorId":"{{anchor}}","label":"note","labelId":"NOTE","color":"#FFEB3B"}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        var annotationId = added.GetProperty("annotationId").GetString()!;

        var listed = Parse(Dispatcher.Call(_store, "docxodus_annotate", J($$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.True(listed.GetProperty("annotations").GetArrayLength() > 0);

        var removed = Parse(Dispatcher.Call(_store, "docxodus_annotate", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove","annotationId":"{{annotationId}}"}""")));
        Assert.True(removed.GetProperty("success").GetBoolean());
    }

    // ─── Comment (native Word comments, issue #300) ────────────────────

    [Fact]
    public void MCP071_Comment_AddUpdateListRemove_IsNative()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"comment on this"}"""));

        var added = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","anchorId":"{{anchor}}","author":"Alice","initials":"AL","markdown":"Needs review."}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        var cmtAnchor = added.GetProperty("created").EnumerateArray()
            .First(a => a.GetProperty("kind").GetString() == "cmt")
            .GetProperty("id").GetString()!;

        // The comment is a real w:comment: the saved bytes carry a comments part entry.
        var session = _store.Get(sessionId);
        var savedBytes = Docxodus.Internal.DocxSessionOps.Save(session.Handle);
        using (var ms = new System.IO.MemoryStream(savedBytes))
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false))
        {
            var part = doc.MainDocumentPart!.WordprocessingCommentsPart;
            Assert.NotNull(part);
            System.Xml.Linq.XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var comment = Assert.Single(part!.GetXDocument().Root!.Elements(w + "comment"));
            Assert.Equal("Alice", (string?)comment.Attribute(w + "author"));
        }

        var updated = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"update","commentAnchorId":"{{cmtAnchor}}","markdown":"Revised."}""")));
        Assert.True(updated.GetProperty("success").GetBoolean());

        var listed = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var entry = Assert.Single(listed.GetProperty("comments").EnumerateArray().ToList());
        Assert.Equal("Alice", entry.GetProperty("author").GetString());
        Assert.Equal("AL", entry.GetProperty("initials").GetString());
        Assert.Equal("Revised.", entry.GetProperty("text").GetString());

        var removed = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove","commentAnchorId":"{{cmtAnchor}}"}""")));
        Assert.True(removed.GetProperty("success").GetBoolean());
        var emptied = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.Equal(0, emptied.GetProperty("comments").GetArrayLength());
    }

    [Fact]
    public void MCP072_Mutations_AcceptsCommentAddStep()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"batched comment target"}"""));

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "apply",
              "steps": [
                { "tool": "docxodus_comment", "args": { "action": "add", "anchorId": "{{anchor}}", "author": "Bot", "markdown": "From a batch." } }
              ]
            }
            """)));
        Assert.Equal("ok", batch.GetProperty("status").GetString());
        Assert.Equal(1, batch.GetProperty("editsApplied").GetInt32());

        var listed = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var entry = Assert.Single(listed.GetProperty("comments").EnumerateArray().ToList());
        Assert.Equal("Bot", entry.GetProperty("author").GetString());

        // The read-only list action is rejected as a batch step.
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "apply",
              "steps": [
                { "tool": "docxodus_comment", "args": { "action": "list" } }
              ]
            }
            """)));
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
    public void MCP100_ToolCatalog_HasFourteenDistinctNamedToolsWithValidSchemas()
    {
        Assert.Equal(14, ToolCatalog.Tools.Count);
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

    // ─── Document store: scoping and isolation ─────────────────────────

    [Fact]
    public void MCP120_Store_ResolvesRelativeLocationUnderRoot()
    {
        var resolved = _store.Documents.Resolve("document.docx");
        Assert.Equal(Path.Combine(_root, "document.docx"), resolved);
    }

    [Fact]
    public void MCP121_Store_AcceptsAbsolutePathInsideRoot()
    {
        // The property that keeps ordinary local use working: an agent may name a file by its
        // natural absolute path, so long as the configured scope contains it.
        Assert.Equal(_tempPath, _store.Documents.Resolve(_tempPath));
    }

    [Fact]
    public void MCP122_Store_RejectsAbsolutePathOutsideRoot()
    {
        var outside = Path.Combine(Path.GetTempPath(), $"outside-{Guid.NewGuid():N}.docx");
        var ex = Assert.Throws<McpToolException>(() => _store.Documents.Resolve(outside));
        Assert.Contains("outside this server's document scope", ex.Message);
    }

    [Fact]
    public void MCP123_Store_RejectsParentTraversal()
    {
        Assert.Throws<McpToolException>(() => _store.Documents.Resolve(Path.Combine("..", "escaped.docx")));
        Assert.Throws<McpToolException>(() =>
            _store.Documents.Resolve(Path.Combine("sub", "..", "..", "escaped.docx")));
    }

    [Fact]
    public void MCP124_Store_RejectsSiblingRootSharingAPrefix()
    {
        // Segment-boundary check: "<root>-sibling" starts with the root as a raw string but is
        // not inside it.
        Assert.Throws<McpToolException>(() => _store.Documents.Resolve(_root + "-sibling/doc.docx"));
    }

    [Fact]
    public void MCP125_Store_RejectsSymlinkEscapingRoot()
    {
        var outsideDir = Path.Combine(Path.GetTempPath(), $"mcp-outside-{Guid.NewGuid():N}");
        Directory.CreateDirectory(outsideDir);
        var secret = Path.Combine(outsideDir, "secret.docx");
        File.WriteAllBytes(secret, DocxSession.CreateBlankDocxBytes());

        var link = Path.Combine(_root, "link");
        try
        {
            Directory.CreateSymbolicLink(link, outsideDir);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
        {
            return; // environment doesn't permit symlink creation; nothing to assert
        }

        try
        {
            // Lexically this is inside the root; only following the link reveals that it isn't.
            var ex = Assert.Throws<McpToolException>(() =>
                _store.Documents.Resolve(Path.Combine("link", "secret.docx")));
            Assert.Contains("outside this server's document scope", ex.Message);
        }
        finally
        {
            Directory.Delete(link);
            Directory.Delete(outsideDir, recursive: true);
        }
    }

    [Fact]
    public void MCP126_Store_ReadWriteRoundTrips()
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        var location = _store.Documents.Resolve(Path.Combine("nested", "created.docx"));

        _store.Documents.Write(location, bytes);       // creates the intermediate directory
        Assert.Equal(bytes, _store.Documents.Read(location));
    }

    [Fact]
    public void MCP127_Open_OutsideScope_IsRejectedBeforeReading()
    {
        // A readable file the store must still refuse, proving the rejection is the scope check
        // rather than an incidental IO failure.
        var outside = Path.Combine(Path.GetTempPath(), $"mcp-outside-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(outside, DocxSession.CreateBlankDocxBytes());
        try
        {
            var ex = Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_open",
                J($$"""{"path":{{JsonSerializer.Serialize(outside)}}}""")));
            Assert.Contains("outside this server's document scope", ex.Message);
        }
        finally
        {
            File.Delete(outside);
        }
    }

    [Fact]
    public void MCP128_Save_ToLocationOutsideScope_IsRejected()
    {
        var sessionId = OpenSession();
        var outside = Path.Combine(Path.GetTempPath(), $"mcp-outside-{Guid.NewGuid():N}.docx");
        var ex = Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_save", J(
            $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"path":{{JsonSerializer.Serialize(outside)}}}""")));
        Assert.Contains("outside this server's document scope", ex.Message);
        Assert.False(File.Exists(outside));
    }

    [Fact]
    public void MCP129_DocumentStores_ScopeSegmentNestsUnderRoot()
    {
        var scoped = DocumentStores.Create(backend: null, root: _root, scope: "tenant-42");
        Assert.Equal(Path.Combine(_root, "tenant-42"), scoped.RootDescription);

        // Two scopes under one root cannot reach each other.
        var other = DocumentStores.Create(backend: null, root: _root, scope: "tenant-7");
        var otherDoc = Path.Combine(scoped.RootDescription, "doc.docx");
        Assert.Throws<McpToolException>(() => other.Resolve(otherDoc));
    }

    [Theory]
    [InlineData("../escape")]
    [InlineData("/absolute")]
    [InlineData("nested/segment")]
    public void MCP130_DocumentStores_RejectsUnsafeScopeSegment(string scope)
    {
        Assert.Throws<McpToolException>(() => DocumentStores.Create(backend: null, root: _root, scope: scope));
    }

    [Fact]
    public void MCP131_DocumentStores_RejectsUnknownBackend()
    {
        var ex = Assert.Throws<McpToolException>(() =>
            DocumentStores.Create(backend: "s3", root: _root, scope: null));
        Assert.Contains("unsupported", ex.Message, StringComparison.OrdinalIgnoreCase);
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

    // ─── Wire framing ─────────────────────────────────────────────────

    [Fact]
    public void MCP132_ToolsListResult_IsSingleLineNdjson()
    {
        // ToolCatalog's InputSchemaJson entries are pretty-printed C# raw-string literals (real
        // embedded newlines) for source readability. If Program spliced one in verbatim rather
        // than re-serializing it, the tools/list response — normally one JSON-RPC message per
        // line per the MCP stdio transport — would spill across multiple physical lines and break
        // any strict NDJSON reader (e.g. a client using readline-based framing).
        var result = Program.BuildToolsListResult();

        Assert.DoesNotContain('\n', result);
        Assert.DoesNotContain('\r', result);

        using var doc = JsonDocument.Parse(result); // still valid, semantically unchanged JSON
        var tools = doc.RootElement.GetProperty("tools");
        Assert.True(tools.GetArrayLength() >= 13);
        foreach (var tool in tools.EnumerateArray())
            Assert.Equal(JsonValueKind.Object, tool.GetProperty("inputSchema").ValueKind);
    }

    // ─── docxodus_track_changes set_mode (issue #304) ───────────────────

    private JsonElement SetMode(string sessionId, string mode, string? author = null)
    {
        var args = $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"set_mode","mode":{{JsonSerializer.Serialize(mode)}}""";
        if (author is not null)
            args += $$""","revisionAuthor":{{JsonSerializer.Serialize(author)}}""";
        args += "}";
        return Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(args)));
    }

    [Fact]
    public void MCP133_SetMode_SwitchesRecordingMidSession()
    {
        var sessionId = OpenSession(); // default: accept
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var direct = ReplaceText(_store, sessionId, anchor, "Direct edit.");
        Assert.True(direct.GetProperty("success").GetBoolean());

        var mode = SetMode(sessionId, "render_inline", "mcp-reviewer");
        Assert.True(mode.GetProperty("success").GetBoolean());

        var tracked = ReplaceText(_store, sessionId, anchor, "Tracked edit.");
        Assert.True(tracked.GetProperty("success").GetBoolean());

        var outPath = Path.Combine(_root, "tracked-out.docx");
        Save(sessionId, outPath);
        var xml = SavedDocumentXml(outPath);
        Assert.Contains("w:ins", xml);
        Assert.Contains("mcp-reviewer", xml);
    }

    [Fact]
    public void MCP134_SetMode_UnknownModeThrows()
    {
        var sessionId = OpenSession();
        var ex = Assert.Throws<McpToolException>(() => SetMode(sessionId, "bogus"));
        Assert.Contains("unknown trackedChanges mode", ex.Message);
    }

    [Fact]
    public void MCP135_SetMode_EchoAndAuthorSemantics()
    {
        var sessionId = OpenSession();

        var r1 = SetMode(sessionId, "render_inline", "Reviewer A");
        Assert.Equal("render_inline", r1.GetProperty("trackedChanges").GetString());
        Assert.Equal("Reviewer A", r1.GetProperty("revisionAuthor").GetString());

        // Absent revisionAuthor leaves the author unchanged.
        var r2 = SetMode(sessionId, "accept");
        Assert.Equal("accept", r2.GetProperty("trackedChanges").GetString());
        Assert.Equal("Reviewer A", r2.GetProperty("revisionAuthor").GetString());

        // Empty string resets to the default (null).
        var r3 = SetMode(sessionId, "accept", "");
        Assert.Equal(JsonValueKind.Null, r3.GetProperty("revisionAuthor").ValueKind);
    }

    // ─── docxodus_track_changes selective accept/reject (issue #318) ────

    [Fact]
    public void MCP136_TrackChanges_SelectiveAcceptAndRejectByRevisionId()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        // Seed real text directly (the blank fixture's paragraph is empty — a tracked
        // edit on it would produce only an insertion), THEN record a tracked rewrite.
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"original text"}"""));
        SetMode(sessionId, "render_inline", "Reviewer A");
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"selective edit"}"""));

        // Markup-native listing: stable ids + the markup's own author.
        var listed = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var revisions = listed.GetProperty("revisions").EnumerateArray().ToList();
        Assert.Equal(2, revisions.Count);
        var deleteRev = Assert.Single(revisions, r => r.GetProperty("type").GetString() == "delete");
        var insertRev = Assert.Single(revisions, r => r.GetProperty("type").GetString() == "insert");
        Assert.Equal("selective edit", insertRev.GetProperty("text").GetString());
        Assert.Equal("Reviewer A", insertRev.GetProperty("author").GetString());
        Assert.StartsWith("rev", insertRev.GetProperty("id").GetString());

        // Accept the insertion; the deletion keeps its id and resolves independently.
        var accepted = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"accept","revisionId":{{insertRev.GetProperty("id").GetRawText()}}}""")));
        Assert.True(accepted.GetProperty("success").GetBoolean());

        var accepted2 = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"accept","revisionId":{{deleteRev.GetProperty("id").GetRawText()}}}""")));
        Assert.True(accepted2.GetProperty("success").GetBoolean());

        var after = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.Equal(0, after.GetProperty("revisions").GetArrayLength());

        var md = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("selective edit", md);
        Assert.DoesNotContain("original text", md);

        // Now record another tracked rewrite and REJECT both halves — the insertion's
        // text goes away and the deletion's text is restored.
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"should not stick"}"""));
        foreach (var rev in Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
                $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("revisions").EnumerateArray())
        {
            var rejected = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
                $$"""{"sessionId":{{sessionArg}},"action":"reject","revisionId":{{rev.GetProperty("id").GetRawText()}}}""")));
            Assert.True(rejected.GetProperty("success").GetBoolean());
        }
        var mdAfterReject = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("selective edit", mdAfterReject);
        Assert.DoesNotContain("should not stick", mdAfterReject);

        // An unknown id surfaces the typed error envelope, not a success.
        var missing = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"accept","revisionId":"rev999999"}""")));
        Assert.False(missing.GetProperty("success").GetBoolean());
        Assert.Equal("revision_not_found",
            missing.GetProperty("error").GetProperty("code").GetString());
    }
}
