#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
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

    private static string SavedSettingsXml(string path)
    {
        using var ms = new MemoryStream(File.ReadAllBytes(path));
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.DocumentSettingsPart?.RootElement?.OuterXml ?? string.Empty;
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

    [Fact]
    public void MCP011_GetContent_IntrospectionFormats_ReturnMutationCompatibleIdsAndSpans()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(ReplaceText(_store, sessionId, anchor, "Alpha beta")
            .GetProperty("success").GetBoolean());

        var styles = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"styles"}""")))
            .GetProperty("styles").EnumerateArray().ToArray();
        var paragraphStyle = styles.First(s => s.GetProperty("type").GetString() == "paragraph");
        var styleId = paragraphStyle.GetProperty("id").GetString()!;
        var styleMutation = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_paragraph_style","anchorId":"{{anchor}}","styleId":{{JsonSerializer.Serialize(styleId)}}}""")));
        Assert.True(styleMutation.GetProperty("success").GetBoolean());

        var formatting = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"formatting","anchorId":"{{anchor}}"}""")))
            .GetProperty("formatting");
        Assert.Equal(anchor, formatting.GetProperty("anchorId").GetString());
        Assert.Equal(JsonValueKind.Object, formatting.GetProperty("directParagraph").ValueKind);
        Assert.Equal(JsonValueKind.Object, formatting.GetProperty("effectiveParagraph").ValueKind);

        var span = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"spans","anchorId":"{{anchor}}"}""")))
            .GetProperty("spans")[0];
        var spanAnchor = span.GetProperty("anchorId").GetString()!;
        var range = span.GetProperty("span");
        var spanArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            action = "apply_format",
            anchorId = spanAnchor,
            span = new
            {
                start = range.GetProperty("start").GetInt32(),
                length = range.GetProperty("length").GetInt32(),
            },
            format = new { bold = true },
        });
        var spanMutation = Parse(Dispatcher.Call(_store, "docxodus_format", J(spanArgs)));
        Assert.True(spanMutation.GetProperty("success").GetBoolean());

        var info = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"info","anchorId":"{{anchor}}"}""")));
        Assert.Equal(anchor, info.GetProperty("sectionInfo").GetProperty("anchorId").GetString());
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

    [Fact]
    public void MCP032_Pagination_RegisterSearchPreviewAndStaleStatus()
    {
        var sessionId = OpenSession();
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(ReplaceText(_store, sessionId, anchor, "citation target")
            .GetProperty("success").GetBoolean());

        var version = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            JsonSerializer.Serialize(new { sessionId, format = "version" }))))
            .GetProperty("version").GetInt64();
        const string fingerprint = "mcp-page-map-v1";
        var pageMap = new
        {
            schemaVersion = 1,
            mode = "paginated",
            availability = "available",
            documentVersion = version,
            rendererFingerprint = fingerprint,
            pages = new[]
            {
                new
                {
                    pageNumber = 1,
                    pageInSection = 1,
                    width = 612,
                    height = 792,
                    sectionIndex = 0,
                    pageName = "docxodus-section-0",
                },
            },
            fragments = new[]
            {
                new
                {
                    fragmentId = $"p1-f0-{anchor}",
                    anchorId = anchor,
                    fragmentIndex = 0,
                    pageNumber = 1,
                    geometry = new { x = 72, y = 90, width = 468, height = 18 },
                    story = "body",
                    inTableCell = false,
                },
            },
        };
        var registered = Parse(Dispatcher.Call(_store, "docxodus_pagination", J(
            JsonSerializer.Serialize(new { sessionId, action = "register", pageMap }))));
        Assert.True(registered.GetProperty("success").GetBoolean());

        var citation = new { documentVersion = version, rendererFingerprint = fingerprint };
        var found = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                mode = "text",
                query = "citation target",
                citation,
            }))));
        Assert.Equal("available", found.GetProperty("matches")[0]
            .GetProperty("citation").GetProperty("availability").GetString());

        var preview = Parse(Dispatcher.Call(_store, "docxodus_preview", J(
            JsonSerializer.Serialize(new { sessionId, anchorId = anchor, citation }))));
        Assert.Equal("available_registered_map",
            preview.GetProperty("pageNavigation").GetString());
        Assert.Equal(1, preview.GetProperty("citation").GetProperty("fragments")[0]
            .GetProperty("pageNumber").GetInt32());
        Assert.Equal(612, preview.GetProperty("citation").GetProperty("pages")[0]
            .GetProperty("width").GetDouble());
        Assert.Contains("pagination-staging", preview.GetProperty("html").GetString());

        Assert.True(ReplaceText(_store, sessionId, anchor, "changed")
            .GetProperty("success").GetBoolean());
        var stale = Parse(Dispatcher.Call(_store, "docxodus_pagination", J(
            JsonSerializer.Serialize(new { sessionId, action = "status", citation }))));
        Assert.Equal("stale_document_version", stale.GetProperty("unavailableReason").GetString());
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
        Assert.Contains(XElement.Parse(htmlWithBorder).DescendantsAndSelf(),
            e => ((string?)e.Attribute("style"))?.Contains("border-bottom", StringComparison.Ordinal) == true);

        var cleared = Parse(Dispatcher.Call(_store, "docxodus_format", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_paragraph_format","anchorId":"{{anchor}}","paragraphFormat":{"clearBorders":true} }""")));
        Assert.True(cleared.GetProperty("success").GetBoolean());

        var htmlAfterClear = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"html"}""")))
            .GetProperty("html").GetString()!;
        Assert.DoesNotContain(XElement.Parse(htmlAfterClear).DescendantsAndSelf(),
            e => ((string?)e.Attribute("style"))?.Contains("border-bottom", StringComparison.Ordinal) == true);
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
    public void MCP137_Create_HeaderFooter_ReturnedAnchorsComposeAndReadBack()
    {
        // Issue #316: the engine and stdio host could author running stories, but the MCP
        // surface could not reach them. Exercise the complete agent workflow: address a section
        // through a BODY anchor, retain each returned story anchor, read both stories back, and
        // use the existing page-field action directly on the returned footer paragraph.
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var bodyAnchor = FirstBodyAnchorId(sessionId, _store);

        var header = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_header_text","bodyAnchorId":"{{bodyAnchor}}","kind":"default","markdown":"**CONFIDENTIAL** running header"}""")));
        Assert.True(header.GetProperty("success").GetBoolean());
        var headerAnchor = header.GetProperty("created")[0].GetProperty("id").GetString()!;
        Assert.StartsWith("p:hdr", headerAnchor);

        var footer = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_footer_text","bodyAnchorId":"{{bodyAnchor}}","kind":"default","markdown":"Running footer page "}""")));
        Assert.True(footer.GetProperty("success").GetBoolean());
        var footerAnchor = footer.GetProperty("created")[0].GetProperty("id").GetString()!;
        Assert.StartsWith("p:ftr", footerAnchor);

        var pageField = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_page_number_field","anchorId":"{{footerAnchor}}","field":"current_page","numberFormat":"lowerRoman"}""")));
        Assert.True(pageField.GetProperty("success").GetBoolean());

        var headerMarkdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown","anchorId":"{{headerAnchor}}"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("CONFIDENTIAL", headerMarkdown);
        Assert.DoesNotContain("Running footer", headerMarkdown);

        var footerText = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"text","anchorId":"{{footerAnchor}}"}""")))
            .GetProperty("text").GetString()!;
        Assert.Contains("Running footer page", footerText);

        var footerHtml = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"html","anchorId":"{{footerAnchor}}"}""")))
            .GetProperty("html").GetString()!;
        Assert.Contains("Running footer page", footerHtml);

        var blocks = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"blocks"}"""))).GetProperty("blocks");
        Assert.True(blocks.TryGetProperty(headerAnchor, out _));
        Assert.True(blocks.TryGetProperty(footerAnchor, out _));

        var visible = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"ensure_header_footer_visible","bodyAnchorId":"{{bodyAnchor}}","kind":"first"}""")));
        Assert.True(visible.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP138_Search_HeaderFooterScope_IsOptInAndComposable()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var bodyAnchor = FirstBodyAnchorId(sessionId, _store);

        var header = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_header_text","bodyAnchorId":"{{bodyAnchor}}","kind":"default","markdown":"scope needle running header"}""")));
        var headerAnchor = header.GetProperty("created")[0].GetProperty("id").GetString()!;
        var footer = Parse(Dispatcher.Call(_store, "docxodus_create", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_footer_text","bodyAnchorId":"{{bodyAnchor}}","kind":"default","markdown":"scope needle running footer"}""")));
        var footerAnchor = footer.GetProperty("created")[0].GetProperty("id").GetString()!;

        // Backward compatibility: absent scope is still body-only.
        var bodyDefault = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"text","query":"scope needle"}""")));
        Assert.Equal(0, bodyDefault.GetProperty("matches").GetArrayLength());

        var headers = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"text","query":"scope needle","scope":"headers"}""")));
        var headerMatch = Assert.Single(headers.GetProperty("matches").EnumerateArray());
        Assert.Equal(headerAnchor,
            headerMatch.GetProperty("enclosingAnchor").GetProperty("id").GetString());
        Assert.StartsWith("hdr",
            headerMatch.GetProperty("enclosingAnchor").GetProperty("scope").GetString());

        var both = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"regex","query":"scope\\s+needle","scope":"header_footer"}""")));
        var bothMatches = both.GetProperty("matches").EnumerateArray().ToList();
        Assert.Equal(2, bothMatches.Count);
        Assert.Contains(bothMatches, m =>
            m.GetProperty("enclosingAnchor").GetProperty("id").GetString() == headerAnchor);
        Assert.Contains(bothMatches, m =>
            m.GetProperty("enclosingAnchor").GetProperty("id").GetString() == footerAnchor);
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

        // Every cell operation consumes the same canonical tc anchor, returned directly by insert.
        var tcAnchor = inserted.GetProperty("created")[0].GetProperty("id").GetString()!;
        Assert.Equal("tc", inserted.GetProperty("created")[0].GetProperty("kind").GetString());
        var otherTcAnchor = inserted.GetProperty("created")[3].GetProperty("id").GetString()!;

        var replaced = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_cell_content","cellAnchorId":"{{tcAnchor}}","markdown":"cell text"}""")));
        Assert.True(replaced.GetProperty("success").GetBoolean());

        var rowAddedJson = Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert_row","cellAnchorId":"{{otherTcAnchor}}","position":"after"}"""));
        var rowAdded = Parse(rowAddedJson);
        Assert.True(rowAdded.GetProperty("success").GetBoolean(), rowAddedJson);

        var tableAnchor = Parse(Dispatcher.Call(_store, "docxodus_search", J(
            $$"""{"sessionId":{{sessionArg}},"mode":"kind","query":"tbl"}""")))
            .GetProperty("matches")[0].GetProperty("id").GetString()!;
        var metadata = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"get_metadata","tableAnchorId":"{{tableAnchor}}"}""")));
        Assert.True(metadata.GetProperty("success").GetBoolean());
        Assert.Equal("col", metadata.GetProperty("metadata").GetProperty("columns")[0]
            .GetProperty("anchor").GetProperty("kind").GetString());
    }

    [Fact]
    public void MCP061_Table_StylingActions()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var inserted = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert","anchorId":"{{anchor}}","position":"after","rows":2,"columns":2}""")));
        Assert.True(inserted.GetProperty("success").GetBoolean());
        var cellAnchor = inserted.GetProperty("created")[0].GetProperty("id").GetString()!;

        var widths = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_column_widths","cellAnchorId":"{{cellAnchor}}","widths":[6000,3000]}""")));
        Assert.True(widths.GetProperty("success").GetBoolean());

        var borders = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_borders","cellAnchorId":"{{cellAnchor}}","borderScope":"outside","borderStyle":"double","borderSize":12,"borderColor":"FF0000"}""")));
        Assert.True(borders.GetProperty("success").GetBoolean());

        var shading = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_shading","cellAnchorId":"{{cellAnchor}}","fill":"D9D9D9","shadingScope":"row"}""")));
        Assert.True(shading.GetProperty("success").GetBoolean());

        var header = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_repeat_header_row","cellAnchorId":"{{cellAnchor}}"}""")));
        Assert.True(header.GetProperty("success").GetBoolean());

        // A width list that doesn't match the column count surfaces the typed error.
        var bad = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_column_widths","cellAnchorId":"{{cellAnchor}}","widths":[6000]}""")));
        Assert.False(bad.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_table_styling", bad.GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void MCP062_Table_MergeAndUnmergeCells()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var inserted = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert","anchorId":"{{anchor}}","position":"after","rows":3,"columns":3}""")));
        Assert.True(inserted.GetProperty("success").GetBoolean());
        var topLeft = inserted.GetProperty("created")[0].GetProperty("id").GetString()!;

        // A 2×2 header block: w:gridSpan across, w:vMerge down.
        var mergedJson = Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"merge_cells","cellAnchorId":"{{topLeft}}","rowSpan":2,"colSpan":2}"""));
        Assert.True(Parse(mergedJson).GetProperty("success").GetBoolean(), mergedJson);

        // Merging the same rectangle again would now clip the vertical run it just created.
        var clipped = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"merge_cells","cellAnchorId":"{{topLeft}}","rowSpan":1,"colSpan":2}""")));
        Assert.False(clipped.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_table_merge", clipped.GetProperty("error").GetProperty("code").GetString());

        var unmerged = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"unmerge_cells","cellAnchorId":"{{topLeft}}"}""")));
        Assert.True(unmerged.GetProperty("success").GetBoolean());

        // Back to unit cells: unmerging a cell that carries no merge markup is now an error.
        var again = Parse(Dispatcher.Call(_store, "docxodus_table", J(
            $$"""{"sessionId":{{sessionArg}},"action":"unmerge_cells","cellAnchorId":"{{topLeft}}"}""")));
        Assert.False(again.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_table_merge", again.GetProperty("error").GetProperty("code").GetString());
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

    [Fact]
    public void MCP073_Comment_ReplyResolveAndReopen_AreNative()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"thread target"}"""));

        var added = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","anchorId":"{{anchor}}","author":"Alice","markdown":"Parent."}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        var parentAnchor = added.GetProperty("created").EnumerateArray()
            .First(a => a.GetProperty("kind").GetString() == "cmt")
            .GetProperty("id").GetString()!;

        var replied = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"reply","commentAnchorId":"{{parentAnchor}}","author":"Bob","initials":"BO","markdown":"Reply."}""")));
        Assert.True(replied.GetProperty("success").GetBoolean());
        var replyAnchor = replied.GetProperty("created").EnumerateArray()
            .First(a => a.GetProperty("kind").GetString() == "cmt")
            .GetProperty("id").GetString()!;

        // Omitting resolved uses the tool's documented resolve=true default.
        var resolved = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"resolve","commentAnchorId":"{{replyAnchor}}"}""")));
        Assert.True(resolved.GetProperty("success").GetBoolean());

        var listed = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var entries = listed.GetProperty("comments").EnumerateArray().ToList();
        var parent = Assert.Single(entries, e => e.GetProperty("anchorId").GetString() == parentAnchor);
        var reply = Assert.Single(entries, e => e.GetProperty("anchorId").GetString() == replyAnchor);
        Assert.False(parent.TryGetProperty("parentAnchorId", out _));
        Assert.False(parent.GetProperty("resolved").GetBoolean());
        Assert.Equal(parentAnchor, reply.GetProperty("parentAnchorId").GetString());
        Assert.True(reply.GetProperty("resolved").GetBoolean());

        var reopened = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"resolve","commentAnchorId":"{{replyAnchor}}","resolved":false}""")));
        Assert.True(reopened.GetProperty("success").GetBoolean());

        var listedAfterReopen = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var reopenedReply = Assert.Single(
            listedAfterReopen.GetProperty("comments").EnumerateArray(),
            e => e.GetProperty("anchorId").GetString() == replyAnchor);
        Assert.Equal(parentAnchor, reopenedReply.GetProperty("parentAnchorId").GetString());
        Assert.False(reopenedReply.GetProperty("resolved").GetBoolean());
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

        // Preserve a live redo cursor across preview; apply-then-undo used to destroy it.
        Assert.True(ReplaceText(_store, sessionId, anchor, "redo target")
            .GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")))
            .GetProperty("success").GetBoolean());

        var before = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        var versionBefore = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"version"}""")))
            .GetProperty("version").GetInt64();

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "preview",
              "previewHtml": "full",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "should not stick" } }
              ]
            }
            """)));
        Assert.Equal("ok", batch.GetProperty("status").GetString());
        Assert.True(batch.GetProperty("preview").GetBoolean());
        Assert.True(batch.GetProperty("success").GetBoolean());
        Assert.Equal(versionBefore, batch.GetProperty("baseVersion").GetInt64());
        Assert.Equal(versionBefore + 1, batch.GetProperty("resultVersion").GetInt64());
        Assert.Equal(64, batch.GetProperty("packageHash").GetString()!.Length);
        Assert.Single(batch.GetProperty("steps").EnumerateArray());
        Assert.True(batch.GetProperty("revisionChanges").TryGetProperty("added", out _));
        Assert.True(batch.GetProperty("commentChanges").TryGetProperty("added", out _));
        Assert.True(batch.GetProperty("annotationChanges").TryGetProperty("added", out _));
        Assert.Contains("should not stick", batch.GetProperty("html").GetString());

        var after = Parse(Dispatcher.Call(_store, "docxodus_get_content", J($$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Equal(before, after);
        var versionAfter = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"version"}""")))
            .GetProperty("version").GetInt64();
        Assert.Equal(versionBefore, versionAfter);

        var undo = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")));
        Assert.False(undo.GetProperty("success").GetBoolean());
        var redo = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"redo"}""")));
        Assert.True(redo.GetProperty("success").GetBoolean());
        var redone = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Contains("redo target", redone);
    }

    [Fact]
    public void MCP098_Mutations_PreviewFlagSupportsExplicitBestEffortWithoutLivePartialApply()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var before = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "best_effort",
              "preview": true,
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "shadow partial" } },
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "p:body:missing", "markdown": "failure" } }
              ]
            }
            """)));

        Assert.True(batch.GetProperty("preview").GetBoolean());
        Assert.Equal("best_effort", batch.GetProperty("mode").GetString());
        Assert.Equal("partial", batch.GetProperty("status").GetString());
        Assert.False(batch.GetProperty("success").GetBoolean());
        Assert.False(batch.GetProperty("rolledBack").GetBoolean());
        Assert.Equal(batch.GetProperty("baseVersion").GetInt64() + 1,
            batch.GetProperty("resultVersion").GetInt64());
        Assert.Contains(batch.GetProperty("warnings").EnumerateArray(),
            warning => warning.GetString()!.Contains("Best-effort", StringComparison.Ordinal));

        var after = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Equal(before, after);
        Assert.Equal(0, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
    }

    [Fact]
    public void MCP093_StalePrecondition_ReturnsStructuredFailureWithoutMutation()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(ReplaceText(_store, sessionId, anchor, "committed").GetProperty("success").GetBoolean());

        var failed = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "action": "replace_text",
              "anchorId": "{{anchor}}",
              "markdown": "must not apply",
              "preconditions": { "expectedVersion": 0 }
            }
            """)));

        Assert.False(failed.GetProperty("success").GetBoolean());
        var error = failed.GetProperty("error");
        Assert.Equal("precondition_failed", error.GetProperty("code").GetString());
        Assert.Equal(1, error.GetProperty("precondition").GetProperty("currentVersion").GetInt64());
        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString()!;
        Assert.Contains("committed", markdown);
        Assert.DoesNotContain("must not apply", markdown);
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

    [Fact]
    public void MCP094_Mutations_AtomicFailureIsStructuredAndLeavesNoVersionOrHistory()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var before = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "speculative" } },
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "p:body:missing", "markdown": "failure" } }
              ]
            }
            """)));

        Assert.Equal("failed", batch.GetProperty("status").GetString());
        Assert.False(batch.GetProperty("success").GetBoolean());
        Assert.True(batch.GetProperty("rolledBack").GetBoolean());
        var failure = batch.GetProperty("failure");
        Assert.Equal(1, failure.GetProperty("index").GetInt32());
        Assert.Equal("docxodus_edit", failure.GetProperty("tool").GetString());
        Assert.Equal("replace_text", failure.GetProperty("action").GetString());
        Assert.Equal("anchor_not_found", failure.GetProperty("error").GetProperty("code").GetString());
        Assert.True(failure.GetProperty("rolledBack").GetBoolean());

        var after = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Equal(before, after);
        var version = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"version"}""")))
            .GetProperty("version").GetInt64();
        Assert.Equal(0, version);
        var undo = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")));
        Assert.False(undo.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP095_Mutations_AtomicSuccessIsOneUndoAndInvalidStepHasCallerErrorCode()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var invalid = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_comment", "args": { "action": "list" } }
              ]
            }
            """)));
        Assert.Equal("invalid_batch_step",
            invalid.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "atomic MCP" } },
                { "tool": "docxodus_format", "args": { "action": "apply_format", "anchorId": "{{anchor}}", "format": { "bold": true } } }
              ]
            }
            """)));
        Assert.Equal("ok", batch.GetProperty("status").GetString());
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));

        var undo = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")));
        Assert.True(undo.GetProperty("success").GetBoolean());
        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.DoesNotContain("atomic MCP", markdown);
    }

    [Fact]
    public void MCP096_AtomicPreflightsLaterArgumentErrorsBeforeStepZeroMutates()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var before = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "must never run" } },
                { "tool": "docxodus_create", "args": { "action": "set_header_text", "bodyAnchorId": "{{anchor}}", "kind": "sideways", "markdown": "invalid header" } }
              ]
            }
            """)));

        Assert.False(batch.GetProperty("success").GetBoolean());
        Assert.True(batch.GetProperty("rolledBack").GetBoolean());
        var failure = batch.GetProperty("failure");
        Assert.Equal(1, failure.GetProperty("index").GetInt32());
        Assert.Equal("docxodus_create", failure.GetProperty("tool").GetString());
        Assert.Equal("set_header_text", failure.GetProperty("action").GetString());
        Assert.Equal("invalid_batch_step", failure.GetProperty("error").GetProperty("code").GetString());
        Assert.Contains("kind", failure.GetProperty("error").GetProperty("message").GetString());

        var after = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Equal(before, after);
        Assert.Equal(0, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
        var undo = Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")));
        Assert.False(undo.GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP097_AtomicStepPreconditionsUseBatchStartState()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var info = Parse(Docxodus.Internal.DocxSessionOps.GetAnchorInfo(
            _store.Get(sessionId).Handle, anchor));
        var originalText = JsonSerializer.Serialize(info.GetProperty("visibleText").GetString());

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "first atomic state" } },
                { "tool": "docxodus_edit", "args": { "action": "replace_text", "anchorId": "{{anchor}}", "markdown": "second atomic state", "preconditions": { "expectedText": {{originalText}} } } }
              ]
            }
            """)));

        Assert.True(batch.GetProperty("success").GetBoolean());
        Assert.Equal("atomic", batch.GetProperty("mode").GetString());
        Assert.Equal(1, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));
        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Contains("second atomic state", markdown);
        Assert.DoesNotContain("first atomic state", markdown);
    }

    [Theory]
    [InlineData("atomic")]
    [InlineData("apply")]
    public void MCP098_BatchedReplaceTextRange_EnforcesExpectedMatchCount(string mode)
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"Company Company Company Company"}""")))
            .GetProperty("success").GetBoolean());

        // The count guard can only be evaluated by the op that enumerated the live matches, so
        // stripping it from the dispatched step made it a silent no-op — all four occurrences
        // were replaced instead of the batch failing. Both the new atomic mode and the legacy
        // "apply" alias must reject the step and leave the anchor untouched.
        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "{{mode}}",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text_range", "anchorId": "{{anchor}}", "find": "Company", "replace": "Acme", "preconditions": { "expectedMatchCount": 1 } } }
              ]
            }
            """)));

        Assert.False(batch.GetProperty("success").GetBoolean());
        var failure = batch.GetProperty("failure");
        Assert.Equal(0, failure.GetProperty("index").GetInt32());
        Assert.Equal("precondition_failed", failure.GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("match_count",
            failure.GetProperty("error").GetProperty("precondition").GetProperty("condition").GetString());

        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.DoesNotContain("Acme", markdown);
        Assert.Contains("Company Company Company Company", markdown);
    }

    [Theory]
    [InlineData("atomic")]
    [InlineData("apply")]
    public void MCP098B_BatchedReplaceTextRange_MatchingExpectedMatchCountStillApplies(string mode)
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"Company Company"}""")))
            .GetProperty("success").GetBoolean());

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "{{mode}}",
              "steps": [
                { "tool": "docxodus_edit", "args": { "action": "replace_text_range", "anchorId": "{{anchor}}", "find": "Company", "replace": "Acme", "preconditions": { "expectedMatchCount": 2 } } }
              ]
            }
            """)));

        Assert.True(batch.GetProperty("success").GetBoolean());
        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}""")))
            .GetProperty("markdown").GetString();
        Assert.Contains("Acme Acme", markdown);
    }

    [Fact]
    public void MCP099_BatchedTableStep_KeepsTableAnchorMappingInItsReceipt()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_table", "args": { "action": "insert", "anchorId": "{{anchor}}", "position": "after", "rows": 2, "columns": 2 } }
              ]
            }
            """)));

        Assert.True(batch.GetProperty("success").GetBoolean());
        // The batch re-serializes every step through the shared EditResult wire shape, so a
        // parser that dropped tableAnchors left an agent with no cell-anchor map for the cells
        // its own step had just created.
        var result = batch.GetProperty("steps")[0].GetProperty("results")[0];
        var mapping = result.GetProperty("tableAnchors");
        var added = mapping.GetProperty("added");
        Assert.NotEmpty(added.EnumerateArray());
        var cells = added.EnumerateArray()
            .Where(x => x.GetProperty("entityKind").GetString() == "cell")
            .ToArray();
        Assert.Equal(4, cells.Length);
        Assert.All(cells, cell =>
        {
            Assert.StartsWith("tc:", cell.GetProperty("anchor").GetProperty("id").GetString());
            Assert.True(cell.TryGetProperty("rowIndex", out _));
            Assert.True(cell.TryGetProperty("columnIndex", out _));
        });
    }

    [Fact]
    public void MCP146_TrackChangesBatchPreviewIsIsolatedAndAtomicApplyResolvesRevision()
    {
        var sessionId = OpenSession(trackedChanges: "render_inline");
        var sessionArg = JsonSerializer.Serialize(sessionId);
        InsertParagraph(sessionId, "tracked paragraph");

        var listed = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        var revision = Assert.Single(listed.GetProperty("revisions").EnumerateArray());
        var revisionId = revision.GetProperty("id").GetString()!;
        Assert.Equal("content_insert", revision.GetProperty("family").GetString());

        var preview = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "preview",
              "steps": [
                { "tool": "docxodus_track_changes", "args": { "action": "reject", "revisionId": {{JsonSerializer.Serialize(revisionId)}} } }
              ]
            }
            """)));
        Assert.Equal("ok", preview.GetProperty("status").GetString());
        Assert.True(preview.GetProperty("success").GetBoolean());
        Assert.True(preview.GetProperty("preview").GetBoolean());
        Assert.True(Assert.Single(preview.GetProperty("steps").EnumerateArray())
            .GetProperty("success").GetBoolean());

        var afterPreview = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.Equal(revisionId,
            Assert.Single(afterPreview.GetProperty("revisions").EnumerateArray()).GetProperty("id").GetString());
        Assert.Contains("tracked paragraph", Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}"""))).GetProperty("markdown").GetString()!);

        var applied = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_track_changes", "args": { "action": "reject", "revisionId": {{JsonSerializer.Serialize(revisionId)}} } }
              ]
            }
            """)));
        Assert.Equal("ok", applied.GetProperty("status").GetString());
        Assert.True(applied.GetProperty("success").GetBoolean());
        Assert.Equal(0, Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}"""))).GetProperty("revisions").GetArrayLength());
        Assert.DoesNotContain("tracked paragraph", Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}"""))).GetProperty("markdown").GetString()!);
    }

    [Fact]
    public void MCP102_TrackChangesBulkAcceptAndRejectAllAreAtomicBatchSteps()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(ReplaceText(_store, sessionId, anchor, "baseline")
            .GetProperty("success").GetBoolean());
        SetMode(sessionId, "render_inline");

        Assert.True(ReplaceText(_store, sessionId, anchor, "accepted replacement")
            .GetProperty("success").GetBoolean());
        var accepted = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_track_changes", "args": { "action": "accept_all" } }
              ]
            }
            """)));
        Assert.Equal("ok", accepted.GetProperty("status").GetString());
        Assert.True(accepted.GetProperty("success").GetBoolean());
        Assert.Contains("accepted replacement", Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}"""))).GetProperty("markdown").GetString()!);

        Assert.True(ReplaceText(_store, sessionId, anchor, "rejected replacement")
            .GetProperty("success").GetBoolean());
        var rejected = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            $$"""
            {
              "sessionId": {{sessionArg}},
              "mode": "atomic",
              "steps": [
                { "tool": "docxodus_track_changes", "args": { "action": "reject_all" } }
              ]
            }
            """)));
        Assert.Equal("ok", rejected.GetProperty("status").GetString());
        Assert.True(rejected.GetProperty("success").GetBoolean());
        var markdown = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}"""))).GetProperty("markdown").GetString()!;
        Assert.Contains("accepted replacement", markdown);
        Assert.DoesNotContain("rejected replacement", markdown);
    }

    [Fact]
    public void MCP103_TrackChangesReadOnlyBatchStepsFailStructuredAndSchemaAdvertisesMutations()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        foreach (var action in new[] { "list", "set_mode" })
        {
            var actionArg = JsonSerializer.Serialize(action);
            var receipt = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
                $$"""
                {
                  "sessionId": {{sessionArg}},
                  "mode": "atomic",
                  "steps": [
                    { "tool": "docxodus_track_changes", "args": { "action": {{actionArg}} } }
                  ]
                }
                """)));
            Assert.Equal("failed", receipt.GetProperty("status").GetString());
            Assert.False(receipt.GetProperty("success").GetBoolean());
            var failure = receipt.GetProperty("failure");
            Assert.Equal("docxodus_track_changes", failure.GetProperty("tool").GetString());
            Assert.Equal(action, failure.GetProperty("action").GetString());
            Assert.Equal("invalid_batch_step",
                failure.GetProperty("error").GetProperty("code").GetString());

            var legacy = Assert.Throws<McpToolException>(() => Dispatcher.Call(
                _store, "docxodus_mutations", J(
                    $$"""
                    {
                      "sessionId": {{sessionArg}},
                      "mode": "apply",
                      "steps": [
                        { "tool": "docxodus_track_changes", "args": { "action": {{actionArg}} } }
                      ]
                    }
                    """)));
            Assert.Contains(action, legacy.Message, StringComparison.Ordinal);
        }

        var mutations = Assert.Single(ToolCatalog.Tools, tool => tool.Name == "docxodus_mutations");
        using var schema = JsonDocument.Parse(mutations.InputSchemaJson);
        var tools = schema.RootElement.GetProperty("properties").GetProperty("steps")
            .GetProperty("items").GetProperty("properties").GetProperty("tool")
            .GetProperty("enum").EnumerateArray().Select(value => value.GetString()).ToList();
        Assert.Contains("docxodus_track_changes", tools);
    }

    // ─── Tool catalog ───────────────────────────────────────────────────

    [Fact]
    public void MCP100_ToolCatalog_HasExpectedDistinctNamedToolsWithValidSchemas()
    {
        string[] expectedNames =
        {
            "docxodus_annotate",
            "docxodus_close",
            "docxodus_comment",
            "docxodus_content_controls",
            "docxodus_create",
            "docxodus_edit",
            "docxodus_format",
            "docxodus_get_content",
            "docxodus_images",
            "docxodus_links",
            "docxodus_list",
            "docxodus_mutations",
            "docxodus_open",
            "docxodus_pagination",
            "docxodus_preview",
            "docxodus_save",
            "docxodus_search",
            "docxodus_table",
            "docxodus_track_changes",
        };

        Assert.Equal(expectedNames.Length, ToolCatalog.Tools.Count);
        var names = new System.Collections.Generic.HashSet<string>();
        foreach (var tool in ToolCatalog.Tools)
        {
            Assert.StartsWith("docxodus_", tool.Name);
            Assert.False(string.IsNullOrWhiteSpace(tool.Description));
            Assert.True(names.Add(tool.Name), $"duplicate tool name: {tool.Name}");
            using var schema = JsonDocument.Parse(tool.InputSchemaJson); // must be valid JSON
            Assert.Equal("object", schema.RootElement.GetProperty("type").GetString());
        }
        Assert.Equal(expectedNames, names.OrderBy(name => name, StringComparer.Ordinal));
    }

    [Fact]
    public void MCP101_PageMapSchemas_DescribeStrictTokensAndActionRequirements()
    {
        static void AssertCitationSchema(JsonElement schema)
        {
            Assert.False(schema.GetProperty("additionalProperties").GetBoolean());
            var required = schema.GetProperty("required").EnumerateArray()
                .Select(value => value.GetString()).ToArray();
            Assert.Contains("documentVersion", required);
            Assert.Contains("rendererFingerprint", required);
            Assert.Equal("integer", schema.GetProperty("properties")
                .GetProperty("documentVersion").GetProperty("type").GetString());
            Assert.Equal(1, schema.GetProperty("properties")
                .GetProperty("rendererFingerprint").GetProperty("minLength").GetInt32());
        }

        foreach (var toolName in new[] { "docxodus_get_content", "docxodus_preview", "docxodus_search" })
        {
            var tool = Assert.Single(ToolCatalog.Tools, item => item.Name == toolName);
            using var schema = JsonDocument.Parse(tool.InputSchemaJson);
            AssertCitationSchema(schema.RootElement.GetProperty("properties").GetProperty("citation"));
        }

        var pagination = Assert.Single(ToolCatalog.Tools, item => item.Name == "docxodus_pagination");
        using var paginationSchema = JsonDocument.Parse(pagination.InputSchemaJson);
        var root = paginationSchema.RootElement;
        AssertCitationSchema(root.GetProperty("properties").GetProperty("citation"));
        var pageMap = root.GetProperty("properties").GetProperty("pageMap");
        Assert.False(pageMap.GetProperty("additionalProperties").GetBoolean());
        Assert.Equal(1, pageMap.GetProperty("properties").GetProperty("schemaVersion")
            .GetProperty("const").GetInt32());
        Assert.False(pageMap.GetProperty("properties").GetProperty("fragments")
            .GetProperty("items").GetProperty("additionalProperties").GetBoolean());
        var variants = root.GetProperty("oneOf").EnumerateArray().ToArray();
        Assert.Contains(variants, variant => variant.GetProperty("properties").GetProperty("action")
            .GetProperty("const").GetString() == "register"
            && variant.GetProperty("required").EnumerateArray().Any(v => v.GetString() == "pageMap"));
        Assert.Contains(variants, variant => variant.GetProperty("properties").GetProperty("action")
            .GetProperty("const").GetString() == "cite"
            && variant.GetProperty("required").EnumerateArray().Any(v => v.GetString() == "citation"));
    }

    [Fact]
    public void MCP139_ToolCatalog_AdvertisesHeaderFooterCreateAndSearchScope()
    {
        var create = Assert.Single(ToolCatalog.Tools, t => t.Name == "docxodus_create");
        using (var schema = JsonDocument.Parse(create.InputSchemaJson))
        {
            var actions = schema.RootElement.GetProperty("properties").GetProperty("action")
                .GetProperty("enum").EnumerateArray().Select(v => v.GetString()).ToList();
            Assert.Contains("set_header_text", actions);
            Assert.Contains("set_footer_text", actions);
            Assert.Contains("ensure_header_footer_visible", actions);
            Assert.True(schema.RootElement.GetProperty("properties").TryGetProperty("bodyAnchorId", out _));
            Assert.True(schema.RootElement.GetProperty("properties").TryGetProperty("kind", out _));
        }

        var search = Assert.Single(ToolCatalog.Tools, t => t.Name == "docxodus_search");
        using var searchSchema = JsonDocument.Parse(search.InputSchemaJson);
        var scopes = searchSchema.RootElement.GetProperty("properties").GetProperty("scope")
            .GetProperty("enum").EnumerateArray().Select(v => v.GetString()).ToList();
        Assert.Contains("headers", scopes);
        Assert.Contains("footers", scopes);
        Assert.Contains("header_footer", scopes);
    }

    // ─── Inline preview (MCP Apps / ChatGPT Apps) ──────────────────────

    [Fact]
    public void MCP140_Preview_RendersDocumentAndSingleBlock()
    {
        var sessionId = OpenSession();
        InsertParagraph(sessionId, "Preview me");
        var sessionArg = JsonSerializer.Serialize(sessionId);

        var whole = Parse(Dispatcher.Call(_store, "docxodus_preview",
            J($$"""{"sessionId":{{sessionArg}}}""")));
        Assert.Equal(sessionId, whole.GetProperty("sessionId").GetString());
        var html = whole.GetProperty("html").GetString()!;
        Assert.Contains("Preview me", html);
        Assert.Contains("<style", html); // whole-document renders carry the converter stylesheet

        var anchor = FirstBodyAnchorId(sessionId, _store);
        var block = Parse(Dispatcher.Call(_store, "docxodus_preview",
            J($$"""{"sessionId":{{sessionArg}},"anchorId":"{{anchor}}"}""")));
        Assert.Equal(anchor, block.GetProperty("anchorId").GetString());
        Assert.False(string.IsNullOrEmpty(block.GetProperty("html").GetString()));
    }

    [Fact]
    public void MCP141_WrapToolResult_RoutesHtmlToMetaNotModelContent()
    {
        var wrapped = Parse(UiResources.WrapToolResult("docxodus_preview",
            """{"sessionId":"s1","anchorId":"p:body:abc","html":"<html><body>big</body></html>"}""",
            isError: false));
        var text = wrapped.GetProperty("content")[0].GetProperty("text").GetString()!;
        Assert.DoesNotContain("<html", text);
        Assert.Equal("<html><body>big</body></html>",
            wrapped.GetProperty("_meta").GetProperty(UiResources.HtmlMetaKey).GetString());
        var structured = wrapped.GetProperty("structuredContent");
        Assert.Equal("s1", structured.GetProperty("sessionId").GetString());
        Assert.Equal("p:body:abc", structured.GetProperty("anchorId").GetString());
        Assert.Equal("<html><body>big</body></html>".Length,
            structured.GetProperty("htmlLength").GetInt32());

        var cited = Parse(UiResources.WrapToolResult("docxodus_preview",
            """{"sessionId":"s1","html":"<p>x</p>","citation":{"availability":"available","pages":[{"pageNumber":3,"pageInSection":1,"width":612,"height":792,"pageName":"docxodus-section-0"}],"fragments":[{"pageNumber":3}]},"pageNavigation":"available_registered_map"}""",
            isError: false)).GetProperty("structuredContent");
        Assert.Equal(3, cited.GetProperty("citation").GetProperty("fragments")[0]
            .GetProperty("pageNumber").GetInt32());
        Assert.Equal("available_registered_map", cited.GetProperty("pageNavigation").GetString());

        // docxodus_open mirrors its result as structuredContent for the widget…
        var open = Parse(UiResources.WrapToolResult("docxodus_open",
            """{"sessionId":"s1","path":"a.docx"}""", isError: false));
        Assert.Equal("s1", open.GetProperty("structuredContent").GetProperty("sessionId").GetString());

        // …while every other tool keeps the original envelope, and errors are never rewrapped.
        var plain = Parse(UiResources.WrapToolResult("docxodus_save", """{"path":"a.docx"}""", isError: false));
        Assert.False(plain.TryGetProperty("structuredContent", out _));
        var error = Parse(UiResources.WrapToolResult("docxodus_preview", """{"success":false}""", isError: true));
        Assert.True(error.GetProperty("isError").GetBoolean());
        Assert.False(error.TryGetProperty("_meta", out _));
    }

    [Fact]
    public void MCP142_UiResources_ServeViewerTemplate()
    {
        var list = Parse(UiResources.BuildResourcesListResult());
        var resource = Assert.Single(list.GetProperty("resources").EnumerateArray());
        Assert.Equal(UiResources.ViewerUri, resource.GetProperty("uri").GetString());
        Assert.Equal(UiResources.ViewerMimeType, resource.GetProperty("mimeType").GetString());

        var read = Parse(UiResources.BuildResourcesReadResult(
            J($$"""{"uri":"{{UiResources.ViewerUri}}"}""")));
        var contents = Assert.Single(read.GetProperty("contents").EnumerateArray());
        var htmlText = contents.GetProperty("text").GetString()!;
        Assert.StartsWith("<!DOCTYPE html>", htmlText.TrimStart());
        Assert.Contains("docxodus_preview", htmlText); // the widget's refresh path
        Assert.Contains("unavailable_continuous_preview", htmlText);
        Assert.Contains("available_registered_map", htmlText);
        Assert.Contains("materializeCitationPage", htmlText);
        Assert.True(contents.GetProperty("_meta").TryGetProperty("ui", out _));

        Assert.Throws<InvalidParamsException>(() =>
            UiResources.BuildResourcesReadResult(J("""{"uri":"ui://docxodus/nope.html"}""")));
    }

    [Fact]
    public void MCP143_ToolsList_StampsUiMetaOnWidgetToolsOnly()
    {
        var tools = Parse(Program.BuildToolsListResult()).GetProperty("tools");
        foreach (var tool in tools.EnumerateArray())
        {
            var name = tool.GetProperty("name").GetString();
            var hasMeta = tool.TryGetProperty("_meta", out var meta);
            if (name is "docxodus_open" or "docxodus_preview")
            {
                Assert.True(hasMeta, $"{name} should carry UI _meta");
                Assert.Equal(UiResources.ViewerUri,
                    meta.GetProperty("ui").GetProperty("resourceUri").GetString());
                Assert.Equal(UiResources.ViewerUri,
                    meta.GetProperty("openai/outputTemplate").GetString());
            }
            else
            {
                Assert.False(hasMeta, $"{name} should not carry _meta");
            }
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
        Assert.StartsWith("rev2-", insertRev.GetProperty("id").GetString());
        Assert.Equal("content_insert", insertRev.GetProperty("family").GetString());
        Assert.Equal("supported", insertRev.GetProperty("resolutionStatus").GetString());
        Assert.Equal("/word/document.xml", insertRev.GetProperty("partUri").GetString());
        Assert.Equal("body", insertRev.GetProperty("scope").GetString());
        Assert.NotEmpty(insertRev.GetProperty("constituentIds").EnumerateArray());
        Assert.NotEmpty(insertRev.GetProperty("affectedAnchors").EnumerateArray());

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

    [Fact]
    public void MCP138_Comment_AddTargetsExactlyOneAnchorOrRevision()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"original"}"""));
        SetMode(sessionId, "render_inline", "Reviewer A");
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"replacement"}"""));

        var revisions = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("revisions").EnumerateArray().ToList();
        var insertion = Assert.Single(revisions, r => r.GetProperty("type").GetString() == "insert");
        var revisionId = insertion.GetProperty("id").GetString()!;

        var added = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","revisionId":"{{revisionId}}","author":"Alice","markdown":"Keep this revision."}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        Assert.Contains(added.GetProperty("created").EnumerateArray(),
            a => a.GetProperty("kind").GetString() == "cmt");

        var rejected = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"reject","revisionId":"{{revisionId}}"}""")));
        Assert.True(rejected.GetProperty("success").GetBoolean());
        var comments = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")));
        Assert.Single(comments.GetProperty("comments").EnumerateArray());

        var stale = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","revisionId":"{{revisionId}}","author":"Alice"}""")));
        Assert.False(stale.GetProperty("success").GetBoolean());
        Assert.Equal("revision_not_found", stale.GetProperty("error").GetProperty("code").GetString());

        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","author":"Alice"}""")));
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","anchorId":"{{anchor}}","revisionId":"rev1","author":"Alice"}""")));
        Assert.Throws<McpToolException>(() => Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","revisionId":"rev1","span":{"start":0,"length":1},"author":"Alice"}""")));
    }

    [Fact]
    public void MCP140_HtmlAndSavedSettings_PreserveTrackedRevisionCommentTarget()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"original"}"""));
        SetMode(sessionId, "render_inline", "Reviewer A");
        Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"replace_text","anchorId":"{{anchor}}","markdown":"replacement"}"""));

        var revisions = Parse(Dispatcher.Call(_store, "docxodus_track_changes", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("revisions").EnumerateArray().ToList();
        var revisionId = Assert.Single(
                revisions, r => r.GetProperty("type").GetString() == "insert")
            .GetProperty("id").GetString()!;
        var comment = Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add","revisionId":"{{revisionId}}","author":"Alice","markdown":"Keep this revision."}""")));
        Assert.True(comment.GetProperty("success").GetBoolean());

        string Render(string? anchorId = null)
        {
            var argsJson = $"{{\"sessionId\":{sessionArg},\"format\":\"html\"";
            if (anchorId is not null)
                argsJson += $",\"anchorId\":{JsonSerializer.Serialize(anchorId)}";
            argsJson += "}";
            return Parse(Dispatcher.Call(_store, "docxodus_get_content", J(argsJson)))
                .GetProperty("html").GetString()!;
        }

        foreach (var html in new[] { Render(), Render(anchor) })
        {
            Assert.Contains("<ins", html);
            Assert.Contains("<del", html);
            Assert.Contains("replacement", html);
            Assert.Contains("original", html);
        }

        var savedPath = Path.Combine(_root, "tracked-comment.docx");
        Save(sessionId, savedPath);
        Assert.Contains("trackRevisions", SavedSettingsXml(savedPath));
    }

    [Fact]
    public void MCP141_NativeLinkAndBookmarkCrud_RoundTripsIdsAndTypedFailures()
    {
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        Assert.True(ReplaceText(_store, sessionId, anchor, "alpha beta")
            .GetProperty("success").GetBoolean());

        var bookmark = Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add_bookmark","name":"Clause","startAnchorId":"{{anchor}}","startOffset":0,"endAnchorId":"{{anchor}}","endOffset":5}""")));
        Assert.True(bookmark.GetProperty("success").GetBoolean());

        var added = Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"add_hyperlink","anchorId":"{{anchor}}","startOffset":6,"length":4,"kind":"external","target":"https://example.test/mcp"}""")));
        Assert.True(added.GetProperty("success").GetBoolean());
        var hyperlinkId = added.GetProperty("hyperlinkId").GetString()!;

        var updated = Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"update_hyperlink","hyperlinkId":{{JsonSerializer.Serialize(hyperlinkId)}},"kind":"internal","target":"Clause"}""")));
        Assert.True(updated.GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"rename_bookmark","name":"Clause","newName":"ClauseTwo"}""")))
            .GetProperty("success").GetBoolean());

        var links = Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list_hyperlinks","scope":"body"}""")));
        var listed = Assert.Single(links.GetProperty("hyperlinks").EnumerateArray());
        Assert.Equal(hyperlinkId, listed.GetProperty("id").GetString());
        Assert.Equal("ClauseTwo", listed.GetProperty("target").GetString());

        var blocked = Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove_bookmark","name":"ClauseTwo"}""")));
        Assert.False(blocked.GetProperty("success").GetBoolean());
        Assert.Equal("bookmark_in_use", blocked.GetProperty("error").GetProperty("code").GetString());

        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove_hyperlink","hyperlinkId":{{JsonSerializer.Serialize(hyperlinkId)}}}""")))
            .GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_links", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove_bookmark","name":"ClauseTwo"}""")))
            .GetProperty("success").GetBoolean());
    }

    [Fact]
    public void MCP144_NativeImageCapabilitiesAndCrud_UseExplicitBase64Boundary()
    {
        var capabilities = Parse(Dispatcher.Call(_store, "docxodus_images",
            J("""{"action":"capabilities"}"""))).GetProperty("capabilities");
        Assert.Equal(96, capabilities.GetProperty("defaultDpi").GetDouble());
        Assert.False(capabilities.GetProperty("supportsNetworkFetch").GetBoolean());
        Assert.DoesNotContain(capabilities.GetProperty("horizontalReferences").EnumerateArray(),
            value => value.GetString() == "unknown");
        var imageTool = Assert.Single(ToolCatalog.Tools, tool => tool.Name == "docxodus_images");
        using (var schema = JsonDocument.Parse(imageTool.InputSchemaJson))
            Assert.Contains("comments", schema.RootElement.GetProperty("properties")
                .GetProperty("scope").GetProperty("enum").EnumerateArray()
                .Select(value => value.GetString()));

        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var png = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(png, 0);
        png[19] = 2;
        png[23] = 3;
        var imageBase64 = JsonSerializer.Serialize(Convert.ToBase64String(png));

        var inserted = Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$$"""{"sessionId":{{{sessionArg}}},"action":"insert","anchorId":"{{{anchor}}}","characterOffset":0,"imageBase64":{{{imageBase64}}},"options":{"altText":"diagram","widthPoints":72}}""")));
        Assert.True(inserted.GetProperty("success").GetBoolean());
        var imageId = inserted.GetProperty("imageId").GetString()!;

        var images = Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list","scope":"body"}""")));
        var image = Assert.Single(images.GetProperty("images").EnumerateArray());
        Assert.Equal(imageId, image.GetProperty("id").GetString());
        Assert.Equal("png", image.GetProperty("format").GetString());

        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$$"""{"sessionId":{{{sessionArg}}},"action":"set_dimensions","imageId":{{{JsonSerializer.Serialize(imageId)}}},"dimensions":{"widthPoints":36}}""")))
            .GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$"""{"sessionId":{{sessionArg}},"action":"set_metadata","imageId":{{JsonSerializer.Serialize(imageId)}},"altText":"updated","title":null}""")))
            .GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$"""{"sessionId":{{sessionArg}},"action":"remove","imageId":{{JsonSerializer.Serialize(imageId)}}}""")))
            .GetProperty("success").GetBoolean());
        Assert.Empty(Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("images").EnumerateArray());

        var urlRejected = Parse(Dispatcher.Call(_store, "docxodus_images", J(
            $$"""{"sessionId":{{sessionArg}},"action":"insert","anchorId":"{{anchor}}","characterOffset":0,"imageBase64":"https://example.test/image.png"}""")));
        Assert.False(urlRejected.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_image_data",
            urlRejected.GetProperty("error").GetProperty("code").GetString());
        var wrongOptions = "{\"sessionId\":" + sessionArg
            + ",\"action\":\"insert\",\"anchorId\":" + JsonSerializer.Serialize(anchor)
            + ",\"characterOffset\":0,\"imageBase64\":" + imageBase64
            + ",\"options\":false}";
        Assert.Throws<McpToolException>(() => Dispatcher.Call(
            _store, "docxodus_images", J(wrongOptions)));
    }

    [Fact]
    public void MCP145_NativeImageBatchPreviewRollsBackParts_AndRejectsReadOnlyActions()
    {
        var sessionId = OpenSession();
        var anchor = FirstBodyAnchorId(sessionId, _store);
        var png = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(png, 0);
        png[19] = 2;
        png[23] = 3;
        var imageBase64 = Convert.ToBase64String(png);
        var previewArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "preview",
            steps = new[]
            {
                new
                {
                    tool = "docxodus_images",
                    args = new
                    {
                        action = "insert", anchorId = anchor, characterOffset = 0,
                        imageBase64, options = new { altText = "preview only" },
                    },
                },
            },
        });
        var preview = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(previewArgs)));
        Assert.Equal("ok", preview.GetProperty("status").GetString());
        Assert.Equal(1, preview.GetProperty("editsApplied").GetInt32());

        var listed = Parse(Dispatcher.Call(_store, "docxodus_images", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "list",
        }))));
        Assert.Empty(listed.GetProperty("images").EnumerateArray());
        var savedPath = Path.Combine(_root, "image-preview-rollback.docx");
        Save(sessionId, savedPath);
        using (var stream = new MemoryStream(File.ReadAllBytes(savedPath)))
        using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false))
            Assert.Empty(document.MainDocumentPart!.ImageParts);

        var readOnlyArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "preview",
            steps = new[] { new { tool = "docxodus_images", args = new { action = "list" } } },
        });
        var invalid = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(readOnlyArgs)));
        Assert.False(invalid.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_batch_step",
            invalid.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void MCP147_ContentControls_ListFillDetachAndBatchPreview_AreFirstClass()
    {
        File.WriteAllBytes(_tempPath, DocxSessionContentControlTests.BuildFixture());
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var listed = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list","scope":"body"}""")))
            .GetProperty("contentControls").EnumerateArray().ToArray();
        Assert.Equal(15, listed.Length);
        var plain = listed.Single(control => control.TryGetProperty("nativeId", out var id)
            && id.GetString() == "101");
        var plainAnchor = plain.GetProperty("anchorId").GetString()!;
        var filled = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"fill_text","anchorId":{{JsonSerializer.Serialize(plainAnchor)}},"text":"MCP value"}""")));
        Assert.True(filled.GetProperty("success").GetBoolean());

        var bound = listed.Single(control => control.TryGetProperty("nativeId", out var id)
            && id.GetString() == "106");
        var boundAnchor = bound.GetProperty("anchorId").GetString()!;
        var refused = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"fill_text","anchorId":{{JsonSerializer.Serialize(boundAnchor)}},"text":"no"}""")));
        Assert.Equal("content_control_bound",
            refused.GetProperty("error").GetProperty("code").GetString());
        var detached = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"fill_text","anchorId":{{JsonSerializer.Serialize(boundAnchor)}},"text":"yes","bindingPolicy":"detach_target"}""")));
        Assert.True(detached.GetProperty("success").GetBoolean());
        Assert.Throws<McpToolException>(() => Dispatcher.Call(
            _store, "docxodus_content_controls", J(JsonSerializer.Serialize(new
            {
                sessionId,
                action = "fill_text",
                anchorId = plainAnchor,
                text = "ignored",
                bindingPolicy = false,
            }))));

        var previewArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "preview",
            steps = new[] { new { tool = "docxodus_content_controls",
                args = new { action = "fill_text", anchorId = plainAnchor, text = "preview" } } },
        });
        Assert.Equal("ok", Parse(Dispatcher.Call(_store, "docxodus_mutations", J(previewArgs)))
            .GetProperty("status").GetString());
        var after = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("contentControls").EnumerateArray().Single(control =>
                control.TryGetProperty("nativeId", out var id) && id.GetString() == "101");
        Assert.Equal("MCP value", after.GetProperty("text").GetString());

        var tool = Assert.Single(ToolCatalog.Tools,
            definition => definition.Name == "docxodus_content_controls");
        using var schema = JsonDocument.Parse(tool.InputSchemaJson);
        Assert.True(schema.RootElement.GetProperty("properties").TryGetProperty("preconditions", out _));
        Assert.Contains("detach_target", schema.RootElement.GetProperty("properties")
            .GetProperty("bindingPolicy").GetProperty("enum").EnumerateArray()
            .Select(value => value.GetString()));
    }

    [Fact]
    public void MCP148_ContentControlPreview_RefusalsNeverConsumePreexistingUndoHistory()
    {
        File.WriteAllBytes(_tempPath, DocxSessionContentControlTests.BuildFixture());
        var sessionId = OpenSession();
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var listed = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("contentControls").EnumerateArray().ToArray();
        string Anchor(string nativeId) => listed.Single(control =>
            control.TryGetProperty("nativeId", out var id) && id.GetString() == nativeId)
            .GetProperty("anchorId").GetString()!;
        var plainAnchor = Anchor("101");
        var boundAnchor = Anchor("106");

        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"fill_text","anchorId":{{JsonSerializer.Serialize(plainAnchor)}},"text":"kept edit"}""")))
            .GetProperty("success").GetBoolean());
        var guarded = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                action = "fill_text",
                anchorId = plainAnchor,
                text = "stale",
                preconditions = new { expectedVersion = 0 },
            }))));
        Assert.Equal("precondition_failed",
            guarded.GetProperty("error").GetProperty("code").GetString());

        string Preview(params object[] steps) => Dispatcher.Call(_store, "docxodus_mutations", J(
            JsonSerializer.Serialize(new { sessionId, mode = "preview", steps })));
        var blocks = Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"blocks"}"""))).GetProperty("blocks");
        var bodyBlock = blocks.EnumerateObject().First(property => property.Value.ValueKind
            == JsonValueKind.Object && property.Value.GetProperty("scope").GetString() == "body"
            && property.Value.GetProperty("kind").GetString() is "p" or "h" or "li" or "tbl").Name;
        var successfulNoOp = Parse(Preview(new
        {
            tool = "docxodus_edit",
            args = new
            {
                action = "move_block",
                sourceAnchorId = bodyBlock,
                targetAnchorId = bodyBlock,
                position = "before",
            },
        }));
        Assert.Equal("ok", successfulNoOp.GetProperty("status").GetString());

        var failedOnly = Parse(Preview(new
        {
            tool = "docxodus_content_controls",
            args = new { action = "fill_text", anchorId = boundAnchor, text = "refused" },
        }));
        Assert.Equal("failed", failedOnly.GetProperty("status").GetString());

        var readOnly = Parse(Preview(new
        {
            tool = "docxodus_content_controls",
            args = new { action = "list" },
        }));
        Assert.Equal("invalid_batch_step",
            readOnly.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());

        var mixed = Parse(Preview(
            new
            {
                tool = "docxodus_content_controls",
                args = new { action = "fill_text", anchorId = plainAnchor, text = "preview" },
            },
            new
            {
                tool = "docxodus_content_controls",
                args = new { action = "fill_text", anchorId = boundAnchor, text = "refused" },
            }));
        Assert.Equal("failed", mixed.GetProperty("status").GetString());
        Assert.Equal(0, mixed.GetProperty("editsApplied").GetInt32());

        var invalidArguments = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                mode = "atomic",
                steps = new object[]
                {
                    new { tool = "docxodus_content_controls",
                        args = new { action = "fill_text", anchorId = plainAnchor, text = "not applied" } },
                    new { tool = "docxodus_content_controls",
                        args = new { action = "set_checked", anchorId = plainAnchor } },
                },
            }))));
        Assert.Equal("invalid_batch_step", invalidArguments.GetProperty("failure")
            .GetProperty("error").GetProperty("code").GetString());
        var after = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("contentControls").EnumerateArray().Single(control =>
                control.TryGetProperty("nativeId", out var id) && id.GetString() == "101");
        Assert.Equal("kept edit", after.GetProperty("text").GetString());

        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}""")))
            .GetProperty("success").GetBoolean());
        var undone = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}""")))
            .GetProperty("contentControls").EnumerateArray().Single(control =>
                control.TryGetProperty("nativeId", out var id) && id.GetString() == "101");
        Assert.Equal("inner", undone.GetProperty("text").GetString());
    }

    [Fact]
    public void MCP149_ContentControlReceiptsAndBestEffortBatch_PreserveSdtIdentities()
    {
        File.WriteAllBytes(_tempPath, DocxSessionContentControlTests.BuildFixture());
        var sessionId = OpenSession();
        var listed = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new { sessionId, action = "list" }))))
            .GetProperty("contentControls").EnumerateArray().ToArray();
        string Anchor(string nativeId) => listed.Single(control =>
            control.TryGetProperty("nativeId", out var id) && id.GetString() == nativeId)
            .GetProperty("anchorId").GetString()!;
        var plainAnchor = Anchor("101");
        var boundAnchor = Anchor("106");
        var sectionAnchor = Anchor("108");

        var fill = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                action = "fill_text",
                anchorId = plainAnchor,
                text = "receipt",
            }))));
        Assert.Equal(plainAnchor, Assert.Single(fill.GetProperty("modified").EnumerateArray())
            .GetProperty("id").GetString());
        Assert.Empty(fill.GetProperty("created").EnumerateArray());
        Assert.Empty(fill.GetProperty("removed").EnumerateArray());

        var added = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                action = "add_repeating_item",
                sectionAnchorId = sectionAnchor,
            }))));
        var createdAnchor = Assert.Single(added.GetProperty("created").EnumerateArray())
            .GetProperty("id").GetString()!;
        Assert.StartsWith("sdt:body:", createdAnchor, StringComparison.Ordinal);
        Assert.Equal(sectionAnchor, Assert.Single(added.GetProperty("modified").EnumerateArray())
            .GetProperty("id").GetString());

        var removed = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                action = "remove_repeating_item",
                itemAnchorId = createdAnchor,
            }))));
        Assert.Equal(createdAnchor, Assert.Single(removed.GetProperty("removed").EnumerateArray())
            .GetProperty("id").GetString());
        Assert.Equal(sectionAnchor, Assert.Single(removed.GetProperty("modified").EnumerateArray())
            .GetProperty("id").GetString());

        var batch = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(
            JsonSerializer.Serialize(new
            {
                sessionId,
                mode = "best_effort",
                steps = new object[]
                {
                    new { tool = "docxodus_content_controls", args = new
                        { action = "fill_text", anchorId = plainAnchor, text = "best effort" } },
                    new { tool = "docxodus_content_controls", args = new
                        { action = "fill_text", anchorId = boundAnchor, text = "refused" } },
                    new { tool = "docxodus_content_controls", args = new
                        { action = "add_repeating_item", sectionAnchorId = sectionAnchor } },
                },
            }))));
        Assert.Equal("partial", batch.GetProperty("status").GetString());
        Assert.Equal(2, batch.GetProperty("editsApplied").GetInt32());
        var steps = batch.GetProperty("steps").EnumerateArray().ToArray();
        Assert.True(steps[0].GetProperty("success").GetBoolean());
        Assert.False(steps[1].GetProperty("success").GetBoolean());
        Assert.True(steps[2].GetProperty("success").GetBoolean());
        Assert.Equal(plainAnchor, Assert.Single(steps[0].GetProperty("results")[0]
            .GetProperty("modified").EnumerateArray()).GetProperty("id").GetString());
        Assert.Equal(sectionAnchor, Assert.Single(steps[2].GetProperty("results")[0]
            .GetProperty("modified").EnumerateArray()).GetProperty("id").GetString());

        var after = Parse(Dispatcher.Call(_store, "docxodus_content_controls", J(
            JsonSerializer.Serialize(new { sessionId, action = "list" }))))
            .GetProperty("contentControls").EnumerateArray().ToArray();
        Assert.Equal("best effort", after.Single(control =>
            control.TryGetProperty("nativeId", out var id) && id.GetString() == "101")
            .GetProperty("text").GetString());
        Assert.Equal(2, after.Count(control =>
            control.GetProperty("type").GetString() == "repeating_section_item"));
    }

    /// <summary>
    /// Issue #468, at the surface it was filed against. The apply-and-undo preview issued one
    /// <c>Undo()</c> per non-throwing step, but a step that returned <c>success: false</c> recorded
    /// no undo entry — so a preview batch carrying one bad anchor popped a pre-batch snapshot and
    /// silently reverted an edit the caller had already committed. #446 replaced that path with an
    /// isolated shadow clone, which removes the whole class; this pins the three live-history
    /// observables the old implementation corrupted, none of which any MCP test asserted together:
    /// a committed edit, the redo cursor, and a batch longer than <c>undoDepth</c>.
    /// </summary>
    [Fact]
    public void MCP150_PreviewBatchWithFailingStep_LeavesCommittedEditsUndoAndRedoIntact()
    {
        // undoDepth 1 is the tight ring from #468's third claim: a preview longer than the ring
        // used to evict the caller's own entry and leave steps permanently applied.
        var sessionId = Parse(Dispatcher.Call(_store, "docxodus_open", J(
            $$"""{"path":{{JsonSerializer.Serialize(_tempPath)}},"undoDepth":1}""")))
            .GetProperty("sessionId").GetString()!;
        var sessionArg = JsonSerializer.Serialize(sessionId);
        var anchor = FirstBodyAnchorId(sessionId, _store);

        // Commit the state that must remain live, then create and undo a second edit. This leaves
        // "committed edit" visible with "redo target" already on the redo stack. No live mutation
        // may occur between this undo and the preview — RecordPreOp would legitimately clear redo.
        Assert.True(ReplaceText(_store, sessionId, anchor, "committed edit")
            .GetProperty("success").GetBoolean());
        Assert.True(ReplaceText(_store, sessionId, anchor, "redo target")
            .GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}"""))).GetProperty("success").GetBoolean());

        var committedVersion = Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle);
        string Markdown() => Parse(Dispatcher.Call(_store, "docxodus_get_content", J(
            $$"""{"sessionId":{{sessionArg}},"format":"markdown"}"""))).GetProperty("markdown").GetString()!;
        var committedMarkdown = Markdown();
        Assert.Contains("committed edit", committedMarkdown);

        // Four mutating steps against a one-deep ring, the last of which fails on a bad anchor.
        object[] StepsEndingInFailure() => new object[]
        {
            new { tool = "docxodus_edit", args = new { action = "replace_text", anchorId = anchor, markdown = "shadow one" } },
            new { tool = "docxodus_create", args = new { action = "set_header_text", bodyAnchorId = anchor, kind = "default", markdown = "shadow header" } },
            new { tool = "docxodus_comment", args = new { action = "add", anchorId = anchor, author = "Preview", markdown = "shadow comment" } },
            new { tool = "docxodus_edit", args = new { action = "replace_text", anchorId = "p:body:missing", markdown = "fails" } },
        };

        var atomic = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(
            new { sessionId, mode = "preview", steps = StepsEndingInFailure() }))));
        Assert.True(atomic.GetProperty("preview").GetBoolean());
        Assert.Equal("failed", atomic.GetProperty("status").GetString());
        Assert.True(atomic.GetProperty("rolledBack").GetBoolean());
        Assert.Equal(3, atomic.GetProperty("failure").GetProperty("index").GetInt32());
        Assert.Equal(committedMarkdown, Markdown());
        Assert.Equal(committedVersion, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));

        // best_effort keeps the three successes in the shadow — the live document still must not
        // move, and the failing step still must not consume a live undo entry.
        var partial = Parse(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(
            new { sessionId, mode = "best_effort", preview = true, steps = StepsEndingInFailure() }))));
        Assert.True(partial.GetProperty("preview").GetBoolean());
        Assert.Equal("partial", partial.GetProperty("status").GetString());
        Assert.False(partial.GetProperty("rolledBack").GetBoolean());
        Assert.Equal(3, partial.GetProperty("editsApplied").GetInt32());
        Assert.Equal(committedMarkdown, Markdown());
        Assert.Equal(committedVersion, Docxodus.Internal.DocxSessionOps.GetVersion(_store.Get(sessionId).Handle));

        // Prove the PRE-EXISTING redo cursor directly, before any live undo can manufacture a new
        // one. Then undo it back to the committed state and prove the one-deep ring has no older
        // entry. The final redo/undo pair leaves the fixture in the committed state for leak checks.
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"redo"}"""))).GetProperty("success").GetBoolean());
        Assert.Contains("redo target", Markdown());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}"""))).GetProperty("success").GetBoolean());
        Assert.Contains("committed edit", Markdown());
        Assert.False(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}"""))).GetProperty("success").GetBoolean());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"redo"}"""))).GetProperty("success").GetBoolean());
        Assert.Contains("redo target", Markdown());
        Assert.True(Parse(Dispatcher.Call(_store, "docxodus_edit", J(
            $$"""{"sessionId":{{sessionArg}},"action":"undo"}"""))).GetProperty("success").GetBoolean());
        Assert.Contains("committed edit", Markdown());

        // Nothing the shadow authored ever reached the live package.
        var live = Markdown();
        Assert.DoesNotContain("shadow one", live);
        Assert.DoesNotContain("shadow header", live);
        Assert.Empty(Parse(Dispatcher.Call(_store, "docxodus_comment", J(
            $$"""{"sessionId":{{sessionArg}},"action":"list"}"""))).GetProperty("comments").EnumerateArray());
    }
}
