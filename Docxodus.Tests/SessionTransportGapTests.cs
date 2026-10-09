// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Session ops that a transport left out by omission, not by design (issue #1026), now reach the
/// facade there too: page setup, page numbering and the first/even header-footer switches on MCP,
/// the horizontal rule on the stdio host, and the note listing on the stdio host and MCP. Each test
/// calls the transport for real and checks the op took effect on the document, so a transport that
/// answered success without calling the facade would still fail.
/// </summary>
public sealed class SessionTransportGapTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"transport-gaps-{Guid.NewGuid():N}");
    private readonly SessionStore _store;

    public SessionTransportGapTests()
    {
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        Directory.Delete(_root, recursive: true);
    }

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    private static JsonElement J(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static void AssertSucceeded(JsonElement result) =>
        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());

    /// <summary>The session's first body paragraph, given some text.</summary>
    private static string Paragraph(int handle)
    {
        var session = SessionRegistry.Get(handle);
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        return session.ReplaceText(first, "Body text for the gap tests").Modified.Select(a => a.Id).FirstOrDefault() ?? first;
    }

    private static SectionInfo Section(int handle, string anchorId) =>
        SessionRegistry.Get(handle).GetSectionInfo(anchorId)!;

    // ─── stdio host ─────────────────────────────────────────────────────

    private sealed class StdioSession : IDisposable
    {
        public int Handle { get; } = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);

        public void Dispose() => DocxSessionOps.CloseSession(Handle);

        public JsonElement Call(string op, Dictionary<string, object?> args)
        {
            args["handle"] = Handle;
            return J(Docxodus.PyHost.Dispatcher.Dispatch(op, J(args)));
        }
    }

    [Fact]
    public void Stdio_InsertHorizontalRule_InsertsARuleParagraphWithTheRequestedEdge()
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);

        var result = s.Call("insert_horizontal_rule", new()
        {
            ["anchorId"] = paragraph, ["position"] = "after", ["rule"] = new { style = "double", size = 4 },
        });

        AssertSucceeded(result);
        var created = result.GetProperty("created").EnumerateArray().First().GetProperty("id").GetString()!;
        var xml = DocxSessionOps.RawGetXml(s.Handle, created);
        Assert.Contains("pBdr", xml);
        Assert.Contains("w:val=\"double\"", xml);
        Assert.Contains("w:sz=\"4\"", xml);
    }

    [Fact]
    public void Stdio_InsertHorizontalRule_WithoutARule_UsesTheDefaultEdge()
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);

        var result = s.Call("insert_horizontal_rule", new() { ["anchorId"] = paragraph, ["position"] = "before" });

        AssertSucceeded(result);
        var created = result.GetProperty("created").EnumerateArray().First().GetProperty("id").GetString()!;
        Assert.Contains("pBdr", DocxSessionOps.RawGetXml(s.Handle, created));
    }

    [Fact]
    public void Stdio_InsertHorizontalRule_WithoutPosition_InsertsAfter_TheFacadeDefault()
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);

        var result = s.Call("insert_horizontal_rule", new() { ["anchorId"] = paragraph });

        AssertSucceeded(result);
        var created = result.GetProperty("created").EnumerateArray().First().GetProperty("id").GetString()!;
        var order = SessionRegistry.Get(s.Handle).Project().AnchorIndex.Keys.ToList();
        Assert.True(order.IndexOf(created) > order.IndexOf(paragraph), "the rule should follow its anchor");
    }

    [Fact]
    public void Stdio_InsertHorizontalRule_HonoursPreconditions()
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);

        var result = s.Call("insert_horizontal_rule", new()
        {
            ["anchorId"] = paragraph, ["position"] = "after",
            ["preconditions"] = new { expectedVersion = 999 },
        });

        Assert.False(result.GetProperty("success").GetBoolean(), result.GetRawText());
        Assert.DoesNotContain("pBdr", DocxSessionOps.RawGetXml(s.Handle, paragraph));
    }

    [Fact]
    public void Stdio_InsertHorizontalRule_RunsAsABatchStep()
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);

        var result = s.Call("execute_batch", new()
        {
            ["steps"] = new object[]
            {
                new { operation = "insert_horizontal_rule", args = new { anchorId = paragraph, position = "after" } },
            },
        });

        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
        Assert.Contains(SessionRegistry.Get(s.Handle).Project().AnchorIndex.Values,
            t => t.Anchor.Scope == "body" && DocxSessionOps.RawGetXml(s.Handle, t.Anchor.Id).Contains("pBdr"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Stdio_ListNotes_ListsTheNotesInCitationOrder(bool endnotes)
    {
        using var s = new StdioSession();
        var paragraph = Paragraph(s.Handle);
        var session = SessionRegistry.Get(s.Handle);
        var inserted = endnotes
            ? session.InsertEndnote(paragraph, 4, "The note")
            : session.InsertFootnote(paragraph, 4, "The note");
        var definition = inserted.Created.First(a => a.Kind is "fn" or "en").Id;

        var notes = s.Call("list_notes", new() { ["endnotes"] = endnotes });
        var other = s.Call("list_notes", new() { ["endnotes"] = !endnotes });

        var note = Assert.Single(notes.EnumerateArray());
        Assert.Equal(1, note.GetProperty("ordinal").GetInt32());
        Assert.Equal(definition, note.GetProperty("defAnchorId").GetString());
        Assert.Equal(0, other.GetArrayLength());
    }

    [Fact]
    public void Stdio_ListNotes_RequiresEndnotes()
    {
        using var s = new StdioSession();
        var e = Assert.ThrowsAny<Exception>(() => s.Call("list_notes", new()));
        Assert.Contains("\"endnotes\"", e.Message);
    }

    // ─── MCP ────────────────────────────────────────────────────────────

    private (string SessionId, int Handle, string Paragraph) OpenMcp(bool captureDeliveryEvidence = false)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path, captureDeliveryEvidence })))
            .GetProperty("sessionId").GetString()!;
        var handle = _store.Get(sessionId).Handle;
        return (sessionId, handle, Paragraph(handle));
    }

    private JsonElement Create(string sessionId, string action, Dictionary<string, object?> args)
    {
        args["sessionId"] = sessionId;
        args["action"] = action;
        return J(Dispatcher.Call(_store, "docxodus_create", J(args)));
    }

    [Fact]
    public void Mcp_SetPageSetup_ChangesTheSectionGeometry()
    {
        var (sessionId, handle, paragraph) = OpenMcp();

        var result = Create(sessionId, "set_page_setup", new()
        {
            ["anchorId"] = paragraph, ["op"] = new { marginLeftTwips = 2000, landscape = true },
        });

        AssertSucceeded(result);
        var section = Section(handle, paragraph);
        Assert.Equal(2000, section.MarginLeftTwips);
        Assert.True(section.Landscape);
    }

    [Fact]
    public void Mcp_SetPageSetup_WithoutOp_LeavesTheSectionUnchanged_TheFacadeDefault()
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        var before = Section(handle, paragraph);
        Create(sessionId, "set_page_setup", new() { ["anchorId"] = paragraph });
        var after = Section(handle, paragraph);
        Assert.Equal(before.PageWidthTwips, after.PageWidthTwips);
        Assert.Equal(before.PageHeightTwips, after.PageHeightTwips);
    }

    [Fact]
    public void Mcp_SetPageSetup_RefusesAnOpThatIsNotAnObject()
    {
        var (sessionId, _, paragraph) = OpenMcp();
        var e = Assert.Throws<McpToolException>(() => Create(sessionId, "set_page_setup", new() { ["anchorId"] = paragraph, ["op"] = "wide" }));
        Assert.Contains("\"op\"", e.Message);
    }

    [Fact]
    public void Mcp_SetPageSetup_RefusesInvalidGeometryWithoutTouchingTheSection()
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        var before = Section(handle, paragraph).PageWidthTwips;

        var result = Create(sessionId, "set_page_setup", new()
        {
            ["anchorId"] = paragraph, ["op"] = new { pageWidthTwips = -5 },
        });

        Assert.False(result.GetProperty("success").GetBoolean());
        Assert.Equal(before, Section(handle, paragraph).PageWidthTwips);
    }

    [Fact]
    public void Mcp_SetAndClearPageNumbering_WriteAndRemoveTheSectionsNumbering()
    {
        var (sessionId, handle, paragraph) = OpenMcp();

        AssertSucceeded(Create(sessionId, "set_page_numbering", new()
        {
            ["anchorId"] = paragraph, ["op"] = new { start = 5, format = "lowerRoman" },
        }));
        var numbered = Section(handle, paragraph);
        Assert.Equal(5, numbered.PageNumberStart);
        Assert.Equal(NumberFormat.LowerRoman, numbered.PageNumberFormat);

        AssertSucceeded(Create(sessionId, "clear_page_numbering", new() { ["anchorId"] = paragraph }));
        var cleared = Section(handle, paragraph);
        Assert.Null(cleared.PageNumberStart);
        Assert.Null(cleared.PageNumberFormat);
    }

    [Theory]
    [InlineData("first")]
    [InlineData("even")]
    public void Mcp_SetHeaderFooterKindEnabled_TogglesWordsSwitch(string kind)
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        bool On() => kind == "first" ? Section(handle, paragraph).TitlePage : Section(handle, paragraph).EvenAndOddHeaders;

        AssertSucceeded(Create(sessionId, "set_header_footer_kind_enabled", new()
        {
            ["anchorId"] = paragraph, ["kind"] = kind, ["enabled"] = true,
        }));
        Assert.True(On());

        AssertSucceeded(Create(sessionId, "set_header_footer_kind_enabled", new()
        {
            ["anchorId"] = paragraph, ["kind"] = kind, ["enabled"] = false,
        }));
        Assert.False(On());
    }

    [Fact]
    public void Mcp_SetHeaderFooterKindEnabled_RequiresEnabled()
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        var e = Assert.Throws<McpToolException>(() => Create(sessionId, "set_header_footer_kind_enabled", new()
        {
            ["anchorId"] = paragraph, ["kind"] = "first",
        }));
        Assert.Contains("\"enabled\"", e.Message);
        Assert.False(Section(handle, paragraph).TitlePage);
    }

    [Fact]
    public void Mcp_ListNotes_ListsTheNotesInCitationOrder()
    {
        var (sessionId, _, paragraph) = OpenMcp();
        var inserted = Create(sessionId, "insert_footnote", new()
        {
            ["anchorId"] = paragraph, ["characterOffset"] = 4, ["markdown"] = "The note",
        });
        AssertSucceeded(inserted);
        var definition = inserted.GetProperty("created").EnumerateArray()
            .First(a => a.GetProperty("kind").GetString() == "fn").GetProperty("id").GetString();

        var footnotes = Create(sessionId, "list_notes", new() { ["endnotes"] = false }).GetProperty("notes");
        var endnotes = Create(sessionId, "list_notes", new() { ["endnotes"] = true }).GetProperty("notes");

        var note = Assert.Single(footnotes.EnumerateArray());
        Assert.Equal(1, note.GetProperty("ordinal").GetInt32());
        Assert.Equal(definition, note.GetProperty("defAnchorId").GetString());
        Assert.Equal(0, endnotes.GetArrayLength());
    }

    [Fact]
    public void Mcp_ListNotes_IsAReadNotABatchStep()
    {
        var (sessionId, _, _) = OpenMcp();
        var result = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId,
            steps = new object[] { new { tool = "docxodus_create", args = new { action = "list_notes", endnotes = false } } },
        })));
        Assert.Equal("invalid_batch_step", result.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
    }

    [Fact]
    public void Mcp_NewLayoutActions_RunAsOneAtomicBatch()
    {
        var (sessionId, handle, paragraph) = OpenMcp();

        var result = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId,
            steps = new object[]
            {
                new { tool = "docxodus_create", args = new { action = "set_page_setup", anchorId = paragraph, op = new { marginTopTwips = 1800 } } },
                new { tool = "docxodus_create", args = new { action = "set_page_numbering", anchorId = paragraph, op = new { start = 3 } } },
                new { tool = "docxodus_create", args = new { action = "set_header_footer_kind_enabled", anchorId = paragraph, kind = "first", enabled = true } },
            },
        })));

        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
        var section = Section(handle, paragraph);
        Assert.Equal(1800, section.MarginTopTwips);
        Assert.Equal(3, section.PageNumberStart);
        Assert.True(section.TitlePage);
    }

    [Fact]
    public void Mcp_AMalformedLayoutStep_FailsTheBatchBeforeAnyStepRuns()
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        var before = Section(handle, paragraph).MarginTopTwips;

        var result = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId,
            steps = new object[]
            {
                new { tool = "docxodus_create", args = new { action = "set_page_setup", anchorId = paragraph, op = new { marginTopTwips = before + 100 } } },
                new { tool = "docxodus_create", args = new { action = "set_header_footer_kind_enabled", anchorId = paragraph, kind = "first" } },
            },
        })));

        Assert.False(result.GetProperty("success").GetBoolean());
        var error = result.GetProperty("failure").GetProperty("error");
        Assert.Equal("invalid_batch_step", error.GetProperty("code").GetString());
        Assert.Contains("\"enabled\"", error.GetProperty("message").GetString());
        Assert.Equal(before, Section(handle, paragraph).MarginTopTwips);
    }

    [Fact]
    public void Mcp_NewLayoutActions_HonourPreconditions()
    {
        var (sessionId, handle, paragraph) = OpenMcp();
        var version = DocxSessionOps.GetVersion(handle);

        AssertSucceeded(Create(sessionId, "set_page_setup", new()
        {
            ["anchorId"] = paragraph, ["op"] = new { marginRightTwips = 1700 },
            ["preconditions"] = new { expectedVersion = version },
        }));
        var stale = Create(sessionId, "clear_page_numbering", new()
        {
            ["anchorId"] = paragraph, ["preconditions"] = new { expectedVersion = version },
        });

        Assert.False(stale.GetProperty("success").GetBoolean(), stale.GetRawText());
        Assert.Equal(1700, Section(handle, paragraph).MarginRightTwips);
    }

    [Fact]
    public void Mcp_ADirectLayoutCall_IsRecordedAsADescribedTransaction_AndListNotesIsNot()
    {
        var (sessionId, handle, paragraph) = OpenMcp(captureDeliveryEvidence: true);
        // Seeding the paragraph ran outside any transport, so it is the one unlabeled step.
        var before = J(DocxSessionOps.GetDeliveryEvidenceStatus(handle));

        AssertSucceeded(Create(sessionId, "set_page_setup", new()
        {
            ["anchorId"] = paragraph, ["op"] = new { marginBottomTwips = 1600 },
        }));
        _ = Create(sessionId, "list_notes", new() { ["endnotes"] = false });

        var after = J(DocxSessionOps.GetDeliveryEvidenceStatus(handle));
        Assert.Equal(before.GetProperty("transactionCount").GetInt32() + 1, after.GetProperty("transactionCount").GetInt32());
        Assert.Equal(before.GetProperty("unlabeledTransactionCount").GetInt32(),
            after.GetProperty("unlabeledTransactionCount").GetInt32());
    }
}
