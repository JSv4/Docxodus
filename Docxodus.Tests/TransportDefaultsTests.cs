// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #960: one session op must behave the same whichever transport calls it. Per-op defaults,
/// argument names, revision filtering and result truncation are owned by the
/// <see cref="DocxSessionOps"/> facade rather than re-implemented in each dispatcher. Test IDs use
/// the TD prefix.
/// </summary>
public class TransportDefaultsTests : IDisposable
{
    private readonly string _root;
    private readonly SessionStore _store;

    public TransportDefaultsTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"transport-defaults-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
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

    private static JsonElement J(object value) => J(JsonSerializer.Serialize(value));

    private static string FirstParagraph(int handle) =>
        SessionRegistry.Get(handle).Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    /// <summary>The body's top-level element names, in order, after the op ran.</summary>
    private static string[] BodyShape(int handle)
    {
        using var ms = new MemoryStream(SessionRegistry.Get(handle).Save());
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!.Elements()
            .Where(e => e.Name != W.sectPr).Select(e => e.Name.LocalName).ToArray();
    }

    private (int Handle, string SessionId) OpenMcp()
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path })))
            .GetProperty("sessionId").GetString()!;
        return (_store.Get(sessionId).Handle, sessionId);
    }

    // ─── Reference-field position (the issue's acceptance case) ─────────

    [Theory]
    [InlineData("insert_table_of_contents", "sdt")]
    [InlineData("insert_table_of_figures", "p")]
    [InlineData("insert_table_of_authorities", "p")]
    public void TD001_ReferenceFieldWithoutPosition_LandsInTheSamePlace_OnEveryTransport(string op, string inserted)
    {
        var (mcpHandle, sessionId) = OpenMcp();
        var mcpAnchor = FirstParagraph(mcpHandle);
        var mcp = J(Dispatcher.Call(_store, "docxodus_create", J(new { sessionId, action = op, anchorId = mcpAnchor })));
        Assert.True(mcp.GetProperty("success").GetBoolean(), mcp.ToString());

        var hostHandle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var hostAnchor = FirstParagraph(hostHandle);
            var host = J(Docxodus.PyHost.Dispatcher.Dispatch(op, J(new { handle = hostHandle, anchorId = hostAnchor })));
            Assert.True(host.GetProperty("success").GetBoolean(), host.ToString());

            var mcpShape = BodyShape(mcpHandle);
            Assert.Equal(mcpShape, BodyShape(hostHandle));
            // The facade's default: the field goes ahead of the anchor paragraph, which ends up last.
            Assert.Equal(inserted, mcpShape[0]);
            Assert.Equal(DocxSessionOps.ReferenceFieldDefaultPosition, Position.Before);
        }
        finally
        {
            DocxSessionOps.CloseSession(hostHandle);
        }
    }

    // ─── One name per argument, the old one as a deprecated alias ───────

    [Theory]
    [InlineData("shadingScope")]
    [InlineData("scope")]
    public void TD010_StdioHost_SetCellShading_AcceptsCanonicalNameAndDeprecatedAlias(string name)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var cells = TableCells(handle);
            var args = $$"""{"handle":{{handle}},"cellAnchorId":"{{cells[0]}}","fill":"FF0000","{{name}}":"row"}""";
            var result = J(Docxodus.PyHost.Dispatcher.Dispatch("set_cell_shading", J(args)));
            Assert.True(result.GetProperty("success").GetBoolean(), result.ToString());
            // Row scope shades the second cell of the row too.
            Assert.Equal(2, ShadedCells(handle));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void TD011_ConflictingCanonicalAndAlias_IsRefused()
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var cells = TableCells(handle);
            var args = $$"""{"handle":{{handle}},"cellAnchorId":"{{cells[0]}}","fill":"FF0000","shadingScope":"row","scope":"cell"}""";
            Assert.ThrowsAny<ArgumentException>(() => Docxodus.PyHost.Dispatcher.Dispatch("set_cell_shading", J(args)));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Theory]
    [InlineData("repeatHeader")]
    [InlineData("repeat")]
    public void TD012_Mcp_SetRowOptions_AcceptsCanonicalNameAndDeprecatedAlias(string name)
    {
        var (handle, sessionId) = OpenMcp();
        var anchor = FirstParagraph(handle);
        var table = J(Dispatcher.Call(_store, "docxodus_table", J(new
        {
            sessionId, action = "insert", anchorId = anchor, position = "after", rows = 2, columns = 2,
        })));
        var cell = table.GetProperty("created")[0].GetProperty("id").GetString()!;
        var args = $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"set_row_options","cellAnchorId":"{{cell}}","{{name}}":true}""";
        var result = J(Dispatcher.Call(_store, "docxodus_table", J(args)));
        Assert.True(result.GetProperty("success").GetBoolean(), result.ToString());
        Assert.Contains("tblHeader", SessionRegistry.Get(handle).Raw.GetXml(
            SessionRegistry.Get(handle).Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "tbl").Anchor.Id));
    }

    // ─── Revision filtering and truncation live in the facade ───────────

    [Fact]
    public void TD020_RevisionFilter_GivesTheSameAnswer_ThroughFacadeHostAndMcp()
    {
        var (handle, sessionId) = OpenMcp();
        var settings = J(Dispatcher.Call(_store, "docxodus_track_changes", J(new
        {
            sessionId, action = "set_mode", mode = "render_inline", revisionAuthor = "Ann",
        })));
        Assert.True(settings.GetProperty("success").GetBoolean());
        var anchor = FirstParagraph(handle);
        Assert.True(J(DocxSessionOps.InsertParagraph(handle, anchor, Position.After, "by Ann")).GetProperty("success").GetBoolean());
        DocxSessionOps.SetRevisionAuthor(handle, "Bob");
        Assert.True(J(DocxSessionOps.InsertParagraph(handle, anchor, Position.After, "by Bob")).GetProperty("success").GetBoolean());

        var all = J(DocxSessionOps.ListRevisions(handle)).EnumerateArray().ToList();
        Assert.Contains(all, r => r.GetProperty("author").GetString() == "Ann");
        Assert.Contains(all, r => r.GetProperty("author").GetString() == "Bob");

        var facade = J(DocxSessionOps.ListRevisions(handle, new RevisionListFilter { Author = "bob" }));
        Assert.NotEmpty(facade.EnumerateArray());
        Assert.All(facade.EnumerateArray(), r => Assert.Equal("Bob", r.GetProperty("author").GetString()));

        var host = J(Docxodus.PyHost.Dispatcher.Dispatch("list_revisions", J(new { handle, author = "bob" })));
        var mcp = J(Dispatcher.Call(_store, "docxodus_track_changes", J(new { sessionId, action = "list", author = "bob" })))
            .GetProperty("revisions");
        Assert.Equal(facade.GetRawText(), host.GetRawText());
        Assert.Equal(facade.GetRawText(), mcp.GetRawText());
    }

    [Fact]
    public void TD021_MaxResults_TruncatesTheSameWay_ThroughHostAndMcp()
    {
        var (handle, sessionId) = OpenMcp();
        var anchor = FirstParagraph(handle);
        for (int i = 0; i < 4; i++)
            Assert.True(J(DocxSessionOps.InsertParagraph(handle, anchor, Position.After, $"needle {i}")).GetProperty("success").GetBoolean());

        var host = J(Docxodus.PyHost.Dispatcher.Dispatch("grep", J(new { handle, pattern = "needle", maxResults = 2 })));
        var mcp = J(Dispatcher.Call(_store, "docxodus_search", J(new { sessionId, mode = "text", query = "needle", maxResults = 2, caseSensitive = true })))
            .GetProperty("matches");
        Assert.Equal(2, host.GetArrayLength());
        Assert.Equal(host.GetRawText(), mcp.GetRawText());
        Assert.Equal(4, J(Docxodus.PyHost.Dispatcher.Dispatch("grep", J(new { handle, pattern = "needle" }))).GetArrayLength());
        Assert.ThrowsAny<ArgumentException>(() => DocxSessionOps.Grep(handle, "needle", new GrepRequest { MaxResults = -1 }));
    }

    [Fact]
    public void TD022_OmittedContextChars_TakesTheCoreDefault()
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var anchor = FirstParagraph(handle);
            var text = new string('a', 200) + " needle " + new string('b', 200);
            Assert.True(J(DocxSessionOps.InsertParagraph(handle, anchor, Position.After, text)).GetProperty("success").GetBoolean());
            var match = J(Docxodus.PyHost.Dispatcher.Dispatch("grep", J(new { handle, pattern = "needle" })))[0];
            Assert.Equal(DocxSession.DefaultContextChars, match.GetProperty("contextBefore").GetString()!.Length);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    private static string[] TableCells(int handle)
    {
        var anchor = FirstParagraph(handle);
        var inserted = J(DocxSessionOps.InsertTable(handle, anchor, Position.After, 1, 2, ""));
        Assert.True(inserted.GetProperty("success").GetBoolean());
        return inserted.GetProperty("created").EnumerateArray().Select(a => a.GetProperty("id").GetString()!).ToArray();
    }

    private static int ShadedCells(int handle)
    {
        using var ms = new MemoryStream(SessionRegistry.Get(handle).Save());
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Descendants(W.tcPr).Count(p => p.Element(W.shd) is not null);
    }
}
