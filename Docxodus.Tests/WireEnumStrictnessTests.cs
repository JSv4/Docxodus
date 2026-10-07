// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #962: a misspelled wire token for a session op must be rejected, not silently mapped to
/// the op's default (which performs a different edit than the caller asked for). Absent tokens
/// still take the default. Test IDs use the WE prefix.
/// </summary>
public class WireEnumStrictnessTests : IDisposable
{
    private readonly int _handle;
    private readonly string[] _cells;
    private readonly string _paragraph;

    public WireEnumStrictnessTests()
    {
        _handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        _paragraph = FirstBlockId(_handle);
        var inserted = Json(DocxSessionOps.InsertTable(_handle, _paragraph, Position.After, 2, 2, ""));
        Assert.True(inserted.GetProperty("success").GetBoolean());
        _cells = inserted.GetProperty("created").EnumerateArray()
            .Select(a => a.GetProperty("id").GetString()!).ToArray();
    }

    public void Dispose() => DocxSessionOps.CloseSession(_handle);

    private static JsonElement Json(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    private static string FirstBlockId(int handle) =>
        SessionRegistry.Get(handle).Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private long Version => DocxSessionOps.GetVersion(_handle);

    /// <summary>Asserts the facade call is refused as a caller error and leaves the session
    /// untouched (no version step, so nothing was edited under a substituted default).</summary>
    private void AssertRejected(Func<string> call, string token)
    {
        var before = Version;
        var ex = Assert.ThrowsAny<ArgumentException>(() => call());
        Assert.Contains(token, ex.Message, StringComparison.Ordinal);
        Assert.Equal(before, Version);
    }

    // ─── Through the facade, one per parser ─────────────────────────────

    [Fact]
    public void WE001_TableBorderScope_Typo_IsRejected_NotAppliedToEveryEdge() =>
        AssertRejected(() => DocxSessionOps.SetTableBorders(_handle, _cells[0], """{"scope":"outsid"}"""), "outsid");

    [Fact]
    public void WE002_ShadingScope_Typo_IsRejected() =>
        AssertRejected(() => DocxSessionOps.SetCellShading(_handle, _cells[0], "FF0000", "rows"), "rows");

    [Fact]
    public void WE003_RowHeightRule_Typo_IsRejected_NotTreatedAsAtLeast() =>
        AssertRejected(() => DocxSessionOps.SetTableRowOptions(
            _handle, _cells[0], null, null, 400, "exactly"), "exactly");

    [Fact]
    public void WE004_MergeContent_Typo_IsRejected() =>
        AssertRejected(() => DocxSessionOps.MergeCells(_handle, _cells[0], 1, 2, "discrad"), "discrad");

    [Fact]
    public void WE005_HeaderFooterKind_Typo_IsRejected() =>
        AssertRejected(() => DocxSessionOps.SetHeaderFooterKindEnabled(_handle, _paragraph, "frist", true), "frist");

    [Fact]
    public void WE006_Position_Typo_IsRejected_NotTreatedAsAfter() =>
        AssertRejected(() => DocxSessionOps.InsertParagraph(
            _handle, _paragraph, DocxSessionJson.ParsePos("befor"), "x"), "befor");

    [Fact]
    public void WE007_TrackedChangeMode_Typo_IsRejected_AtOpen() =>
        Assert.Contains("render-inline", Assert.ThrowsAny<ArgumentException>(
            () => DocxSessionJson.ParseSettings("""{"trackedChanges":"render-inline"}""")).Message);

    [Fact]
    public void WE008_ParagraphAlignment_And_LineRule_Typos_AreRejected_NotIgnored()
    {
        AssertRejected(() => DocxSessionOps.SetParagraphFormat(
            _handle, _paragraph, DocxSessionJson.ParseParagraphFormatOp("""{"alignment":"centre"}""")), "centre");
        AssertRejected(() => DocxSessionOps.SetParagraphFormat(
            _handle, _paragraph, DocxSessionJson.ParseParagraphFormatOp(
                """{"lineSpacing":300,"lineSpacingRule":"atleats"}""")), "atleats");
    }

    [Fact]
    public void WE009_PageNumberField_Typo_IsRejected_NotTreatedAsCurrentPage() =>
        Assert.Contains("totalpage", Assert.ThrowsAny<ArgumentException>(
            () => DocxSessionJson.ParsePageNumberField("totalpage")).Message);

    [Fact]
    public void WE010_NumberFormat_And_AuthorityCategory_Typos_AreRejected()
    {
        Assert.ThrowsAny<ArgumentException>(() => DocxSessionJson.ParseNumberFormatOrNull("upperroma"));
        Assert.ThrowsAny<ArgumentException>(() => DocxSessionJson.ParseAuthorityCategory("statues"));
        Assert.ThrowsAny<ArgumentException>(() => DocxSessionJson.ParseTableInsertOptions(
            """{"cellAlignment":"middle"}"""));
    }

    [Fact]
    public void WE011_ListFormat_Typo_IsRejected_NotReadAsRemoveTheList() =>
        AssertRejected(() => DocxSessionOps.ApplyListFormat(
            _handle, _paragraph, DocxSessionJson.ParseListFormat("bulet")), "bulet");

    // ─── Absent and every documented spelling still parse ───────────────

    [Fact]
    public void WE020_AbsentTokens_TakeTheDefault()
    {
        Assert.Equal(Position.After, DocxSessionJson.ParsePos(null));
        Assert.Equal(Position.After, DocxSessionJson.ParsePos(""));
        Assert.Equal(HeaderFooterKind.Default, DocxSessionJson.ParseHeaderFooterKind(null));
        Assert.Equal(PageNumberField.CurrentPage, DocxSessionJson.ParsePageNumberField(null));
        Assert.Equal(TrackedChangeMode.Accept, DocxSessionJson.ParseTrackedChangeMode(null));
        Assert.Equal(TrackedChangeMode.Accept, DocxSessionJson.ParseSettings("{}").TrackedChanges);
        Assert.Equal(TableShadingScope.Cell, DocxSessionJson.ParseTableShadingScope(null));
        Assert.Equal(TableMergeContent.Append, DocxSessionJson.ParseTableMergeContent(null));
        Assert.Equal(TableRowHeightRule.AtLeast, DocxSessionJson.ParseTableRowHeightRule(null));
        Assert.Equal(TableBorderScope.All, DocxSessionJson.ParseTableBorderSpec("{}").Scope);
        Assert.Null(DocxSessionJson.ParseNumberFormatOrNull(""));
        Assert.Null(DocxSessionJson.ParseAuthorityCategory(null));
        Assert.Null(DocxSessionJson.ParseParagraphFormatOp("{}").Alignment);
        Assert.Equal(ListFormat.None, DocxSessionJson.ParseListFormat(null));
        Assert.Equal(ListFormat.None, DocxSessionJson.ParseListFormat("none"));
    }

    [Theory]
    [InlineData("current_page", PageNumberField.CurrentPage)]   // MCP schema spelling
    [InlineData("currentPage", PageNumberField.CurrentPage)]    // TS / Python spelling
    [InlineData("total_pages", PageNumberField.TotalPages)]
    [InlineData("totalPages", PageNumberField.TotalPages)]
    [InlineData("numpages", PageNumberField.TotalPages)]
    [InlineData("page_of_total", PageNumberField.PageOfTotal)]
    [InlineData("pageOfTotal", PageNumberField.PageOfTotal)]
    public void WE021_PageNumberField_AcceptsEveryClientSpelling(string token, PageNumberField expected) =>
        Assert.Equal(expected, DocxSessionJson.ParsePageNumberField(token));

    [Fact]
    public void WE022_KnownTokens_ParseCaseInsensitively()
    {
        Assert.Equal(Position.Before, DocxSessionJson.ParsePos("BEFORE"));
        Assert.Equal(HeaderFooterKind.Even, DocxSessionJson.ParseHeaderFooterKind("Even"));
        Assert.Equal(TableRowHeightRule.AtLeast, DocxSessionJson.ParseTableRowHeightRule("atLeast"));
        Assert.Equal(TableRowHeightRule.Exact, DocxSessionJson.ParseTableRowHeightRule("exact"));
        Assert.Equal(TableShadingScope.Row, DocxSessionJson.ParseTableShadingScope("Row"));
        Assert.Equal(TableMergeContent.Reject, DocxSessionJson.ParseTableMergeContent("reject"));
        Assert.Equal(TableBorderScope.Inside, DocxSessionJson.ParseTableBorderSpec("""{"scope":"inside"}""").Scope);
        Assert.Equal(TrackedChangeMode.StripDeletions, DocxSessionJson.ParseTrackedChangeMode("strip_deletions"));
        Assert.Equal(ParagraphAlignment.Justify, DocxSessionJson.ParseParagraphFormatOp("""{"alignment":"both"}""").Alignment);
        Assert.Equal(NumberFormat.UpperRoman, DocxSessionJson.ParseNumberFormatOrNull("upperRoman"));
        Assert.Equal(AuthorityCategory.OtherAuthorities, DocxSessionJson.ParseAuthorityCategory("other_authorities"));
    }

    // ─── Through the transports ─────────────────────────────────────────

    [Fact]
    public void WE030_StdioHost_ReportsTypo_AsArgumentError()
    {
        // Program.cs maps an ArgumentException escaping Dispatch to the invalid_argument envelope.
        var args = Json($$$"""{"handle":{{{_handle}}},"cellAnchorId":"{{{_cells[0]}}}","spec":{"scope":"outsid"}}""");
        var before = Version;
        Assert.ThrowsAny<ArgumentException>(() => Docxodus.PyHost.Dispatcher.Dispatch("set_table_borders", args));
        Assert.Equal(before, Version);
    }

    [Fact]
    public void WE031_Mcp_ReportsTypo_AsToolError()
    {
        var root = Path.Combine(Path.GetTempPath(), $"we031-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new Docxodus.McpServer.SessionStore(new Docxodus.McpServer.LocalFileDocumentStore(root));
        try
        {
            var path = Path.Combine(root, "d.docx");
            File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
            var sessionId = Json(Docxodus.McpServer.Dispatcher.Call(store, "docxodus_open",
                Json($$"""{"path":{{JsonSerializer.Serialize(path)}}}"""))).GetProperty("sessionId").GetString()!;
            var sid = JsonSerializer.Serialize(sessionId);
            var anchor = Json(Docxodus.McpServer.Dispatcher.Call(store, "docxodus_search",
                Json($$"""{"sessionId":{{sid}},"mode":"kind","query":"p"}""")))
                .GetProperty("matches")[0].GetProperty("id").GetString()!;
            var inserted = Json(Docxodus.McpServer.Dispatcher.Call(store, "docxodus_table", Json(
                $$"""{"sessionId":{{sid}},"action":"insert","anchorId":"{{anchor}}","position":"after","rows":2,"columns":2}""")));
            var cell = inserted.GetProperty("created")[0].GetProperty("id").GetString()!;

            // Program.cs turns any exception escaping Call into an isError tool result.
            var ex = Assert.ThrowsAny<Exception>(() => Docxodus.McpServer.Dispatcher.Call(store, "docxodus_table", Json(
                $$"""{"sessionId":{{sid}},"action":"set_borders","cellAnchorId":"{{cell}}","borderScope":"outsid"}""")));
            Assert.Contains("outsid", ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }
}
