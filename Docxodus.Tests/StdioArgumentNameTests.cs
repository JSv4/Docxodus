// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Every session op takes one name per argument on the stdio host and on MCP (issue #1023). The
/// stdio host takes the canonical (MCP) name and keeps its old spelling as a deprecated alias.
/// Each test calls the host for real and checks the argument took effect: several of the old
/// spellings were optional, so a host that ignored the canonical name would still have answered
/// success.
/// </summary>
public sealed class StdioArgumentNameTests : IDisposable
{
    private readonly int _handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);

    public void Dispose() => DocxSessionOps.CloseSession(_handle);

    private DocxSession Session => SessionRegistry.Get(_handle);

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    private JsonElement Call(string op, Dictionary<string, object?> args)
    {
        args["handle"] = _handle;
        using var doc = JsonDocument.Parse(Docxodus.PyHost.Dispatcher.Dispatch(op, J(args)));
        var result = doc.RootElement.Clone();
        if (result.ValueKind == JsonValueKind.Object && result.TryGetProperty("success", out var ok))
            Assert.True(ok.GetBoolean(), $"{op} failed: {result}");
        return result;
    }

    private string Paragraph(string text)
    {
        var first = Session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        return Session.ReplaceText(first, text).Modified.Select(a => a.Id).FirstOrDefault() ?? first;
    }

    private string Append(string after, string text) =>
        Session.InsertParagraph(after, Position.After, text).Created.First().Id;

    private string Xml(string anchorId) => DocxSessionOps.RawGetXml(_handle, anchorId);

    // ─── Scope (list_hyperlinks, list_bookmarks, list_images, list_content_controls) ───

    [Theory]
    [InlineData("list_hyperlinks")]
    [InlineData("list_bookmarks")]
    [InlineData("list_images")]
    [InlineData("list_content_controls")]
    public void Listings_TakeScope_AsMcpsTokenOrAsAScopeMask(string op)
    {
        var p = Paragraph("Visit the site today");
        DocxSessionOps.AddHyperlink(_handle, p, 6, 4, "external", "https://example.com");
        DocxSessionOps.AddBookmark(_handle, "mark", p, 0, p, 5);

        int Count(object scopeArgs) => Call(op, (Dictionary<string, object?>)scopeArgs).GetArrayLength();
        var body = op is "list_hyperlinks" or "list_bookmarks" ? 1 : 0;

        Assert.Equal(body, Count(new Dictionary<string, object?> { ["scope"] = "body" }));
        Assert.Equal(0, Count(new Dictionary<string, object?> { ["scope"] = "headers" }));
        Assert.Equal(body, Count(new Dictionary<string, object?> { ["scope"] = (int)ProjectionScopes.Body }));
        Assert.Equal(0, Count(new Dictionary<string, object?> { ["scope"] = (int)ProjectionScopes.Headers }));
        // The old spelling, a scope mask under "scopes", still works.
        Assert.Equal(0, Count(new Dictionary<string, object?> { ["scopes"] = (int)ProjectionScopes.Footers }));
        Assert.Equal(body, Count(new Dictionary<string, object?>()));
    }

    [Fact]
    public void Scope_AndItsDeprecatedAlias_MustAgree()
    {
        Paragraph("text");
        Assert.Throws<ArgumentException>(() => Call("list_hyperlinks",
            new Dictionary<string, object?> { ["scope"] = 1, ["scopes"] = 2 }));
    }

    [Fact]
    public void Scope_RefusesAnUnknownToken()
    {
        Paragraph("text");
        Assert.Throws<FormatException>(() => Call("list_hyperlinks",
            new Dictionary<string, object?> { ["scope"] = "everywhere" }));
    }

    // ─── Formatting ─────────────────────────────────────────────────────

    [Fact]
    public void ApplyFormat_TakesFormat()
    {
        var p = Paragraph("Hello world");
        Call("apply_format", new() { ["anchorId"] = p, ["format"] = new { bold = true } });
        Assert.Contains("<w:b", Xml(p));
    }

    [Fact]
    public void ApplyFormat_StillTakesItsDeprecatedSpelling()
    {
        var p = Paragraph("Hello world");
        Call("apply_format", new() { ["anchorId"] = p, ["op"] = new { italic = true } });
        Assert.Contains("<w:i", Xml(p));
    }

    [Fact]
    public void ApplyFormatBySubstring_TakesFormat()
    {
        var p = Paragraph("Hello world");
        Call("apply_format_by_substring", new() { ["anchorId"] = p, ["substring"] = "world", ["format"] = new { bold = true } });
        Assert.Contains("<w:b", Xml(p));
    }

    [Fact]
    public void SetParagraphFormat_TakesParagraphFormat()
    {
        var p = Paragraph("Hello world");
        Call("set_paragraph_format", new() { ["anchorId"] = p, ["paragraphFormat"] = new { alignment = "center" } });
        Assert.Contains("w:val=\"center\"", Xml(p));
    }

    [Fact]
    public void SetParagraphFormat_StillTakesItsDeprecatedSpelling()
    {
        var p = Paragraph("Hello world");
        Call("set_paragraph_format", new() { ["anchorId"] = p, ["op"] = new { alignment = "right" } });
        Assert.Contains("w:val=\"right\"", Xml(p));
    }

    // ─── Headers and footers ────────────────────────────────────────────

    [Theory]
    [InlineData("set_header_text", "bodyAnchorId")]
    [InlineData("set_footer_text", "bodyAnchorId")]
    [InlineData("set_header_text", "anchorId")]
    [InlineData("set_footer_text", "anchorId")]
    public void SetHeaderFooterText_TakesBodyAnchorId_OrItsDeprecatedSpelling(string op, string name)
    {
        var p = Paragraph("Body");
        Call(op, new() { [name] = p, ["kind"] = "default", ["markdown"] = "Running text" });
        var scope = op == "set_header_text" ? "hdr" : "ftr";
        Assert.Contains(Session.Project().AnchorIndex.Values, t => t.Anchor.Scope.StartsWith(scope, StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("bodyAnchorId")]
    [InlineData("anchorId")]
    public void EnsureHeaderFooterVisible_TakesBodyAnchorId_OrItsDeprecatedSpelling(string name)
    {
        var p = Paragraph("Body");
        Call("ensure_header_footer_visible", new() { [name] = p, ["kind"] = "first" });
        Assert.True(Session.GetSectionInfo(p)!.TitlePage);
    }

    [Theory]
    [InlineData("numberFormat")]
    [InlineData("format")]
    public void InsertPageNumberField_TakesNumberFormat_OrItsDeprecatedSpelling(string name)
    {
        var p = Paragraph("Body");
        Session.SetFooterText(p, HeaderFooterKind.Default, "Page ");
        var footer = Session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope.StartsWith("ftr", StringComparison.Ordinal) && t.Anchor.Kind == "p").Anchor.Id;
        Call("insert_page_number_field", new() { ["anchorId"] = footer, ["field"] = "currentPage", [name] = "upperRoman" });
        Assert.Contains("ROMAN", Xml(footer));
    }

    // ─── Links, lists, tables, text, revisions ──────────────────────────

    [Theory]
    [InlineData("startOffset")]
    [InlineData("start")]
    public void AddHyperlink_TakesStartOffset_OrItsDeprecatedSpelling(string name)
    {
        var p = Paragraph("Visit the site today");
        Call("add_hyperlink", new() { ["anchorId"] = p, [name] = 6, ["length"] = 4, ["kind"] = "external", ["target"] = "https://example.com" });
        var link = Assert.Single(Session.ListHyperlinks());
        Assert.Equal("the ", link.Text);
    }

    [Theory]
    [InlineData("startValue")]
    [InlineData("value")]
    public void SetListStartOverride_TakesStartValue_OrItsDeprecatedSpelling(string name)
    {
        var p = Paragraph("First item");
        Session.ApplyListFormat(p, ListFormat.Decimal);
        Call("set_list_start_override", new() { ["anchorId"] = p, [name] = 7 });
        Assert.Contains("w:startOverride w:val=\"7\"", RawNumbering());
    }

    private string RawNumbering()
    {
        using var ms = new MemoryStream(Session.Save());
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml;
    }

    [Theory]
    [InlineData("anchorId")]
    [InlineData("firstAnchorId")]
    public void MergeParagraphs_TakesAnchorId_OrItsDeprecatedSpelling(string name)
    {
        var first = Paragraph("One");
        var second = Append(first, "Two");
        Call("merge_paragraphs", new() { [name] = first, ["secondAnchorId"] = second });
        Assert.DoesNotContain(Session.Project().AnchorIndex.Values, t => t.Anchor.Id == second);
    }

    [Fact]
    public void MergeCells_TakesColSpanAndMergeContent()
    {
        var cell = SeedTable();
        Call("merge_cells", new() { ["cellAnchorId"] = cell, ["rowSpan"] = 1, ["colSpan"] = 2, ["mergeContent"] = "discard" });
        var table = Xml(TableAnchor());
        Assert.Contains("w:gridSpan w:val=\"2\"", table);
        Assert.DoesNotContain("B1", table);
    }

    [Fact]
    public void MergeCells_StillTakesItsDeprecatedSpellings()
    {
        var cell = SeedTable();
        Call("merge_cells", new() { ["cellAnchorId"] = cell, ["rowSpan"] = 1, ["columnSpan"] = 2, ["content"] = "discard" });
        var table = Xml(TableAnchor());
        Assert.Contains("w:gridSpan w:val=\"2\"", table);
        Assert.DoesNotContain("B1", table);
    }

    private string SeedTable()
    {
        var p = Paragraph("Before");
        Session.InsertTable(p, Position.After, 1, 2, new TableInsertOptions { CellContents = new[] { "A1", "B1" } });
        return Session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "tc").Anchor.Id;
    }

    private string TableAnchor() => Session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "tbl").Anchor.Id;

    [Theory]
    [InlineData("revisionAuthor")]
    [InlineData("author")]
    public void SetRevisionAuthor_TakesRevisionAuthor_OrItsDeprecatedSpelling(string name)
    {
        Call("set_revision_author", new() { [name] = "Ada" });
        Assert.Equal("Ada", Session.RevisionAuthor);
    }

    [Fact]
    public void SetRevisionAuthor_WithNoAuthor_ClearsIt()
    {
        Call("set_revision_author", new() { ["revisionAuthor"] = "Ada" });
        Call("set_revision_author", new());
        Assert.Null(Session.RevisionAuthor);
    }

    // ─── The descriptions ───────────────────────────────────────────────

    /// <summary>The descriptions record no stdio spelling that differs from the canonical name.</summary>
    [Fact]
    public void NoDescriptionRecordsAStdioArgumentSpelling()
    {
        var recorded = new List<string>();
        foreach (var file in Directory.GetFiles(Path.Combine("../../../..", "tools", "op-descriptions"), "*.json"))
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file));
            if (!doc.RootElement.TryGetProperty("ops", out var ops)) continue;
            foreach (var op in ops.EnumerateArray())
                if (op.GetProperty("transports").TryGetProperty("stdio", out var stdio) && stdio.TryGetProperty("argNames", out var names))
                    recorded.Add($"{Path.GetFileName(file)}/{op.GetProperty("op").GetString()}: {names}");
        }
        Assert.True(recorded.Count == 0, "stdio argNames still recorded:\n" + string.Join("\n", recorded));
    }
}
