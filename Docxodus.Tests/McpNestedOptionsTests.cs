// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1025: fifteen session ops take one object argument (<c>options</c>, <c>rule</c>,
/// <c>spec</c>) on the facade, the stdio host, WASM and npm, but MCP spread the object's fields
/// across top-level properties and could not reach some of them at all. MCP now takes the same
/// nested object. The flat properties stay as deprecated aliases, and a call that names a field
/// both ways with different values is refused. Each test calls the MCP dispatcher and checks the
/// option took effect, not just that the call succeeded: every option is optional, so a server
/// that ignored the nested object would still answer success.
/// </summary>
[Collection("MCP session registry isolation")]
public sealed class McpNestedOptionsTests : IDisposable
{
    private const string RepoRoot = "../../../..";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly string _root;
    private readonly SessionStore _store;

    public McpNestedOptionsTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-nested-options-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    // ─── Harness ────────────────────────────────────────────────────────

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(value as string ?? JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    private string Open(byte[]? bytes = null)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, bytes ?? DocxSession.CreateBlankDocxBytes());
        return J(Dispatcher.Call(_store, "docxodus_open", J(new { path }))).GetProperty("sessionId").GetString()!;
    }

    private DocxSession Session(string sessionId) => SessionRegistry.Get(_store.Get(sessionId).Handle);

    /// <summary>Call a tool with <paramref name="args"/> plus the session id.</summary>
    private JsonElement Call(string sessionId, string tool, Dictionary<string, object?> args)
    {
        args["sessionId"] = sessionId;
        return J(Dispatcher.Call(_store, tool, J(args)));
    }

    /// <summary>Whether an edit result (or every result of an op that answers a list) succeeded.</summary>
    private static bool Success(JsonElement result) => result.ValueKind == JsonValueKind.Array
        ? result.GetArrayLength() > 0 && result.EnumerateArray().All(Success)
        : result.GetProperty("success").GetBoolean();

    private JsonElement Succeeds(string sessionId, string tool, Dictionary<string, object?> args)
    {
        var result = Call(sessionId, tool, args);
        Assert.True(Success(result), result.ToString());
        return result;
    }

    private static string ErrorCode(JsonElement result) =>
        result.GetProperty("error").GetProperty("code").GetString()!;

    /// <summary>A body paragraph with <paramref name="text"/>.</summary>
    private string Paragraph(string sessionId, string text)
    {
        var session = Session(sessionId);
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        return session.ReplaceText(first, text).Modified.Select(a => a.Id).FirstOrDefault() ?? first;
    }

    private string Text(string sessionId, string anchorId) =>
        Session(sessionId).Project().AnchorIndex.Values.Single(t => t.Anchor.Id == anchorId).TextPreview;

    /// <summary>The session's main document part as saved.</summary>
    private XDocument Document(string sessionId)
    {
        using var zip = new ZipArchive(new MemoryStream(Session(sessionId).Save()), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("word/document.xml")!.Open();
        return XDocument.Load(stream);
    }

    /// <summary>Every field instruction in the body, with runs of a field joined.</summary>
    private List<string> FieldInstructions(string sessionId)
    {
        var body = Document(sessionId).Root!.Element(W + "body")!;
        var simple = body.Descendants(W + "fldSimple").Select(f => ((string)f.Attribute(W + "instr")!).Trim());
        var complex = new List<string>();
        string? current = null;
        foreach (var e in body.Descendants().Where(e => e.Name == W + "fldChar" || e.Name == W + "instrText"))
        {
            if (e.Name == W + "instrText") { if (current is not null) current += e.Value; continue; }
            var type = (string?)e.Attribute(W + "fldCharType");
            if (type == "begin") current = "";
            else if (type == "separate" && current is not null) { complex.Add(current.Trim()); current = null; }
        }
        return simple.Concat(complex).ToList();
    }

    private static Dictionary<string, object?> A(params (string Key, object? Value)[] pairs) =>
        pairs.ToDictionary(p => p.Key, p => p.Value);

    // ─── replaceTextRange ───────────────────────────────────────────────

    [Fact]
    public void ReplaceTextRange_NestedMaxReplacements_CapsTheReplacements()
    {
        var sid = Open();
        var p = Paragraph(sid, "one cat two cat three cat");
        Succeeds(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", p),
            ("find", "cat"), ("replace", "dog"), ("options", new { maxReplacements = 1 })));
        Assert.Equal("one dog two cat three cat", Text(sid, p));
    }

    [Fact]
    public void ReplaceTextRange_NestedExpectedMatchCount_RefusesAMismatch()
    {
        var sid = Open();
        var p = Paragraph(sid, "one cat two cat three cat");
        var refused = Call(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", p),
            ("find", "cat"), ("replace", "dog"), ("options", new { expectedMatchCount = 2 })));
        Assert.False(Success(refused), refused.ToString());
        Assert.Equal("one cat two cat three cat", Text(sid, p));
        Succeeds(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", p),
            ("find", "cat"), ("replace", "dog"), ("options", new { expectedMatchCount = 3 })));
        Assert.Equal("one dog two dog three dog", Text(sid, p));
    }

    [Fact]
    public void ReplaceTextRange_NestedIgnoreCase_AndMcpsCaseInsensitiveDefault()
    {
        var sid = Open();
        var p = Paragraph(sid, "Cat cat");
        // MCP has always matched case-insensitively when told nothing.
        Succeeds(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", p),
            ("find", "CAT"), ("replace", "dog")));
        Assert.Equal("dog dog", Text(sid, p));

        var q = Paragraph(sid, "Cat cat");
        var caseSensitive = Call(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", q),
            ("find", "CAT"), ("replace", "dog"), ("options", new { ignoreCase = false })));
        Assert.False(Success(caseSensitive), caseSensitive.ToString());
        Assert.Equal("Cat cat", Text(sid, q));
    }

    [Fact]
    public void ReplaceTextRange_FlatCaseSensitive_IsTheNegatedAliasOfIgnoreCase()
    {
        var sid = Open();
        var p = Paragraph(sid, "Cat cat");
        Succeeds(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", p),
            ("find", "cat"), ("replace", "dog"), ("caseSensitive", true)));
        Assert.Equal("Cat dog", Text(sid, p));

        // ignoreCase true and caseSensitive false say the same thing.
        var q = Paragraph(sid, "Cat cat");
        Succeeds(sid, "docxodus_edit", A(("action", "replace_text_range"), ("anchorId", q),
            ("find", "cat"), ("replace", "dog"), ("caseSensitive", false), ("options", new { ignoreCase = true })));
        Assert.Equal("dog dog", Text(sid, q));

        // ignoreCase true and caseSensitive true contradict each other.
        var ex = Assert.Throws<McpToolException>(() => Call(sid, "docxodus_edit", A(("action", "replace_text_range"),
            ("anchorId", q), ("find", "dog"), ("replace", "x"), ("caseSensitive", true), ("options", new { ignoreCase = true }))));
        Assert.Contains("options.ignoreCase", ex.Message, StringComparison.Ordinal);
        Assert.Equal("dog dog", Text(sid, q));
    }

    // ─── insertHorizontalRule ───────────────────────────────────────────

    private XElement RuleEdge(string sessionId, JsonElement result)
    {
        var created = result.GetProperty("created")[0].GetProperty("id").GetString()!;
        var p = XElement.Parse(DocxSessionOps.RawGetXml(_store.Get(sessionId).Handle, created));
        return p.Descendants(W + "pBdr").Single().Element(W + "bottom")!;
    }

    [Fact]
    public void InsertHorizontalRule_NestedRule_ReachesSizeColourAndSpacing()
    {
        var sid = Open();
        var p = Paragraph(sid, "Above the rule");
        var result = Succeeds(sid, "docxodus_create", A(("action", "insert_horizontal_rule"), ("anchorId", p),
            ("rule", new { style = "dotted", size = 24, color = "FF0000", space = 4 })));
        var edge = RuleEdge(sid, result);
        Assert.Equal("dotted", (string?)edge.Attribute(W + "val"));
        Assert.Equal("24", (string?)edge.Attribute(W + "sz"));
        Assert.Equal("FF0000", (string?)edge.Attribute(W + "color"));
        Assert.Equal("4", (string?)edge.Attribute(W + "space"));
    }

    [Fact]
    public void InsertHorizontalRule_FlatRuleStyle_StillWorks_AndConflictsAreRefused()
    {
        var sid = Open();
        var p = Paragraph(sid, "Above the rule");
        var result = Succeeds(sid, "docxodus_create", A(("action", "insert_horizontal_rule"), ("anchorId", p),
            ("ruleStyle", "double")));
        Assert.Equal("double", (string?)RuleEdge(sid, result).Attribute(W + "val"));

        Assert.Throws<McpToolException>(() => Call(sid, "docxodus_create", A(("action", "insert_horizontal_rule"),
            ("anchorId", p), ("ruleStyle", "double"), ("rule", new { style = "single" }))));
    }

    [Fact]
    public void InsertHorizontalRule_ABatchStepAcceptsWhatADirectCallAccepts()
    {
        // The flat ruleStyle enum (single | double | thick) does not restrict rule.style.
        var sid = Open();
        var p = Paragraph(sid, "Above the rule");
        var batch = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId = sid,
            steps = new[]
            {
                new { tool = "docxodus_create", args = (object)new { action = "insert_horizontal_rule", anchorId = p, rule = new { style = "dotted" } } },
            },
        })));
        Assert.True(batch.GetProperty("success").GetBoolean(), batch.ToString());
        var rule = Document(sid).Descendants(W + "pBdr").Single().Element(W + "bottom")!;
        Assert.Equal("dotted", (string?)rule.Attribute(W + "val"));
    }

    // ─── insertTable (docxodus_create and docxodus_table) ───────────────

    [Theory]
    [InlineData("docxodus_create", "insert_table")]
    [InlineData("docxodus_table", "insert")]
    public void InsertTable_NestedOptions_TakeEffect(string tool, string action)
    {
        var sid = Open();
        var p = Paragraph(sid, "Before the table");
        Succeeds(sid, tool, A(("action", action), ("anchorId", p), ("rows", 1), ("columns", 2),
            ("options", new { cellContents = new[] { "left cell", "right cell" }, columnWidths = new[] { 2000, 3000 }, cellAlignment = "center" })));
        var table = Document(sid).Descendants(W + "tbl").Single();
        Assert.Equal(new[] { "2000", "3000" },
            table.Element(W + "tblGrid")!.Elements(W + "gridCol").Select(c => (string)c.Attribute(W + "w")!));
        Assert.Equal(new[] { "left cell", "right cell" },
            table.Descendants(W + "tc").Select(c => string.Concat(c.Descendants(W + "t").Select(t => t.Value))));
        Assert.All(table.Descendants(W + "tc"), c =>
            Assert.Equal("center", (string?)c.Descendants(W + "jc").Single().Attribute(W + "val")));
    }

    [Theory]
    [InlineData("docxodus_create", "insert_table")]
    [InlineData("docxodus_table", "insert")]
    public void InsertTable_FlatAliases_StillWork_AndConflictsAreRefused(string tool, string action)
    {
        var sid = Open();
        var p = Paragraph(sid, "Before the table");
        Succeeds(sid, tool, A(("action", action), ("anchorId", p), ("rows", 1), ("columns", 2),
            ("columnWidths", new[] { 2500, 2500 })));
        Assert.Equal(new[] { "2500", "2500" }, Document(sid).Descendants(W + "gridCol").Select(c => (string)c.Attribute(W + "w")!));

        Assert.Throws<McpToolException>(() => Call(sid, tool, A(("action", action), ("anchorId", p), ("rows", 1),
            ("columns", 2), ("borderless", true), ("options", new { borderless = false }))));
        Assert.Single(Document(sid).Descendants(W + "tbl"));
    }

    // ─── setTableBorders ────────────────────────────────────────────────

    private string Table(string sid)
    {
        var p = Paragraph(sid, "Before the table");
        var inserted = Succeeds(sid, "docxodus_table", A(("action", "insert"), ("anchorId", p), ("rows", 2), ("columns", 2)));
        return inserted.GetProperty("created")[0].GetProperty("id").GetString()!;
    }

    [Fact]
    public void SetTableBorders_NestedSpec_TakesEffect()
    {
        var sid = Open();
        var cell = Table(sid);
        Succeeds(sid, "docxodus_table", A(("action", "set_borders"), ("cellAnchorId", cell),
            ("spec", new { scope = "outside", style = "double", size = 12, color = "FF0000" })));
        var borders = Document(sid).Descendants(W + "tblBorders").Single();
        var top = borders.Element(W + "top")!;
        Assert.Equal("double", (string?)top.Attribute(W + "val"));
        Assert.Equal("12", (string?)top.Attribute(W + "sz"));
        Assert.Equal("FF0000", (string?)top.Attribute(W + "color"));
        // outside: the inside edges were not written with the spec.
        Assert.NotEqual("FF0000", (string?)borders.Element(W + "insideH")?.Attribute(W + "color"));
    }

    [Fact]
    public void SetTableBorders_SpecScopeAndItsFlatAliasBorderScope_MustAgree()
    {
        var sid = Open();
        var cell = Table(sid);
        Succeeds(sid, "docxodus_table", A(("action", "set_borders"), ("cellAnchorId", cell),
            ("borderScope", "inside"), ("spec", new { scope = "inside", color = "00FF00" })));
        var borders = Document(sid).Descendants(W + "tblBorders").Single();
        Assert.Equal("00FF00", (string?)borders.Element(W + "insideH")!.Attribute(W + "color"));
        Assert.NotEqual("00FF00", (string?)borders.Element(W + "top")?.Attribute(W + "color"));

        var ex = Assert.Throws<McpToolException>(() => Call(sid, "docxodus_table", A(("action", "set_borders"),
            ("cellAnchorId", cell), ("borderScope", "outside"), ("spec", new { scope = "inside" }))));
        Assert.Contains("spec.scope", ex.Message, StringComparison.Ordinal);
    }

    // ─── Reference fields ───────────────────────────────────────────────

    [Fact]
    public void InsertTableOfContents_NestedOptions_ReachEverySwitch()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Succeeds(sid, "docxodus_create", A(("action", "insert_table_of_contents"), ("anchorId", p),
            ("options", new { levels = "1-2", hyperlinks = false, hideTabAndPageNumbersInWeb = false, useOutlineLevels = false, title = "Index" })));
        var toc = Assert.Single(FieldInstructions(sid), i => i.StartsWith("TOC", StringComparison.Ordinal));
        Assert.Contains("\\o \"1-2\"", toc, StringComparison.Ordinal);
        Assert.DoesNotContain("\\h", toc, StringComparison.Ordinal);
        Assert.DoesNotContain("\\z", toc, StringComparison.Ordinal);
        Assert.DoesNotContain("\\u", toc, StringComparison.Ordinal);
        Assert.Contains(Document(sid).Descendants(W + "t"), t => t.Value == "Index");
    }

    [Fact]
    public void InsertTableOfContents_UnadvertisedFlatAliases_StillWork()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Succeeds(sid, "docxodus_create", A(("action", "insert_table_of_contents"), ("anchorId", p),
            ("hideTabAndPageNumbersInWeb", false), ("useOutlineLevels", false), ("levels", "2")));
        var toc = Assert.Single(FieldInstructions(sid), i => i.StartsWith("TOC", StringComparison.Ordinal));
        Assert.Contains("\\o \"2-2\"", toc, StringComparison.Ordinal);
        Assert.DoesNotContain("\\z", toc, StringComparison.Ordinal);
        Assert.DoesNotContain("\\u", toc, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertTableOfContents_ExplicitNullTitle_InsertsNoHeading()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Succeeds(sid, "docxodus_create", A(("action", "insert_table_of_contents"), ("anchorId", p),
            ("options", new Dictionary<string, object?> { ["title"] = null })));
        Assert.DoesNotContain(Document(sid).Descendants(W + "t"), t => t.Value == "Contents");
    }

    [Fact]
    public void InsertTableOfFigures_NestedOptions_TakeEffect()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Succeeds(sid, "docxodus_create", A(("action", "insert_table_of_figures"), ("anchorId", p),
            ("options", new { captionLabel = "Exhibit", hyperlinks = false })));
        var tof = Assert.Single(FieldInstructions(sid), i => i.StartsWith("TOC", StringComparison.Ordinal));
        Assert.Contains("\\c \"Exhibit\"", tof, StringComparison.Ordinal);
        Assert.DoesNotContain("\\h", tof, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertTableOfAuthorities_NestedOptions_TakeEffect()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Succeeds(sid, "docxodus_create", A(("action", "insert_table_of_authorities"), ("anchorId", p),
            ("options", new { category = "statutes", entryPageSeparator = "; ", hyperlinks = false })));
        var toa = Assert.Single(FieldInstructions(sid), i => i.StartsWith("TOA", StringComparison.Ordinal));
        Assert.Contains("\\c \"2\"", toa, StringComparison.Ordinal);
        Assert.Contains("\\e \"; \"", toa, StringComparison.Ordinal);
        Assert.DoesNotContain("\\h", toa, StringComparison.Ordinal);
    }

    [Fact]
    public void ReferenceTables_NestedAndFlatHyperlinks_MustAgree()
    {
        var sid = Open();
        var p = Paragraph(sid, "Body text");
        Assert.Throws<McpToolException>(() => Call(sid, "docxodus_create", A(("action", "insert_table_of_figures"),
            ("anchorId", p), ("hyperlinks", true), ("options", new { hyperlinks = false }))));
        Assert.Empty(FieldInstructions(sid));
    }

    [Fact]
    public void InsertCrossReference_NestedOptions_ReachEverySwitch()
    {
        var sid = Open();
        var p = Paragraph(sid, "Target text and a reference here");
        Assert.Contains("\"success\":true", DocxSessionOps.AddBookmark(_store.Get(sid).Handle, "target", p, 0, p, 6), StringComparison.Ordinal);
        Succeeds(sid, "docxodus_links", A(("action", "insert_cross_reference"), ("anchorId", p), ("characterOffset", 32),
            ("bookmarkName", "target"), ("options", new { referenceNumber = true, hyperlink = true, includePosition = true })));
        var reference = Assert.Single(FieldInstructions(sid), i => i.StartsWith("REF", StringComparison.Ordinal));
        Assert.Contains("\\r", reference, StringComparison.Ordinal);
        Assert.Contains("\\h", reference, StringComparison.Ordinal);
        Assert.Contains("\\p", reference, StringComparison.Ordinal);
    }

    [Fact]
    public void InsertCrossReference_FlatAliases_StillWork()
    {
        var sid = Open();
        var p = Paragraph(sid, "Target text and a reference here");
        DocxSessionOps.AddBookmark(_store.Get(sid).Handle, "target", p, 0, p, 6);
        Succeeds(sid, "docxodus_links", A(("action", "insert_cross_reference"), ("anchorId", p), ("characterOffset", 32),
            ("bookmarkName", "target"), ("hyperlink", true)));
        var reference = Assert.Single(FieldInstructions(sid), i => i.StartsWith("REF", StringComparison.Ordinal));
        Assert.Contains("\\h", reference, StringComparison.Ordinal);
        Assert.DoesNotContain("\\r", reference, StringComparison.Ordinal);
    }

    // ─── Content controls ───────────────────────────────────────────────

    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    /// <summary>The content-control fixture, plus a picture control (113), with the controls named
    /// in <paramref name="bound"/> data-bound so a mutation fails closed without
    /// <c>bindingPolicy: detach_target</c>.</summary>
    private static byte[] ControlsFixture(params string[] bound)
    {
        var stream = new MemoryStream();
        stream.Write(DocxSessionContentControlTests.BuildFixture());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var body = main.GetXDocument().Root!.Element(W + "body")!;
            body.Add(new XElement(W + "sdt",
                new XElement(W + "sdtPr", new XElement(W + "id", new XAttribute(W + "val", "113")), new XElement(W + "picture")),
                new XElement(W + "sdtContent", new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", "picture placeholder"))))));
            foreach (var id in bound)
            {
                var properties = body.Descendants(W + "sdtPr")
                    .Single(p => (string?)p.Element(W + "id")?.Attribute(W + "val") == id);
                properties.Elements().Last().AddBeforeSelf(new XElement(W + "dataBinding",
                    new XAttribute(W + "storeItemID", "{11111111-1111-1111-1111-111111111111}"),
                    new XAttribute(W + "xpath", $"/root/c{id}"),
                    new XAttribute(W + "prefixMappings", "xmlns:x='urn:test'")));
            }
            main.PutXDocument();
        }
        using var seed = new DocxSession(stream.ToArray());
        var placeholder = seed.Project().AnchorIndex.Values.Single(v =>
            v.Anchor.Kind == "p" && v.TextPreview.Contains("picture placeholder", StringComparison.Ordinal));
        Assert.True(seed.InsertImage(placeholder.Anchor.Id, 0, Png(2, 3)).Success);
        return seed.Save();
    }

    private static byte[] Png(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(bytes, 0);
        bytes[16] = (byte)(width >> 24); bytes[17] = (byte)(width >> 16); bytes[18] = (byte)(width >> 8); bytes[19] = (byte)width;
        bytes[20] = (byte)(height >> 24); bytes[21] = (byte)(height >> 16); bytes[22] = (byte)(height >> 8); bytes[23] = (byte)height;
        return bytes;
    }

    private JsonElement Control(string sid, string nativeId) =>
        Call(sid, "docxodus_content_controls", A(("action", "list"))).GetProperty("contentControls").EnumerateArray()
            .Single(c => c.TryGetProperty("nativeId", out var id) && id.GetString() == nativeId);

    private string ControlAnchor(string sid, string nativeId) => Control(sid, nativeId).GetProperty("anchorId").GetString()!;

    /// <summary>One mutating content-control action against a bound control of its type.</summary>
    public static IEnumerable<object[]> BoundControlActions() => new[]
    {
        new object[] { "fill_text", "106" },
        new object[] { "fill_rich_text", "100" },
        new object[] { "set_checked", "102" },
        new object[] { "set_date", "103" },
        new object[] { "select_item", "104" },
        new object[] { "fill_picture", "113" },
        new object[] { "add_repeating_item", "108" },
    };

    private Dictionary<string, object?> ControlArgs(string sid, string action, string nativeId)
    {
        var anchor = ControlAnchor(sid, nativeId);
        var args = A(("action", action));
        switch (action)
        {
            case "fill_text": args["anchorId"] = anchor; args["text"] = "filled"; break;
            case "fill_rich_text": args["anchorId"] = anchor; args["markdown"] = "**filled**"; break;
            case "set_checked": args["anchorId"] = anchor; args["checked"] = true; break;
            case "set_date": args["anchorId"] = anchor; args["value"] = "2026-01-02"; break;
            case "select_item": args["anchorId"] = anchor; args["value"] = "b"; break;
            case "fill_picture": args["anchorId"] = anchor; args["imageBase64"] = Convert.ToBase64String(Png(5, 7)); break;
            case "add_repeating_item": args["sectionAnchorId"] = anchor; break;
            default: throw new ArgumentOutOfRangeException(nameof(action));
        }
        return args;
    }

    private bool IsBound(string sid, string nativeId) => Control(sid, nativeId).GetProperty("isBound").GetBoolean();

    [Theory]
    [MemberData(nameof(BoundControlActions))]
    public void ContentControls_NestedBindingPolicy_DetachesTheTarget(string action, string nativeId)
    {
        var sid = Open(ControlsFixture(nativeId == "106" ? Array.Empty<string>() : new[] { nativeId }));
        Assert.True(IsBound(sid, nativeId));

        var refused = Call(sid, "docxodus_content_controls", ControlArgs(sid, action, nativeId));
        Assert.Equal("content_control_bound", ErrorCode(refused));

        var args = ControlArgs(sid, action, nativeId);
        args["options"] = action == "fill_rich_text"
            ? new { bindingPolicy = "detach_target", nestedControls = "preserve" }
            : new { bindingPolicy = "detach_target" };
        Succeeds(sid, "docxodus_content_controls", args);
        Assert.False(IsBound(sid, nativeId));
    }

    [Theory]
    [MemberData(nameof(BoundControlActions))]
    public void ContentControls_FlatBindingPolicy_StillDetachesTheTarget(string action, string nativeId)
    {
        var sid = Open(ControlsFixture(nativeId == "106" ? Array.Empty<string>() : new[] { nativeId }));
        var args = ControlArgs(sid, action, nativeId);
        args["bindingPolicy"] = "detach_target";
        if (action == "fill_rich_text") args["nestedControls"] = "preserve";
        Succeeds(sid, "docxodus_content_controls", args);
        Assert.False(IsBound(sid, nativeId));
    }

    [Fact]
    public void ContentControls_NestedChildFills_FillTheNestedControl()
    {
        var sid = Open(ControlsFixture());
        var inner = ControlAnchor(sid, "101");
        Succeeds(sid, "docxodus_content_controls", A(("action", "fill_text"), ("anchorId", ControlAnchor(sid, "100")),
            ("text", "Outer nested"), ("options", new { nestedControls = "preserve", childFills = new Dictionary<string, string> { [inner] = "Inner nested" } })));
        Assert.Equal("Inner nested", Control(sid, "101").GetProperty("text").GetString());
    }

    [Fact]
    public void ContentControls_ConflictingNestedAndFlatPolicy_IsRefusedWithoutMutating()
    {
        var sid = Open(ControlsFixture());
        var before = Session(sid).UndoCount;
        var ex = Assert.Throws<McpToolException>(() => Call(sid, "docxodus_content_controls", A(("action", "fill_text"),
            ("anchorId", ControlAnchor(sid, "106")), ("text", "x"), ("bindingPolicy", "detach_target"),
            ("options", new { bindingPolicy = "preserve" }))));
        Assert.Contains("options.bindingPolicy", ex.Message, StringComparison.Ordinal);
        Assert.Equal(before, Session(sid).UndoCount);
        Assert.True(IsBound(sid, "106"));
    }

    [Fact]
    public void ContentControls_AgreeingNestedAndFlatPolicy_IsAccepted()
    {
        var sid = Open(ControlsFixture());
        Succeeds(sid, "docxodus_content_controls", A(("action", "fill_text"), ("anchorId", ControlAnchor(sid, "106")),
            ("text", "x"), ("bindingPolicy", "detach_target"), ("options", new { bindingPolicy = "detach_target" })));
        Assert.False(IsBound(sid, "106"));
    }

    [Fact]
    public void ContentControls_NestedUnknownPolicy_IsATypedEditError()
    {
        var sid = Open(ControlsFixture());
        var result = Call(sid, "docxodus_content_controls", A(("action", "fill_text"), ("anchorId", ControlAnchor(sid, "106")),
            ("text", "x"), ("options", new { bindingPolicy = "sometimes" })));
        Assert.Equal("invalid_content_control_value", ErrorCode(result));
    }

    // ─── Batches and malformed objects ──────────────────────────────────

    [Fact]
    public void ABatchStepWithAConflictingPair_FailsBeforeAnyStepRuns()
    {
        var sid = Open();
        var p = Paragraph(sid, "one cat");
        var batch = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId = sid,
            steps = new object[]
            {
                new { tool = "docxodus_edit", args = new { action = "replace_text", anchorId = p, markdown = "changed" } },
                new { tool = "docxodus_edit", args = new { action = "replace_text_range", anchorId = p, find = "cat", replace = "dog", caseSensitive = true, options = new { ignoreCase = true } } },
            },
        })));
        Assert.False(batch.GetProperty("success").GetBoolean(), batch.ToString());
        Assert.Equal("invalid_batch_step", batch.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        Assert.Equal("one cat", Text(sid, p));
    }

    [Fact]
    public void ABatchStepWithNestedOptions_TakesEffect()
    {
        var sid = Open();
        var p = Paragraph(sid, "one cat two cat");
        var batch = J(Dispatcher.Call(_store, "docxodus_mutations", J(new
        {
            sessionId = sid,
            steps = new[]
            {
                new { tool = "docxodus_edit", args = (object)new { action = "replace_text_range", anchorId = p, find = "cat", replace = "dog", options = new { maxReplacements = 1 } } },
            },
        })));
        Assert.True(batch.GetProperty("success").GetBoolean(), batch.ToString());
        Assert.Equal("one dog two cat", Text(sid, p));
    }

    [Fact]
    public void ANonObjectOptionsArgument_IsRefused()
    {
        var sid = Open();
        var p = Paragraph(sid, "one cat");
        var ex = Assert.Throws<McpToolException>(() => Call(sid, "docxodus_edit", A(("action", "replace_text_range"),
            ("anchorId", p), ("find", "cat"), ("replace", "dog"), ("options", "maxReplacements=1"))));
        Assert.Contains("options", ex.Message, StringComparison.Ordinal);
        Assert.Equal("one cat", Text(sid, p));
    }

    // ─── The schema and the op descriptions ─────────────────────────────

    [Fact]
    public void TheSchemaAdvertisesEachObjectArgumentWithEveryAliasedField()
    {
        foreach (var ((tool, action), argument) in ObjectArguments.ByAction)
        {
            using var schema = JsonDocument.Parse(Assert.Single(ToolCatalog.Tools, t => t.Name == tool).InputSchemaJson);
            var properties = schema.RootElement.GetProperty("properties");
            Assert.True(properties.TryGetProperty(argument.Name, out var nested), $"{tool} does not advertise {argument.Name} for {action}");
            Assert.Equal("object", nested.GetProperty("type").GetString());
            var fields = nested.GetProperty("properties");
            foreach (var alias in argument.Aliases)
            {
                Assert.True(fields.TryGetProperty(alias.Field, out _), $"{tool} {argument.Name} does not advertise {alias.Field}");
                // A flat alias the schema still lists says it is deprecated in favour of the field.
                if (properties.TryGetProperty(alias.Property, out var flat))
                    Assert.Contains($"Deprecated: use {argument.Name}.{alias.Field}", flat.GetProperty("description").GetString(), StringComparison.Ordinal);
            }
        }
    }

    private static IEnumerable<string> DescriptionFiles =>
        Directory.GetFiles(Path.Combine(RepoRoot, "tools", "op-descriptions"), "*.json");

    private static IEnumerable<string> PropertyNames(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Object => element.EnumerateObject().SelectMany(p => PropertyNames(p.Value).Prepend(p.Name)),
        JsonValueKind.Array => element.EnumerateArray().SelectMany(PropertyNames),
        _ => Enumerable.Empty<string>(),
    };

    [Fact]
    public void NoDescriptionRecordsAFlattenedObjectArgument()
    {
        Assert.NotEmpty(DescriptionFiles);
        foreach (var file in DescriptionFiles)
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file));
            Assert.DoesNotContain("flattens", PropertyNames(doc.RootElement));
        }
    }

    [Fact]
    public void TheDescribedFlatAliasesAreTheOnesTheServerReads()
    {
        var described = new Dictionary<(string, string), (string Name, Dictionary<string, string> Fields)>();
        foreach (var file in DescriptionFiles)
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file));
            if (!doc.RootElement.TryGetProperty("ops", out var ops)) continue;
            foreach (var op in ops.EnumerateArray())
            {
                var mcp = op.GetProperty("transports").GetProperty("mcp");
                if (!mcp.TryGetProperty("flatAliases", out var aliases)) continue;
                var argument = Assert.Single(aliases.EnumerateObject());
                var record = (argument.Name, argument.Value.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!));
                var pairs = new List<(string, string)> { (mcp.GetProperty("tool").GetString()!, mcp.GetProperty("action").GetString()!) };
                if (mcp.TryGetProperty("aliases", out var more))
                    pairs.AddRange(more.EnumerateArray().Select(a => (a.GetProperty("tool").GetString()!, a.GetProperty("action").GetString()!)));
                foreach (var pair in pairs) described[pair] = record;
            }
        }

        Assert.Equal(ObjectArguments.ByAction.Keys.OrderBy(k => k), described.Keys.OrderBy(k => k));
        foreach (var (key, argument) in ObjectArguments.ByAction)
        {
            var (name, fields) = described[key];
            Assert.Equal(argument.Name, name);
            Assert.Equal(argument.Aliases.ToDictionary(a => a.Field, a => a.Property).OrderBy(kv => kv.Key), fields.OrderBy(kv => kv.Key));
        }
    }
}
