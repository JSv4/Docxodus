// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1024: an argument a caller may leave out takes the default <see cref="DocxSessionOps"/>
/// declares, on every transport. Before, the stdio host, MCP, npm and Python each supplied their
/// own (an insert's position, a list format, a query's flags), so two transports could disagree,
/// and a transport that supplied none refused the call. Each test here calls the stdio host or the
/// MCP server for real with the argument left out, and checks the effect is exactly what the
/// facade does when handed its own constant: the same document, or the same query answer.
/// </summary>
public sealed class FacadeDefaultTests : IDisposable
{
    private readonly string _root;
    private readonly SessionStore _store;

    public FacadeDefaultTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"facade-defaults-{Guid.NewGuid():N}");
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

    // ─── Fixture ────────────────────────────────────────────────────────

    /// <summary>The anchors one fixture session hands its op: a heading (whose section a projection
    /// covers only at the default depth), a second
    /// paragraph, a bulleted list item, the first cell of a 2×2 table (with text, shaded red, so a
    /// row or column inserted before it differs from one inserted after), and an annotation id.</summary>
    private sealed record Fixture(string First, string Second, string ListItem, string Cell, string Annotation);

    private static Fixture Seed(int handle)
    {
        var session = SessionRegistry.Get(handle);
        var first = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        first = session.ReplaceText(first, "# Alpha [NAME] and alpha").Modified.Select(a => a.Id).FirstOrDefault() ?? first;
        var second = session.InsertParagraph(first, Position.After, "Beta").Created.First().Id;
        var item = session.InsertParagraph(second, Position.After, "Gamma").Created.First().Id;
        item = session.ApplyListFormat(item, ListFormat.Bullet).Modified.First().Id;
        var cell = session.InsertTable(item, Position.After, 2, 2).Created.First().Id;
        Assert.True(session.ReplaceCellContent(cell, "Corner").Success);
        Assert.True(session.SetCellShading(cell, "FF0000").Success);
        Assert.True(J(DocxSessionOps.AddAnnotation(handle, second, null,
            """{"id":"a1","labelId":"L","label":"Label"}""")).GetProperty("success").GetBoolean());
        return new Fixture(first, second, item, cell, "a1");
    }

    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    /// <summary>The saved main document, with the per-session identity (anchor unids, rsids,
    /// paragraph ids) removed, so two sessions that made the same edit compare equal.</summary>
    private static string Fingerprint(int handle)
    {
        using var ms = new MemoryStream(SessionRegistry.Get(handle).Save());
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        var root = new XElement(doc.MainDocumentPart!.GetXDocument().Root!);
        foreach (var e in root.DescendantsAndSelf())
        {
            e.Attributes().Where(a => a.IsNamespaceDeclaration
                    || a.Name.Namespace == PtOpenXml.pt
                    || a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)
                    || a.Name == W14 + "paraId" || a.Name == W14 + "textId")
                .Remove();
        }
        return root.ToString(SaveOptions.DisableFormatting);
    }

    private (int Handle, string SessionId) OpenMcp()
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path })))
            .GetProperty("sessionId").GetString()!;
        return (_store.Get(sessionId).Handle, sessionId);
    }

    private static string Stdio(int handle, string op, Dictionary<string, object?> args)
    {
        args["handle"] = handle;
        return Docxodus.PyHost.Dispatcher.Dispatch(op, J(args));
    }

    // ─── Mutations: the omitted argument edits the document as the facade's default does ───

    /// <summary>A mutation whose argument is left out: the stdio op and the MCP tool/action that
    /// expose it (null where the transport does not), the arguments without the defaulted one, and
    /// the facade call that passes the facade's own default explicitly.</summary>
    private sealed record Mutation(
        string? StdioOp, string? McpTool, string? McpAction,
        Func<Fixture, Dictionary<string, object?>> Args,
        Func<int, Fixture, string> Facade);

    private static readonly Dictionary<string, Mutation> Mutations = new()
    {
        ["insertParagraph.position"] = new("insert_paragraph", "docxodus_edit", "insert_paragraph",
            f => new() { ["anchorId"] = f.First, ["markdown"] = "Inserted" },
            (h, f) => DocxSessionOps.InsertParagraph(h, f.First, DocxSessionOps.DefaultInsertPosition, "Inserted")),
        ["insertParagraph.position (create alias)"] = new(null, "docxodus_create", "insert_paragraph",
            f => new() { ["anchorId"] = f.First, ["markdown"] = "Inserted" },
            (h, f) => DocxSessionOps.InsertParagraph(h, f.First, DocxSessionOps.DefaultInsertPosition, "Inserted")),
        ["moveBlock.position"] = new("move_block", "docxodus_edit", "move_block",
            f => new() { ["sourceAnchorId"] = f.First, ["targetAnchorId"] = f.Second },
            (h, f) => DocxSessionOps.MoveBlock(h, f.First, f.Second, DocxSessionOps.DefaultInsertPosition)),
        ["insertHorizontalRule.position"] = new(null, "docxodus_create", "insert_horizontal_rule",
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.InsertHorizontalRule(h, f.First, DocxSessionOps.DefaultInsertPosition, "")),
        ["insertTable.position"] = new("insert_table", "docxodus_table", "insert",
            f => new() { ["anchorId"] = f.First, ["rows"] = 1, ["columns"] = 1 },
            (h, f) => DocxSessionOps.InsertTable(h, f.First, DocxSessionOps.DefaultInsertPosition, 1, 1, "")),
        ["insertTableRow.position"] = new("insert_table_row", "docxodus_table", "insert_row",
            f => new() { ["cellAnchorId"] = f.Cell },
            (h, f) => DocxSessionOps.InsertTableRow(h, f.Cell, DocxSessionOps.DefaultInsertPosition)),
        ["insertTableColumn.position"] = new("insert_table_column", "docxodus_table", "insert_column",
            f => new() { ["cellAnchorId"] = f.Cell },
            (h, f) => DocxSessionOps.InsertTableColumn(h, f.Cell, DocxSessionOps.DefaultInsertPosition)),
        ["mergeCells.rowSpan"] = new("merge_cells", "docxodus_table", "merge_cells",
            f => new() { ["cellAnchorId"] = f.Cell, ["colSpan"] = 2 },
            (h, f) => DocxSessionOps.MergeCells(h, f.Cell, DocxSessionOps.DefaultMergeSpan, 2, null)),
        ["mergeCells.colSpan"] = new("merge_cells", "docxodus_table", "merge_cells",
            f => new() { ["cellAnchorId"] = f.Cell, ["rowSpan"] = 2 },
            (h, f) => DocxSessionOps.MergeCells(h, f.Cell, 2, DocxSessionOps.DefaultMergeSpan, null)),
        ["setCellShading.fill"] = new("set_cell_shading", "docxodus_table", "set_shading",
            f => new() { ["cellAnchorId"] = f.Cell },
            (h, f) => DocxSessionOps.SetCellShading(h, f.Cell, DocxSessionOps.DefaultCellFill, null)),
        ["setRepeatHeaderRow.repeat"] = new("set_repeat_header_row", "docxodus_table", "set_repeat_header_row",
            f => new() { ["cellAnchorId"] = f.Cell },
            (h, f) => DocxSessionOps.SetRepeatHeaderRow(h, f.Cell, DocxSessionOps.DefaultRepeatHeaderRow)),
        ["applyListFormat.listFormat"] = new("apply_list_format", "docxodus_list", "apply_format",
            f => new() { ["anchorId"] = f.ListItem },
            (h, f) => DocxSessionOps.ApplyListFormat(h, f.ListItem, DocxSessionOps.DefaultListFormat)),
        ["applyListFormat.listFormat (format alias)"] = new(null, "docxodus_format", "apply_list_format",
            f => new() { ["anchorId"] = f.ListItem },
            (h, f) => DocxSessionOps.ApplyListFormat(h, f.ListItem, DocxSessionOps.DefaultListFormat)),
        ["applyListFormatRange.listFormat"] = new("apply_list_format_range", "docxodus_list", "apply_format_range",
            f => new() { ["firstAnchorId"] = f.ListItem, ["lastAnchorId"] = f.ListItem },
            (h, f) => DocxSessionOps.ApplyListFormatRange(h, f.ListItem, f.ListItem, DocxSessionOps.DefaultListFormat)),
        ["applyFormat.format"] = new("apply_format", "docxodus_format", "apply_format",
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.ApplyFormat(h, f.First, null, null)),
        ["applyFormatBySubstring.format"] = new("apply_format_by_substring", "docxodus_format", "apply_format_by_substring",
            f => new() { ["anchorId"] = f.First, ["substring"] = "Alpha" },
            (h, f) => DocxSessionOps.ApplyFormatBySubstring(h, f.First, "Alpha", null)),
        ["setParagraphFormat.paragraphFormat"] = new("set_paragraph_format", "docxodus_format", "set_paragraph_format",
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.SetParagraphFormat(h, f.First, null)),
        ["replaceTextAtSpanWithFormat.format"] = new("replace_text_at_span_with_format", "docxodus_edit", "replace_text_at_span_with_format",
            f => new() { ["anchorId"] = f.First, ["spanStart"] = 0, ["spanLength"] = 5, ["replace"] = "Omega" },
            (h, f) => DocxSessionOps.ReplaceTextAtSpanWithFormat(h, f.First, 0, 5, "Omega", null)),
        ["insertPageNumberField.field"] = new("insert_page_number_field", "docxodus_create", "insert_page_number_field",
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.InsertPageNumberField(h, f.First, DocxSessionOps.DefaultPageNumberField)),
        ["setPageNumbering.op"] = new("set_page_numbering", null, null,
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.SetPageNumbering(h, f.First, (PageNumberingOp?)null)),
        ["setPageSetup.op"] = new("set_page_setup", null, null,
            f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.SetPageSetup(h, f.First, (PageSetupOp?)null)),
        ["updateAnnotation.update"] = new("update_annotation", "docxodus_annotate", "update",
            f => new() { ["annotationId"] = f.Annotation },
            (h, f) => DocxSessionOps.UpdateAnnotation(h, f.Annotation, null)),
    };

    public static IEnumerable<object[]> StdioMutations =>
        Mutations.Where(m => m.Value.StdioOp is not null).Select(m => new object[] { m.Key });

    public static IEnumerable<object[]> McpMutations =>
        Mutations.Where(m => m.Value.McpTool is not null).Select(m => new object[] { m.Key });

    /// <summary>The op's result with the per-session anchors and the volatile fields blanked, so
    /// two sessions' answers to the same edit compare equal.</summary>
    private static string Outcome(string resultJson) =>
        Regex.Replace(resultJson, "[0-9a-f]{32}", "#");

    [Theory]
    [MemberData(nameof(StdioMutations))]
    public void StdioHost_OmittedArgument_EditsAsTheFacadeDefaultDoes(string name)
    {
        var m = Mutations[name];
        var omitted = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        var explicitDefault = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var viaHost = Stdio(omitted, m.StdioOp!, m.Args(Seed(omitted)));
            var viaFacade = m.Facade(explicitDefault, Seed(explicitDefault));
            Assert.Equal(Outcome(viaFacade), Outcome(viaHost));
            Assert.Equal(Fingerprint(explicitDefault), Fingerprint(omitted));
        }
        finally
        {
            DocxSessionOps.CloseSession(omitted);
            DocxSessionOps.CloseSession(explicitDefault);
        }
    }

    [Theory]
    [MemberData(nameof(McpMutations))]
    public void McpServer_OmittedArgument_EditsAsTheFacadeDefaultDoes(string name)
    {
        var m = Mutations[name];
        var (omitted, sessionId) = OpenMcp();
        var (explicitDefault, _) = OpenMcp();
        var args = m.Args(Seed(omitted));
        args["sessionId"] = sessionId;
        args["action"] = m.McpAction;
        var viaMcp = Dispatcher.Call(_store, m.McpTool!, J(args));
        var viaFacade = m.Facade(explicitDefault, Seed(explicitDefault));
        Assert.Equal(J(viaFacade).GetProperty("success").GetBoolean(), J(viaMcp).GetProperty("success").GetBoolean());
        Assert.Equal(Fingerprint(explicitDefault), Fingerprint(omitted));
    }

    // ─── Queries: the omitted argument answers as the facade's default does ───

    private static readonly Dictionary<string, (string Op, Func<Fixture, Dictionary<string, object?>> Args, Func<int, Fixture, string> Facade)> Queries = new()
    {
        ["getDiff.format"] = ("get_diff", _ => new(),
            (h, _) => DocxSessionOps.GetDiff(h, DocxSessionOps.DefaultDiffFormat)),
        ["projectAnchor.depth"] = ("project_anchor", f => new() { ["anchorId"] = f.First },
            (h, f) => DocxSessionOps.ProjectAnchor(h, f.First, DocxSessionOps.DefaultProjectionDepth)),
        ["findByRegex.regexOptions"] = ("find_by_regex", _ => new() { ["pattern"] = "beta" },
            (h, _) => DocxSessionOps.FindByRegex(h, "beta", DocxSessionOps.DefaultRegexOptions, null)),
        ["findPlaceholders.kinds,scope,boundary"] = ("find_placeholders", _ => new(),
            (h, _) => DocxSessionOps.FindPlaceholders(h, DocxSessionOps.DefaultPlaceholderKinds,
                DocxSessionOps.DefaultPlaceholderScope, null, DocxSessionOps.DefaultContextBoundary)),
        ["remainingPlaceholders.kinds"] = ("remaining_placeholders", _ => new(),
            (h, _) => DocxSessionOps.RemainingPlaceholders(h, DocxSessionOps.DefaultPlaceholderKinds)),
    };

    public static IEnumerable<object[]> QueryNames => Queries.Keys.Select(k => new object[] { k });

    [Theory]
    [MemberData(nameof(QueryNames))]
    public void StdioHost_OmittedArgument_AnswersAsTheFacadeDefaultDoes(string name)
    {
        var q = Queries[name];
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var fixture = Seed(handle);
            Assert.Equal(q.Facade(handle, fixture), Stdio(handle, q.Op, q.Args(fixture)));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    /// <summary>A defaulted argument sent with the wrong type is refused, not read as absent: a
    /// mistyped span or position must not quietly become the default edit.</summary>
    [Theory]
    [InlineData("insert_paragraph", "position", 1)]
    [InlineData("insert_page_number_field", "field", 2)]
    [InlineData("merge_cells", "rowSpan", "2")]
    [InlineData("merge_cells", "colSpan", "2")]
    public void StdioHost_RefusesADefaultedArgumentOfTheWrongType(string op, string argument, object value)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var f = Seed(handle);
            var args = op switch
            {
                "insert_paragraph" => new Dictionary<string, object?> { ["anchorId"] = f.First, ["markdown"] = "x" },
                "insert_page_number_field" => new Dictionary<string, object?> { ["anchorId"] = f.First },
                _ => new Dictionary<string, object?> { ["cellAnchorId"] = f.Cell, ["rowSpan"] = 2, ["colSpan"] = 2 },
            };
            args[argument] = value;
            var before = Fingerprint(handle);
            var refusal = Assert.ThrowsAny<FormatException>(() => Stdio(handle, op, args));
            Assert.Contains(argument, refusal.Message);
            Assert.Equal(before, Fingerprint(handle));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // ─── Arguments with no safe default are required on every transport ───

    private static string AltText(int handle) =>
        J(DocxSessionOps.ListImages(handle))[0].TryGetProperty("altText", out var alt) ? alt.GetString() ?? "" : "";

    private static string SeedImage(int handle)
    {
        var fixture = Seed(handle);
        var inserted = J(DocxSessionOps.InsertImage(handle, fixture.First, 0,
            Convert.ToBase64String(Ir.IrTestDocuments.TinyPng), "{\"altText\":\"A chart\",\"title\":\"Chart\"}"));
        Assert.True(inserted.GetProperty("success").GetBoolean(), inserted.ToString());
        return inserted.GetProperty("imageId").GetString()!;
    }

    /// <summary>A null alt text or title removes it, so leaving one out must not: the stdio host
    /// used to read an omitted value as null and silently clear the picture's alt text, where MCP
    /// refused the call. Both now require each, as a string or as an explicit null.</summary>
    [Theory]
    [InlineData("altText")]
    [InlineData("title")]
    public void StdioHost_SetImageMetadata_RequiresEachValue_AndLeavesTheImageAloneWithout(string omitted)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var imageId = SeedImage(handle);
            var args = new Dictionary<string, object?> { ["imageId"] = imageId, ["altText"] = "New", ["title"] = "New" };
            args.Remove(omitted);
            Assert.ThrowsAny<FormatException>(() => Stdio(handle, "set_image_metadata", args));
            Assert.Equal("A chart", AltText(handle));

            // An explicit null still clears it.
            var cleared = J(Stdio(handle, "set_image_metadata",
                new Dictionary<string, object?> { ["imageId"] = imageId, ["altText"] = null, ["title"] = "Chart" }));
            Assert.True(cleared.GetProperty("success").GetBoolean(), cleared.ToString());
            Assert.Equal("", AltText(handle));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    /// <summary>An empty dimensions object or repair list is refused by the session ("at least one
    /// rendered dimension is required", "no repairs requested"), so neither is a default: the stdio
    /// host refuses the call that leaves it out, as MCP does, instead of filling it in.</summary>
    [Theory]
    [InlineData("set_image_dimensions", "dimensions")]
    [InlineData("repair_revisions", "repairs")]
    public void StdioHost_RefusesAnOmittedArgumentWhoseEmptyValueTheSessionRefuses(string op, string argument)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            var args = op == "set_image_dimensions"
                ? new Dictionary<string, object?> { ["imageId"] = SeedImage(handle) }
                : new Dictionary<string, object?>();
            var refusal = Assert.ThrowsAny<FormatException>(() => Stdio(handle, op, args));
            Assert.Contains(argument, refusal.Message);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // ─── The descriptions record each default once, on the argument ───

    private static readonly string DescriptionDirectory =
        Path.Combine("../../../..", "tools", "op-descriptions");

    private static IEnumerable<(string File, JsonElement Op)> DescribedOps()
    {
        foreach (var file in Directory.GetFiles(DescriptionDirectory, "*.json").Order(StringComparer.Ordinal))
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file));
            if (!doc.RootElement.TryGetProperty("ops", out var ops)) continue;
            foreach (var op in ops.EnumerateArray()) yield return (Path.GetFileName(file), op.Clone());
        }
    }

    [Fact]
    public void NoDescription_RecordsATransportDefault_OrATransportRequirement()
    {
        var divergent = DescribedOps()
            .SelectMany(o => o.Op.GetProperty("transports").EnumerateObject()
                .Where(t => t.Value.TryGetProperty("defaults", out _) || t.Value.TryGetProperty("requires", out _))
                .Select(t => $"{o.File}: {o.Op.GetProperty("op").GetString()} ({t.Name})"))
            .ToList();
        Assert.Empty(divergent);
    }

    /// <summary>The facade constant each moved default now comes from, as the description should
    /// record it on the argument (the wire value: a token for a named enum, the number for a flag
    /// set or a numeric enum).</summary>
    public static IEnumerable<object[]> FacadeDefaults => new (string Op, string Arg, object Default)[]
    {
        ("insertParagraph", "position", "after"),
        ("moveBlock", "position", "after"),
        ("insertHorizontalRule", "position", "after"),
        ("insertTable", "position", "after"),
        ("insertTableRow", "position", "after"),
        ("insertTableColumn", "position", "after"),
        ("mergeCells", "rowSpan", DocxSessionOps.DefaultMergeSpan),
        ("mergeCells", "colSpan", DocxSessionOps.DefaultMergeSpan),
        ("setCellShading", "fill", DocxSessionOps.DefaultCellFill),
        ("setRepeatHeaderRow", "repeat", DocxSessionOps.DefaultRepeatHeaderRow),
        ("applyListFormat", "listFormat", "none"),
        ("applyListFormatRange", "listFormat", "none"),
        ("insertPageNumberField", "field", "currentPage"),
        ("getDiff", "format", (int)DocxSessionOps.DefaultDiffFormat),
        ("projectAnchor", "depth", (int)DocxSessionOps.DefaultProjectionDepth),
        ("findByRegex", "regexOptions", (int)DocxSessionOps.DefaultRegexOptions),
        ("findPlaceholders", "kinds", (int)DocxSessionOps.DefaultPlaceholderKinds),
        ("findPlaceholders", "scope", (int)DocxSessionOps.DefaultPlaceholderScope),
        ("findPlaceholders", "boundary", (int)DocxSessionOps.DefaultContextBoundary),
        ("remainingPlaceholders", "kinds", (int)DocxSessionOps.DefaultPlaceholderKinds),
        ("renderBlockHtml", "cssPrefix", DocxSessionOps.DefaultBlockCssPrefix),
        ("renderBlockHtml", "fabricateClasses", DocxSessionOps.DefaultFabricateClasses),
    }.Select(d => new object[] { d.Op, d.Arg, d.Default });

    [Theory]
    [MemberData(nameof(FacadeDefaults))]
    public void TheDescription_RecordsTheFacadeConstant_AsTheArgumentsDefault(string op, string arg, object expected)
    {
        var described = DescribedOps().Single(o => o.Op.GetProperty("op").GetString() == op).Op
            .GetProperty("args").EnumerateArray().Single(a => a.GetProperty("name").GetString() == arg);
        Assert.False(described.GetProperty("required").GetBoolean(), $"{op}.{arg} is described as required");
        Assert.True(described.TryGetProperty("default", out var recorded), $"{op}.{arg} records no default");
        Assert.Equal(JsonSerializer.Serialize(expected), recorded.GetRawText());
    }

    /// <summary>The token defaults above are the facade's enum constants, spelled on the wire.</summary>
    [Fact]
    public void TheTokenDefaults_AreTheFacadeConstants()
    {
        Assert.Equal(Position.After, DocxSessionOps.DefaultInsertPosition);
        Assert.Equal(Position.After, DocxSessionJson.ParsePos(null));
        Assert.Equal(DocxSessionOps.DefaultListFormat, DocxSessionJson.ParseListFormat("none"));
        Assert.Equal(DocxSessionOps.DefaultPageNumberField, DocxSessionJson.ParsePageNumberField("currentPage"));
    }

    /// <summary>The object arguments whose empty value is a successful no-op default to it, which
    /// the description records as <c>{}</c>.</summary>
    [Theory]
    [InlineData("applyFormat", "format")]
    [InlineData("applyFormatBySubstring", "format")]
    [InlineData("replaceTextAtSpanWithFormat", "format")]
    [InlineData("setParagraphFormat", "paragraphFormat")]
    [InlineData("setPageNumbering", "op")]
    [InlineData("setPageSetup", "op")]
    [InlineData("updateAnnotation", "update")]
    public void TheDescription_RecordsAnEmptyObjectDefault(string op, string arg)
    {
        var described = DescribedOps().Single(o => o.Op.GetProperty("op").GetString() == op).Op
            .GetProperty("args").EnumerateArray().Single(a => a.GetProperty("name").GetString() == arg);
        Assert.False(described.GetProperty("required").GetBoolean());
        Assert.Equal("{}", described.GetProperty("default").GetRawText());
    }
}
