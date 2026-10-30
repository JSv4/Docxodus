using System.Text.RegularExpressions;
using System.Xml.Linq;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class MarkdownShapeProjectionTests
{
    private static string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
    private static string Sdt(string content) => $"<w:sdt><w:sdtPr/><w:sdtContent>{content}</w:sdtContent></w:sdt>";
    private static string Shape(string content) => "<v:shape style=\"width:200pt;height:100pt\">" +
        $"<v:textbox><w:txbxContent>{content}</w:txbxContent></v:textbox></v:shape>";
    private static string Box(string content) => "<w:pict>" + Shape(content) + "</w:pict>";
    private static string Drawing(string content) =>
        "<w:drawing><wp:inline><wp:extent cx=\"2540000\" cy=\"1270000\"/><wp:docPr id=\"1\" name=\"Textbox\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr/><wps:txbx><w:txbxContent>" + content +
        "</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";
    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXmlWithDrawingNamespaces(body);
    private static string Host(string carrier) => "<w:p><w:r><w:t>Host</w:t>" + carrier + "</w:r></w:p>";

    private static MarkdownProjection Project(WmlDocument doc, WmlToMarkdownConverterSettings? settings = null)
    {
        settings ??= new();
        var result = WmlToMarkdownConverter.Convert(new WmlDocument(doc), settings);
        foreach (bool retain in new[] { true, false })
            Assert.Equal(result.Markdown, IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc),
                new IrReaderOptions { RetainSources = retain }), settings).Markdown);
        return result;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextboxesProjectAnchoredBlocksAfterTheirHost(bool suppressEmpty)
    {
        var doc = Doc(Host(Box(Sdt(P("Inside")) + P("Second"))) + P("Following"));
        var result = Project(doc, new() { EmptyParagraphs = suppressEmpty ? EmptyParagraphMode.Suppress : EmptyParagraphMode.AnchorOnly });
        Assert.Matches(@"Host\s+\{#p:body:[^}]+\} Inside\s+\{#p:body:[^}]+\} Second\s+\{#p:body:[^}]+\} Following", result.Markdown);
        using var session = new DocxSession(doc.DocumentByteArray);
        var anchor = result.AnchorIndex.Values.Single(a => a.Anchor.Kind == "p" && a.TextPreview == "Inside").Anchor.Id;
        Assert.Contains("Inside", session.ProjectAnchor(anchor, ProjectionDepth.SelfOnly).Markdown);
        Assert.DoesNotContain("Second", session.ProjectAnchor(anchor, ProjectionDepth.SelfOnly).Markdown);
    }

    [Theory]
    [InlineData("wps", "Choice")]
    [InlineData("unknown", "Fallback")]
    [InlineData("wps unknown", "Fallback")]
    public void AlternateContentProjectsOnlyTheSupportedBranch(string requires, string expected)
    {
        var alternate = $"<mc:AlternateContent xmlns:unknown=\"urn:unknown\"><mc:Choice Requires=\"{requires}\">" +
            Drawing(P("Choice")) + "</mc:Choice><mc:Fallback>" + Box(P("Fallback")) + "</mc:Fallback></mc:AlternateContent>";
        var result = Project(Doc(Host(alternate)));
        Assert.Contains(expected, result.Markdown);
        Assert.DoesNotContain(expected == "Choice" ? "Fallback" : "Choice", result.Markdown);
    }

    [Fact]
    public void NestedAndGroupedTextboxesRenderOnceWithFormatting()
    {
        var nested = P("Outer") + Host(Box(P("Nested")));
        var group = "<w:pict><v:group>" + Shape(Sdt(nested)) +
            Shape("<w:p><w:hyperlink w:anchor=\"target\"><w:r><w:t>Linked</w:t></w:r></w:hyperlink></w:p>") +
            "</v:group></w:pict>";
        var markdown = Project(Doc(Host(group))).Markdown;
        foreach (var text in new[] { "Outer", "Nested", "[Linked](#target)" })
            Assert.Equal(1, Regex.Matches(markdown, Regex.Escape(text)).Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedRowsAndCellsKeepTheirTextAndTableGeometry(bool opaque)
    {
        var cell = "<w:tc><w:tcPr/>" + P("Cell") + "</w:tc>";
        var table = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>" +
            "<w:tr>" + cell + Sdt(cell.Replace("Cell", "Wrapped")) + "</w:tr>" +
            Sdt("<w:tr>" + Sdt(cell.Replace("Cell", "Last")) + cell + "</w:tr>") + "</w:tbl>";
        var markdown = Project(Doc(table), new() { TableMode = opaque ? TableRenderMode.AlwaysOpaque : TableRenderMode.GfmWithOpaqueFallback }).Markdown;
        if (opaque) Assert.Contains("rows: 2\ncols: 2", markdown.Replace("\r", ""));
        else Assert.Contains("| Cell | Wrapped |\n| --- | --- |\n| Last | Cell |", markdown.Replace("\r", ""));
    }

    [Fact]
    public void TextboxOnlyHostsAndCellsDoNotDuplicateAlternativeText()
    {
        var alternate = "<mc:AlternateContent><mc:Choice Requires=\"wps\">" + Drawing(P("Unique")) +
            "</mc:Choice><mc:Fallback>" + Box(P("Unique")) + "</mc:Fallback></mc:AlternateContent>";
        foreach (var inCell in new[] { false, true })
        {
            var body = "<w:p><w:r>" + alternate + "</w:r></w:p>";
            if (inCell) body = "<w:tbl><w:tr><w:tc>" + body + "</w:tc></w:tr></w:tbl>";
            var markdown = Project(Doc(body), new() { EmptyParagraphs = EmptyParagraphMode.Suppress }).Markdown;
            Assert.Equal(1, Regex.Matches(markdown, "Unique").Count);
            Assert.Matches(@"\{#p:body:[^}]+\} Unique", markdown);
        }
    }

    [Fact]
    public void TableInsideTextboxRetainsItsOwnCellText()
    {
        var table = "<w:tbl><w:tr><w:tc>" + P("TablePayload") + "</w:tc></w:tr></w:tbl>";
        Assert.Contains("| TablePayload |", Project(Doc(Host(Box(table + P("Tail"))))).Markdown);
    }

    [Theory]
    [InlineData(64, true)]
    [InlineData(65, false)]
    public void TextboxDiscoveryHonorsCellWrapperDepth(int depth, bool visible)
    {
        var content = Host(Box(P("Deep")));
        for (int i = 0; i < depth; i++) content = "<w:customXml>" + content + "</w:customXml>";
        var doc = Doc("<w:tbl><w:tr><w:tc>" + content + "</w:tc></w:tr></w:tbl>");
        Assert.Equal(visible, Project(doc).Markdown.Contains("Deep"));
    }

    [Fact]
    public void TextboxesProjectInEveryStory()
    {
        using var memory = new OpenXmlMemoryStreamDocument(IrTestDocuments.WithMixedReviewParts(
            "Header", "Footer", "Footnote", "Endnote", "Comment"));
        using (var package = memory.GetWordprocessingDocument())
        {
            foreach (var part in package.MainDocumentPart!.Parts.Select(p => p.OpenXmlPart))
            {
                var xml = part.GetXDocument();
                foreach (var p in xml.Descendants(W.p).Where(p => p.Descendants(W.t).Any()).ToList())
                    p.ReplaceWith(XElement.Parse($"<w:p xmlns:w=\"{W.w}\" xmlns:v=\"urn:schemas-microsoft-com:vml\"><w:r>" +
                        Box(p.ToString()) + "</w:r></w:p>"));
                part.PutXDocument();
            }
        }
        var markdown = Project(memory.GetModifiedWmlDocument()).Markdown;
        foreach (var scope in new[] { "hdr1", "ftr1", "fn", "en", "cmt" })
            Assert.Matches($@"\{{#p:{scope}:[^}}]+\}} ", markdown);
    }

    [Theory]
    [InlineData(TrackedChangeMode.Accept, false)]
    [InlineData(TrackedChangeMode.StripDeletions, false)]
    [InlineData(TrackedChangeMode.RenderInline, true)]
    public void DeletedTextboxFollowsTheRequestedRevisionView(TrackedChangeMode mode, bool visible)
    {
        var doc = Doc("<w:p><w:del w:id=\"1\" w:author=\"Reviewer\"><w:r>" + Box(P("Removed")) + "</w:r></w:del></w:p>");
        var markdown = WmlToMarkdownConverter.Convert(doc, new() { TrackedChanges = mode }).Markdown;
        Assert.Equal(visible, markdown.Contains("Removed"));
        if (visible) Assert.Contains("{-Removed-}", markdown);
    }

    [Fact]
    public void DeletionWinsOverInsertionInsideANestedTextbox()
    {
        var nested = "<w:p><w:ins w:id=\"2\" w:author=\"Reviewer\"><w:r>" + Box(P("Nested")) + "</w:r></w:ins></w:p>";
        var doc = Doc("<w:p><w:del w:id=\"1\" w:author=\"Reviewer\"><w:r>" + Box(nested) + "</w:r></w:del></w:p>");
        var markdown = WmlToMarkdownConverter.Convert(doc, new() { TrackedChanges = TrackedChangeMode.RenderInline }).Markdown;
        Assert.Contains("{-Nested-}", markdown);
        Assert.DoesNotContain("{+Nested+}", markdown);
    }

    [Theory]
    [InlineData("edit")]
    [InlineData("delete")]
    [InlineData("following")]
    public void TextboxEditsAndNeighbourChangesReturnCompletePatches(string operation)
    {
        using var session = new DocxSession(Doc(Host(Box(P("Inside"))) + P("Following")).DocumentByteArray);
        var projection = session.Project();
        string Id(string text) => projection.AnchorIndex.Values.Single(a => a.TextPreview == text).Anchor.Id;
        var result = operation switch
        {
            "edit" => session.ReplaceText(Id("Inside"), "Updated"),
            "delete" => session.DeleteBlock(Id("HostInside")),
            _ => session.ReplaceText(Id("Following"), "Later"),
        };
        Assert.True(result.Success, result.Error?.Message);
        var patch = Assert.IsType<MarkdownPatch>(result.Patch);
        Assert.True(patch.IsFullDocument);
        Assert.Equal(session.Project().Markdown, patch.Markdown);
        if (operation == "edit") Assert.Contains("Updated", patch.Markdown);
        if (operation == "delete") Assert.DoesNotContain("Inside", patch.Markdown);
    }
}
