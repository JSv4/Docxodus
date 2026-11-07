using System.Xml.Linq;
using Docxodus.Tests.Ir;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionAlternateTextboxTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Grep_ReturnsOneHitForEachVisibleTextbox(bool alternate)
    {
        using var session = new DocxSession(Document(alternate).DocumentByteArray);
        Assert.Single(session.Grep("Alice Example"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReplaceMatch_UpdatesEveryRepresentation(bool alternate)
    {
        var input = Document(alternate);
        using var session = new DocxSession(input.DocumentByteArray);
        var match = session.Grep("Alice Example")[0];
        Assert.True(session.ReplaceMatch(match, "Bob Example").Success);
        using var stream = new MemoryStream(session.Save());
        using var package = WordprocessingDocument.Open(stream, false);
        var boxes = package.MainDocumentPart!.GetXDocument().Descendants(W + "txbxContent").ToArray();
        Assert.Equal(alternate ? 2 : 1, boxes.Length);
        Assert.All(boxes, box => Assert.Equal("Approved by: Bob Example", box.Value));
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
    }

    [Theory]
    [InlineData(AnchorIdRendering.FullUnid)]
    [InlineData(AnchorIdRendering.Abbreviated)]
    [InlineData(AnchorIdRendering.Sequential)]
    public void AnchorAliases_DoNotDuplicateSearchOrBreakEdits(AnchorIdRendering rendering)
    {
        using var session = new DocxSession(Document(true).DocumentByteArray,
            new DocxSessionSettings { ProjectionSettings = new WmlToMarkdownConverterSettings { AnchorIdRendering = rendering } });
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("Alice")), "Bob").Success);
        AssertBoxes(session.Save(), "Approved by: Bob Example");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RangeReplacement_FromEitherAnchor_IsOneUndo(bool hiddenAnchor)
    {
        using var session = new DocxSession(Document(true, paragraph: Paragraph("Alice and Alice")).DocumentByteArray);
        var visible = Assert.Single(session.Grep("Alice and Alice")).EnclosingAnchor.Anchor.Id;
        var anchor = hiddenAnchor
            ? session.Project().AnchorIndex.Values.Single(t => t.TextPreview == "Alice and Alice" && t.Anchor.Id != visible).Anchor.Id
            : visible;
        var results = session.ReplaceTextRange(anchor, "Alice", "Bob", new ReplaceOptions { ExpectedMatchCount = 2 });
        Assert.Equal(2, results.Count);
        Assert.All(results, r => Assert.True(r.Success, r.Error?.Message));
        Assert.Equal(1, session.Version);
        AssertBoxes(session.Save(), "Bob and Bob");
        Assert.True(session.Undo());
        AssertBoxes(session.Save(), "Alice and Alice");
        Assert.False(session.Undo());
        Assert.True(session.Redo());
        AssertBoxes(session.Save(), "Bob and Bob");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TrackedReplacement_ResolvesBothBranches(bool accept)
    {
        using var session = new DocxSession(Document(true).DocumentByteArray,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Reviewer" });
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("Alice")), "Bob").Success);
        var revisions = Xml(session.Save()).Descendants().Where(e => e.Name == W + "ins" || e.Name == W + "del").ToArray();
        Assert.Equal(4, revisions.Length);
        Assert.Equal(4, revisions.Select(e => (string?)e.Attribute(W + "id")).Distinct().Count());
        Assert.Single(revisions.Select(e => (string?)e.Attribute(W + "date")).Distinct());
        Assert.True((accept ? session.AcceptAllRevisions() : session.RejectAllRevisions()).Success);
        AssertBoxes(session.Save(), accept ? "Approved by: Bob Example" : "Approved by: Alice Example");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormattedReplacement_FormatsBothBranches(bool insertion)
    {
        using var session = new DocxSession(Document(true).DocumentByteArray);
        var match = Assert.Single(session.Grep("Alice"));
        var result = session.ReplaceTextAtSpanWithFormat(match.EnclosingAnchor.Anchor.Id,
            match.Span.Start, insertion ? 0 : match.Span.Length, "Bob", new FormatOp { Bold = true });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(1, session.Version);
        AssertBoxes(session.Save(), insertion ? "Approved by: BobAlice Example" : "Approved by: Bob Example");
        Assert.All(Xml(session.Save()).Descendants(W + "txbxContent"), box =>
            Assert.NotNull(box.Descendants(W + "r").Single(r => r.Element(W + "t")?.Value == "Bob")
                .Element(W + "rPr")?.Element(W + "b")));
        Assert.True(session.Undo());
        AssertBoxes(session.Save(), "Approved by: Alice Example");
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("<w:p><w:r><w:t>Different text</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t>Approved by: Alice Example</w:t></w:r></w:p><w:p/>")]
    [InlineData("<w:sdt><w:sdtContent><w:p><w:r><w:t>Approved by: Alice Example</w:t></w:r></w:p></w:sdtContent></w:sdt>")]
    public void DivergentFallback_RefusesWithoutChangingState(string fallback)
    {
        using var session = new DocxSession(Document(true, fallback: fallback).DocumentByteArray);
        var before = session.Save();
        var result = session.ReplaceMatch(Assert.Single(session.Grep("Alice")), "Bob");
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Equal(0, session.Version);
        Assert.Equal(before, session.Save());
        Assert.False(session.Undo());
    }

    [Fact]
    public void FallbackBoundaryFailure_RollsBackSelectedBranchAndPreservesRedo()
    {
        var fallback = "<w:p><w:fldSimple w:instr=\"AUTHOR\"><w:r><w:t>Approved by: Alice Example</w:t></w:r></w:fldSimple></w:p>";
        using var session = new DocxSession(Document(true, fallback: fallback).DocumentByteArray);
        var match = Assert.Single(session.Grep("Alice"));
        Assert.True(session.ReplaceMatch(match, "Bob").Success);
        Assert.True(session.Undo());
        var before = session.Save();
        var version = session.Version;
        var result = session.ReplaceTextAtSpan(match.EnclosingAnchor.Anchor.Id, match.Span.Start, 0, "X");
        Assert.False(result.Success);
        Assert.Equal(before, session.Save());
        Assert.Equal(version, session.Version);
        Assert.True(session.Redo());
        AssertBoxes(session.Save(), "Approved by: Bob Example");
    }

    [Fact]
    public void DifferentRunSplits_PreserveEachBranchFormatting()
    {
        var fallback = "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Approved by: Ali</w:t></w:r><w:r><w:t>ce Example</w:t></w:r></w:p>";
        using var session = new DocxSession(Document(true, fallback: fallback).DocumentByteArray);
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("Alice")), "Bob").Success);
        AssertBoxes(session.Save(), "Approved by: Bob Example");
        var boxes = Xml(session.Save()).Descendants(W + "txbxContent").ToArray();
        Assert.Empty(boxes[0].Descendants(W + "i"));
        Assert.Single(boxes[1].Descendants(W + "i"));
    }

    [Fact]
    public void UnsupportedChoice_SearchesFallbackAndSynchronizesBothBranches()
    {
        using var session = new DocxSession(Document(true, unsupportedChoice: true).DocumentByteArray);
        var match = Assert.Single(session.Grep("Alice"));
        Assert.True(session.ReplaceMatch(match, "Bob").Success);
        AssertBoxes(session.Save(), "Approved by: Bob Example");
    }

    [Fact]
    public void ReorderedTextboxes_RefuseWhenSurroundingTextDisprovesCorrespondence()
    {
        using var session = new DocxSession(Document(true, reordered: true).DocumentByteArray);
        var before = session.Save();
        var matches = session.Grep("Alice");
        Assert.Equal(2, matches.Count);
        var result = session.ReplaceMatch(matches[0], "Bob");
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Equal(before, session.Save());
        Assert.Equal(0, session.Version);
        Assert.False(session.Undo());
    }

    [Fact]
    public void NestedAlternatives_RefuseWithoutLeavingAnOuterFallbackStale()
    {
        using var session = new DocxSession(Document(true, nested: true).DocumentByteArray);
        var before = session.Save();
        var result = session.ReplaceMatch(Assert.Single(session.Grep("Alice")), "Bob");
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Equal(before, session.Save());
        Assert.Equal(0, session.Version);
        Assert.False(session.Undo());
    }

    [Fact]
    public void CrossBlockSearch_UsesOnlyVisibleBranch()
    {
        var paragraphs = Paragraph("Alice") + Paragraph("Example");
        using var session = new DocxSession(Document(true, paragraph: paragraphs).DocumentByteArray);
        Assert.Single(session.GrepCrossBlock("Alice\\nExample"));
    }

    private static string Paragraph(string text) => "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>";

    private static XDocument Xml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        return new XDocument(package.MainDocumentPart!.GetXDocument());
    }

    private static void AssertBoxes(byte[] bytes, string text) =>
        Assert.All(Xml(bytes).Descendants(W + "txbxContent"), box => Assert.Equal(text, box.Value));

    private static WmlDocument Document(bool alternate, string? paragraph = null, string? fallback = null, bool unsupportedChoice = false, bool nested = false, bool reordered = false)
    {
        paragraph ??= Paragraph("Approved by: Alice Example");
        var text = "<w:txbxContent>" + paragraph + "</w:txbxContent>";
        var fallbackText = "<w:txbxContent>" + (fallback ?? paragraph) + "</w:txbxContent>";
        var vml = "<w:pict><v:shape id=\"Box\" style=\"width:157pt;height:39pt\"><v:textbox>" + fallbackText + "</v:textbox></v:shape></w:pict>";
        var drawing = "<w:drawing><wp:inline><wp:extent cx=\"2000000\" cy=\"500000\"/><wp:docPr id=\"1\" name=\"Box\"/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp>" +
            "<wps:cNvSpPr txBox=\"1\"/><wps:spPr/><wps:txbx>" + text + "</wps:txbx><wps:bodyPr/>" +
            "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";
        var content = alternate
            ? "<mc:AlternateContent><mc:Choice Requires=\"wps\">" + drawing + "</mc:Choice><mc:Fallback>" + vml + "</mc:Fallback></mc:AlternateContent>"
            : vml;
        if (reordered)
        {
            var first = "<w:txbxContent>" + Paragraph("Section A") + paragraph + "</w:txbxContent>";
            var second = "<w:txbxContent>" + Paragraph("Section B") + paragraph + "</w:txbxContent>";
            content = "<mc:AlternateContent><mc:Choice Requires=\"wps\">" + drawing.Replace(text, first) +
                drawing.Replace(text, second).Replace("id=\"1\"", "id=\"2\"") + "</mc:Choice><mc:Fallback>" +
                vml.Replace(fallbackText, second) + vml.Replace(fallbackText, first).Replace("id=\"Box\"", "id=\"OtherBox\"") +
                "</mc:Fallback></mc:AlternateContent>";
        }
        if (nested)
            content = "<mc:AlternateContent><mc:Choice Requires=\"wps\">" + content +
                "</mc:Choice><mc:Fallback>" + vml + "</mc:Fallback></mc:AlternateContent>";
        if (unsupportedChoice)
            content = content.Replace("Requires=\"wps\"", "xmlns:unsupported=\"urn:unsupported\" Requires=\"unsupported\"");
        var document = IrTestDocuments.FromBodyXmlWithDrawingNamespaces("<w:p><w:r>" + content + "</w:r></w:p>");
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(package));
        return document;
    }
}
