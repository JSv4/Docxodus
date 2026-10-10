using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxDiffMisplacedBreakTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Theory]
    [InlineData("", false, false)]
    [InlineData("page", false, false)]
    [InlineData("column", false, false)]
    [InlineData("", true, false)]
    [InlineData("page", true, false)]
    [InlineData("column", true, false)]
    [InlineData("", false, true)]
    [InlineData("page", false, true)]
    [InlineData("column", false, true)]
    [InlineData("", true, true)]
    [InlineData("page", true, true)]
    [InlineData("column", true, true)]
    public void WholeParagraphBreaks_AreRunChildrenAndRoundTrip(string type, bool deletion, bool nested)
    {
        var empty = IrTestDocuments.Create("");
        var populated = IrTestDocuments.FromBodyXml(Paragraph(type, nested));
        var originalBytes = populated.DocumentByteArray.ToArray();
        var result = DocxCompare.Compare(deletion ? populated : empty, deletion ? empty : populated);

        var xml = Read(result);
        var breaks = xml.Descendants(W + "br").ToArray();
        Assert.Equal(2, breaks.Length);
        Assert.All(breaks, b =>
        {
            Assert.Equal(W + "r", b.Parent!.Name);
            Assert.Contains(b.Ancestors(), a => a.Name == W + (deletion ? "del" : "ins"));
        });
        AssertClean(result);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        AssertView(deletion ? rejected : accepted, type);
        Assert.Empty(Read(deletion ? accepted : rejected).Descendants(W + "br"));
        AssertClean(accepted);
        AssertClean(rejected);
        Assert.Equal(originalBytes, populated.DocumentByteArray);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BreakAddedToPairedParagraph_RoundTrips(bool deletion)
    {
        var plain = IrTestDocuments.Create("Copper seed traySilver watering canIndigo garden label");
        var broken = IrTestDocuments.FromBodyXml(Paragraph("", false));
        var result = DocxCompare.Compare(deletion ? broken : plain, deletion ? plain : broken);
        Assert.All(Read(result).Descendants(W + "br"), b => Assert.Equal(W + "r", b.Parent!.Name));
        AssertView(deletion ? RevisionProcessor.RejectRevisions(result) : RevisionProcessor.AcceptRevisions(result), "");
        Assert.Empty(Read(deletion ? RevisionProcessor.AcceptRevisions(result) : RevisionProcessor.RejectRevisions(result)).Descendants(W + "br"));
        AssertClean(result);
    }

    private static string Paragraph(string type, bool nested)
    {
        var br = $"<w:br{(type == "" ? "" : $" w:type=\"{type}\"")} w:clear=\"all\"/>";
        if (nested) br = "<w:r>" + br + "</w:r>";
        return "<w:p><w:r><w:t>Copper seed tray</w:t></w:r>" + br +
            "<w:r><w:t>Silver watering can</w:t></w:r>" + br +
            "<w:r><w:t>Indigo garden label</w:t></w:r></w:p>";
    }

    [Fact]
    public void Normalization_PreservesStructuralChildrenAndAlreadyNestedBreaks()
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"inventory\"/>" +
            "<w:r><w:t>One</w:t><w:br/></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>");
        Assert.Same(doc, MarkupCompatibilityNormalizer.Normalize(doc));

        var bare = IrTestDocuments.FromBodyXml("<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"inventory\"/><w:br/>" +
            "<w:bookmarkEnd w:id=\"1\"/></w:p>");
        var normalized = MarkupCompatibilityNormalizer.Normalize(bare);
        Assert.Equal(new[] { "pPr", "bookmarkStart", "r", "bookmarkEnd" },
            Read(normalized).Descendants(W + "p").Single().Elements().Select(e => e.Name.LocalName));
        AssertClean(normalized);
        Assert.Same(normalized, MarkupCompatibilityNormalizer.Normalize(normalized));
    }

    private static void AssertView(WmlDocument doc, string type)
    {
        var xml = Read(doc);
        Assert.Equal("Copper seed tray\nSilver watering can\nIndigo garden label",
            string.Concat(xml.Descendants().Where(e => e.Name == W + "t" || e.Name == W + "br")
                .Select(e => e.Name == W + "br" ? "\n" : e.Value)));
        Assert.All(xml.Descendants(W + "br"), b =>
        {
            Assert.Equal(W + "r", b.Parent!.Name);
            Assert.Equal(type, (string?)b.Attribute(W + "type") ?? "");
            Assert.Equal("all", (string?)b.Attribute(W + "clear"));
        });
    }

    private static XDocument Read(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        using var part = package.MainDocumentPart!.GetStream();
        return XDocument.Load(part);
    }

    private static void AssertClean(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package));
    }
}
