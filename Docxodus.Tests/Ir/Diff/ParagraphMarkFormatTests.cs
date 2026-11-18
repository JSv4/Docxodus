using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A change to a modeled run property of a paragraph mark (its size, bold, colour, …) changes how the paragraph
/// looks, an empty paragraph's height among it, so the comparison tracks it as a paragraph-mark run-property change;
/// an unmodeled mark property such as its language stays untracked under the default policy (issue #1794).
/// </summary>
public class ParagraphMarkFormatTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Document(string mark) => IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>Empty mark.</w:t></w:r></w:p>" +
        $"<w:p>{(mark.Length == 0 ? "" : $"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>")}</w:p>" +
        "<w:p><w:r><w:t>After.</w:t></w:r></w:p>");

    private static XElement? MarkProperties(WmlDocument doc)
    {
        using var stream = new System.IO.MemoryStream(doc.DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        var body = package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
        return body.Elements(W + "p").ElementAt(1).Element(W + "pPr")?.Element(W + "rPr");
    }

    private static string? Size(WmlDocument doc) => (string?)MarkProperties(doc)?.Element(W + "sz")?.Attribute(W + "val");

    [Theory]
    [InlineData("", "<w:sz w:val=\"72\"/>")]
    [InlineData("<w:sz w:val=\"72\"/>", "")]
    [InlineData("<w:sz w:val=\"24\"/>", "<w:sz w:val=\"72\"/>")]
    public void A_paragraph_mark_size_change_is_tracked_and_rejects_to_the_original(string original, string revised)
    {
        var left = Document(original);
        var right = Document(revised);

        var redline = DocxCompare.Compare(left, right);
        Assert.NotNull(MarkProperties(redline)?.Element(W + "rPrChange"));
        Assert.Equal(Size(left), Size(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(Size(right), Size(RevisionProcessor.AcceptRevisions(redline)));
        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void A_paragraph_mark_change_is_one_revision_under_either_policy()
    {
        foreach (var comparison in new[] { DocxDiffFormatComparison.ModeledOnly, DocxDiffFormatComparison.Full })
        {
            var revision = Assert.Single(DocxDiff.GetRevisions(Document(""), Document("<w:sz w:val=\"72\"/>"),
                new DocxDiffSettings { FormatComparison = comparison }));
            Assert.Equal(DocxDiffRevisionType.FormatChanged, revision.Type);
            Assert.Equal(DocxDiffFormatChangeScope.Paragraph, revision.FormatChange!.Scope);
            Assert.NotEmpty(revision.FormatChange.ChangedPropertyNames);
        }
    }

    [Fact]
    public void A_paragraph_mark_language_change_stays_untracked()
    {
        // Control: the default policy does not model a mark's language, as it does not model a run's.
        var (left, right) = (Document(""), Document("<w:lang w:val=\"nb-NO\"/>"));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
        var redline = DocxCompare.Compare(left, right);
        using var stream = new System.IO.MemoryStream(redline.DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        Assert.Empty(package.MainDocumentPart!.GetXDocument().Descendants(W + "rPrChange"));
    }
}
