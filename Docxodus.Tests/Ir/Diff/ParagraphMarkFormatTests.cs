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

    [Fact]
    public void A_run_of_blank_paragraphs_whose_marks_change_rejects_to_the_original()
    {
        // Blank paragraphs at the start of the document whose marks change size: the redline keeps one shared
        // mark per blank, and it must record the original's.
        static WmlDocument Blanks(string mark) => IrTestDocuments.FromBodyXml(
            string.Concat(Enumerable.Repeat($"<w:p>{(mark.Length == 0 ? "" : $"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>")}</w:p>", 3)) +
            "<w:p><w:r><w:t>After the blanks.</w:t></w:r></w:p>");
        static string?[] Sizes(WmlDocument doc)
        {
            using var stream = new System.IO.MemoryStream(doc.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            return package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Elements(W + "p")
                .Select(p => (string?)p.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val")).ToArray();
        }
        var (left, right) = (Blanks(""), Blanks("<w:sz w:val=\"72\"/>"));
        var redline = DocxCompare.Compare(left, right);
        Assert.Equal(Sizes(left), Sizes(RevisionProcessor.RejectRevisions(redline)));
        Assert.Equal(Sizes(right), Sizes(RevisionProcessor.AcceptRevisions(redline)));
    }

    [Fact]
    public void A_mark_change_reports_mark_prefixed_property_names()
    {
        // A mark's run properties never read as the paragraph's own properties of the same name.
        var revision = Assert.Single(DocxDiff.GetRevisions(Document(""), Document("<w:rStyle w:val=\"Strong\"/><w:sz w:val=\"72\"/>")));
        Assert.All(revision.FormatChange!.ChangedPropertyNames, name => Assert.StartsWith("mark", name));
        Assert.DoesNotContain("style", revision.FormatChange.ChangedPropertyNames);
    }

    [Fact]
    public void Consolidating_a_mark_change_beside_another_reviewers_text_edit_keeps_the_text_edit()
    {
        // The composite leaves marks out of its pairing, so one reviewer's mark change cannot conflict with and
        // cancel another reviewer's text edit in the same paragraph.
        static WmlDocument Paragraph(string text, string mark) => IrTestDocuments.FromBodyXml(
            $"<w:p>{(mark.Length == 0 ? "" : $"<w:pPr><w:rPr>{mark}</w:rPr></w:pPr>")}<w:r><w:t>{text}</w:t></w:r></w:p>");
        var original = Paragraph("Hello world here.", "");
        var merged = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Document = Paragraph("Hello world here.", "<w:sz w:val=\"72\"/>"), Author = "Reviewer A" },
            new DocxDiffReviewer { Document = Paragraph("Hello there here.", ""), Author = "Reviewer B" },
        });
        using var stream = new System.IO.MemoryStream(RevisionProcessor.AcceptRevisions(merged).DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        Assert.Equal("Hello there here.", package.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!.Value);
    }

    [Theory]
    [InlineData("WC007-Unmodified.docx", "WC007-Longest-At-End.docx")]
    [InlineData("WC007-Longest-At-End.docx", "WC007-Unmodified.docx")]
    public void A_replaced_paragraph_whose_marks_never_change_gets_no_mark_change(string left, string right)
    {
        // A deleted paragraph's mark carries the deletion marker, not formatting; no mark run-property change may
        // be invented from it where neither document's mark has any properties.
        var redline = DocxDiff.Compare(new WmlDocument(System.IO.Path.Combine(System.AppContext.BaseDirectory, "../../../../TestFiles", "WC", left)),
            new WmlDocument(System.IO.Path.Combine(System.AppContext.BaseDirectory, "../../../../TestFiles", "WC", right)));
        using var stream = new System.IO.MemoryStream(redline.DocumentByteArray);
        using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        Assert.DoesNotContain(package.MainDocumentPart!.GetXDocument().Descendants(W + "rPrChange"),
            change => change.Parent!.Parent!.Name == W + "pPr");
    }
}
