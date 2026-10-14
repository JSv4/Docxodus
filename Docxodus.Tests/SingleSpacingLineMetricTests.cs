using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1123: single spacing records its unit multiple, so the paginated text renderer can place
// its lines on Word's twip grid, while the paragraph's line box and inline content stay unchanged.
public class SingleSpacingLineMetricTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original paragraphs; no external document or template is read.
    private static XElement Render(string line)
    {
        var content = $"<w:document xmlns:w=\"{W}\"><w:body>"
            + $"<w:p><w:pPr><w:spacing w:after=\"0\" w:line=\"{line}\" w:lineRule=\"auto\"/></w:pPr><w:r><w:t>Chestnut margin record.</w:t></w:r></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Chestnut.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
    }

    private static XElement Paragraph(XElement html) =>
        html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value.Contains("Chestnut"));

    [Fact]
    public void SingleSpacingRecordsItsUnitMultiple()
    {
        var paragraph = Paragraph(Render("240"));
        var style = (string?)paragraph.Attribute("style") ?? "";

        Assert.Contains("--docx-auto-line-spacing: 1", style);
        Assert.DoesNotContain("line-height", style);
        // The run is not rewritten: a unit multiple neither scales nor lifts its line.
        var run = paragraph.Elements().Single();
        Assert.DoesNotContain("calc(1lh", (string?)run.Attribute("style") ?? "");
        Assert.DoesNotContain("position", (string?)run.Attribute("style") ?? "");
    }

    [Fact]
    public void OtherMultiplesKeepTheirLineSpacingRewrite()
    {
        var paragraph = Paragraph(Render("276"));
        Assert.Contains("--docx-auto-line-spacing: 1.15", (string?)paragraph.Attribute("style") ?? "");
        Assert.Contains("calc(1lh * var(--docx-auto-line-spacing))", (string?)paragraph.Elements().Single().Attribute("style") ?? "");
    }
}
