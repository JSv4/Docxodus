using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Page size, page margins and tab advances keep twip precision in the HTML (issue #1153). They were
/// written as inches rounded to hundredths, a 0.72pt step: A4 came out about a quarter point short and a
/// 1417-twip tab stop (0.984in) became 0.98in.
/// </summary>
public class PageAndTabLengthPrecisionTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Document(string body, string sectPr)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(
                $"<w:document xmlns:w=\"{W}\"><w:body>{body}{sectPr}</w:body></w:document>"));
            main.FeedData(xml);
            using var styles = new MemoryStream(Encoding.UTF8.GetBytes($"<w:styles xmlns:w=\"{W}\"/>"));
            main.AddNewPart<StyleDefinitionsPart>().FeedData(styles);
        }
        return new WmlDocument("Larch.docx", stream.ToArray());
    }

    private const string A4 =
        "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
        "<w:pgMar w:top=\"1417\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>";

    private const string Paragraph = "<w:p><w:r><w:t>Larch ledger.</w:t></w:r></w:p>";

    private static string PageCss(WmlDocument doc, PaginationMode mode) => string.Concat(
        WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings { RenderPagination = mode, GeneratePageCss = true })
            .Descendants().Where(e => e.Name.LocalName == "style").Select(e => e.Value));

    [Fact]
    public void ContinuousPageRule_WritesA4SizeAndMarginsInExactPoints()
    {
        var css = PageCss(Document(Paragraph, A4), PaginationMode.None);

        // 11906 x 16838 twips; margins 1417 / 1134 / 1134 / 1701 twips (top, right, bottom, left).
        Assert.Contains("size: 595.3pt 841.9pt;", css);
        Assert.Contains("margin: 70.85pt 56.7pt 56.7pt 85.05pt;", css);
        Assert.DoesNotContain("in;", css);
    }

    [Fact]
    public void PaginatedPageRule_WritesA4SizeInExactPoints()
    {
        var css = PageCss(Document(Paragraph, A4), PaginationMode.Paginated);

        Assert.Contains("size: 595.3pt 841.9pt;", css);
        Assert.Contains("margin: 0;", css);
    }

    [Fact]
    public void NamedPageRules_WriteEachSectionSizeInExactPoints()
    {
        var landscape =
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>" +
            "<w:pgMar w:top=\"1417\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1701\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>" +
            "</w:sectPr></w:pPr><w:r><w:t>Landscape section.</w:t></w:r></w:p>";
        var css = PageCss(Document(landscape + Paragraph, A4), PaginationMode.Paginated);

        Assert.Contains("@page docxodus-section-0", css);
        Assert.Contains("size: 841.9pt 595.3pt;", css);
        Assert.Contains("@page docxodus-section-1", css);
        Assert.Contains("size: 595.3pt 841.9pt;", css);
    }

    private static XElement Html(string paragraph) => WmlToHtmlConverter.ConvertToHtml(
        Document(paragraph, A4), new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    private static string Style(XElement element) => (string?)element.Attribute("style") ?? "";

    [Theory]
    [InlineData("")]
    [InlineData(" w:leader=\"dot\"")]
    public void TabAdvance_KeepsTwipPrecision(string leader)
    {
        // The paragraph starts at the margin, so the tab advances exactly to its 1417-twip stop: 70.85pt.
        var html = Html(
            $"<w:p><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1417\"{leader}/></w:tabs></w:pPr>" +
            "<w:r><w:tab/><w:t>Larch</w:t></w:r></w:p>");

        var tab = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-tab") != null);
        Assert.Contains("width: 70.85pt", Style(tab));
    }

    [Fact]
    public void LineBreakInATabbedParagraph_KeepsTwipPrecision()
    {
        // In a paragraph whose tabs are laid out, a line break carries the first-line start (1417 twips)
        // as its leading margin.
        var html = Html(
            "<w:p><w:pPr><w:ind w:firstLine=\"1417\"/></w:pPr>" +
            "<w:r><w:tab/><w:t>Larch</w:t><w:br/><w:t>ledger</w:t></w:r></w:p>");

        Assert.Single(html.Descendants(), e => e.Name.LocalName == "br");
        Assert.Contains(html.Descendants(), e => Style(e).Contains("margin: 0 0 0 70.85pt"));
    }
}
