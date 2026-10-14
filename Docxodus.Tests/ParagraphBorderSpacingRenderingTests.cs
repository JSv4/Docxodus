using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1126: Word draws left and right paragraph borders w:space points OUTSIDE the text edge, so
// the text keeps its indentation. The border wrapper's content box must therefore be the text box.
public class ParagraphBorderSpacingRenderingTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original paragraphs; no external document or template is read.
    private static byte[] AuthoredDocument(params string[] paragraphs)
    {
        var content = $"<w:document xmlns:w=\"{W}\"><w:body>" + string.Concat(paragraphs)
            + "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid><w:tr><w:tc>"
            + "<w:tcPr><w:tcW w:w=\"9360\" w:type=\"dxa\"/><w:tcBorders><w:left w:val=\"single\" w:sz=\"8\" w:space=\"4\"/></w:tcBorders></w:tcPr>"
            + "<w:p><w:r><w:t>Juniper cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
        }
        return stream.ToArray();
    }

    private static string Card(string text, string space, string ind = "", string extra = "") =>
        $"<w:p><w:pPr>{extra}<w:pBdr>" + string.Concat(new[] { "top", "left", "bottom", "right" }.Select(side =>
            $"<w:{side} w:val=\"single\" w:sz=\"8\" w:space=\"{space}\" w:color=\"246B7D\"/>"))
        + $"</w:pBdr>{ind}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static XElement Convert(byte[] bytes) => WmlToHtmlConverter.ConvertToHtml(
        new WmlDocument("Juniper.docx", bytes), new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    private static XElement Wrapper(XElement html, string text) => html.Descendants()
        .First(e => e.Name.LocalName == "div" && Style(e).ContainsKey("border-left")
                    && e.Descendants().Any(p => p.Name.LocalName == "p" && p.Value.Contains(text)));

    private static XElement Paragraph(XElement wrapper, string text) =>
        wrapper.Descendants().First(e => e.Name.LocalName == "p" && e.Value.Contains(text));

    [Fact]
    public void BordersSitOutsideAnUnindentedParagraphsText()
    {
        var wrapper = Wrapper(Convert(AuthoredDocument(Card("Plain", "4"))), "Plain");
        var style = Style(wrapper);

        // 4pt space + the 0.75pt (one CSS pixel) laid-out border outside the text band on each
        // side; the text stays at the margin.
        Assert.Equal("-4.75pt", style["margin-left"]);
        Assert.Equal("-4.75pt", style["margin-right"]);
        Assert.Equal("solid #246B7D 0.75pt", style["border-left"]);
        Assert.Equal("4.0pt", style["padding-left"]);
        Assert.Equal("4.0pt", style["padding-right"]);
    }

    [Fact]
    public void IndentedParagraphsKeepTheirTextEdgesInsideTheBorders()
    {
        var html = Convert(AuthoredDocument(Card("Indented", "12", "<w:ind w:left=\"1440\" w:right=\"720\"/>")));
        var wrapper = Wrapper(html, "Indented");

        // Left text edge 72pt, right text edge 36pt: each border is 12pt + 0.75pt beyond them.
        Assert.Equal("59.25pt", Style(wrapper)["margin-left"]);
        Assert.Equal("23.25pt", Style(wrapper)["margin-right"]);
        var paragraph = Style(Paragraph(wrapper, "Indented"));
        Assert.Equal("0", paragraph["margin-right"]);
        Assert.Equal("0", paragraph["margin-left"]);
    }

    [Fact]
    public void AdjacentGroupedParagraphsShareOneOutsetBox()
    {
        var ind = "<w:ind w:left=\"720\"/>";
        var html = Convert(AuthoredDocument(Card("GroupOne", "8", ind), Card("GroupTwo", "8", ind)));
        var wrapper = Wrapper(html, "GroupOne");

        Assert.Same(wrapper, Wrapper(html, "GroupTwo"));
        Assert.Equal("27.25pt", Style(wrapper)["margin-left"]);
        Assert.Equal("-8.75pt", Style(wrapper)["margin-right"]);
    }

    [Fact]
    public void HangingIndentBorderStartsOutsideTheFirstLine()
    {
        var wrapper = Wrapper(Convert(AuthoredDocument(
            Card("Hanging", "6", "<w:ind w:left=\"1440\" w:hanging=\"720\"/>"))), "Hanging");

        // The first line starts at 36pt, so the box starts 6pt + 0.75pt before it.
        Assert.Equal("29.25pt", Style(wrapper)["margin-left"]);
    }

    [Fact]
    public void RightToLeftGroupsKeepTheirTrailingIndentOnTheParagraph()
    {
        var wrapper = Wrapper(Convert(AuthoredDocument(
            Card("Logical", "4", "<w:ind w:start=\"720\" w:end=\"720\"/>", "<w:bidi/>"))), "Logical");

        Assert.Equal("-4.75pt", Style(wrapper)["margin-right"]);
    }

    [Fact]
    public void WiderBordersLayOutInWholeCssPixels()
    {
        // 2.25pt (sz 18) is exactly three CSS pixels; 1.5pt (sz 12) is two.
        var wide = Style(Wrapper(Convert(AuthoredDocument(Card("Plain", "4").Replace("w:sz=\"8\"", "w:sz=\"18\""))), "Plain"));
        Assert.Equal("-6.25pt", wide["margin-left"]);
        var medium = Style(Wrapper(Convert(AuthoredDocument(Card("Plain", "4").Replace("w:sz=\"8\"", "w:sz=\"12\""))), "Plain"));
        Assert.Equal("-5.5pt", medium["margin-left"]);
    }

    [Fact]
    public void CellBorderSpacingIsUnchanged()
    {
        var cell = Convert(AuthoredDocument(Card("Plain", "4"))).Descendants()
            .First(e => e.Name.LocalName == "td");
        Assert.Equal("5.4pt", Style(cell)["padding-left"]);
        Assert.False(Style(cell).ContainsKey("margin-left"));
    }
}
