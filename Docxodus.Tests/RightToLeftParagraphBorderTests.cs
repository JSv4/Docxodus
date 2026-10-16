using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Paragraph borders on right-to-left (<c>w:bidi</c>) paragraphs. Unlike <c>w:ind</c>'s
/// <c>w:left</c>/<c>w:right</c>, which a bidi paragraph reads as start/end, <c>w:pBdr</c>'s
/// <c>w:left</c> and <c>w:right</c> are physical: ISO/IEC 29500-1 §17.3.1.17 says the left border
/// "shall be displayed on the left side of the page … This shall not change based on the paragraph
/// direction", and LibreOffice draws it there (issue #1154).
/// </summary>
public class RightToLeftParagraphBorderTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement Convert(params string[] paragraphs)
    {
        var content = $"<w:document xmlns:w=\"{W}\"><w:body>" + string.Concat(paragraphs)
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            document.AddMainDocumentPart().FeedData(xml);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Rowan.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
    }

    // A thick double border on the left and a thin single one on the right, so the sides are distinguishable.
    private const string UnequalBorders =
        "<w:pBdr><w:left w:val=\"double\" w:sz=\"24\" w:space=\"4\" w:color=\"C0392B\"/>" +
        "<w:right w:val=\"single\" w:sz=\"4\" w:space=\"4\" w:color=\"1F4E79\"/></w:pBdr>";

    private static string BorderedParagraph(string text, string ind, bool bidi) =>
        $"<w:p><w:pPr>{(bidi ? "<w:bidi/>" : "")}{UnequalBorders}{ind}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    private static List<XElement> Wrappers(XElement html) => html.Descendants()
        .Where(e => e.Name.LocalName == "div" && Style(e).ContainsKey("border-left")).ToList();

    private static XElement WrapperOf(XElement html, string text) =>
        Wrappers(html).Single(e => e.Descendants().Any(p => p.Name.LocalName == "p" && p.Value.Contains(text)));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LeftAndRightBordersStayOnTheirPhysicalSides(bool bidi)
    {
        var wrapper = WrapperOf(Convert(BorderedParagraph("Rowan", "", bidi)), "Rowan");

        Assert.StartsWith("double", Style(wrapper)["border-left"]);
        Assert.StartsWith("solid", Style(wrapper)["border-right"]);
    }

    [Fact]
    public void RightToLeftIndentsMirrorWhileBordersStayPhysical()
    {
        // Leading (start) indent 72pt, trailing (end) indent 18pt. On the right-to-left paragraph the
        // leading indent is on the right, but the right border is still the thin single one, so each
        // margin subtracts the outset of the border that is physically on that side.
        var ind = "<w:ind w:start=\"1440\" w:end=\"360\"/>";
        var ltr = Style(WrapperOf(Convert(BorderedParagraph("Forward", ind, bidi: false)), "Forward"));
        var rtl = Style(WrapperOf(Convert(BorderedParagraph("Mirror", ind, bidi: true)), "Mirror"));

        Assert.StartsWith("double", rtl["border-left"]);
        Assert.StartsWith("solid", rtl["border-right"]);
        // Each outset is w:space (4pt) plus the border's laid-out width: 7.5pt for the double, 0.75pt (one CSS
        // pixel) for the single.
        Assert.Equal("60.5pt", ltr["margin-left"]);
        Assert.Equal("13.25pt", ltr["margin-right"]);
        Assert.Equal("67.25pt", rtl["margin-right"]);
        Assert.Equal("6.5pt", rtl["margin-left"]);
    }
    // Issue #1155: consecutive bordered paragraphs share a box only when their boxes coincide.
    private const string SingleBorders =
        "<w:pBdr><w:top w:val=\"single\" w:sz=\"8\" w:space=\"4\"/><w:left w:val=\"single\" w:sz=\"8\" w:space=\"4\"/>" +
        "<w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"4\"/><w:right w:val=\"single\" w:sz=\"8\" w:space=\"4\"/>" +
        "<w:between w:val=\"single\" w:sz=\"8\" w:space=\"4\"/></w:pBdr>";

    private static string Boxed(string text, string ind, bool bidi) =>
        $"<w:p><w:pPr>{(bidi ? "<w:bidi/>" : "")}{SingleBorders}{ind}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private const string Unequal = "<w:ind w:start=\"720\" w:end=\"1440\"/>";

    [Fact]
    public void MixedDirectionsWithUnequalIndentsGetMirroredBoxes()
    {
        var html = Convert(Boxed("Eastward", Unequal, bidi: false), Boxed("Westward", Unequal, bidi: true));

        Assert.Equal(2, Wrappers(html).Count);
        var ltr = Style(WrapperOf(html, "Eastward"));
        var rtl = Style(WrapperOf(html, "Westward"));
        // Start 36pt and end 72pt, each less the 4pt space and the 0.75pt border.
        Assert.Equal("31.25pt", ltr["margin-left"]);
        Assert.Equal("67.25pt", ltr["margin-right"]);
        Assert.Equal("31.25pt", rtl["margin-right"]);
        Assert.Equal("67.25pt", rtl["margin-left"]);
    }

    [Fact]
    public void MixedDirectionsWithEqualIndentsStillShareOneBox()
    {
        var equal = "<w:ind w:start=\"720\" w:end=\"720\"/>";
        var html = Convert(Boxed("Eastward", equal, bidi: false), Boxed("Westward", equal, bidi: true),
            Boxed("Onward", equal, bidi: false));

        var wrapper = Assert.Single(Wrappers(html));
        Assert.Equal(3, wrapper.Descendants().Count(e => e.Name.LocalName == "p"));
    }

    [Fact]
    public void RightToLeftParagraphsWithUnequalIndentsShareOneBox()
    {
        var html = Convert(Boxed("Westward", Unequal, bidi: true), Boxed("Seaward", Unequal, bidi: true));

        var wrapper = Assert.Single(Wrappers(html));
        Assert.Equal("31.25pt", Style(wrapper)["margin-right"]);
        Assert.Equal("67.25pt", Style(wrapper)["margin-left"]);
    }
    [Fact]
    public void MixedDirectionsShareOneBoxWhenAHangingIndentMakesTheBoxSymmetric()
    {
        // Start 72pt less a 36pt hanging indent puts the box's leading edge 36pt in, the same as the
        // 36pt end: the box is the same in either direction, so the paragraphs join (as LibreOffice joins them).
        var hanging = "<w:ind w:start=\"1440\" w:hanging=\"720\" w:end=\"720\"/>";
        var html = Convert(Boxed("Eastward", hanging, bidi: false), Boxed("Westward", hanging, bidi: true));

        var wrapper = Assert.Single(Wrappers(html));
        Assert.Equal(2, wrapper.Descendants().Count(e => e.Name.LocalName == "p"));
    }
}
