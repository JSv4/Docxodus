using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1677: Word draws a paragraph border around the text and w:space only. The space before the
// first bordered paragraph and after the last lie outside the box, where they meet the neighbours'
// spacing; the border wrapper must therefore carry them, not the paragraphs inside it.
public class ParagraphBorderOuterSpacingTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Box = "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"1\"/>"
        + "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"1\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"1\"/></w:pBdr>";

    private const string Frame = "<w:framePr w:w=\"4000\" w:wrap=\"notBeside\" w:vAnchor=\"text\" w:hAnchor=\"text\" w:y=\"1\" w:yAlign=\"inline\"/>";

    private static string Paragraph(string text, int before, int after, string properties = "") =>
        $"<w:p><w:pPr>{properties}<w:spacing w:before=\"{before}\" w:after=\"{after}\"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    // Literal original paragraphs; no external document or template is read.
    private static XElement Convert(string body, string compat = "")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            Feed(main, $"<w:document xmlns:w=\"{W}\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
                + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>");
            Feed(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{W}\"><w:compat>{compat}</w:compat></w:settings>");
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Larch.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

        static void Feed(OpenXmlPart part, string xml)
        {
            using var content = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(content);
        }
    }

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    private static XElement Block(XElement html, string text) =>
        html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == text);

    private static (string Top, string Bottom) Spacing(XElement block) =>
        (Style(block).GetValueOrDefault("margin-top", "none"), Style(block).GetValueOrDefault("margin-bottom", "none"));

    private static (string Top, string Bottom) Padding(XElement block) =>
        (Style(block).GetValueOrDefault("padding-top", "none"), Style(block).GetValueOrDefault("padding-bottom", "none"));

    [Fact]
    public void ABorderedParagraphsSpacingLiesOutsideItsBorder()
    {
        var html = Convert(Paragraph("LarchOpen", 0, 0) + Paragraph("LarchBoxed", 280, 200, Box) + Paragraph("LarchNext", 280, 0));
        var boxed = Block(html, "LarchBoxed");
        var box = boxed.Parent!;

        Assert.Contains("border-top", Style(box).Keys);
        Assert.Equal(("14pt", "10pt"), Spacing(box));
        Assert.Equal(("0", "0"), Spacing(boxed));
    }

    [Fact]
    public void OnlyTheOuterSpacingOfAGroupLeavesItsParagraphs()
    {
        var html = Convert(Paragraph("LarchOne", 280, 100, Box) + Paragraph("LarchTwo", 120, 160, Box) + Paragraph("LarchThree", 180, 200, Box));
        var box = Block(html, "LarchOne").Parent!;

        Assert.Equal(3, box.Elements().Count());
        Assert.Equal(("14pt", "10pt"), Spacing(box));
        // Between the members the spacing stays inside the box as the paragraphs' padding (issue
        // #1692): each keeps its space after, and its space before only where that exceeds the
        // space after above it. The first and last also hold the box's 1pt w:space.
        foreach (var text in new[] { "LarchOne", "LarchTwo", "LarchThree" })
            Assert.Equal(("0", "0"), Spacing(Block(html, text)));
        Assert.Equal(("1pt", "5pt"), Padding(Block(html, "LarchOne")));
        Assert.Equal(("1pt", "8pt"), Padding(Block(html, "LarchTwo")));
        Assert.Equal(("1pt", "1pt"), Padding(Block(html, "LarchThree")));
    }

    [Fact]
    public void AnUnborderedParagraphKeepsItsOwnSpacing()
    {
        var html = Convert(Paragraph("LarchOpen", 0, 0) + Paragraph("LarchPlain", 280, 200) + Paragraph("LarchNext", 280, 0));
        var plain = Block(html, "LarchPlain");

        Assert.Equal(("14pt", "10pt"), Spacing(plain));
        Assert.DoesNotContain(plain.Ancestors(), a => Style(a).ContainsKey("margin-top") || Style(a).ContainsKey("border-top"));
    }

    // w:doNotUseHTMLParagraphAutoSpacing adds the space after above to a paragraph's space before
    // (#1553). The sum is carried by the following paragraph, so the box's own space after must be
    // outside the border to collapse into it instead of counting a second time.
    [Fact]
    public void TheSummedSpacingBelowABoxAbsorbsTheBoxsSpaceAfter()
    {
        var html = Convert(Paragraph("LarchBoxed", 280, 280, Box) + Paragraph("LarchNext", 280, 0),
            "<w:doNotUseHTMLParagraphAutoSpacing/>");

        Assert.Equal("14pt", Spacing(Block(html, "LarchBoxed").Parent!).Bottom);
        Assert.Equal("0", Spacing(Block(html, "LarchBoxed")).Bottom);
        Assert.Equal("28pt", Spacing(Block(html, "LarchNext")).Top);
    }

    // The frame slot already took its last paragraph's space after; it now takes it from the box
    // that paragraph sits in, through the same rule.
    [Theory]
    [InlineData("")]
    [InlineData(Box)]
    public void AnInFlowFramesTrailingSpacingStaysOnItsSlot(string border)
    {
        var html = Convert(Paragraph("LarchFramed", 280, 200, Frame + border) + Paragraph("LarchNext", 0, 0));
        var framed = Block(html, "LarchFramed");
        var slot = framed.Ancestors().Single(a => a.Attribute("data-docx-frame-flow") != null);

        Assert.Equal("10pt", Spacing(slot).Bottom);
        Assert.All(framed.AncestorsAndSelf().TakeWhile(a => a != slot), inner => Assert.NotEqual("10pt", Spacing(inner).Bottom));
        Assert.Equal("14pt", Spacing(border == "" ? framed : framed.Parent!).Top);
    }
}
