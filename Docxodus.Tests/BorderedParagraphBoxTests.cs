using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1692: what Word 2024 does with the box around bordered paragraphs. It draws each
// paragraph's w:between below it, applies contextual spacing across the edge of the box, honours a
// page start or a keep on a bordered paragraph, and fills each paragraph's part of the box with that
// paragraph's shading. The box itself is only a border: its w:space and the spacing between its
// paragraphs are the paragraphs' padding, which is what their shading fills.
public class BorderedParagraphBoxTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Border(string sides) => "<w:pBdr>" + string.Concat(
        new[] { "top", "left", "bottom", "right", "between" }.Where(side => sides.Split(' ').Contains(side)).Select(side =>
            $"<w:{side} w:val=\"single\" w:sz=\"4\" w:space=\"{(side is "left" or "right" ? 4 : 1)}\" w:color=\"000000\"/>")) + "</w:pBdr>";

    private const string Box = "top left bottom right";
    private const string Ruled = "top left bottom between right";
    private const string Fill = "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"CCCCCC\"/>";
    private const string Contextual = "<w:contextualSpacing/>";

    // Properties in schema order: keep, border, shading, spacing, contextual spacing.
    private static string Paragraph(string text, string sides = "", string keep = "", string shading = "", string after = "",
        int space = 280, string style = "", int? spaceAfter = null) =>
        $"<w:p><w:pPr>{(style == "" ? "" : $"<w:pStyle w:val=\"{style}\"/>")}{keep}{(sides == "" ? "" : Border(sides))}{shading}"
        + $"<w:spacing w:before=\"{space}\" w:after=\"{spaceAfter ?? space}\"/>{after}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    // Literal original paragraphs; no external document or template is read.
    private static XElement Convert(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            Feed(main, $"<w:document xmlns:w=\"{W}\"><w:body>{string.Concat(paragraphs)}<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
                + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>");
            Feed(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{W}\"><w:docDefaults><w:pPrDefault><w:pPr>"
                + "<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault></w:docDefaults>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Other\"><w:name w:val=\"Other\"/><w:basedOn w:val=\"Normal\"/></w:style></w:styles>");
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Sedge.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated });

        static void Feed(OpenXmlPart part, string xml)
        {
            using var content = new MemoryStream(Encoding.UTF8.GetBytes(xml));
            part.FeedData(content);
        }
    }

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2)).ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    private static string Value(XElement element, string property) => Style(element).GetValueOrDefault(property, "unset");

    private static XElement Block(XElement html, string text) =>
        html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == text);

    /// <summary>The paragraph's box, or the paragraph itself when nothing wraps it.</summary>
    private static XElement BoxOf(XElement html, string text) =>
        Block(html, text).Parent is { } parent && Style(parent).Keys.Any(key => key.StartsWith("border-")) ? parent : Block(html, text);

    // ---- w:between -----------------------------------------------------------------------------

    [Fact]
    public void EachBetweenRuleIsTheTopBorderOfTheNextParagraphInTheBox()
    {
        var html = Convert(Paragraph("SedgeOne", Ruled), Paragraph("SedgeTwo", Ruled), Paragraph("SedgeThree", Ruled));
        var (one, two, three) = (Block(html, "SedgeOne"), Block(html, "SedgeTwo"), Block(html, "SedgeThree"));

        Assert.Same(one.Parent, three.Parent);
        Assert.Equal("unset", Value(one, "border-top"));
        foreach (var lower in new[] { two, three })
        {
            Assert.Equal("solid #000000 0.75pt", Value(lower, "border-top"));
            Assert.Equal("0.5pt", Value(lower, "--docx-border-top-width"));
            Assert.Equal("true", (string?)lower.Attribute("data-docx-exact-border"));
            // w:space below the rule. The rule runs to the side borders because the box's 4pt
            // w:space is the paragraph's own padding.
            Assert.Equal(("1.0pt", "4pt", "4pt"), (Value(lower, "padding-top"), Value(lower, "padding-left"), Value(lower, "padding-right")));
            Assert.DoesNotContain("-", Value(lower, "margin-left") + Value(lower, "margin-right"));
        }
        // Above each rule the space after (14pt) and the rule's w:space; below the last paragraph
        // its space after is outside the box and only the box's own w:space remains.
        Assert.Equal(("15pt", "15pt", "1pt"), (Value(one, "padding-bottom"), Value(two, "padding-bottom"), Value(three, "padding-bottom")));
    }

    [Theory]
    [InlineData(100, 400, "6pt", "16pt", "15pt")]  // 5pt after, 20pt before: 5 and w:space above the rule, w:space and the other 15 below
    [InlineData(400, 100, "21pt", "1.0pt", null)]    // 20pt after, 5pt before: all 20 above the rule
    public void TheRuleSitsBelowTheSpaceAfterAndAboveWhatTheSpaceBeforeAdds(int after, int before, string above, string below, string? dropped)
    {
        // Word 2024: space after, w:space, rule, w:space, then max(0, space before - space after).
        var html = Convert(Paragraph("SedgeOne", Ruled, spaceAfter: after), Paragraph("SedgeTwo", Ruled, space: before));
        var (one, two) = (Block(html, "SedgeOne"), Block(html, "SedgeTwo"));

        Assert.Equal((above, below), (Value(one, "padding-bottom"), Value(two, "padding-top")));
        Assert.Equal(("0", "0"), (Value(one, "margin-bottom"), Value(two, "margin-top")));
        // What the paginator takes off the padding when a page starts with the paragraph.
        Assert.Equal(dropped, (string?)two.Attribute("data-docx-space-before"));
    }

    [Fact]
    public void ARuleInABoxWithNoRightBorderStillReachesTheLeftBorder()
    {
        // The left border's w:space band is not filled (see the shading tests), so it stays the
        // box's padding and the ruled paragraph reaches through it.
        var html = Convert(Paragraph("SedgeOne", "top left bottom between"), Paragraph("SedgeTwo", "top left bottom between"));

        Assert.Equal("4.0pt", Value(BoxOf(html, "SedgeOne"), "padding-left"));
        Assert.Equal(("-4.0pt", "4.0pt"), (Value(Block(html, "SedgeTwo"), "margin-left"), Value(Block(html, "SedgeTwo"), "padding-left")));
        Assert.Equal("unset", Value(Block(html, "SedgeOne"), "padding-left"));
    }

    [Fact]
    public void ParagraphsThatDifferOnlyInBetweenShareOneBoxAndEachDrawsItsOwnRule()
    {
        // Word: the rule and the w:space below it come from the paragraph above, the w:space above it from the one below.
        var html = Convert(Paragraph("SedgeOne", Ruled), Paragraph("SedgeTwo", Box), Paragraph("SedgeThree", Ruled));
        var (one, two, three) = (Block(html, "SedgeOne"), Block(html, "SedgeTwo"), Block(html, "SedgeThree"));

        Assert.Same(one.Parent, three.Parent);
        Assert.Equal(("14pt", "solid #000000 0.75pt", "1.0pt", "15pt"),
            (Value(one, "padding-bottom"), Value(two, "border-top"), Value(two, "padding-top"), Value(two, "padding-bottom")));
        Assert.Equal(("unset", "unset"), (Value(three, "border-top"), Value(three, "padding-top")));
    }

    [Fact]
    public void ABoxWithoutBetweenKeepsItsParagraphsUnruled()
    {
        var html = Convert(Paragraph("SedgeOne", Box), Paragraph("SedgeTwo", Box));

        Assert.Same(Block(html, "SedgeOne").Parent, Block(html, "SedgeTwo").Parent);
        foreach (var text in new[] { "SedgeOne", "SedgeTwo" })
            Assert.DoesNotContain(Style(Block(html, text)).Keys, key => key.StartsWith("border"));
        // Equal spacing collapses to the upper paragraph's space after; the lower one adds nothing.
        Assert.Equal(("14pt", "unset"), (Value(Block(html, "SedgeOne"), "padding-bottom"), Value(Block(html, "SedgeTwo"), "padding-top")));
    }

    // ---- w:contextualSpacing at the edge of a box ----------------------------------------------

    [Theory]
    [InlineData("", "0", "0", "0")]           // one style: the box and its neighbours give up their spacing
    [InlineData("Other", "14pt", "14pt", "14pt")] // another style beside the box: nothing is removed
    public void ContextualSpacingReachesAcrossTheEdgeOfABox(string neighbourStyle, string above, string below, string next)
    {
        var html = Convert(Paragraph("SedgeOpen", after: Contextual, space: 0, style: neighbourStyle),
            Paragraph("SedgeBoxed", Box, after: Contextual), Paragraph("SedgeNext", after: Contextual, style: neighbourStyle),
            Paragraph("SedgeClose", space: 0, style: "Other"));
        var box = BoxOf(html, "SedgeBoxed");

        Assert.NotSame(box, Block(html, "SedgeBoxed"));
        Assert.Equal((above, below, next), (Value(box, "margin-top"), Value(box, "margin-bottom"), Value(Block(html, "SedgeNext"), "margin-top")));
    }

    [Fact]
    public void TheSpaceAfterABoxGivesUpStillAbsorbsTheNextSpaceBefore()
    {
        // Only the boxed paragraph has the flag: Word keeps 14pt above it (the neighbour's space
        // after) and leaves max(0, 14 - 14) of the next paragraph's space before below it.
        var html = Convert(Paragraph("SedgeOpen"), Paragraph("SedgeBoxed", Box, after: Contextual), Paragraph("SedgeNext"));
        var box = BoxOf(html, "SedgeBoxed");

        Assert.Equal(("14pt", "0", "0", "0"), (Value(Block(html, "SedgeOpen"), "margin-bottom"), Value(box, "margin-top"),
            Value(box, "margin-bottom"), Value(Block(html, "SedgeNext"), "margin-top")));
    }

    // ---- page start and keep -------------------------------------------------------------------

    [Fact]
    public void TheBoxStartsThePageItsFirstParagraphStarts()
    {
        var html = Convert(Paragraph("SedgeOpen"), Paragraph("SedgeBoxed", Box, keep: "<w:pageBreakBefore/>"));

        Assert.Equal("true", (string?)BoxOf(html, "SedgeBoxed").Attribute("data-page-break-before"));
        Assert.Null(Block(html, "SedgeOpen").Attribute("data-page-break-before"));
    }

    [Fact]
    public void APageStartInsideARunOfBorderedParagraphsClosesTheBoxBeforeIt()
    {
        var html = Convert(Paragraph("SedgeOne", Box), Paragraph("SedgeTwo", Box, keep: "<w:pageBreakBefore/>"), Paragraph("SedgeThree", Box));
        var (first, second) = (BoxOf(html, "SedgeOne"), BoxOf(html, "SedgeTwo"));

        Assert.NotSame(first, second);
        Assert.Same(second, BoxOf(html, "SedgeThree"));
        Assert.Null(first.Attribute("data-page-break-before"));
        Assert.Equal("true", (string?)second.Attribute("data-page-break-before"));
    }

    [Theory]
    [InlineData(false, true, "true")]  // the last paragraph keeps with the next block: so does the box
    [InlineData(true, false, null)]    // a keep on an earlier paragraph is between members of the box
    [InlineData(false, false, null)]
    public void TheBoxKeepsWithTheNextBlockWhenItsLastParagraphDoes(bool first, bool last, string? expected)
    {
        var html = Convert(Paragraph("SedgeOne", Box, keep: first ? "<w:keepNext/>" : ""),
            Paragraph("SedgeTwo", Box, keep: last ? "<w:keepNext/>" : ""), Paragraph("SedgeNext"));

        Assert.Equal(expected, (string?)BoxOf(html, "SedgeOne").Attribute("data-keep-with-next"));
    }

    // ---- shading -------------------------------------------------------------------------------

    // Each paragraph's shading fills its own part of the box: Word fills from the border (or the
    // paragraph above) down to the next paragraph's first line, and across to the side borders. That
    // region is the paragraph's padding, so the fill is the paragraph's own background and the box
    // has none. `top`, `side` and `bottom` are the first paragraph's padding; w:space is 1pt above
    // and below and 4pt at the sides, and the paragraphs are 14pt apart.
    [Theory]
    [InlineData(Box, true, true, "1pt", "4pt", "14pt")]
    [InlineData(Box, true, false, "1pt", "4pt", "14pt")]          // mixed shading: the same slabs, one of them unfilled
    [InlineData(Box, false, true, "1pt", "4pt", "14pt")]
    [InlineData("top", true, true, "1pt", "unset", "14pt")]
    [InlineData("left right", true, true, "unset", "4pt", "14pt")]
    [InlineData("top right bottom", true, true, "1pt", "4pt", "14pt")]
    public void ShadingFillsEachParagraphsPartOfTheBox(string sides, bool first, bool second, string top, string side, string bottom)
    {
        var html = Convert(Paragraph("SedgeOne", sides, shading: first ? Fill : ""), Paragraph("SedgeTwo", sides, shading: second ? Fill : ""));
        var (box, one, two) = (BoxOf(html, "SedgeOne"), Block(html, "SedgeOne"), Block(html, "SedgeTwo"));

        Assert.Same(box, BoxOf(html, "SedgeTwo"));
        Assert.DoesNotContain(Style(box).Keys, key => key.StartsWith("background") || key.StartsWith("padding"));
        Assert.Equal((first ? "#CCCCCC" : "unset", second ? "#CCCCCC" : "unset"), (Value(one, "background"), Value(two, "background")));
        Assert.Equal((top, side, bottom), (Value(one, "padding-top"), Value(one, "padding-right"), Value(one, "padding-bottom")));
        // The second paragraph starts where the first one's space after ends and closes the box.
        Assert.Equal(("unset", side, top == "unset" || sides == "top" ? "unset" : "1pt"),
            (Value(two, "padding-top"), Value(two, "padding-right"), Value(two, "padding-bottom")));
    }

    [Theory]
    [InlineData("left")]             // Word leaves a left border's w:space unshaded unless the box has a right border too
    [InlineData("top left bottom")]
    public void TheSpaceBesideALeftBorderWithNoRightBorderIsNotFilled(string sides)
    {
        var html = Convert(Paragraph("SedgeOne", sides, shading: Fill), Paragraph("SedgeTwo", sides, shading: Fill));
        var box = BoxOf(html, "SedgeOne");

        // The band stays the box's padding, outside the paragraphs' backgrounds; above and below, the
        // fill still reaches the borders.
        Assert.Equal(("4.0pt", "unset"), (Value(box, "padding-left"), Value(box, "background")));
        Assert.Equal(("unset", sides == "left" ? "unset" : "1pt"), (Value(Block(html, "SedgeOne"), "padding-left"), Value(Block(html, "SedgeOne"), "padding-top")));
    }

    [Fact]
    public void AnUnborderedShadedParagraphIsUnchanged()
    {
        var html = Convert(Paragraph("SedgePlain", shading: Fill));

        Assert.Equal("#CCCCCC", Value(Block(html, "SedgePlain"), "background"));
        Assert.Same(Block(html, "SedgePlain"), BoxOf(html, "SedgePlain"));
        Assert.DoesNotContain(Block(html, "SedgePlain").Ancestors(), ancestor => Style(ancestor).ContainsKey("background"));
    }
}
