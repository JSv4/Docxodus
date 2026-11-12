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
// paragraph's shading. The spacing between the paragraphs of a box is their padding, and a
// paragraph that paints at an edge of the box (a fill or a rule) reaches through the box's w:space
// with a negative margin and as much padding.
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
            // w:space below the rule, and the rule runs through the box's 4pt w:space to its side borders.
            Assert.Equal(("1.0pt", "-4pt", "4pt", "-4pt", "4pt"), (Value(lower, "padding-top"),
                Value(lower, "margin-left"), Value(lower, "padding-left"), Value(lower, "margin-right"), Value(lower, "padding-right")));
        }
        // Above each rule the space after (14pt) and the rule's w:space; below the last paragraph
        // its space after is outside the box, where the box's own border follows.
        Assert.Equal(("15pt", "15pt", "unset"), (Value(one, "padding-bottom"), Value(two, "padding-bottom"), Value(three, "padding-bottom")));
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
        // The left border's w:space band is not filled (see the shading tests), but the rule crosses it.
        var html = Convert(Paragraph("SedgeOne", "top left bottom between", shading: Fill), Paragraph("SedgeTwo", "top left bottom between", shading: Fill));

        Assert.Equal(("-4pt", "4pt"), (Value(Block(html, "SedgeTwo"), "margin-left"), Value(Block(html, "SedgeTwo"), "padding-left")));
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
    // paragraph above) down to the next paragraph's first line, and across to the side borders. The
    // fill is the paragraph's own background and the box has none: the gap below a paragraph is its
    // padding, and a shaded paragraph reaches through the box's w:space (1pt above and below, 4pt at
    // the sides) with a negative margin and as much padding. An unshaded one does not.
    [Theory]
    [InlineData(Box, true, true)]
    [InlineData(Box, true, false)]          // mixed shading: the same slabs, one of them unfilled
    [InlineData(Box, false, true)]
    [InlineData("top", true, true)]
    [InlineData("left right", true, true)]
    [InlineData("top right bottom", true, true)]
    public void ShadingFillsEachParagraphsPartOfTheBox(string sides, bool first, bool second)
    {
        var html = Convert(Paragraph("SedgeOne", sides, shading: first ? Fill : ""), Paragraph("SedgeTwo", sides, shading: second ? Fill : ""));
        var (box, one, two) = (BoxOf(html, "SedgeOne"), Block(html, "SedgeOne"), Block(html, "SedgeTwo"));
        // Whether the paragraph's margin is negative, and its padding, on one side.
        (bool, string) Reach(XElement paragraph, string side) =>
            (Value(paragraph, "margin-" + side).StartsWith('-'), Value(paragraph, "padding-" + side));
        (bool, string) Through(bool shaded, string side, string space) =>
            shaded && sides.Split(' ').Contains(side) ? (true, space) : (false, "unset");

        Assert.Same(box, BoxOf(html, "SedgeTwo"));
        Assert.Equal("unset", Value(box, "background"));
        Assert.Equal((first ? "#CCCCCC" : "unset", second ? "#CCCCCC" : "unset"), (Value(one, "background"), Value(two, "background")));
        // The gap between them is the first paragraph's, shaded or not.
        Assert.Equal(("0", "14pt", "0"), (Value(one, "margin-bottom"), Value(one, "padding-bottom"), Value(two, "margin-top")));
        Assert.Equal(Through(first, "top", "1pt"), Reach(one, "top"));
        Assert.Equal(Through(second, "bottom", "1pt"), Reach(two, "bottom"));
        Assert.Equal((false, "unset"), Reach(two, "top"));
        foreach (var (paragraph, shaded) in new[] { (one, first), (two, second) })
            Assert.Equal(Through(shaded, "right", "4pt"), Reach(paragraph, "right"));
    }

    [Theory]
    [InlineData("left")]             // Word leaves a left border's w:space unshaded unless the box has a right border too
    [InlineData("top left bottom")]
    public void TheSpaceBesideALeftBorderWithNoRightBorderIsNotFilled(string sides)
    {
        var html = Convert(Paragraph("SedgeOne", sides, shading: Fill), Paragraph("SedgeTwo", sides, shading: Fill));
        var box = BoxOf(html, "SedgeOne");

        // No paragraph reaches into the band; above and below, the fill still reaches the borders.
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
