// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="WmlToHtmlConverterSettings.SemanticLists"/>: Word lists rendered as
/// <c>ol</c>/<c>ul</c>/<c>li</c> instead of one <c>p</c> per item, with the marker drawn by CSS
/// where CSS can draw it and kept as the computed marker span where it cannot (issue #895).
/// </summary>
public class WmlSemanticListsTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // numId 1: decimal / lowerLetter / lowerRoman.   numId 2: Symbol U+F0B7 / Courier New "o".
    // numId 3: "%1." over "%1.%2." (outline).       numId 4: "(%1)".
    // numId 5: numId 1's definition restarted at 5. numId 6: a second instance of numId 1's.
    // numId 7: Wingdings "l" (filled circle) / "ü" (check mark) / U+F0A7 (small square).
    // numId 8: "%1." with no hanging indent.       numId 9: "%1." followed by a space, not a tab.
    private const int Numbered = 1;
    private const int Bullets = 2;
    private const int Outline = 3;
    private const int Parenthesized = 4;
    private const int StartsAtFive = 5;
    private const int NumberedAgain = 6;
    private const int WingdingsBullets = 7;
    private const int NoHang = 8;
    private const int SpaceAfterMarker = 9;

    private static string Lvl(int ilvl, string numFmt, string lvlText, int left, string rPr = "",
        int hanging = 360, string suff = "tab") =>
        $"<w:lvl w:ilvl=\"{ilvl}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{numFmt}\"/>" +
        $"<w:suff w:val=\"{suff}\"/><w:lvlText w:val=\"{lvlText}\"/><w:lvlJc w:val=\"left\"/>" +
        $"<w:pPr><w:ind w:left=\"{left}\" w:hanging=\"{hanging}\"/></w:pPr>{rPr}</w:lvl>";

    private static string Font(string name) =>
        $"<w:rPr><w:rFonts w:ascii=\"{name}\" w:hAnsi=\"{name}\" w:hint=\"default\"/></w:rPr>";

    private static readonly string NumberingXml =
        $"<w:numbering xmlns:w=\"{WNs}\">" +
        "<w:abstractNum w:abstractNumId=\"0\">" +
        Lvl(0, "decimal", "%1.", 720) + Lvl(1, "lowerLetter", "%2.", 1440) + Lvl(2, "lowerRoman", "%3.", 2160) +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"1\">" +
        Lvl(0, "bullet", "\uF0B7", 720, Font("Symbol")) + Lvl(1, "bullet", "o", 1440, Font("Courier New")) +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"2\">" +
        Lvl(0, "decimal", "%1.", 720) + Lvl(1, "decimal", "%1.%2.", 1440) +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"3\">" +
        Lvl(0, "decimal", "(%1)", 720) +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"4\">" +
        Lvl(0, "bullet", "l", 720, Font("Wingdings")) + Lvl(1, "bullet", "\u00FC", 1440, Font("Wingdings")) +
        Lvl(2, "bullet", "\uF0A7", 2160, Font("Wingdings")) +
        "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"5\">" + Lvl(0, "decimal", "%1.", 0, hanging: 0) + "</w:abstractNum>" +
        "<w:abstractNum w:abstractNumId=\"6\">" + Lvl(0, "decimal", "%1.", 720, suff: "space") + "</w:abstractNum>" +
        "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>" +
        "<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>" +
        "<w:num w:numId=\"3\"><w:abstractNumId w:val=\"2\"/></w:num>" +
        "<w:num w:numId=\"4\"><w:abstractNumId w:val=\"3\"/></w:num>" +
        "<w:num w:numId=\"5\"><w:abstractNumId w:val=\"0\"/>" +
        "<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"5\"/></w:lvlOverride></w:num>" +
        "<w:num w:numId=\"6\"><w:abstractNumId w:val=\"0\"/>" +
        "<w:lvlOverride w:ilvl=\"0\"><w:startOverride w:val=\"1\"/></w:lvlOverride></w:num>" +
        "<w:num w:numId=\"7\"><w:abstractNumId w:val=\"4\"/></w:num>" +
        "<w:num w:numId=\"8\"><w:abstractNumId w:val=\"5\"/></w:num>" +
        "<w:num w:numId=\"9\"><w:abstractNumId w:val=\"6\"/></w:num>" +
        "</w:numbering>";

    private static readonly string StylesXml =
        $"<w:styles xmlns:w=\"{WNs}\">" +
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/>" +
        "<w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"Spaced\"><w:name w:val=\"Spaced\"/>" +
        "<w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"160\"/></w:pPr></w:style>" +
        "</w:styles>";

    private static byte[] BuildDoc(params string[] body)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{string.Concat(body)}</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(), StylesXml);
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{WNs}\"/>");
            Write(main.AddNewPart<NumberingDefinitionsPart>(), NumberingXml);
        }
        return ms.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var s = part.GetStream(FileMode.Create);
        using var w = new StreamWriter(s);
        w.Write(xml);
    }

    private static string Item(string text, int numId, int ilvl = 0, string pPr = "") =>
        $"<w:p><w:pPr>{pPr}<w:numPr><w:ilvl w:val=\"{ilvl}\"/><w:numId w:val=\"{numId}\"/></w:numPr></w:pPr>" +
        $"<w:r><w:t>{text}</w:t></w:r></w:p>";

    private static string Para(string text, string pPr = "") =>
        $"<w:p><w:pPr>{pPr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static XElement Convert(byte[] docx, bool semanticLists = true,
        PaginationMode pagination = PaginationMode.None) =>
        WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument("lists.docx", docx),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                SemanticLists = semanticLists,
                RenderPagination = pagination,
            });

    private static List<XElement> All(XElement html, string localName) =>
        html.Descendants().Where(e => e.Name.LocalName == localName).ToList();

    private static List<XElement> Children(XElement element, string localName) =>
        element.Elements().Where(e => e.Name.LocalName == localName).ToList();

    private static Dictionary<string, string> StyleOf(XElement element) =>
        ((string?)element.Attribute("style") ?? string.Empty)
            .Split(';', System.StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2))
            .Where(kv => kv.Length == 2)
            .ToDictionary(kv => kv[0].Trim(), kv => kv[1].Trim());

    private static bool IsMarker(XElement e) => e.Attribute("data-list-marker") != null;

    /// <summary>Text of <paramref name="block"/> outside any marker span and any nested list.</summary>
    private static string OwnText(XElement block) =>
        string.Concat(block.DescendantNodes().OfType<XText>()
            .Where(t => t is not XEntity)
            .Where(t => !t.Ancestors().TakeWhile(a => a != block)
                .Any(a => a.Name.LocalName is "ol" or "ul" || IsMarker(a)))
            .Select(t => t.Value));

    /// <summary>Text of <paramref name="block"/>'s own marker spans (not a nested list's). An
    /// entity reads as a space, so the bidi marks around an RTL marker trim away.</summary>
    private static string MarkerText(XElement block) =>
        string.Concat(block.Elements().Where(IsMarker).DescendantNodes().OfType<XText>()
            .Select(t => t is XEntity ? " " : t.Value)).Trim();

    [Fact]
    public void Off_ByDefault_EachItemStaysAParagraph()
    {
        var html = Convert(BuildDoc(Item("One", Numbered), Item("Two", Numbered)), semanticLists: false);

        Assert.Empty(All(html, "ol"));
        Assert.Empty(All(html, "li"));
        Assert.False(new WmlToHtmlConverterSettings().SemanticLists);
    }

    [Fact]
    public void ConsecutiveItems_BecomeOneOrderedList_WithCssMarkers()
    {
        var html = Convert(BuildDoc(Item("One", Numbered), Item("Two", Numbered), Item("Three", Numbered)));

        var ol = Assert.Single(All(html, "ol"));
        Assert.Equal("decimal", StyleOf(ol)["list-style-type"]);
        Assert.Null(ol.Attribute("start"));
        var items = Children(ol, "li");
        Assert.Equal(new[] { "One", "Two", "Three" }, items.Select(OwnText));
        // CSS draws the number, so the generated marker span is gone.
        Assert.All(items, li => Assert.DoesNotContain(li.Descendants(), IsMarker));
        Assert.Empty(All(html, "p"));
    }

    [Fact]
    public void DeeperLevel_NestsInsideThePrecedingItem()
    {
        var html = Convert(BuildDoc(
            Item("One", Numbered), Item("One-a", Numbered, 1), Item("One-b", Numbered, 1), Item("Two", Numbered)));

        var outer = Assert.Single(html.Descendants(), e => e.Name.LocalName == "ol" &&
            !e.Ancestors().Any(a => a.Name.LocalName == "ol"));
        var items = Children(outer, "li");
        Assert.Equal(new[] { "One", "Two" }, items.Select(OwnText));
        var nested = Assert.Single(Children(items[0], "ol"));
        Assert.Equal("lower-alpha", StyleOf(nested)["list-style-type"]);
        Assert.Equal(new[] { "One-a", "One-b" }, Children(nested, "li").Select(OwnText));
        Assert.Empty(Children(items[1], "ol"));
    }

    [Fact]
    public void Bullets_BecomeAnUnorderedList()
    {
        var html = Convert(BuildDoc(Item("Dot", Bullets), Item("Circle", Bullets, 1), Item("Dot again", Bullets)));

        var ul = Assert.Single(html.Descendants(), e => e.Name.LocalName == "ul" &&
            !e.Ancestors().Any(a => a.Name.LocalName == "ul"));
        Assert.Equal("disc", StyleOf(ul)["list-style-type"]);
        var nested = Assert.Single(All(ul, "ul"));
        Assert.Equal("circle", StyleOf(nested)["list-style-type"]);
        Assert.Empty(All(html, "ol"));
        Assert.Null(ul.Attribute("start"));
    }

    [Fact]
    public void SymbolFontBullets_UseCss_OnlyForAShapeCssHas()
    {
        var html = Convert(BuildDoc(
            Item("Circle", WingdingsBullets), Item("Check", WingdingsBullets, 1), Item("Square", WingdingsBullets, 2)));

        // Wingdings "l" is a filled circle, not the letter.
        var outer = html.Descendants().First(e => e.Name.LocalName == "ul");
        Assert.Equal("disc", StyleOf(outer)["list-style-type"]);
        // CSS has no check mark, and as a string marker it would draw "\u00FC": the marker stays.
        var check = Assert.Single(Children(Children(outer, "li")[0], "ul"));
        Assert.Equal("none", StyleOf(check)["list-style-type"]);
        var checkItem = Assert.Single(Children(check, "li"));
        Assert.Equal("\u00FC", MarkerText(checkItem));
        // U+F0A7 maps to U+25AA, which CSS draws as a square.
        var square = Assert.Single(Children(checkItem, "ul"));
        Assert.Equal("square", StyleOf(square)["list-style-type"]);
    }

    [Fact]
    public void MarkerCssCannotDraw_KeepsTheComputedMarker_UnderListStyleNone()
    {
        var html = Convert(BuildDoc(Item("First", Parenthesized), Item("Second", Parenthesized)));

        var ol = Assert.Single(All(html, "ol"));
        Assert.Equal("none", StyleOf(ol)["list-style-type"]);
        var items = Children(ol, "li");
        Assert.Equal(new[] { "(1)", "(2)" }, items.Select(MarkerText));
        Assert.Equal(new[] { "First", "Second" }, items.Select(OwnText));
    }

    [Fact]
    public void MultiLevelNumbers_KeepTheirMarkers_WhileTheOuterLevelUsesCss()
    {
        var html = Convert(BuildDoc(Item("Part", Outline), Item("Clause", Outline, 1), Item("Clause", Outline, 1)));

        var outer = html.Descendants().First(e => e.Name.LocalName == "ol");
        Assert.Equal("decimal", StyleOf(outer)["list-style-type"]);
        var nested = Assert.Single(All(outer, "ol"));
        Assert.Equal("none", StyleOf(nested)["list-style-type"]);
        Assert.Equal(new[] { "1.1.", "1.2." }, Children(nested, "li").Select(MarkerText));
    }

    [Fact]
    public void InterruptedList_ResumesWithStart()
    {
        var html = Convert(BuildDoc(Item("One", Numbered), Para("Interruption."), Item("Two", Numbered)));

        var lists = All(html, "ol");
        Assert.Equal(2, lists.Count);
        Assert.Null(lists[0].Attribute("start"));
        Assert.Equal("2", (string?)lists[1].Attribute("start"));
        var between = Assert.Single(All(html, "p"));
        Assert.Equal("Interruption.", OwnText(between));
        Assert.True(lists[0].IsBefore(between) && between.IsBefore(lists[1]));
    }

    [Fact]
    public void StartOverride_SetsStart()
    {
        var html = Convert(BuildDoc(Item("Five", StartsAtFive), Item("Six", StartsAtFive)));

        var ol = Assert.Single(All(html, "ol"));
        Assert.Equal("5", (string?)ol.Attribute("start"));
        Assert.All(Children(ol, "li"), li => Assert.Null(li.Attribute("value")));
    }

    [Fact]
    public void AdjacentItemsOfDifferentLists_StaySeparateLists()
    {
        var html = Convert(BuildDoc(Item("A1", Numbered), Item("A2", Numbered), Item("B1", NumberedAgain)));

        var lists = All(html, "ol");
        Assert.Equal(2, lists.Count);
        Assert.Equal(new[] { "A1", "A2" }, Children(lists[0], "li").Select(OwnText));
        Assert.Equal(new[] { "B1" }, Children(lists[1], "li").Select(OwnText));
    }

    [Fact]
    public void LevelIndent_MovesFromTheItemOntoItsList()
    {
        var html = Convert(BuildDoc(Item("One", Numbered), Item("One-a", Numbered, 1)));

        var outer = html.Descendants().First(e => e.Name.LocalName == "ol");
        var outerStyle = StyleOf(outer);
        Assert.Equal("0.50in", outerStyle["padding-inline-start"]);
        Assert.Equal("0", outerStyle["margin-top"]);
        Assert.Equal("0", outerStyle["margin-bottom"]);

        var item = Children(outer, "li")[0];
        Assert.Equal("0", StyleOf(item)["margin-left"]);
        Assert.Equal("0", StyleOf(item)["text-indent"]);

        // Level 1 sits at 1.00in: half an inch inside the level-0 item it nests in.
        var nested = Assert.Single(Children(item, "ol"));
        Assert.Equal("0.50in", StyleOf(nested)["padding-inline-start"]);
    }

    [Theory]
    [InlineData(NoHang)] // the number sits inline before a first-line tab: no hanging indent
    [InlineData(SpaceAfterMarker)] // the text follows the number, not the hanging-indent edge
    public void MarkerThatDoesNotFillTheHangingIndent_KeepsItsSpan(int numId)
    {
        // CSS draws an outside marker in the list's padding and starts every line of the item at
        // the content edge, which would move this first line's text.
        var html = Convert(BuildDoc(Item("First", numId), Item("Second", numId)));

        var ol = Assert.Single(All(html, "ol"));
        Assert.Equal("none", StyleOf(ol)["list-style-type"]);
        Assert.Equal(new[] { "1.", "2." }, Children(ol, "li").Select(MarkerText));
    }

    [Fact]
    public void NestedList_KeepsTheSpaceAfterTheItemItNestsIn()
    {
        // 8pt after each item. Flat, it separates "Dot" from "One" below it; nested, the host
        // item's bottom margin would land below the whole nested list instead. (Within one
        // numbering definition the converter keeps only the last item's space after, so the
        // nested item comes from another list, as a numbered step under a bullet often does.)
        const string spaced = "<w:pStyle w:val=\"Spaced\"/>";
        var body = new[]
        {
            Item("Dot", Bullets, pPr: spaced),
            Item("One", Numbered, 1, spaced),
            Item("Dot again", Bullets, pPr: spaced),
        };
        var flat = All(Convert(BuildDoc(body), semanticLists: false), "p");
        var html = Convert(BuildDoc(body));

        var host = Children(html.Descendants().First(e => e.Name.LocalName == "ul"), "li")[0];
        var nested = Assert.Single(Children(host, "ol"));
        Assert.Equal("8pt", StyleOf(flat[0])["margin-bottom"]);
        Assert.Equal(StyleOf(flat[0])["margin-bottom"], StyleOf(nested)["margin-top"]);
        Assert.Equal("0", StyleOf(host)["margin-bottom"]);
        Assert.Equal(StyleOf(flat[1])["margin-bottom"], StyleOf(Children(nested, "li")[0])["margin-bottom"]);
    }

    [Fact]
    public void KeptMarkers_KeepTheHangingIndent()
    {
        var html = Convert(BuildDoc(Item("First", Parenthesized)));

        var li = Assert.Single(All(html, "li"));
        Assert.Equal("-0.25in", StyleOf(li)["text-indent"]);
        Assert.Equal("0", StyleOf(li)["margin-left"]);
    }

    [Fact]
    public void NumberedHeading_StaysAHeading_AndEndsTheList()
    {
        var html = Convert(BuildDoc(
            Item("One", Numbered),
            Item("Heading", Numbered, pPr: "<w:pStyle w:val=\"Heading1\"/>"),
            Item("Three", Numbered)));

        var h1 = Assert.Single(All(html, "h1"));
        Assert.Equal("Heading", OwnText(h1));
        Assert.Equal("2.", MarkerText(h1));
        Assert.DoesNotContain(h1.Ancestors(), a => a.Name.LocalName is "ol" or "li");
        var lists = All(html, "ol");
        Assert.Equal(2, lists.Count);
        Assert.Equal("3", (string?)lists[1].Attribute("start"));
    }

    [Fact]
    public void ListInATableCell_StaysInThatCell()
    {
        var html = Convert(BuildDoc(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" +
            Item("Cell one", Numbered) + Item("Cell two", Numbered) +
            "</w:tc></w:tr></w:tbl>",
            Para("After the table.")));

        var ol = Assert.Single(All(html, "ol"));
        Assert.Contains(ol.Ancestors(), a => a.Name.LocalName == "td");
        Assert.Equal(new[] { "Cell one", "Cell two" }, Children(ol, "li").Select(OwnText));
    }

    [Fact]
    public void PaginatedOutput_KeepsParagraphs()
    {
        var html = Convert(BuildDoc(Item("One", Numbered), Item("Two", Numbered)),
            pagination: PaginationMode.Paginated);

        Assert.Empty(All(html, "li"));
        Assert.Equal(2, All(html, "p").Count(p => p.Elements().Any(IsMarker)));
    }

    [Fact]
    public void DocumentWithoutLists_RendersTheSameEitherWay()
    {
        var docx = BuildDoc(Para("Just a paragraph."), Para("And another."));

        Assert.Equal(
            Convert(docx, semanticLists: false).ToString(SaveOptions.DisableFormatting),
            Convert(docx).ToString(SaveOptions.DisableFormatting));
    }

    [Fact]
    public void HtmlConversionOps_PassesTheOption_AndAnchorsStampEachItem()
    {
        var docx = BuildDoc(Item("One", Numbered), Item("Two", Numbered));

        var plain = HtmlConversionOps.ConvertToHtml(docx, new HtmlConversionOptions { StampAnchors = true });
        var lists = HtmlConversionOps.ConvertToHtml(docx,
            new HtmlConversionOptions { StampAnchors = true, SemanticLists = true });

        Assert.DoesNotMatch("<li[ >]", plain);
        Assert.Single(Regex.Matches(lists, "<ol[ >]"));
        Assert.Equal(2, Regex.Matches(lists, "<li [^>]*data-anchor=\"").Count);
    }

    // Real documents: regrouping must not lose, duplicate or reorder content, every list
    // paragraph must become exactly one item, and a CSS-numbered list must count to the same
    // numbers Word shows.
    public static IEnumerable<object[]> ListCorpus() => new[]
    {
        "DB012-Lists-With-Different-Numberings.docx",
        "DB013b-Blue-List-English.docx",
        "DB013b-Orange-List-Danish.docx",
        "HC010-Test-05.docx",
        "HC012-Test-07.docx",
        "DB006-Source1.docx",
        "DB006-Source3.docx",
        "DB007-Spec.docx",
    }.Select(f => new object[] { f });

    private static XElement ConvertFixture(string fileName, bool semanticLists) =>
        WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument(Path.Combine("..", "..", "..", "..", "TestFiles", fileName)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, SemanticLists = semanticLists });

    private static string TextWithoutMarkers(XElement html)
    {
        var body = html.Descendants().First(e => e.Name.LocalName == "body");
        return string.Concat(body.DescendantNodes().OfType<XText>()
            .Where(t => !t.Ancestors().Any(IsMarker))
            .Select(t => t.Value));
    }

    private static int Blocks(XElement html) =>
        html.Descendants().Count(e => e.Name.LocalName is "p" or "li" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6");

    [Theory]
    [MemberData(nameof(ListCorpus))]
    public void Corpus_KeepsEveryParagraph_InOrder(string fileName)
    {
        var off = ConvertFixture(fileName, semanticLists: false);
        var on = ConvertFixture(fileName, semanticLists: true);

        Assert.Equal(TextWithoutMarkers(off), TextWithoutMarkers(on));
        // Every paragraph is still exactly one block: a <p> or heading, or now an <li>.
        Assert.Equal(Blocks(off), Blocks(on));
        Assert.DoesNotContain(All(on, "p"), p => p.Elements().Any(IsMarker));
    }

    [Theory]
    [MemberData(nameof(ListCorpus))]
    public void Corpus_CssNumberedLists_CountLikeWord(string fileName)
    {
        // The markers Word shows, in document order, for every list paragraph.
        var wordMarkers = All(ConvertFixture(fileName, semanticLists: false), "p")
            .Where(p => p.Elements().Any(IsMarker))
            .Select(p => (Text: OwnText(p), Marker: MarkerText(p)))
            .ToList();
        var markerByText = wordMarkers
            .GroupBy(m => m.Text)
            .Where(g => g.Count() == 1)
            .ToDictionary(g => g.Key, g => g.Single().Marker);

        var on = ConvertFixture(fileName, semanticLists: true);
        foreach (var ol in All(on, "ol").Where(l => StyleOf(l).GetValueOrDefault("list-style-type") == "decimal"))
        {
            var number = (int?)ol.Attribute("start") ?? 1;
            foreach (var li in Children(ol, "li"))
            {
                number = (int?)li.Attribute("value") ?? number;
                if (markerByText.TryGetValue(OwnText(li), out var marker))
                    Assert.Equal(marker, number + ".");
                number++;
            }
        }
    }
}
