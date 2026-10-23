using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Leader-and-page-number lines (a left tab after a label, a right dot-leader tab before the
/// number) must render the right tab as a browser-sized flex segment that ends on the stop, even
/// when the right indent ends before the stop or the entry carries a bookmark, a non-breaking
/// hyphen, or underline/strike on its tab runs (issue #1275).
/// </summary>
public class RightTabLeaderLineTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly (string Label, string Title, string Number)[] Entries =
    {
        ("2.1.", "Quarry Access Roads", "3"),
        ("2.2.", "Spillway", "12"),
        ("2.3.", "Culvert Inspection and Annual Load Ratings", "27"),
        ("2.4.", "Riprap", "140"),
    };

    /// <summary>Four entries on a Letter page with one-inch margins (a 9360-twip column): a left
    /// stop at 1296 and a right dot-leader stop at <paramref name="rightStop"/>.</summary>
    private static byte[] Fixture(string variant, int? rightIndent, int rightStop = 9360, bool fieldLink = false)
    {
        var tabRPr = variant switch
        {
            "underline" or "toc-link-underline" => "<w:rPr><w:u w:val=\"single\"/></w:rPr>",
            "double-underline" => "<w:rPr><w:u w:val=\"double\" w:color=\"FF0000\"/></w:rPr>",
            "words-underline" => "<w:rPr><w:u w:val=\"words\"/></w:rPr>",
            "strike" => "<w:rPr><w:strike/></w:rPr>",
            "underline-strike" => "<w:rPr><w:u w:val=\"single\"/><w:strike/></w:rPr>",
            _ => "",
        };
        var right = rightIndent is { } r ? $" w:right=\"{r}\"" : "";
        var body = string.Concat(Entries.Select((e, i) =>
            "<w:p><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1296\"/>" +
            $"<w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"{rightStop}\"/></w:tabs>" +
            $"<w:ind w:left=\"1296\"{right} w:hanging=\"648\"/></w:pPr>" +
            // A cached TOC: the field opens in the first entry and closes in the last.
            (variant == "toc-link-underline" && i == 0
                ? "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
                : "") +
            (variant is "linked-bookmark" or "toc-link-underline" ? fieldLink
                ? $"<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> HYPERLINK \"https://example.com/entry{i}\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
                : $"<w:hyperlink w:anchor=\"_Ref{i}\">" : "") +
            (variant is "bookmark" or "linked-bookmark" ? $"<w:bookmarkStart w:id=\"{i}\" w:name=\"_Toc{i}\"/><w:bookmarkEnd w:id=\"{i}\"/>" : "") +
            $"<w:r><w:t>{e.Label}</w:t></w:r><w:r>{tabRPr}<w:tab/></w:r><w:r><w:t>{e.Title}</w:t></w:r>" +
            (variant == "hyphen" ? "<w:r><w:noBreakHyphen/><w:t>West</w:t></w:r>" : "") +
            $"<w:r>{tabRPr}<w:tab/><w:t>{e.Number}</w:t></w:r>" +
            (variant is "linked-bookmark" or "toc-link-underline" ? fieldLink
                ? "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" : "</w:hyperlink>" : "") +
            (variant == "toc-link-underline" && i == Entries.Length - 1 ? "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" : "") +
            "</w:p>"));

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            Write(main, $"<w:document xmlns:w=\"{WNs}\"><w:body>{body}" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>" +
                "</w:body></w:document>");
            Write(main.AddNewPart<StyleDefinitionsPart>(),
                $"<w:styles xmlns:w=\"{WNs}\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/>" +
                "<w:sz w:val=\"24\"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings xmlns:w=\"{WNs}\"/>");
        }
        return ms.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create);
        var bytes = Encoding.UTF8.GetBytes(xml);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static XElement Html(byte[] docx) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("leader-lines.docx", docx),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });

    private static XElement[] RightTabSegments(byte[] docx)
    {
        var html = Html(docx);
        var tabs = html.Descendants().Where(e => (string?)e.Attribute("data-docx-tab") == "right").ToArray();
        Assert.Equal(Entries.Length, tabs.Length);
        // A flexible tab fills its wrapper; the fixed fallback carries a server-measured width.
        Assert.All(tabs, tab => Assert.Contains("width: 100%", (string?)tab.Attribute("style") ?? ""));
        return tabs.Select(tab => tab.Ancestors().First(e => Style(e, "display") == "inline-flex")).ToArray();
    }

    private static string? Style(XElement element, string property) =>
        ((string?)element.Attribute("style") ?? "").Split(';').Select(d => d.Split(':', 2))
            .Where(d => d.Length == 2 && d[0].Trim() == property).Select(d => d[1].Trim()).SingleOrDefault();

    [Theory]
    [InlineData("plain")]
    [InlineData("bookmark")]
    [InlineData("linked-bookmark")]
    [InlineData("hyphen")]
    [InlineData("underline")]
    [InlineData("strike")]
    public void RightTab_IsAFlexSegmentEndingOnTheStop(string variant)
    {
        foreach (var segment in RightTabSegments(Fixture(variant, rightIndent: null)))
        {
            // 1296 twips to 9360 twips: the segment ends exactly on the right stop.
            Assert.Equal("403.2pt", Style(segment, "width"));
            Assert.Null(Style(segment, "margin-right"));
        }
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("bookmark")]
    [InlineData("underline")]
    public void RightTabSegment_OverhangsTheRightIndentToReachTheStop(string variant)
    {
        // The line ends 864 twips (43.2pt) before the stop; the segment keeps its full width and
        // gives that overhang back to the line so it fits beside the label.
        foreach (var segment in RightTabSegments(Fixture(variant, rightIndent: 864)))
        {
            Assert.Equal("403.2pt", Style(segment, "width"));
            Assert.Equal("-43.2pt", Style(segment, "margin-right"));
        }
    }

    [Theory]
    // A stop at 8000 lies inside a line that ends at 9360 - 864 = 8496: no overhang.
    [InlineData(864, 8000, null)]
    // A stop past the column edge overhangs into the page margin, as LibreOffice lays it out:
    // 10000 - 9360 = 640 twips (32pt), or 10000 - 8496 = 1504 twips (75.2pt) with the indent.
    [InlineData(null, 10000, "-32pt")]
    [InlineData(864, 10000, "-75.2pt")]
    public void RightTabSegment_OverhangRunsFromTheLineEndToTheStop(int? rightIndent, int rightStop, string? margin)
    {
        foreach (var segment in RightTabSegments(Fixture("plain", rightIndent, rightStop)))
            Assert.Equal(margin, Style(segment, "margin-right"));
    }

    [Fact]
    public void RightTabSegment_OverhangEndsAtTheStop()
    {
        // Where the longest entry's number is estimated to end (font metrics vary by host): the
        // stop at 9360 less the estimated leader width of its right tab.
        var textEnd = 9360 - (int)(1440 * (decimal)RightTabSegments(Fixture("plain", rightIndent: null))[2]
            .Descendants().Single(e => (string?)e.Attribute("data-docx-tab") == "right").Attribute("data-docx-tab-width")!);
        // End the line 30 twips before that and put the stop 20 twips past the line end, so that
        // number runs past the stop. Every segment gives back only the 20 twips (1pt) to the stop:
        // text estimated to run past a stop gets no extra room, the same as at a stop inside the line.
        foreach (var segment in RightTabSegments(Fixture("plain", rightIndent: 9360 - (textEnd - 30), rightStop: textEnd - 10)))
            Assert.Equal("-1pt", Style(segment, "margin-right"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BookmarkInsideALink_IsATargetThatDoesNotNestAnAnchor(bool fieldLink)
    {
        // HTML forbids <a> inside <a>: a parser closes the link at the inner anchor and leaves the
        // label, title and number outside it (issue #1285).
        var html = Html(Fixture("linked-bookmark", rightIndent: 864, fieldLink: fieldLink));
        var links = html.Descendants()
            .Where(e => e.Name.LocalName == "a").ToArray();
        Assert.Equal(Entries.Length, links.Length);
        foreach (var (link, i) in links.Select((link, i) => (link, i)))
        {
            Assert.Equal(fieldLink ? $"https://example.com/entry{i}" : $"#_Ref{i}", (string?)link.Attribute("href"));
            Assert.DoesNotContain(link.Descendants(), e => e.Name.LocalName == "a");
            var target = link.Descendants().SingleOrDefault(e => (string?)e.Attribute("id") == $"_Toc{i}");
            Assert.NotNull(target);
            Assert.Equal("span", target.Name.LocalName);
            Assert.Contains(Entries[i].Number, link.Value);
        }
    }

    [Fact]
    public void BookmarkOutsideALink_StaysAnAnchorTarget()
    {
        var html = Html(Fixture("bookmark", rightIndent: null));
        Assert.All(Enumerable.Range(0, Entries.Length), i => Assert.Equal("a",
            html.Descendants().Single(e => (string?)e.Attribute("id") == $"_Toc{i}").Name.LocalName));
    }

    /// <summary>The border-top lines drawn inside each tab box, by (top offset, border).</summary>
    private static (string? Top, string? Border)[][] TabLines(string variant) =>
        Html(Fixture(variant, rightIndent: null)).Descendants().Where(e => e.Attribute("data-docx-tab") != null)
            .Select(tab => tab.Elements().Select(line => (Style(line, "top"), Style(line, "border-top"))).ToArray())
            .ToArray();

    [Theory]
    // CSS does not paint a run's text decoration across an atomic inline box, so the tab box draws
    // the run's lines itself, at the browser's underline and line-through positions (issue #1285).
    [InlineData("underline", "0.05em", "1px solid currentColor")]
    [InlineData("double-underline", "0.05em", "3px double currentColor")]
    [InlineData("strike", "-0.3em", "1px solid currentColor")]
    public void DecoratedTab_DrawsItsRunsLineAcrossTheTab(string variant, string top, string border)
    {
        var tabs = TabLines(variant);
        Assert.Equal(2 * Entries.Length, tabs.Length);
        Assert.All(tabs, lines => Assert.Equal(new (string?, string?)[] { (top, border) }, lines));
    }

    [Fact]
    public void UnderlinedStruckTab_DrawsBothLines() =>
        Assert.All(TabLines("underline-strike"), lines => Assert.Equal(
            new (string?, string?)[] { ("0.05em", "1px solid currentColor"), ("-0.3em", "1px solid currentColor") }, lines));

    [Theory]
    // Controls: a plain tab draws nothing, and a words-only underline skips tabs as it skips spaces.
    [InlineData("plain")]
    [InlineData("words-underline")]
    // A cached TOC result drops the link underline from its tabs, as from its text.
    [InlineData("toc-link-underline")]
    public void UndecoratedTab_DrawsNoLine(string variant) =>
        Assert.All(TabLines(variant), Assert.Empty);

    [Fact]
    public void ColoredUnderline_KeepsItsColorOnTheTab()
    {
        var line = Html(Fixture("double-underline", rightIndent: null)).Descendants()
            .First(e => e.Attribute("data-docx-tab") != null).Elements().Single();
        Assert.Equal("#FF0000", Style(line, "border-top-color"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DecoratedTab_KeepsItsRunFontAndColorWithoutPrecedingText(bool leading)
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("underline", rightIndent: null));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            foreach (var paragraph in xml.Descendants(W.p))
            {
                if (leading) paragraph.Elements(W.r).First().Remove();
                var tabRun = paragraph.Elements(W.r).First(r => r.Element(W.tab) != null);
                tabRun.Element(W.rPr)!.ReplaceWith(new XElement(W.rPr,
                    new XElement(W.color, new XAttribute(W.val, "123456")),
                    new XElement(W.sz, new XAttribute(W.val, "48")),
                    new XElement(W.u, new XAttribute(W.val, "single"))));
            }
            part.PutXDocument(xml);
        }
        var tabs = Html(stream.ToArray()).Descendants().Where(e => (string?)e.Attribute("data-docx-tab") == "left");
        Assert.Equal(Entries.Length, tabs.Count());
        foreach (var tab in tabs)
        {
            string? Inherited(string name) => tab.AncestorsAndSelf().Select(e => Style(e, name)).FirstOrDefault(s => s != null);
            Assert.Equal("24pt", Inherited("font-size"));
            Assert.Equal("#123456", Inherited("color"));
        }
    }

    [Fact]
    public void AnnotationBookmark_KeepsTheRunAwarePath()
    {
        // Control: a bookmark that opens a rendered annotation range is not an inert marker.
        var docx = Fixture("plain", rightIndent: null);
        var annotated = AnnotationManager.AddAnnotation(new WmlDocument("leader-lines.docx", docx),
            new DocumentAnnotation("cedar-1", "SECTION", "Section", "#FFEB3B"),
            new AnnotationRange { SearchText = "Spillway" });
        var html = WmlToHtmlConverter.ConvertToHtml(annotated,
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderAnnotations = true });
        var spillway = html.Descendants().First(e => e.Name.LocalName == "p" && e.Value.Contains("Spillway"));
        var tab = spillway.Descendants().Single(e => (string?)e.Attribute("data-docx-tab") == "right");
        Assert.DoesNotContain("width: 100%", (string?)tab.Attribute("style") ?? "");
    }
}
