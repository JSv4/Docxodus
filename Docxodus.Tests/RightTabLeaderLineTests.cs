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
    private static byte[] Fixture(string variant, int? rightIndent, int rightStop = 9360)
    {
        var tabRPr = variant switch
        {
            "underline" => "<w:rPr><w:u w:val=\"single\"/></w:rPr>",
            "strike" => "<w:rPr><w:strike/></w:rPr>",
            _ => "",
        };
        var right = rightIndent is { } r ? $" w:right=\"{r}\"" : "";
        var body = string.Concat(Entries.Select((e, i) =>
            "<w:p><w:pPr><w:tabs><w:tab w:val=\"left\" w:pos=\"1296\"/>" +
            $"<w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"{rightStop}\"/></w:tabs>" +
            $"<w:ind w:left=\"1296\"{right} w:hanging=\"648\"/></w:pPr>" +
            (variant == "bookmark" ? $"<w:bookmarkStart w:id=\"{i}\" w:name=\"_Toc{i}\"/><w:bookmarkEnd w:id=\"{i}\"/>" : "") +
            $"<w:r><w:t>{e.Label}</w:t></w:r><w:r>{tabRPr}<w:tab/></w:r><w:r><w:t>{e.Title}</w:t></w:r>" +
            (variant == "hyphen" ? "<w:r><w:noBreakHyphen/><w:t>West</w:t></w:r>" : "") +
            $"<w:r>{tabRPr}<w:tab/><w:t>{e.Number}</w:t></w:r></w:p>"));

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

    private static XElement[] RightTabSegments(byte[] docx)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("leader-lines.docx", docx),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
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
        // The line ends at 9360 - 3360 = 6000 and the stop is at 6040. The longest title is estimated
        // to end near 5920 twips, so its number is estimated to run past the stop. Every segment
        // gives back only the 40 twips (2pt) up to the stop: text estimated to run past a stop gets
        // no extra room, the same as at a stop inside the line.
        foreach (var segment in RightTabSegments(Fixture("plain", rightIndent: 3360, rightStop: 6040)))
            Assert.Equal("-2pt", Style(segment, "margin-right"));
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
