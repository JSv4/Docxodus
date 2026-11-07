using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class HtmlPageBreakTests
{
    private const string Break = "<w:br w:type=\"page\"/>";

    private static XElement Render(string content, PaginationMode mode = PaginationMode.Paginated, string flag = "")
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}"><w:body>
                  <w:p><w:r><w:t>First</w:t></w:r></w:p>{content}
                  <w:p><w:r><w:t>Last</w:t></w:r></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"""
                <w:styles xmlns:w="{W.w}"><w:docDefaults><w:rPrDefault><w:rPr>
                  <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/>
                </w:rPr></w:rPrDefault></w:docDefaults></w:styles>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"><w:compat>{flag}</w:compat></w:settings>"));
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("break.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { RenderPagination = mode, FabricateCssClasses = false, StampCanonicalSourceAnchors = true, RenderTrackedChanges = true });
    }

    [Theory]
    [InlineData("", "")]
    [InlineData("Before", "")]
    [InlineData("", "After")]
    [InlineData("Before", "After")]
    public void HardPageBreaksAreSiblingBlocksWithoutSyntheticParagraphLines(string before, string after)
    {
        var html = Render($"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{before}</w:t>{Break}<w:t>{after}</w:t></w:r></w:p>");
        var marker = html.Descendants().Single(e => (string?)e.Attribute("data-page-break") == "true");
        Assert.DoesNotContain(marker.Ancestors(), e => e.Name.LocalName is "p" or "span");
        Assert.EndsWith("></div>", marker.ToString(SaveOptions.DisableFormatting));
        var paragraphs = html.Descendants().Where(e => e.Name.LocalName == "p").ToArray();
        Assert.Equal(new[] { "First", before, after, "Last" }.Where(s => s.Length > 0), paragraphs.Select(p => p.Value).Where(s => s.Length > 0));
        Assert.All(paragraphs.Where(p => p.Value.Length == 0), p =>
            Assert.Equal("true", (string?)p.ElementsAfterSelf().First().Attribute("data-page-break")));
        var fragments = paragraphs.Where(p => p.Value is "Before" or "After").ToArray();
        Assert.All(fragments, p => Assert.NotNull(p.Attribute("data-source-anchor-id")));
        Assert.All(fragments, p => Assert.Contains("font-weight: bold", p.ToString()));
        if (fragments.Length == 2)
            Assert.Equal(fragments[0].Attribute("data-source-anchor-id")?.Value, fragments[1].Attribute("data-source-anchor-id")?.Value);
    }

    [Fact]
    public void ConsecutivePageBreaksRetainEveryExplicitBoundary()
    {
        var html = Render($"<w:p><w:r>{Break}{Break}</w:r></w:p>");
        var markers = html.Descendants().Where(e => e.Attribute("data-page-break") != null).ToArray();
        Assert.Equal(2, markers.Length);
        Assert.All(markers, m => Assert.DoesNotContain(m.Ancestors(), e => e.Name.LocalName == "p"));
        Assert.Equal(new[] { "First", "Last" }, html.Descendants().Where(e => e.Name.LocalName == "p" && e.Value.Length > 0).Select(p => p.Value));
    }

    [Theory]
    [InlineData("<w:p/>")]
    [InlineData("<w:p><w:r><w:br/></w:r></w:p>")]
    public void AuthoredBlankParagraphsAndLineBreaksKeepTheirLine(string content)
    {
        var html = Render(content);
        Assert.Equal(3, html.Descendants().Count(e => e.Name.LocalName == "p"));
        Assert.DoesNotContain(html.Descendants(), e => e.Attribute("data-page-break") != null);
    }

    [Fact]
    public void FlowModeRetainsItsInlineBreak()
    {
        var html = Render($"<w:p><w:r><w:t>Before</w:t>{Break}<w:t>After</w:t></w:r></w:p>", PaginationMode.None);
        Assert.Single(html.Descendants(Xhtml.br));
        Assert.Contains(html.Descendants(Xhtml.p), p => p.Value.Contains("Before") && p.Value.Contains("After"));
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("<w:splitPgBreakAndParaMark w:val=\"0\"/>", false)]
    [InlineData("<w:splitPgBreakAndParaMark/>", true)]
    public void TerminalBreakMarkSettingIgnoresEmptyLinkAndBookmarkShells(string flag, bool blankLine)
    {
        foreach (var content in new[]
        {
            $"<w:r><w:t>Before</w:t>{Break}</w:r>",
            $"<w:hyperlink w:anchor=\"Target\"><w:r><w:t>Before</w:t>{Break}</w:r></w:hyperlink>",
            $"<w:r><w:t>Before</w:t>{Break}</w:r><w:bookmarkStart w:id=\"1\" w:name=\"Target\"/><w:bookmarkEnd w:id=\"1\"/>",
            $"<w:ins w:id=\"2\" w:author=\"Reviewer\"><w:r><w:t>Before</w:t>{Break}</w:r></w:ins>",
        })
        {
            var html = Render($"<w:p><w:pPr><w:spacing w:after=\"240\"/></w:pPr>{content}</w:p>", flag: flag);
            var marker = html.Descendants().Single(e => e.Attribute("data-page-break") != null);
            var next = marker.ElementsAfterSelf().First();
            Assert.Equal(blankLine ? "" : "Last", next.Value);
            Assert.Equal(blankLine, next.Descendants(Xhtml.br).Any());
            if (content.Contains("bookmarkStart"))
            {
                var target = Assert.Single(html.Descendants(), e => (string?)e.Attribute("id") == "Target");
                Assert.DoesNotContain(target.Ancestors(), e => e.Attribute("data-page-break") != null);
            }
        }
    }

    [Fact]
    public void MixedPageAndColumnBreaksKeepTheirOrderAndCanonicalSource()
    {
        var html = Render($"<w:p><w:r><w:t>North</w:t>{Break}<w:t>East</w:t><w:br w:type=\"column\"/><w:t>South</w:t></w:r></w:p>");
        var first = html.Descendants(Xhtml.p).Single(p => p.Value == "North");
        var east = html.Descendants(Xhtml.p).Single(p => p.Value == "East");
        var south = html.Descendants(Xhtml.p).Single(p => p.Value == "South");
        Assert.Equal("true", (string?)first.ElementsAfterSelf().First().Attribute("data-page-break"));
        Assert.Equal("true", (string?)south.Attribute("data-column-break-before"));
        Assert.Equal(first.Attribute("data-source-anchor-id")?.Value, east.Attribute("data-source-anchor-id")?.Value);
        Assert.Equal(first.Attribute("data-source-anchor-id")?.Value, south.Attribute("data-source-anchor-id")?.Value);
    }

    [Fact]
    public void TextboxBreakStaysInItsOwnStoryWhenTheHostAlsoBreaks()
    {
        var html = Render($"""
            <w:p><w:r><w:t>Host</w:t><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"
              id="Box" style="width:120pt;height:60pt"><v:textbox><w:txbxContent>
                <w:p><w:r><w:t>BoxStart</w:t>{Break}<w:t>BoxEnd</w:t></w:r></w:p>
              </w:txbxContent></v:textbox></v:shape></w:pict>{Break}<w:t>Tail</w:t></w:r></w:p>
            """);
        var markers = html.Descendants().Where(e => e.Attribute("data-page-break") != null).ToArray();
        Assert.Equal(2, markers.Length);
        Assert.Single(markers, e => e.Ancestors().Any(a => a.Attribute("data-docx-textbox") != null));
        Assert.Contains(html.Descendants(Xhtml.p), p => p.Value == "Tail");
    }
}
