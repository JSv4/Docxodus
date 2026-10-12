using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class RenderingFidelityInheritanceTests
{
    private static XElement Render(string body, string? notes = null, bool endnote = false,
        PaginationMode mode = PaginationMode.Paginated)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}"><w:body>{body}<w:sectPr><w:cols w:num="2"/></w:sectPr></w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"""
                <w:styles xmlns:w="{W.w}"><w:docDefaults><w:rPrDefault><w:rPr>
                  <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
                </w:rPr></w:rPrDefault></w:docDefaults></w:styles>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
            if (notes != null)
            {
                var kind = endnote ? "endnote" : "footnote";
                OpenXmlPart part = endnote ? main.AddNewPart<EndnotesPart>() : main.AddNewPart<FootnotesPart>();
                part.PutXDocument(XDocument.Parse($"<w:{kind}s xmlns:w=\"{W.w}\"><w:{kind} w:id=\"1\">{notes}</w:{kind}></w:{kind}s>"));
            }
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("authored.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings
            {
                RenderPagination = mode, RenderFootnotesAndEndnotes = true,
                FabricateCssClasses = false, StampCanonicalSourceAnchors = true,
            });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NoteNumberInheritsTheNoteFont(bool endnote)
    {
        var kind = endnote ? "endnote" : "footnote";
        var html = Render($"<w:p><w:r><w:{kind}Reference w:id=\"1\"/></w:r></w:p>",
            $"<w:p><w:r><w:{kind}Ref/></w:r><w:r><w:t> Note body.</w:t></w:r></w:p>", endnote);
        var item = html.Descendants().First(e => endnote ? e.Attribute("data-endnote-number") != null : (string?)e.Attribute("class") == "footnote-number");
        Assert.Contains("'Liberation Serif'", (string?)item.Attribute("style"));
        Assert.DoesNotContain("↩", html.Value);
        if (endnote)
        {
            Assert.Equal("i", html.Descendants().Single(e => e.Attribute("data-endnote-number") != null).Value);
            Assert.Single(html.Descendants().Single(e => (string?)e.Attribute("class") == "endnotes").Elements().Where(e => e.Name.LocalName == "hr"));
        }
    }

    [Fact]
    public void LeadingColumnBreakKeepsTextInsideItsCanonicalParagraph()
    {
        var html = Render("<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:br w:type=\"column\"/><w:t>Right column.</w:t></w:r></w:p>");
        var p = html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value.Contains("Right column."));
        Assert.NotNull(p.Attribute("data-source-anchor-id"));
        Assert.Equal("true", (string?)p.Attribute("data-column-break-before"));
        Assert.Contains("break-before: column", (string?)p.Attribute("style"));
        Assert.Contains("break-inside: avoid", (string?)p.Attribute("style"));
        Assert.Empty(p.Descendants().Where(e => e.Name.LocalName == "div"));
    }

    [Fact]
    public void MidRunColumnBreakPreservesFormattingAndSourceIdentity()
    {
        var html = Render("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Left.</w:t><w:br w:type=\"column\"/><w:t>Right.</w:t></w:r></w:p>");
        var paragraphs = html.Descendants().Where(e => e.Name.LocalName == "p").ToList();
        Assert.Equal(new[] { "Left.", "Right." }, paragraphs.Select(e => e.Value));
        Assert.Equal(paragraphs[0].Attribute("data-source-anchor-id")?.Value, paragraphs[1].Attribute("data-source-anchor-id")?.Value);
        Assert.Contains("font-weight: bold", paragraphs[1].ToString());
        Assert.Equal("true", (string?)paragraphs[1].Attribute("data-column-break-before"));
    }

    [Fact]
    public void CellMarginsResolveEachSideWithCellOverridesAndExplicitZero()
    {
        var html = Render("""
            <w:tbl><w:tblPr><w:tblCellMar>
              <w:top w:w="240" w:type="dxa"/><w:left w:w="240" w:type="dxa"/>
              <w:bottom w:w="240" w:type="dxa"/><w:right w:w="240" w:type="dxa"/>
            </w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
            <w:tr><w:tc><w:tcPr><w:tcMar><w:top w:w="0" w:type="dxa"/><w:start w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>
              <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
            </w:tc></w:tr></w:tbl>
            """);
        var style = html.Descendants().Single(e => e.Name.LocalName == "td").Attribute("style")!.Value;
        Assert.Contains("padding-top: 0pt", style);
        Assert.Contains("padding-left: 6pt", style);
        Assert.Contains("padding-bottom: 12pt", style);
        Assert.Contains("padding-right: 12pt", style);
    }

    private static string NestedCell(string tail, string following = "") => $$"""
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>
          <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>
            <w:p><w:r><w:t>Inner</w:t></w:r></w:p>
          </w:tc></w:tr></w:tbl>{{tail}}{{following}}
        </w:tc></w:tr></w:tbl>
        """;

    [Theory]
    [InlineData("<w:p/>")]
    [InlineData("<w:p><w:pPr/><w:r><w:rPr/><w:t/></w:r></w:p>")]
    public void UnformattedNestedTableEndParagraphKeepsACompactSourceAnchor(string tail)
    {
        var html = Render(NestedCell(tail));
        var paragraph = html.Descendants().Single(e => e.Name.LocalName == "p" && string.IsNullOrWhiteSpace(e.Value));
        var style = (string?)paragraph.Attribute("style");
        Assert.Contains("line-height: 0.05pt", style);
        Assert.Contains("margin-top: 0", style);
        Assert.Contains("margin-bottom: 0", style);
        Assert.DoesNotContain("--docx-auto-line-spacing", style);
        Assert.NotNull(paragraph.Attribute("data-source-anchor-id"));
    }

    [Theory]
    [InlineData("<w:p><w:pPr><w:spacing w:after=\"240\"/></w:pPr></w:p>")]
    [InlineData("<w:p><w:r><w:rPr><w:sz w:val=\"40\"/></w:rPr><w:t/></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:t>Visible tail</w:t></w:r></w:p>")]
    [InlineData("<w:p><w:r><w:br/></w:r></w:p>")]
    public void IntentionalNestedTableEndParagraphRetainsItsLine(string tail)
    {
        var html = Render(NestedCell(tail));
        var paragraph = html.Descendants().Last(e => e.Name.LocalName == "p");
        Assert.DoesNotContain("line-height: 0.05pt", (string?)paragraph.Attribute("style"));
        Assert.NotNull(paragraph.Attribute("data-source-anchor-id"));
        if (tail.Contains("after=")) Assert.Contains("margin-bottom: 12pt", (string?)paragraph.Attribute("style"));
        if (tail.Contains("sz ")) Assert.Contains("font-size: 20pt", paragraph.ToString());
        if (tail.Contains("Visible tail")) Assert.Equal("Visible tail", paragraph.Value);
    }

    [Fact]
    public void BlankBodyAndNonterminalCellParagraphsKeepTheirLines()
    {
        var html = Render(NestedCell("<w:p/>", "<w:p><w:r><w:t>Later</w:t></w:r></w:p>") + "<w:p/>");
        var blanks = html.Descendants().Where(e => e.Name.LocalName == "p" && string.IsNullOrWhiteSpace(e.Value)).ToList();
        Assert.Equal(2, blanks.Count);
        Assert.All(blanks, p => Assert.DoesNotContain("line-height: 0.05pt", (string?)p.Attribute("style")));
    }
}
