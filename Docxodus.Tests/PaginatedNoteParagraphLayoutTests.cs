using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1120: paginated note paragraphs keep their authored layout. Generic web-note CSS must not
// inset or space later paragraphs, and the number shares its first paragraph's exact line.
// Issue #1130: the number is a run of the first paragraph, as in Word, so no paragraph is forced
// inline (an inline paragraph ignores its own margins and text-indent).
public class PaginatedNoteParagraphLayoutTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original note; no external document or template is read.
    private static byte[] AuthoredDocument(string marker = "<w:r><w:footnoteRef/></w:r>")
    {
        const string props = "<w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"360\" w:lineRule=\"exact\"/></w:pPr>";
        var document = $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t xml:space=\"preserve\">Willow note </w:t></w:r>"
            + "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        var notes = $"<w:footnotes xmlns:w=\"{W}\">"
            + "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
            + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
            + $"<w:footnote w:id=\"1\"><w:p>{props}{marker}<w:r><w:t xml:space=\"preserve\"> Ledger001</w:t></w:r></w:p>"
            + $"<w:p>{props}<w:r><w:t>Ledger002</w:t></w:r></w:p></w:footnote></w:footnotes>";
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using (var xml = new MemoryStream(Encoding.UTF8.GetBytes(document))) main.FeedData(xml);
            var part = main.AddNewPart<FootnotesPart>();
            using (var xml = new MemoryStream(Encoding.UTF8.GetBytes(notes))) part.FeedData(xml);
        }
        return stream.ToArray();
    }

    private static XElement Convert(string marker = "<w:r><w:footnoteRef/></w:r>") => WmlToHtmlConverter.ConvertToHtml(
        new WmlDocument("Willow.docx", AuthoredDocument(marker)), new WmlToHtmlConverterSettings
        {
            RenderFootnotesAndEndnotes = true,
            RenderPagination = PaginationMode.Paginated,
        });

    private static string Css(XElement html) =>
        Regex.Match(html.ToString(), "<style[^>]*>(.*?)</style>", RegexOptions.Singleline).Groups[1].Value;

    [Fact]
    public void NoteParagraphsGetNoGenericDisplayInsetOrGap()
    {
        // Every rule that targets a note paragraph: none may override the paragraph's own
        // display, margins or indent, which its resolved classes carry. The one exception drops
        // only the spacing after the note area's last paragraph, at the bottom margin.
        var rules = Regex.Matches(Css(Convert()), @"([^{}]*\.footnote-content p[^{,]*)\{([^}]*)\}")
            .Select(m => (selector: m.Groups[1].Value.Trim(), body: m.Groups[2].Value))
            .ToList();

        var bandLast = Assert.Single(rules, r => r.selector.Contains(".footnote-item:last-child"));
        Assert.Equal("margin-bottom: 0 !important;", bandLast.body.Trim());
        Assert.All(rules.Where(r => r != bandLast), rule =>
        {
            Assert.DoesNotContain("display: inline", rule.body);
            Assert.DoesNotContain("margin", rule.body);
            Assert.DoesNotContain("text-indent", rule.body);
        });
    }

    [Fact]
    public void NumberOpensTheFirstNoteParagraphOnly()
    {
        var item = Convert().Descendants().First(e => (string?)e.Attribute("class") == "footnote-item");
        var paragraphs = item.Descendants().Where(e => e.Name.LocalName == "p").ToList();

        Assert.Equal(2, paragraphs.Count);
        var number = Assert.Single(item.Descendants(), e => (string?)e.Attribute("class") == "footnote-number");
        Assert.Same(paragraphs[0], number.Parent);
        Assert.Same(number, paragraphs[0].Elements().First());
        Assert.Matches(@"^1\s+Ledger001$", paragraphs[0].Value);
        Assert.Equal("Ledger002", paragraphs[1].Value);
    }

    [Fact]
    public void GenericNumberOpensTheFirstNoteParagraph()
    {
        // A note with no w:footnoteRef run still gets its number on the first paragraph's line.
        var item = Convert("").Descendants().First(e => (string?)e.Attribute("class") == "footnote-item");
        var number = Assert.Single(item.Descendants(), e => (string?)e.Attribute("class") == "footnote-number");

        Assert.Equal("p", number.Parent!.Name.LocalName);
        Assert.Null(number.Attribute("data-note-run"));
    }

    [Fact]
    public void GenericNumberCssDoesNotGrowItsLine()
    {
        // Inside the paragraph the number no longer inherits the item's zero line height, so the
        // generic rule declares it (issue #955): a raised number must not grow the line.
        var rule = Regex.Match(Css(Convert()), @"\.footnote-number\s*\{([^}]*)\}").Groups[1].Value;
        Assert.Contains("line-height: 0", rule);
    }

    [Fact]
    public void NoteNumberTakesItsFirstParagraphsExactLineHeight()
    {
        var number = Convert().Descendants().First(e => (string?)e.Attribute("class") == "footnote-number");

        Assert.Equal("true", (string?)number.Attribute("data-note-run"));
        Assert.Contains("line-height: 18.0pt", (string?)number.Attribute("style") ?? "");
    }

    [Theory]
    [InlineData("")]
    [InlineData("<w:r><w:rPr><w:position w:val=\"6\"/></w:rPr><w:footnoteRef/></w:r>")]
    public void ShiftedOrGenericNumbersKeepTheZeroLineHeightStrut(string marker)
    {
        var number = Convert(marker).Descendants().First(e => (string?)e.Attribute("class") == "footnote-number");
        Assert.DoesNotContain("line-height", (string?)number.Attribute("style") ?? "");
    }
}
