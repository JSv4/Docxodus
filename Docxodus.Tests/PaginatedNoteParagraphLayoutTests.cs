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

    [Fact]
    public void LaterNoteParagraphsGetNoGenericInsetOrGap()
    {
        var css = Regex.Match(Convert().ToString(), "<style[^>]*>(.*?)</style>", RegexOptions.Singleline).Groups[1].Value;
        var rule = Regex.Match(css, @"\.footnote-content p:not\(:first-of-type\)\s*\{([^}]*)\}").Groups[1].Value;

        Assert.Contains("display: block", rule);
        Assert.DoesNotContain("margin", rule);
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
