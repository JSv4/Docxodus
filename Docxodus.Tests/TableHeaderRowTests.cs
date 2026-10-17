using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1171: Word repeats a table's header rows (w:tblHeader) at the top of every page the table
// continues on, whatever their height rule. The paginator does the repeating; the converter tells it
// which rows are header rows.
public class TableHeaderRowTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original ledger; no external document or template is read.
    private static XElement Render(string headerTrPr)
    {
        static string Row(string trPr, string text) =>
            $"<w:tr><w:trPr>{trPr}</w:trPr><w:tc><w:tcPr><w:tcW w:w=\"3600\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>";
        var content = $"<w:document xmlns:w=\"{W}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"3600\" w:type=\"dxa\"/></w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"3600\"/></w:tblGrid>"
            + Row(headerTrPr, "Heading ledger") + Row("<w:trHeight w:val=\"400\"/>", "Rowan entry")
            + "</w:tbl><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Rowan.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderPagination = PaginationMode.Paginated });
    }

    private static string?[] HeaderMarks(XElement html) =>
        html.Descendants(Xhtml.tr).Select(tr => (string?)tr.Attribute("data-docx-header-row")).ToArray();

    [Theory]
    [InlineData("<w:tblHeader/><w:trHeight w:val=\"400\" w:hRule=\"exact\"/>")]
    [InlineData("<w:tblHeader w:val=\"1\"/><w:trHeight w:val=\"400\" w:hRule=\"atLeast\"/>")]
    [InlineData("<w:tblHeader w:val=\"true\"/>")]
    public void HeaderRowIsMarkedWhateverItsHeightRule(string trPr) =>
        Assert.Equal(new[] { "true", null }, HeaderMarks(Render(trPr)));

    [Theory]
    [InlineData("<w:tblHeader w:val=\"0\"/><w:trHeight w:val=\"400\" w:hRule=\"exact\"/>")]
    [InlineData("<w:tblHeader w:val=\"false\"/>")]
    [InlineData("<w:trHeight w:val=\"400\" w:hRule=\"exact\"/>")]
    public void RowThatIsNotAHeaderIsNotMarked(string trPr) =>
        Assert.Equal(new string?[] { null, null }, HeaderMarks(Render(trPr)));
}
