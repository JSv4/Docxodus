#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class ParagraphBorderSizeRenderingTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original paragraph and cell; no external document or template is read.
    private static byte[] AuthoredDocument(string style, string? size)
    {
        var borders = string.Concat(new[] { "top", "left", "bottom", "right" }.Select(side =>
            $"<w:{side} w:val=\"{style}\" w:color=\"246B7D\"{(size is null ? "" : $" w:sz=\"{size}\"")}/>"));
        var content = $"<w:document xmlns:w=\"{W}\"><w:body>"
            + $"<w:p><w:pPr><w:pBdr>{borders}</w:pBdr></w:pPr><w:r><w:t>Cypress paragraph border.</w:t></w:r></w:p>"
            + "<w:tbl><w:tblPr><w:tblW w:w=\"9360\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9360\"/></w:tblGrid><w:tr><w:tc>"
            + $"<w:tcPr><w:tcW w:w=\"9360\" w:type=\"dxa\"/><w:tcBorders>{borders}</w:tcBorders></w:tcPr>"
            + "<w:p><w:r><w:t>Cypress cell border.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
            if (size != "bad")
                Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
                    .Validate(document));
        }
        return stream.ToArray();
    }

    private static XElement Convert(byte[] bytes) => WmlToHtmlConverter.ConvertToHtml(
        new WmlDocument("Cypress.docx", bytes), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
        });

    [Theory]
    [InlineData("single")]
    [InlineData("double")]
    [InlineData("dashed")]
    [InlineData("nil")]
    [InlineData("none")]
    public void OmittedSizesUseTheExplicitZeroWidthMappingForParagraphAndCellBorders(string style)
    {
        var actual = Convert(AuthoredDocument(style, null));
        var control = Convert(AuthoredDocument(style, "0"));
        Assert.Equal(control.Descendants().Attributes("style").Select(a => a.Value),
                     actual.Descendants().Attributes("style").Select(a => a.Value));
        Assert.Equal(control.Value, actual.Value);
    }

    [Fact]
    public void MalformedPresentSizeStillFailsNumericParsing()
    {
        Assert.Throws<FormatException>(() => Convert(AuthoredDocument("single", "bad")));
    }
}
