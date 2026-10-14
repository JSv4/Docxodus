using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1116: explicit single and dashed paragraph border sizes are recorded in physical units, so
// the export can lay out and paint a 0.25pt border instead of the one CSS pixel CSS allows.
public class ExactParagraphBorderWidthTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Literal original card; no external document or template is read.
    private static XElement Wrapper(string style, string? size)
    {
        var borders = string.Concat(new[] { "top", "left", "bottom", "right" }.Select(side =>
            $"<w:{side} w:val=\"{style}\" w:space=\"4\" w:color=\"246B7D\"{(size is null ? "" : $" w:sz=\"{size}\"")}/>"));
        var content = $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:pPr><w:pBdr>{borders}</w:pBdr></w:pPr><w:r><w:t>BorderCard Juniper.</w:t></w:r></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            main.FeedData(xml);
        }
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Juniper.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        return html.Descendants().First(e => e.Name.LocalName == "div" && ((string?)e.Attribute("style") ?? "").Contains("border-left"));
    }

    [Theory]
    [InlineData("single", "2", "0.25pt")]
    [InlineData("dashed", "2", "0.25pt")]
    [InlineData("single", "12", "1.5pt")]
    public void ExplicitSizesRecordTheirPhysicalWidth(string style, string size, string width)
    {
        var wrapper = Wrapper(style, size);
        var css = (string?)wrapper.Attribute("style") ?? "";

        Assert.Equal("true", (string?)wrapper.Attribute("data-docx-exact-border"));
        foreach (var side in new[] { "top", "right", "bottom", "left" })
            Assert.Contains($"--docx-border-{side}-width: {width}", css);
    }

    [Theory]
    [InlineData("single", null)]
    [InlineData("single", "0")]
    [InlineData("dashed", null)]
    [InlineData("double", "6")]
    public void OmittedSizesAndOtherStylesKeepTheirCssRendering(string style, string? size)
    {
        var wrapper = Wrapper(style, size);
        Assert.Null(wrapper.Attribute("data-docx-exact-border"));
        Assert.DoesNotContain("--docx-border-", (string?)wrapper.Attribute("style") ?? "");
    }
}
