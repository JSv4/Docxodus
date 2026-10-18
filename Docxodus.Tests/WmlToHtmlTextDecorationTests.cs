using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class WmlToHtmlTextDecorationTests
{
    private static byte[] Fixture(string properties)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(
                    new Wp.RunProperties("<w:rPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + properties + "</w:rPr>"),
                    new Wp.Text("Cedar rule sample."))), new Wp.SectionProperties()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    private static string Styles(string properties) => string.Join(" ", XElement.Parse(
        HtmlConversionOps.ConvertToHtml(Fixture(properties), new HtmlConversionOptions { FabricateCssClasses = false }))
        .Descendants().Where(e => e.Name.LocalName == "span").Select(e => (string?)e.Attribute("style")));

    [Theory]
    [InlineData("<w:u w:val=\"double\"/>", "underline double")]
    [InlineData("<w:dstrike/>", "line-through double")]
    [InlineData("<w:strike/><w:u w:val=\"single\"/>", "underline line-through")]
    [InlineData("<w:dstrike/><w:u w:val=\"double\"/>", "underline line-through double")]
    [InlineData("<w:u w:val=\"single\"/>", "underline")]
    [InlineData("<w:strike/>", "line-through")]
    public void Decorations_RetainTheirLineAndStyle(string properties, string decoration) =>
        Assert.Contains("text-decoration: " + decoration, Styles(properties));

    [Fact]
    public void DottedUnderline_RetainsItsExplicitColor()
    {
        var styles = Styles("<w:u w:val=\"dotted\" w:color=\"0000FF\"/>");
        Assert.Contains("text-decoration: underline dotted", styles);
        Assert.Contains("text-decoration-color: #0000FF", styles);
    }

    [Fact]
    public void DifferentUnderlineAndStrikeStyles_KeepIndependentDecorations()
    {
        var styles = Styles("<w:dstrike/><w:u w:val=\"dotted\" w:color=\"0000FF\"/>");
        Assert.Contains("text-decoration: underline dotted", styles);
        Assert.Contains("text-decoration-color: #0000FF", styles);
        Assert.Contains("text-decoration: line-through double", styles);
    }

    [Fact]
    public void DisabledDecorations_DoNotDrawLines() => Assert.DoesNotContain("text-decoration:",
        Styles("<w:strike w:val=\"0\"/><w:dstrike w:val=\"false\"/><w:u w:val=\"none\"/>"));
}
