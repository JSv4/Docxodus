using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class EmptyRunLineMetricTests
{
    private static XElement Render(string emptyContent, bool blank = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Wp.Paragraph(new Wp.ParagraphProperties(
                new Wp.SpacingBetweenLines { Line = "480", LineRule = Wp.LineSpacingRuleValues.Auto }));
            paragraph.Append(new Wp.Run($"<w:r xmlns:w=\"{W.w.NamespaceName}\"><w:rPr><w:sz w:val=\"48\"/></w:rPr>{emptyContent}</w:r>"));
            if (!blank)
                paragraph.Append(new Wp.Run(new Wp.Text("Juniper inventory remains unchanged.")));
            main.Document = new Wp.Document(new Wp.Body(paragraph));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" },
                    new Wp.FontSize { Val = "24" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document.Save();
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Juniper.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false })
            .Descendants().Single(e => e.Name.LocalName == "p");
    }

    [Theory]
    [InlineData("")]
    [InlineData("<w:t/>")]
    public void EmptyRunDoesNotSizeAnOrdinaryParagraph(string content)
    {
        var paragraph = Render(content);
        Assert.Contains("font-size: 12pt", (string?)paragraph.Attribute("style"));
        Assert.Equal("Juniper inventory remains unchanged.", paragraph.Value);
        Assert.DoesNotContain(paragraph.Descendants(), span =>
            ((string?)span.Attribute("style"))?.Contains("font-size: 24pt") == true);
    }

    [Theory]
    [InlineData("<w:t xml:space=\"preserve\"> </w:t>")]
    [InlineData("<w:br/>")]
    [InlineData("<w:tab/>")]
    public void StructuralOrWhitespaceContentKeepsItsExistingMetrics(string content)
    {
        Assert.Contains("font-size: 24pt", (string?)Render(content).Attribute("style"));
    }

    [Fact]
    public void BlankParagraphKeepsItsExistingFontAndPlaceholder()
    {
        var paragraph = Render("<w:t/>", blank: true);
        Assert.Contains("font-size: 24pt", (string?)paragraph.Attribute("style"));
        Assert.Equal("\u00a0", paragraph.Value);
    }
}
