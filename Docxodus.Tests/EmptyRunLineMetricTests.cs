using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class EmptyRunLineMetricTests
{
    private static XElement Render(string emptyContent, bool blank = false, string properties = "",
        Wp.LineSpacingRuleValues? rule = null, string extraRun = "", bool paginated = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Wp.Paragraph(new Wp.ParagraphProperties(
                new Wp.SpacingBetweenLines { Line = "480", LineRule = rule ?? Wp.LineSpacingRuleValues.Auto }));
            paragraph.Append(new Wp.Run($"<w:r xmlns:w=\"{W.w.NamespaceName}\"><w:rPr><w:sz w:val=\"48\"/>{properties}</w:rPr>{emptyContent}</w:r>"));
            if (!blank)
                paragraph.Append(new Wp.Run(new Wp.Text("Juniper inventory remains unchanged.")));
            if (extraRun.Length > 0) paragraph.Append(new Wp.Run(extraRun));
            main.Document = new Wp.Document(new Wp.Body(paragraph));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(
                new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" },
                    new Wp.FontSize { Val = "24" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document.Save();
        }
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Juniper.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false,
                RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None })
            .Descendants().First(e => e.Name.LocalName == "p");
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
    [InlineData("<w:t xml:space=\"preserve\">&#160;</w:t>")]
    [InlineData("<w:br/>")]
    [InlineData("<w:tab/>")]
    public void StructuralOrWhitespaceContentKeepsItsExistingMetrics(string content)
    {
        Assert.Contains("font-size: 24pt", (string?)Render(content).Attribute("style"));
    }

    [Fact]
    public void LargerOrdinarySpaceKeepsItsWidthWithoutRaisingTextLines()
    {
        var paragraph = Render("<w:t xml:space=\"preserve\"> </w:t>");
        Assert.Contains("font-size: 12pt", (string?)paragraph.Attribute("style"));
        Assert.Equal(" Juniper inventory remains unchanged.", paragraph.Value);
        var space = paragraph.Descendants().Single(e => e.Value == " ");
        Assert.Contains("font-size: 24pt", (string?)space.Attribute("style"));
        Assert.Contains("line-height: 0", (string?)space.Attribute("style"));
        Assert.Contains("vertical-align: top", (string?)space.Attribute("style"));
    }

    [Theory]
    [InlineData("<w:u w:val=\"single\"/>")]
    [InlineData("<w:highlight w:val=\"yellow\"/>")]
    [InlineData("<w:shd w:fill=\"FFFF00\"/>")]
    [InlineData("<w:bdr w:val=\"single\" w:sz=\"8\"/>")]
    [InlineData("<w:position w:val=\"4\"/>")]
    [InlineData("<w:vertAlign w:val=\"superscript\"/>")]
    public void PaintedAndRaisedSpacesKeepTheirExistingMetrics(string properties)
    {
        Assert.Contains("font-size: 24pt", (string?)Render("<w:t> </w:t>", properties: properties).Attribute("style"));
    }

    [Fact]
    public void SpaceOnlyParagraphAndFixedSpacingKeepTheirExistingMetrics()
    {
        Assert.Contains("font-size: 24pt", (string?)Render("<w:t> </w:t>", blank: true).Attribute("style"));
        foreach (var rule in new[] { Wp.LineSpacingRuleValues.Exact, Wp.LineSpacingRuleValues.AtLeast })
            Assert.Contains("font-size: 24pt", (string?)Render("<w:t> </w:t>", rule: rule).Attribute("style"));
    }

    [Fact]
    public void TextboxTextDoesNotChangeItsSpaceOnlyHostsMetrics()
    {
        var paragraph = Render("<w:t> </w:t>", blank: true, extraRun:
            $"<w:r xmlns:w=\"{W.w.NamespaceName}\" xmlns:v=\"urn:schemas-microsoft-com:vml\">"
            + "<w:pict><v:shape id=\"Box\" style=\"width:120pt;height:36pt\"><v:textbox>"
            + "<w:txbxContent><w:p><w:r><w:t>Nested story.</w:t></w:r></w:p></w:txbxContent>"
            + "</v:textbox></v:shape></w:pict></w:r>");
        Assert.Contains("font-size: 24pt", (string?)paragraph.Attribute("style"));
        var space = paragraph.Descendants().Single(e => e.Value == " ");
        Assert.DoesNotContain("line-height: 0", (string?)space.Attribute("style"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WebHiddenTextKeepsTheSpaceOnlyParagraphsExistingMetrics(bool paginated)
    {
        var paragraph = Render("<w:t> </w:t>", blank: true, paginated: paginated, extraRun:
            $"<w:r xmlns:w=\"{W.w.NamespaceName}\"><w:rPr><w:sz w:val=\"24\"/><w:webHidden/></w:rPr>"
            + "<w:t>Hidden in web layout.</w:t></w:r>");
        Assert.Contains("font-size: 24pt", (string?)paragraph.Attribute("style"));
        Assert.Equal(paginated, paragraph.Value.Contains("Hidden in web layout."));
    }

    [Fact]
    public void BlankParagraphKeepsItsExistingFontAndPlaceholder()
    {
        var paragraph = Render("<w:t/>", blank: true);
        Assert.Contains("font-size: 24pt", (string?)paragraph.Attribute("style"));
        Assert.Equal("\u00a0", paragraph.Value);
    }
}
