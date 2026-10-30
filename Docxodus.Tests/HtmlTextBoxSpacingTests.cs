using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlTextBoxSpacingTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SameStyleContextualParagraphsShareTheOrdinarySpacingRules(bool legacy)
    {
        var paragraphs = Render(legacy, contextual: true, differentStyle: false);
        Assert.Contains("margin-bottom: 0;", (string?)paragraphs[0].Attribute("style"));
        Assert.Contains("margin-top: 0;", (string?)paragraphs[1].Attribute("style"));
        Assert.Contains("margin-bottom: 9pt;", (string?)paragraphs[1].Attribute("style"));
        Assert.All(paragraphs, paragraph => Assert.Contains("display: block;", (string?)paragraph.Attribute("style")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImplicitAndExplicitDefaultStyleShareContextualSpacing(bool legacy)
    {
        var paragraphs = Render(legacy, contextual: true, differentStyle: false, implicitDefault: true);
        Assert.Contains("margin-bottom: 0;", (string?)paragraphs[0].Attribute("style"));
        Assert.Contains("margin-top: 0;", (string?)paragraphs[1].Attribute("style"));
    }

    [Fact]
    public void BodyParagraphsRecognizeTheResolvedDefaultStyle()
    {
        var paragraphs = Render(false, contextual: true, differentStyle: false, implicitDefault: true, inBox: false);
        Assert.Contains("margin-bottom: 0;", (string?)paragraphs[0].Attribute("style"));
        Assert.Contains("margin-top: 0;", (string?)paragraphs[1].Attribute("style"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void UnrelatedParagraphsKeepTheirSpacing(bool legacy, bool differentStyle)
    {
        var paragraphs = Render(legacy, contextual: differentStyle, differentStyle);
        Assert.Contains("margin-bottom: 9pt;", (string?)paragraphs[0].Attribute("style"));
        Assert.Contains("margin-top: 6pt;", (string?)paragraphs[1].Attribute("style"));
    }

    private static XElement[] Render(bool legacy, bool contextual, bool differentStyle, bool implicitDefault = false, bool inBox = true)
    {
        string Paragraph(string text, string style) =>
            "<w:p><w:pPr>" + (implicitDefault && text == "Swift copper lantern"
                ? "" : $"<w:pStyle w:val=\"{style}\"/>") + $"</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";
        var content = Paragraph("Swift copper lantern", "Swift") +
            Paragraph("Swift violet meadow", differentStyle ? "Other" : "Swift");
        var drawing = legacy
            ? "<w:pict><v:shape style=\"width:180pt;height:64pt\"><v:textbox><w:txbxContent>" + content +
                "</w:txbxContent></v:textbox></v:shape></w:pict>"
            : "<w:drawing><wp:inline><wp:extent cx=\"2286000\" cy=\"812800\"/>" +
                "<wp:docPr id=\"1\" name=\"Swift frame\"/><a:graphic><a:graphicData " +
                "uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
                "<wps:wsp><wps:txbx><w:txbxContent>" + content + "</w:txbxContent></wps:txbx>" +
                "<wps:bodyPr lIns=\"0\" tIns=\"0\" rIns=\"0\" bIns=\"0\"><a:noAutofit/>" +
                "</wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write($"<w:document xmlns:w=\"{W}\" xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
                    "<w:body>" + (inBox ? $"<w:p><w:r>{drawing}</w:r></w:p>" : content) + "<w:sectPr/></w:body></w:document>");
            using (var writer = new StreamWriter(main.AddNewPart<StyleDefinitionsPart>().GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write($"<w:styles xmlns:w=\"{W}\">" + string.Concat(new[] { "Swift", "Other" }.Select(id =>
                    $"<w:style w:type=\"paragraph\" w:styleId=\"{id}\"" + (id == "Swift" ? " w:default=\"1\"" : "") + $"><w:name w:val=\"{id}\"/>" +
                    "<w:pPr><w:spacing w:before=\"120\" w:after=\"180\"/>" +
                    (contextual ? "<w:contextualSpacing/>" : "") + "</w:pPr></w:style>")) + "</w:styles>");
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            document.Save();
        }
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)PaginationMode.Paginated }));
        return html.Descendants().Where(e => e.Name.LocalName == (inBox ? "span" : "p") &&
            e.Elements().Any() && e.Value.StartsWith("Swift ") &&
            (!inBox || ((string?)e.Attribute("style"))?.Contains("display: block") == true)).ToArray();
    }
}
