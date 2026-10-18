using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class FormatChangeDescriptionTests
{
    private static byte[] Fixture(string current, string previous)
    {
        const string ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(new Wp.Run(
                new Wp.RunProperties($"<w:rPr xmlns:w=\"{ns}\">{current}<w:rPrChange w:id=\"1\" w:author=\"Cedar Reviewer\" w:date=\"2026-09-10T12:00:00Z\"><w:rPr>{previous}</w:rPr></w:rPrChange></w:rPr>"),
                new Wp.Text("Cedar revision sample."))), new Wp.SectionProperties()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Liberation Serif", HighAnsi = "Liberation Serif" },
                    new Wp.FontSize { Val = "24" }))),
                new Wp.Style(new Wp.StyleName { Val = "Cedar Strong" }, new Wp.StyleRunProperties(new Wp.Bold()))
                    { Type = Wp.StyleValues.Character, StyleId = "CedarStrong" });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("<w:b/>", "", "Bold added")]
    [InlineData("<w:rFonts w:ascii=\"Liberation Serif\" w:hAnsi=\"Liberation Serif\"/><w:b/><w:sz w:val=\"24\"/>",
        "<w:rFonts w:ascii=\"Liberation Serif\" w:hAnsi=\"Liberation Serif\"/><w:sz w:val=\"24\"/>", "Bold added")]
    [InlineData("<w:rStyle w:val=\"CedarStrong\"/>", "", "Bold added")]
    [InlineData("", "<w:rStyle w:val=\"CedarStrong\"/>", "Bold removed")]
    [InlineData("", "<w:b/>", "Bold removed")]
    [InlineData("<w:b w:val=\"0\"/>", "<w:b/>", "Bold removed")]
    [InlineData("<w:sz w:val=\"28\"/>", "<w:sz w:val=\"24\"/>", "Font size changed")]
    [InlineData("<w:rFonts w:ascii=\"Liberation Sans\" w:hAnsi=\"Liberation Sans\"/>",
        "<w:rFonts w:ascii=\"Liberation Serif\" w:hAnsi=\"Liberation Serif\"/>", "Font changed")]
    public void Description_ComparesResolvedPreviousAndCurrentFormatting(string current, string previous, string expected)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-revision.docx", Fixture(current, previous)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true });
        var revision = html.Descendants().Single(e => ((string?)e.Attribute("class") ?? "").Split(' ').Contains("rev-format-change"));
        Assert.Equal(expected, (string?)revision.Attribute("title"));
    }

    [Fact]
    public void PreviousFontResolution_DoesNotOverwriteTheRenderedCurrentFont()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-font.docx", Fixture(
            "<w:rFonts w:ascii=\"Liberation Sans\" w:hAnsi=\"Liberation Sans\"/>",
            "<w:rFonts w:ascii=\"Liberation Serif\" w:hAnsi=\"Liberation Serif\"/>")),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true });
        Assert.Contains("font-family: 'Liberation Sans'", string.Join(" ", html.Descendants().Attributes("style").Select(a => a.Value)));
    }

    [Fact]
    public void GeneralFormattingAssembly_PreservesTheOriginalRevisionSnapshot()
    {
        using var stream = new MemoryStream();
        stream.Write(Fixture("<w:b/>", ""));
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, true);
        FormattingAssembler.AssembleFormatting(doc, new FormattingAssemblerSettings { ClearStyles = false });
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var previous = doc.MainDocumentPart!.GetXDocument().Descendants(w + "rPrChange").Single().Element(w + "rPr");
        Assert.Empty(previous?.Elements() ?? Enumerable.Empty<XElement>());
    }
}
