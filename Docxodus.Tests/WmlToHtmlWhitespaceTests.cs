using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class WmlToHtmlWhitespaceTests
{
    private static byte[] Fixture(string text, bool splitRun = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var paragraph = new Wp.Paragraph(new Wp.Run(new Wp.Text(text)
                { Space = SpaceProcessingModeValues.Preserve }));
            if (splitRun)
                paragraph.Append(new Wp.Run(new Wp.RunProperties(new Wp.Bold()),
                    new Wp.Text(" Oak ") { Space = SpaceProcessingModeValues.Preserve }));
            main.Document = new Wp.Document(new Wp.Body(paragraph,
                new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 })));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(new Wp.RunPropertiesDefault(new Wp.RunPropertiesBaseStyle(
                    new Wp.RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                    new Wp.FontSize { Val = "22" }))),
                new Wp.Style { Type = Wp.StyleValues.Paragraph, Default = true, StyleId = "Normal" });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
            PackageValidation.AssertValid(doc, FileFormatVersions.Microsoft365);
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(false, "Alder  Oak")]
    [InlineData(true, "Alder  Oak")]
    [InlineData(false, " Alder ")]
    [InlineData(true, " Alder ")]
    [InlineData(false, "Alder\u00a0 Oak")]
    [InlineData(true, "Alder\u00a0 Oak")]
    public void Source_spaces_keep_their_original_breaking_semantics(bool classes, string text)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("spaces.docx", Fixture(text)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = classes });
        var parsed = XElement.Parse(html.ToString(SaveOptions.DisableFormatting));
        Assert.Equal(text, Assert.Single(parsed.Descendants(Xhtml.p)).Value);
        Assert.Contains("white-space: pre-wrap", string.Concat(parsed.Descendants(Xhtml.style).Select(s => s.Value)));
    }

    [Fact]
    public void Spaces_at_run_boundaries_stay_breakable_without_losing_their_width()
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("split-spaces.docx", Fixture("Alder ", true)),
            new WmlToHtmlConverterSettings());
        Assert.Equal("Alder  Oak ", Assert.Single(html.Descendants(Xhtml.p)).Value);
    }
}
